from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import msal
import requests

from sistema1_gantt_builder import derive_project_identity, build_gantt_workbook


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_ACTIVE_PROJECTS_ROOT = "Proyectos/02_Activos"


@dataclass(frozen=True)
class Settings:
    tenant_id: str
    client_id: str
    client_secret: str
    site_hostname: str
    site_path: str
    queue_list_name: str
    queue_list_id: str | None
    control_list_name: str
    control_list_id: str | None
    active_projects_root: str


@dataclass(frozen=True)
class ProcessResult:
    item_id: str
    title: str
    status: str
    project_id: str = ""
    project_name: str = ""
    gantt_url: str = ""
    budget_url: str = ""
    message: str = ""


def env_value(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def load_settings() -> Settings:
    missing = [
        name
        for name in ("MS_GRAPH_TENANT_ID", "MS_GRAPH_CLIENT_ID", "MS_GRAPH_CLIENT_SECRET")
        if not env_value(name)
    ]
    if missing:
        raise RuntimeError(
            "Faltan secrets/variables para Microsoft Graph: "
            + ", ".join(missing)
            + ". GitHub Actions necesita autenticacion no interactiva."
        )

    return Settings(
        tenant_id=env_value("MS_GRAPH_TENANT_ID"),
        client_id=env_value("MS_GRAPH_CLIENT_ID"),
        client_secret=env_value("MS_GRAPH_CLIENT_SECRET"),
        site_hostname=env_value("SP_SITE_HOSTNAME", "sciprowall.sharepoint.com"),
        site_path=env_value("SP_SITE_PATH", "/sites/PROYECTOSPROWALL").strip("/"),
        queue_list_name=env_value("SP_QUEUE_LIST_NAME", "Cola_Automatizacion_Proyectos"),
        queue_list_id=env_value("SP_QUEUE_LIST_ID") or None,
        control_list_name=env_value("SP_CONTROL_LIST_NAME", "Control_Gantt_Asignaciones"),
        control_list_id=env_value("SP_CONTROL_LIST_ID") or None,
        active_projects_root=env_value("SP_ACTIVE_PROJECTS_ROOT", DEFAULT_ACTIVE_PROJECTS_ROOT).strip("/"),
    )


def acquire_token(settings: Settings) -> str:
    authority = f"https://login.microsoftonline.com/{settings.tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=settings.client_id,
        client_credential=settings.client_secret,
        authority=authority,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(
            "No se pudo obtener token Microsoft Graph: "
            + json.dumps(
                {
                    "error": result.get("error"),
                    "error_description": result.get("error_description"),
                    "correlation_id": result.get("correlation_id"),
                },
                ensure_ascii=False,
            )
        )
    return str(result["access_token"])


def decode_token_claims(token: str) -> dict[str, Any]:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        decoded = base64.urlsafe_b64decode(payload.encode("ascii"))
        return json.loads(decoded)
    except Exception:
        return {}


def print_token_diagnostics(token: str) -> None:
    claims = decode_token_claims(token)
    diagnostic = {
        "aud": claims.get("aud"),
        "appid": claims.get("appid") or claims.get("azp"),
        "tid": claims.get("tid"),
        "roles": claims.get("roles") or [],
        "scp": claims.get("scp") or "",
    }
    print("Graph token diagnostic:")
    print(json.dumps(diagnostic, ensure_ascii=False, sort_keys=True))
    if not diagnostic["roles"] and not diagnostic["scp"]:
        print(
            "WARNING: El token no trae roles ni scopes. Para GitHub Actions con client secret, "
            "la app de Entra necesita permisos de aplicacion de Microsoft Graph y admin consent."
        )


def graph_request(
    token: str,
    method: str,
    url: str,
    *,
    json_body: dict[str, Any] | None = None,
    data: bytes | None = None,
    expected: tuple[int, ...] = (200,),
) -> Any:
    headers = {"Authorization": f"Bearer {token}"}
    if json_body is not None:
        headers["Content-Type"] = "application/json"
    response = requests.request(method, url, headers=headers, json=json_body, data=data, timeout=120)
    if response.status_code not in expected:
        raise RuntimeError(f"Graph {method} {response.status_code}: {response.text[:3000]}")
    if response.status_code == 204 or not response.text:
        return {}
    return response.json()


def graph_get(token: str, url: str) -> dict[str, Any]:
    return graph_request(token, "GET", url, expected=(200,))


def graph_patch(token: str, url: str, payload: dict[str, Any]) -> dict[str, Any]:
    return graph_request(token, "PATCH", url, json_body=payload, expected=(200, 204))


def graph_post(token: str, url: str, payload: dict[str, Any]) -> dict[str, Any]:
    return graph_request(token, "POST", url, json_body=payload, expected=(200, 201))


def graph_put_bytes(token: str, url: str, content: bytes) -> dict[str, Any]:
    return graph_request(token, "PUT", url, data=content, expected=(200, 201))


def resolve_site(token: str, settings: Settings) -> dict[str, Any]:
    url = f"{GRAPH_BASE}/sites/{settings.site_hostname}:/{settings.site_path}"
    return graph_get(token, url)


def resolve_list(token: str, site_id: str, list_name: str, list_id: str | None = None) -> dict[str, Any]:
    if list_id:
        return graph_get(token, f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}")
    escaped = list_name.replace("'", "''")
    data = graph_get(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,displayName,name,webUrl&$filter=displayName eq '{escaped}'",
    )
    values = data.get("value") or []
    if not values:
        raise RuntimeError(f"No se encontro la lista SharePoint: {list_name}")
    return values[0]


def list_columns(token: str, site_id: str, list_id: str) -> list[dict[str, Any]]:
    data = graph_get(token, f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/columns?$select=name,displayName")
    return data.get("value") or []


def normalized_key(value: str) -> str:
    text = value.strip().casefold()
    text = text.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    return re.sub(r"[^a-z0-9]+", "", text)


def build_field_map(columns: list[dict[str, Any]]) -> dict[str, str]:
    field_map: dict[str, str] = {}
    for column in columns:
        name = str(column.get("name") or "")
        display_name = str(column.get("displayName") or "")
        if name:
            field_map[normalized_key(name)] = name
        if display_name:
            field_map[normalized_key(display_name)] = name
    return field_map


def pick_field(field_map: dict[str, str], candidates: tuple[str, ...]) -> str | None:
    for candidate in candidates:
        match = field_map.get(normalized_key(candidate))
        if match:
            return match
    return None


def list_queue_items(token: str, site_id: str, list_id: str, top: int) -> list[dict[str, Any]]:
    select_fields = (
        "Title,EventType,Estado,Filename,FileID,FolderPath,FileLink,"
        "CreatedByEmail,CreatedTime,Intentos,UltimoError,FechaProcesado,ProyectoID,Notas"
    )
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
        f"?$top={top}&$orderby=createdDateTime asc&expand=fields($select={select_fields})"
    )
    data = graph_get(token, url)
    return data.get("value") or []


def item_status(item: dict[str, Any]) -> str:
    fields = item.get("fields") or {}
    return str(fields.get("Estado") or "").strip()


def item_event_type(item: dict[str, Any]) -> str:
    fields = item.get("fields") or {}
    return str(fields.get("EventType") or "").strip()


def pending_budget_items(items: list[dict[str, Any]], max_items: int) -> list[dict[str, Any]]:
    pending = [
        item
        for item in items
        if item_status(item).casefold() == "pendiente"
        and item_event_type(item).casefold() == "presupuesto_aprobado"
    ]
    return pending[:max(1, max_items)]


def parse_attempts(value: Any) -> int:
    try:
        return int(str(value or "0").strip() or "0")
    except ValueError:
        return 0


def update_queue_fields(token: str, site_id: str, list_id: str, item_id: str, fields: dict[str, Any]) -> None:
    graph_patch(token, f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields", fields)


def trim_note(value: str, max_len: int = 1800) -> str:
    value = re.sub(r"\s+", " ", value).strip()
    return value[:max_len]


def drive_path_from_queue_fields(fields: dict[str, Any]) -> str:
    file_name = str(fields.get("Filename") or fields.get("FileName") or fields.get("Title") or "").strip()
    folder_path = str(fields.get("FolderPath") or "").replace("\\", "/").strip("/")
    for prefix in ("Documentos compartidos/", "Shared Documents/", "Documents/"):
        if folder_path.casefold().startswith(prefix.casefold()):
            folder_path = folder_path[len(prefix) :]
            break
    if not file_name:
        raise RuntimeError("El item de cola no contiene Filename/FileName.")
    return f"{folder_path}/{file_name}".strip("/")


def encoded_drive_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def get_drive_item_by_path(token: str, site_id: str, drive_path: str) -> dict[str, Any]:
    return graph_get(token, f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(drive_path)}")


def download_drive_file(token: str, site_id: str, drive_path: str, output_path: Path) -> dict[str, Any]:
    item = get_drive_item_by_path(token, site_id, drive_path)
    response = requests.get(
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(drive_path)}:/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=120,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Graph download {response.status_code}: {response.text[:2000]}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(response.content)
    return item


def path_exists(token: str, site_id: str, drive_path: str) -> bool:
    response = requests.get(
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(drive_path)}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    if response.status_code == 404:
        return False
    if response.status_code >= 400:
        raise RuntimeError(f"Graph exists {response.status_code}: {response.text[:2000]}")
    return True


def ensure_drive_folder(token: str, site_id: str, folder_path: str) -> dict[str, Any]:
    current_path = ""
    current_item: dict[str, Any] | None = None
    for part in [part for part in folder_path.strip("/").split("/") if part]:
        parent_path = current_path
        current_path = f"{current_path}/{part}".strip("/")
        if path_exists(token, site_id, current_path):
            current_item = get_drive_item_by_path(token, site_id, current_path)
            continue
        payload = {"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
        if parent_path:
            url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(parent_path)}:/children"
        else:
            url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children"
        current_item = graph_post(token, url, payload)
    if current_item is None:
        current_item = graph_get(token, f"{GRAPH_BASE}/sites/{site_id}/drive/root")
    return current_item


def upload_file_replace(token: str, site_id: str, local_path: Path, folder_path: str, file_name: str) -> dict[str, Any]:
    ensure_drive_folder(token, site_id, folder_path)
    target_path = f"{folder_path.strip('/')}/{file_name}".strip("/")
    return graph_put_bytes(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(target_path)}:/content",
        local_path.read_bytes(),
    )


def hyperlink_value(url: str, label: str) -> dict[str, str]:
    return {"Url": url, "Description": label}


def create_control_assignment(
    token: str,
    site_id: str,
    control_list: dict[str, Any],
    identity_project_id: str,
    identity_project_name: str,
    source_web_url: str,
    gantt_web_url: str,
    notes: str,
) -> str:
    columns = list_columns(token, site_id, control_list["id"])
    field_map = build_field_map(columns)
    fields: dict[str, Any] = {}

    assignments: list[tuple[tuple[str, ...], Any]] = [
        (("Title",), f"{identity_project_id} - {identity_project_name}"),
        (("ProyectoID",), identity_project_id),
        (("NombreProyecto",), identity_project_name),
        (("PresupuestoLink",), hyperlink_value(source_web_url, "Presupuesto aprobado")),
        (("GanttWorkingLink", "GanntWorkingLink"), hyperlink_value(gantt_web_url, "Gantt WORKING")),
        (("EstadoGantt", "EstadoGannt"), "Pendiente de asignación"),
        (("FechaGanttGenerado",), datetime.now(timezone.utc).isoformat()),
        (("Notas",), notes),
    ]
    for candidates, value in assignments:
        field_name = pick_field(field_map, candidates)
        if field_name:
            fields[field_name] = value

    if not fields:
        raise RuntimeError("No se pudo mapear ningun campo de Control_Gantt_Asignaciones.")

    try:
        created = graph_post(
            token,
            f"{GRAPH_BASE}/sites/{site_id}/lists/{control_list['id']}/items",
            {"fields": fields},
        )
    except RuntimeError:
        # Fallback: algunos tenants no aceptan objetos Hyperlink por Graph en create item.
        safe_fields = {
            key: value
            for key, value in fields.items()
            if not isinstance(value, dict)
        }
        created = graph_post(
            token,
            f"{GRAPH_BASE}/sites/{site_id}/lists/{control_list['id']}/items",
            {"fields": safe_fields},
        )
    return str(created.get("id") or "")


def print_summary(site: dict[str, Any], queue_list: dict[str, Any], items: list[dict[str, Any]]) -> None:
    pending = pending_budget_items(items, max_items=10_000)
    print("Autosys Sistema 1 - Queue poll")
    print(f"Run UTC: {datetime.now(timezone.utc).isoformat()}")
    print(f"Site: {site.get('webUrl') or site.get('displayName')}")
    print(f"Queue list: {queue_list.get('displayName')} ({queue_list.get('id')})")
    print(f"Items scanned: {len(items)}")
    print(f"Pending presupuesto_aprobado items: {len(pending)}")
    if pending:
        print("Pending detail:")
        for item in pending[:20]:
            fields = item.get("fields") or {}
            print(
                "- "
                f"id={item.get('id')} "
                f"title={fields.get('Title')!r} "
                f"event={fields.get('EventType')!r} "
                f"estado={fields.get('Estado')!r} "
                f"filename={fields.get('Filename')!r} "
                f"fileid={fields.get('FileID')!r}"
            )


def process_queue_item(
    token: str,
    site_id: str,
    queue_list: dict[str, Any],
    control_list: dict[str, Any] | None,
    settings: Settings,
    item: dict[str, Any],
    work_dir: Path,
) -> ProcessResult:
    item_id = str(item.get("id") or "")
    fields = item.get("fields") or {}
    title = str(fields.get("Title") or fields.get("Filename") or item_id)
    attempts = parse_attempts(fields.get("Intentos")) + 1
    queue_list_id = queue_list["id"]

    update_queue_fields(
        token,
        site_id,
        queue_list_id,
        item_id,
        {
            "Estado": "Procesando",
            "Intentos": str(attempts),
            "UltimoError": "",
            "Notas": "Procesando por GitHub Actions.",
        },
    )

    try:
        source_drive_path = drive_path_from_queue_fields(fields)
        source_file_name = Path(source_drive_path).name
        identity = derive_project_identity(source_file_name)
        project_folder = f"{settings.active_projects_root}/{identity.folder_name}"
        gantt_folder = f"{project_folder}/gantts"

        local_budget = work_dir / "input" / source_file_name
        local_gantt = work_dir / "output" / identity.gantt_file_name

        source_item = download_drive_file(token, site_id, source_drive_path, local_budget)
        build_result = build_gantt_workbook(local_budget, local_gantt, source_file_name)

        budget_upload = upload_file_replace(
            token,
            site_id,
            local_budget,
            project_folder,
            identity.budget_copy_name,
        )
        gantt_upload = upload_file_replace(
            token,
            site_id,
            local_gantt,
            gantt_folder,
            identity.gantt_file_name,
        )

        source_url = str(source_item.get("webUrl") or budget_upload.get("webUrl") or "")
        budget_url = str(budget_upload.get("webUrl") or "")
        gantt_url = str(gantt_upload.get("webUrl") or "")
        control_note = ""
        if control_list:
            try:
                control_id = create_control_assignment(
                    token=token,
                    site_id=site_id,
                    control_list=control_list,
                    identity_project_id=identity.project_id,
                    identity_project_name=identity.project_name,
                    source_web_url=budget_url or source_url,
                    gantt_web_url=gantt_url,
                    notes="Creado automaticamente por GitHub Actions Sistema 1.",
                )
                control_note = f" Control_Gantt_Asignaciones item={control_id}."
            except Exception as exc:
                control_note = f" Advertencia: no se pudo crear Control_Gantt_Asignaciones: {exc}"

        note = trim_note(
            "Procesado por GitHub Actions. "
            f"Proyecto={identity.project_id}. "
            f"Hoja={build_result.selected_sheet}. "
            f"Filas={build_result.rows_written}. "
            f"Revision={build_result.review_rows}. "
            f"Gantt={gantt_url}."
            f"{control_note}"
        )
        update_queue_fields(
            token,
            site_id,
            queue_list_id,
            item_id,
            {
                "Estado": "Procesado",
                "FechaProcesado": datetime.now(timezone.utc).isoformat(),
                "ProyectoID": identity.project_id,
                "UltimoError": "",
                "Notas": note,
            },
        )
        return ProcessResult(
            item_id=item_id,
            title=title,
            status="Procesado",
            project_id=identity.project_id,
            project_name=identity.project_name,
            gantt_url=gantt_url,
            budget_url=budget_url,
            message=note,
        )
    except Exception as exc:
        error = trim_note(str(exc), 1500)
        update_queue_fields(
            token,
            site_id,
            queue_list_id,
            item_id,
            {
                "Estado": "Error",
                "UltimoError": error,
                "Notas": "Error procesando por GitHub Actions. Revisar UltimoError.",
            },
        )
        return ProcessResult(item_id=item_id, title=title, status="Error", message=error)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Lee y procesa la cola SharePoint del Sistema 1 para GitHub Actions."
    )
    parser.add_argument("--top", type=int, default=50, help="Cantidad maxima de items recientes a leer.")
    parser.add_argument("--max-items", type=int, default=5, help="Cantidad maxima de pendientes a procesar por corrida.")
    parser.add_argument(
        "--process",
        action="store_true",
        help="Procesa items Pendiente. Sin este flag solo lee y resume la cola.",
    )
    parser.add_argument(
        "--fail-on-pending",
        action="store_true",
        help="Falla si hay pendientes. Util para pruebas, no para schedule normal.",
    )
    args = parser.parse_args()

    settings = load_settings()
    token = acquire_token(settings)
    print_token_diagnostics(token)
    site = resolve_site(token, settings)
    queue_list = resolve_list(token, site["id"], settings.queue_list_name, settings.queue_list_id)
    items = list_queue_items(token, site["id"], queue_list["id"], max(1, args.top))
    print_summary(site, queue_list, items)

    pending = pending_budget_items(items, args.max_items)
    if args.process and pending:
        try:
            control_list = resolve_list(token, site["id"], settings.control_list_name, settings.control_list_id)
        except Exception as exc:
            print(f"WARNING: no se pudo resolver Control_Gantt_Asignaciones; se procesara solo la cola. {exc}")
            control_list = None

        results: list[ProcessResult] = []
        with tempfile.TemporaryDirectory(prefix="autosys_s1_") as temp_dir:
            work_dir = Path(temp_dir)
            for item in pending:
                result = process_queue_item(token, site["id"], queue_list, control_list, settings, item, work_dir)
                results.append(result)
                print(
                    f"Processed item id={result.item_id} status={result.status} "
                    f"project={result.project_id or '-'} gantt={result.gantt_url or '-'}"
                )
        errors = [result for result in results if result.status == "Error"]
        if errors:
            print("Errores de procesamiento:")
            for result in errors:
                print(f"- id={result.item_id} title={result.title!r}: {result.message}")
            return 1
    elif args.process:
        print("No hay items Pendiente/presupuesto_aprobado para procesar.")

    if args.fail_on_pending and pending:
        return 2
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
