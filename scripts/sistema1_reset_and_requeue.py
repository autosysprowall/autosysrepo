from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime, timezone
from typing import Any

from sistema1_poll_queue import (
    GRAPH_BASE,
    acquire_token,
    encoded_drive_path,
    graph_get,
    graph_post,
    graph_request,
    load_settings,
    print_token_diagnostics,
    resolve_list,
    resolve_site,
)
from sistema1_gantt_builder import derive_project_identity


DEFAULT_APPROVED_ROOT = "Proyectos/Presupuestos Aprobados"
EXPECTED_ACTIVE_ROOT = "Proyectos/Proyectos Activos"
FORBIDDEN_FRAGMENT = "proyectos/proyectos terminados"


def normalized_path(value: str) -> str:
    return "/".join(part.strip() for part in value.replace("\\", "/").strip("/").split("/"))


def validate_roots(active_root: str, approved_root: str, confirmation: str) -> None:
    active = normalized_path(active_root)
    approved = normalized_path(approved_root)
    confirmed = normalized_path(confirmation)
    if active.casefold() != EXPECTED_ACTIVE_ROOT.casefold():
        raise RuntimeError(
            "Reinicio rechazado: SP_ACTIVE_PROJECTS_ROOT debe ser exactamente "
            f"{EXPECTED_ACTIVE_ROOT!r}; recibido {active!r}."
        )
    if confirmed.casefold() != active.casefold():
        raise RuntimeError(
            "Reinicio rechazado: --confirm-active-root no coincide con la carpeta activa."
        )
    if FORBIDDEN_FRAGMENT in active.casefold() or FORBIDDEN_FRAGMENT in approved.casefold():
        raise RuntimeError("Reinicio rechazado: Proyectos Terminados está fuera del alcance.")
    if approved.casefold() != DEFAULT_APPROVED_ROOT.casefold():
        raise RuntimeError(
            "Reencolado rechazado: SP_APPROVED_BUDGETS_ROOT debe ser exactamente "
            f"{DEFAULT_APPROVED_ROOT!r}; recibido {approved!r}."
        )


def paged_values(token: str, url: str) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    next_url = url
    while next_url:
        payload = graph_get(token, next_url)
        values.extend(payload.get("value") or [])
        next_url = str(payload.get("@odata.nextLink") or "")
    return values


def list_drive_children(token: str, site_id: str, folder_path: str) -> list[dict[str, Any]]:
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(folder_path)}:/children"
        "?$select=id,name,webUrl,file,folder,size,lastModifiedDateTime&$top=999"
    )
    return paged_values(token, url)


def list_all_queue_items(token: str, site_id: str, list_id: str) -> list[dict[str, Any]]:
    return paged_values(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
        "?$select=id,createdDateTime&$expand=fields&$top=999",
    )


def delete_drive_item(token: str, site_id: str, item_id: str) -> None:
    graph_request(
        token,
        "DELETE",
        f"{GRAPH_BASE}/sites/{site_id}/drive/items/{item_id}",
        expected=(204,),
    )


def delete_queue_item(token: str, site_id: str, list_id: str, item_id: str) -> None:
    graph_request(
        token,
        "DELETE",
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}",
        expected=(204,),
    )


def approved_budget_files(children: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        [
            item
            for item in children
            if item.get("file") is not None
            and str(item.get("name") or "").lower().endswith(".xlsx")
            and not str(item.get("name") or "").startswith("~$")
        ],
        key=lambda item: str(item.get("name") or "").casefold(),
    )


def active_items_to_delete(
    children: list[dict[str, Any]],
    queue_only: bool,
) -> list[dict[str, Any]]:
    return [] if queue_only else children


def queue_budget(
    token: str,
    site_id: str,
    list_id: str,
    approved_root: str,
    item: dict[str, Any],
) -> str:
    name = str(item.get("name") or "").strip()
    now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    fields: dict[str, Any] = {
        "Title": name,
        "EventType": "presupuesto_aprobado",
        "Estado": "Pendiente",
        "Filename": name,
        "FileID": str(item.get("id") or ""),
        "FolderPath": f"Documentos compartidos/{normalized_path(approved_root)}/",
        "CreatedTime": now,
        "Intentos": "0",
        "UltimoError": "",
        "Notas": "Reencolado por reinicio controlado de Sistema 1.",
    }
    created = graph_post(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items",
        {"fields": fields},
    )
    return str(created.get("id") or "")


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Vacía la cola técnica y Proyectos Activos, y reencola los presupuestos "
            "oficiales para una regeneración limpia de Sistema 1."
        )
    )
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--confirm-active-root", default="")
    parser.add_argument(
        "--queue-only",
        action="store_true",
        help="Vacía solamente la cola y reencola presupuestos; no toca Proyectos Activos.",
    )
    parser.add_argument(
        "--reset-related-tracking",
        action="store_true",
        help=(
            "Elimina control y notificaciones solamente para los ProyectoID "
            "derivados de los presupuestos reencolados."
        ),
    )
    args = parser.parse_args()

    settings = load_settings()
    active_root = normalized_path(settings.active_projects_root)
    approved_root = normalized_path(
        os.getenv("SP_APPROVED_BUDGETS_ROOT", DEFAULT_APPROVED_ROOT)
    )
    validate_roots(active_root, approved_root, args.confirm_active_root)

    token = acquire_token(settings)
    print_token_diagnostics(token)
    site = resolve_site(token, settings)
    site_id = str(site["id"])
    queue_list = resolve_list(
        token,
        site_id,
        settings.queue_list_name,
        settings.queue_list_id,
    )
    active_children = list_drive_children(token, site_id, active_root)
    queue_items = list_all_queue_items(token, site_id, str(queue_list["id"]))
    source_children = list_drive_children(token, site_id, approved_root)
    budgets = approved_budget_files(source_children)
    project_ids = {
        derive_project_identity(str(item.get("name") or "")).project_id.casefold()
        for item in budgets
    }
    related_control_items: list[dict[str, Any]] = []
    related_notification_items: list[dict[str, Any]] = []
    control_list = None
    notification_list = None
    if args.reset_related_tracking:
        control_list = resolve_list(
            token,
            site_id,
            settings.control_list_name,
            settings.control_list_id,
        )
        notification_list = resolve_list(
            token,
            site_id,
            os.getenv("SP_NOTIFICATION_LIST_NAME", "Cola_Notificaciones_Gantt"),
            os.getenv("SP_NOTIFICATION_LIST_ID") or None,
        )
        related_control_items = [
            item
            for item in list_all_queue_items(
                token,
                site_id,
                str(control_list["id"]),
            )
            if str((item.get("fields") or {}).get("ProyectoID") or "")
            .strip()
            .casefold()
            in project_ids
        ]
        related_notification_items = [
            item
            for item in list_all_queue_items(
                token,
                site_id,
                str(notification_list["id"]),
            )
            if str((item.get("fields") or {}).get("ProyectoID") or "")
            .strip()
            .casefold()
            in project_ids
        ]

    print("Sistema 1 controlled reset plan:")
    print(f"- Active root: {active_root}")
    active_to_delete = active_items_to_delete(active_children, args.queue_only)
    print(f"- Active children to delete: {len(active_to_delete)}")
    for item in active_to_delete:
        print(f"  - {item.get('name')}")
    print(f"- Queue items to delete: {len(queue_items)}")
    print(f"- Official budgets to requeue: {len(budgets)}")
    for item in budgets:
        print(f"  - {item.get('name')}")
    print(f"- Related control items to delete: {len(related_control_items)}")
    print(
        "- Related notification items to delete: "
        f"{len(related_notification_items)}"
    )

    if not args.execute:
        print("Dry run only. Use --execute with the exact --confirm-active-root value.")
        return 0
    if not budgets:
        raise RuntimeError(
            "Reinicio cancelado: no se encontraron presupuestos .xlsx en Presupuestos Aprobados."
        )

    for item in active_to_delete:
        delete_drive_item(token, site_id, str(item["id"]))
    for item in queue_items:
        delete_queue_item(token, site_id, str(queue_list["id"]), str(item["id"]))
    if control_list:
        for item in related_control_items:
            delete_queue_item(
                token,
                site_id,
                str(control_list["id"]),
                str(item["id"]),
            )
    if notification_list:
        for item in related_notification_items:
            delete_queue_item(
                token,
                site_id,
                str(notification_list["id"]),
                str(item["id"]),
            )

    created_ids = [
        queue_budget(token, site_id, str(queue_list["id"]), approved_root, item)
        for item in budgets
    ]
    print(
        "Controlled reset completed: "
        f"deleted_active={len(active_to_delete)} "
        f"deleted_queue={len(queue_items)} "
        f"deleted_control={len(related_control_items)} "
        f"deleted_notifications={len(related_notification_items)} "
        f"requeued={len(created_ids)} "
        f"queue_ids={','.join(created_ids)}"
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
