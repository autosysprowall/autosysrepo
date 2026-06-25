from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

import msal
import requests


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


@dataclass(frozen=True)
class Settings:
    tenant_id: str
    client_id: str
    client_secret: str
    site_hostname: str
    site_path: str
    queue_list_name: str
    queue_list_id: str | None


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
            + ". GitHub Actions necesita autenticacion no interactiva; el device login local no sirve en runner."
        )

    return Settings(
        tenant_id=env_value("MS_GRAPH_TENANT_ID"),
        client_id=env_value("MS_GRAPH_CLIENT_ID"),
        client_secret=env_value("MS_GRAPH_CLIENT_SECRET"),
        site_hostname=env_value("SP_SITE_HOSTNAME", "sciprowall.sharepoint.com"),
        site_path=env_value("SP_SITE_PATH", "/sites/PROYECTOSPROWALL").strip("/"),
        queue_list_name=env_value("SP_QUEUE_LIST_NAME", "Cola_Automatizacion_Proyectos"),
        queue_list_id=env_value("SP_QUEUE_LIST_ID") or None,
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


def graph_get(token: str, url: str) -> dict[str, Any]:
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if response.status_code >= 400:
        raise RuntimeError(f"Graph GET {response.status_code}: {response.text[:2000]}")
    return response.json()


def resolve_site(token: str, settings: Settings) -> dict[str, Any]:
    url = f"{GRAPH_BASE}/sites/{settings.site_hostname}:/{settings.site_path}"
    return graph_get(token, url)


def resolve_list(token: str, site_id: str, settings: Settings) -> dict[str, Any]:
    if settings.queue_list_id:
        return graph_get(token, f"{GRAPH_BASE}/sites/{site_id}/lists/{settings.queue_list_id}")
    escaped = settings.queue_list_name.replace("'", "''")
    data = graph_get(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,displayName,name,webUrl&$filter=displayName eq '{escaped}'",
    )
    values = data.get("value") or []
    if not values:
        raise RuntimeError(f"No se encontro la lista SharePoint: {settings.queue_list_name}")
    return values[0]


def list_queue_items(token: str, site_id: str, list_id: str, top: int) -> list[dict[str, Any]]:
    select_fields = (
        "Title,EventType,Estado,Filename,FileID,FolderPath,FileLink,"
        "CreatedByEmail,CreatedTime,Intentos,UltimoError,FechaProcesado,ProyectoID,Notas"
    )
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
        f"?$top={top}&$orderby=createdDateTime desc&expand=fields($select={select_fields})"
    )
    data = graph_get(token, url)
    return data.get("value") or []


def item_status(item: dict[str, Any]) -> str:
    fields = item.get("fields") or {}
    return str(fields.get("Estado") or "").strip()


def print_summary(site: dict[str, Any], queue_list: dict[str, Any], items: list[dict[str, Any]]) -> None:
    pending = [item for item in items if item_status(item).casefold() == "pendiente"]
    print("Autosys Sistema 1 - Queue poll")
    print(f"Run UTC: {datetime.now(timezone.utc).isoformat()}")
    print(f"Site: {site.get('webUrl') or site.get('displayName')}")
    print(f"Queue list: {queue_list.get('displayName')} ({queue_list.get('id')})")
    print(f"Items scanned: {len(items)}")
    print(f"Pending items: {len(pending)}")
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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Lee la cola SharePoint del Sistema 1 para GitHub Actions."
    )
    parser.add_argument("--top", type=int, default=50, help="Cantidad maxima de items recientes a leer.")
    parser.add_argument(
        "--fail-on-pending",
        action="store_true",
        help="Falla si hay pendientes. Util para pruebas, no para schedule normal.",
    )
    args = parser.parse_args()

    settings = load_settings()
    token = acquire_token(settings)
    site = resolve_site(token, settings)
    queue_list = resolve_list(token, site["id"], settings)
    items = list_queue_items(token, site["id"], queue_list["id"], max(1, args.top))
    print_summary(site, queue_list, items)
    if args.fail_on_pending and any(item_status(item).casefold() == "pendiente" for item in items):
        return 2
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
