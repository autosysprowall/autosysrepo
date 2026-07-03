from __future__ import annotations

import os
import sys
from urllib.parse import quote

from sistema1_poll_queue import (
    GRAPH_BASE,
    acquire_token,
    graph_get,
    list_columns,
    load_settings,
    resolve_list,
    resolve_site,
)


SCRIPT_NAME = "GetGanttActivityStatus"
REQUIRED_COLUMNS = {
    "PlanificacionFingerprint",
    "EstadoActividadesFingerprint",
    "AlertasActividadesFingerprint",
    "ActividadesAtrasadas",
    "ResumenActividadesAtrasadas",
    "FechaLecturaActividades",
    "UltimaAlertaActividadesFingerprint",
    "FechaUltimoCorreoActividades",
    "UltimoErrorActividades",
}


def main() -> int:
    settings = load_settings()
    token = acquire_token(settings)
    site = resolve_site(token, settings)
    control = resolve_list(
        token,
        str(site["id"]),
        settings.control_list_name,
        settings.control_list_id,
    )
    columns = list_columns(token, str(site["id"]), str(control["id"]))
    available = {
        str(value)
        for column in columns
        for value in (column.get("name"), column.get("displayName"))
        if value
    }
    missing = sorted(REQUIRED_COLUMNS - available)
    print(f"Activity status columns missing: {','.join(missing) or 'none'}")
    for column in columns:
        internal_name = str(column.get("name") or "")
        display_name = str(column.get("displayName") or "")
        if (
            internal_name in REQUIRED_COLUMNS
            or display_name in REQUIRED_COLUMNS
            or any(
                internal_name.startswith(required[:20])
                for required in REQUIRED_COLUMNS
            )
        ):
            print(
                "ACTIVITY_STATUS_COLUMN "
                f"display={display_name!r} internal={internal_name!r}"
            )

    owner = os.getenv(
        "OFFICE_SCRIPT_OWNER_UPN",
        "auto.sys@prowallpanama.com",
    ).strip()
    search_url = (
        f"{GRAPH_BASE}/users/{quote(owner, safe='')}/drive/root/delta"
        "?$select=id,name,webUrl,lastModifiedDateTime,parentReference"
    )
    items: list[dict[str, object]] = []
    while search_url:
        page = graph_get(token, search_url)
        items.extend(page.get("value") or [])
        search_url = str(page.get("@odata.nextLink") or "")
    matches = [
        item
        for item in items
        if SCRIPT_NAME.casefold()
        in str(item.get("name") or "").casefold()
    ]
    if not matches:
        candidates = sorted(
            [
                item
                for item in items
                if (
                    str(item.get("name") or "").casefold().endswith(".osts")
                    or "office scripts"
                    in str(
                        (item.get("parentReference") or {}).get("path")
                        if isinstance(item.get("parentReference"), dict)
                        else ""
                    ).casefold()
                )
            ],
            key=lambda item: str(item.get("lastModifiedDateTime") or ""),
            reverse=True,
        )
        for item in candidates[:20]:
            print(
                "OFFICE_SCRIPT_CANDIDATE "
                f"name={item.get('name')!r} "
                f"id={item.get('id')} "
                f"modified={item.get('lastModifiedDateTime')}"
            )
        print(
            f"ERROR: Office Script {SCRIPT_NAME!r} no encontrado "
            f"en el OneDrive de {owner}.",
            file=sys.stderr,
        )
        return 2
    for item in matches:
        print(
            "OFFICE_SCRIPT_MATCH "
            f"name={item.get('name')!r} "
            f"id={item.get('id')} "
            f"modified={item.get('lastModifiedDateTime')}"
        )
    return 0 if not missing else 3


if __name__ == "__main__":
    raise SystemExit(main())
