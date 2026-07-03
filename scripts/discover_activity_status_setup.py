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
        str(column.get("name") or column.get("displayName") or "")
        for column in columns
    }
    missing = sorted(REQUIRED_COLUMNS - available)
    print(f"Activity status columns missing: {','.join(missing) or 'none'}")

    owner = os.getenv(
        "OFFICE_SCRIPT_OWNER_UPN",
        "auto.sys@prowallpanama.com",
    ).strip()
    search_url = (
        f"{GRAPH_BASE}/users/{quote(owner, safe='')}/drive/root/"
        f"search(q='{quote(SCRIPT_NAME, safe='')}')"
        "?$select=id,name,webUrl,lastModifiedDateTime,parentReference"
    )
    data = graph_get(token, search_url)
    matches = [
        item
        for item in data.get("value") or []
        if SCRIPT_NAME.casefold()
        in str(item.get("name") or "").casefold()
    ]
    if not matches:
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
