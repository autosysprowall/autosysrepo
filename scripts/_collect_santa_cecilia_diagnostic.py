from __future__ import annotations

import os
from pathlib import Path

from sistema1_poll_queue import (
    GRAPH_BASE,
    acquire_token,
    download_drive_file,
    encoded_drive_path,
    graph_get,
    load_settings,
    resolve_site,
)


PROJECT_ID = "2025-135"
SOURCE_ROOT = "Proyectos/Presupuestos Aprobados"
ACTIVE_ROOT = "Proyectos/Proyectos Activos"


def children(token: str, site_id: str, path: str) -> list[dict]:
    data = graph_get(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(path)}:/children",
    )
    return data.get("value") or []


def main() -> None:
    settings = load_settings()
    token = acquire_token(settings)
    site_id = str(resolve_site(token, settings)["id"])
    source = next(
        item
        for item in children(token, site_id, SOURCE_ROOT)
        if str(item.get("name") or "").startswith(PROJECT_ID)
        and str(item.get("name") or "").lower().endswith(".xlsx")
    )
    project = next(
        item
        for item in children(token, site_id, ACTIVE_ROOT)
        if str(item.get("name") or "").startswith(f"{PROJECT_ID}_")
        and item.get("folder") is not None
    )
    source_path = f"{SOURCE_ROOT}/{source['name']}"
    gantt_path = (
        f"{ACTIVE_ROOT}/{project['name']}/gantts/working/"
        f"{PROJECT_ID}_gantt_WORKING.xlsx"
    )
    output = Path(os.getenv("DIAGNOSTIC_OUTPUT", "diagnostic"))
    download_drive_file(token, site_id, source_path, output / source["name"])
    download_drive_file(
        token,
        site_id,
        gantt_path,
        output / f"{PROJECT_ID}_gantt_WORKING.xlsx",
    )
    print("Downloaded source and generated workbook for diagnostic.")


if __name__ == "__main__":
    main()
