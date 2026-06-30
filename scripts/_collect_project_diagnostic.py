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


SOURCE_ROOT = "Proyectos/Presupuestos Aprobados"
ACTIVE_ROOT = "Proyectos/Proyectos Activos"


def children(token: str, site_id: str, path: str) -> list[dict]:
    data = graph_get(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(path)}:/children",
    )
    return data.get("value") or []


def main() -> None:
    project_id = os.environ["DIAGNOSTIC_PROJECT_ID"].strip()
    if not project_id or "/" in project_id or "\\" in project_id:
        raise RuntimeError("DIAGNOSTIC_PROJECT_ID inválido.")
    settings = load_settings()
    token = acquire_token(settings)
    site_id = str(resolve_site(token, settings)["id"])
    source = next(
        item
        for item in children(token, site_id, SOURCE_ROOT)
        if str(item.get("name") or "").startswith(project_id)
        and str(item.get("name") or "").lower().endswith(".xlsx")
    )
    project = next(
        item
        for item in children(token, site_id, ACTIVE_ROOT)
        if str(item.get("name") or "").startswith(f"{project_id}_")
        and item.get("folder") is not None
    )
    source_path = f"{SOURCE_ROOT}/{source['name']}"
    gantt_path = (
        f"{ACTIVE_ROOT}/{project['name']}/gantts/working/"
        f"{project_id}_gantt_WORKING.xlsx"
    )
    output = Path(os.getenv("DIAGNOSTIC_OUTPUT", "diagnostic"))
    download_drive_file(token, site_id, source_path, output / source["name"])
    download_drive_file(
        token,
        site_id,
        gantt_path,
        output / f"{project_id}_gantt_WORKING.xlsx",
    )


if __name__ == "__main__":
    main()
