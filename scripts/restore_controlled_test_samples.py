from __future__ import annotations

import base64
import os
from typing import Any

from sistema1_poll_queue import (
    GRAPH_BASE,
    acquire_token,
    encoded_drive_path,
    graph_put_bytes,
    load_settings,
    resolve_site,
)


APPROVED_ROOT = "Proyectos/Presupuestos Aprobados"
SAMPLES: tuple[tuple[str, str, int], ...] = (
    (
        "2024-047 Desarrollo Urbanistico MILLA 9 PTAR - Ing Jorge Perez.xlsx",
        "RESTORE_SAMPLE_1_B64",
        3,
    ),
    (
        "2026-058 Casa de las Baterias, Banco General ZAPATAS PEDESTALES - Ing. Jorge Perez.xlsx",
        "RESTORE_SAMPLE_2_B64",
        2,
    ),
)


def decode_chunks(prefix: str, count: int) -> bytes:
    chunks = [
        os.getenv(f"{prefix}_{index}", "").strip()
        for index in range(1, count + 1)
    ]
    if any(not chunk for chunk in chunks):
        raise RuntimeError(f"Faltan secretos temporales para {prefix}.")
    return base64.b64decode("".join(chunks), validate=True)


def upload_sample(
    token: str,
    site_id: str,
    name: str,
    content: bytes,
) -> dict[str, Any]:
    if not content.startswith(b"PK"):
        raise RuntimeError(f"{name} no contiene un workbook XLSX válido.")
    path = f"{APPROVED_ROOT}/{name}"
    return graph_put_bytes(
        token,
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{encoded_drive_path(path)}:/content",
        content,
    )


def main() -> int:
    settings = load_settings()
    token = acquire_token(settings)
    site = resolve_site(token, settings)
    for name, prefix, count in SAMPLES:
        content = decode_chunks(prefix, count)
        uploaded = upload_sample(token, str(site["id"]), name, content)
        print(
            "Restored controlled sample: "
            f"name={name} size={len(content)} id={uploaded.get('id', '')}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
