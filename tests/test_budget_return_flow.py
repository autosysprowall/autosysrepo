from __future__ import annotations

import sys
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from sistema1_poll_queue import (  # noqa: E402
    drive_item_uploader_email,
    rejection_queue_updates,
)


def test_rejection_keeps_actionable_reason_and_power_automate_marker() -> None:
    updates = rejection_queue_updates(
        "CONTENIDO_DE_COLUMNAS_INVALIDO: Cantidad contiene texto."
    )

    assert updates["Estado"] == "RequiereRevision"
    assert updates["UltimoError"].startswith("CONTENIDO_DE_COLUMNAS_INVALIDO")
    assert updates["Notas"].startswith("DEVOLVER_PRESUPUESTO")
    assert "presupuesto original" in updates["Notas"]


def test_rejection_reason_is_bounded_for_sharepoint() -> None:
    updates = rejection_queue_updates("x" * 500)
    assert len(updates["UltimoError"]) == 240


def test_queue_processor_recovers_uploader_from_drive_item() -> None:
    assert drive_item_uploader_email(
        {
            "createdBy": {"user": {"displayName": "Invitado"}},
            "lastModifiedBy": {
                "user": {"email": "uploader@example.com"}
            },
        }
    ) == "uploader@example.com"
