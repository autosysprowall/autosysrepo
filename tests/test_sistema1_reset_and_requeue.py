from __future__ import annotations

import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from sistema1_reset_and_requeue import (  # noqa: E402
    active_items_to_delete,
    approved_budget_files,
    drive_item_created_by_email,
    queue_created_by_by_filename,
    sharepoint_file_identifier,
    validate_roots,
)


def test_queue_only_never_selects_active_items_for_deletion() -> None:
    children = [{"id": "active-project", "name": "Proyecto existente"}]
    assert active_items_to_delete(children, queue_only=True) == []
    assert active_items_to_delete(children, queue_only=False) == children


def test_validate_roots_accepts_only_expected_system1_paths() -> None:
    validate_roots(
        "Proyectos/Proyectos Activos",
        "Proyectos/Presupuestos Aprobados",
        "Proyectos/Proyectos Activos",
    )


@pytest.mark.parametrize(
    ("active_root", "approved_root", "confirmation"),
    [
        (
            "Proyectos/Proyectos Terminados",
            "Proyectos/Presupuestos Aprobados",
            "Proyectos/Proyectos Terminados",
        ),
        (
            "Proyectos/Proyectos Activos",
            "Proyectos/Proyectos Terminados",
            "Proyectos/Proyectos Activos",
        ),
        (
            "Proyectos/Proyectos Activos",
            "Proyectos/Presupuestos Aprobados",
            "Proyectos/02_Activos",
        ),
    ],
)
def test_validate_roots_rejects_unapproved_or_forbidden_paths(
    active_root: str,
    approved_root: str,
    confirmation: str,
) -> None:
    with pytest.raises(RuntimeError):
        validate_roots(active_root, approved_root, confirmation)


def test_approved_budget_files_only_returns_real_xlsx_files() -> None:
    children = [
        {"id": "3", "name": "B.xlsx", "file": {}},
        {"id": "2", "name": "~$A.xlsx", "file": {}},
        {"id": "1", "name": "A.xlsx", "file": {}},
        {"id": "4", "name": "notes.pdf", "file": {}},
        {"id": "5", "name": "folder", "folder": {}},
    ]

    assert [item["name"] for item in approved_budget_files(children)] == [
        "A.xlsx",
        "B.xlsx",
    ]


def test_sharepoint_file_identifier_is_double_url_encoded() -> None:
    identifier = sharepoint_file_identifier(
        "Proyectos/Presupuestos Aprobados",
        "2025-111 Presupuesto prueba.xlsx",
    )

    assert identifier == (
        "%252FDocumentos%2Bcompartidos%252FProyectos%252F"
        "Presupuestos%2BAprobados%252F2025-111%2BPresupuesto%2Bprueba.xlsx"
    )


def test_drive_item_created_by_email_uses_graph_metadata() -> None:
    assert drive_item_created_by_email(
        {"createdBy": {"user": {"email": "comercial@example.com"}}}
    ) == "comercial@example.com"


def test_drive_item_created_by_email_falls_back_to_last_modifier() -> None:
    assert drive_item_created_by_email(
        {
            "createdBy": {"user": {"displayName": "Invitado"}},
            "lastModifiedBy": {
                "user": {"userPrincipalName": "uploader@example.com"}
            },
        }
    ) == "uploader@example.com"


def test_drive_item_uploader_prefers_last_modifier_over_original_creator() -> None:
    assert drive_item_created_by_email(
        {
            "createdBy": {
                "user": {"email": "original-creator@example.com"}
            },
            "lastModifiedBy": {
                "user": {"email": "folder-uploader@example.com"}
            },
        }
    ) == "folder-uploader@example.com"


def test_reset_preserves_existing_uploader_by_filename() -> None:
    assert queue_created_by_by_filename(
        [
            {
                "fields": {
                    "Filename": "Presupuesto.xlsx",
                    "CreatedByEmail": "uploader@example.com",
                }
            }
        ]
    ) == {"presupuesto.xlsx": "uploader@example.com"}
