from __future__ import annotations

import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from sistema1_reset_and_requeue import (  # noqa: E402
    approved_budget_files,
    validate_roots,
)


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
