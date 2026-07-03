from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "power_automate" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from deploy_activity_status_alerts_flow import (  # noqa: E402
    ACTIVITY_SCRIPT_ID,
    LAST_ALERT_INTERNAL,
    build_definition,
)


def test_activity_status_flow_runs_every_15_minutes() -> None:
    definition = build_definition()
    recurrence = definition["triggers"]["Every_15_minutes"]["recurrence"]

    assert recurrence["frequency"] == "Minute"
    assert recurrence["interval"] == 15
    assert recurrence["timeZone"] == "SA Pacific Standard Time"


def test_activity_status_flow_reads_script_and_sends_email() -> None:
    definition = build_definition()
    serialized = str(definition)

    assert ACTIVITY_SCRIPT_ID in serialized
    assert "RunScriptProd" in serialized
    assert "SendEmailV2" in serialized
    assert "IngenieroEmail" in serialized
    assert "SupervisoresEmail" in serialized
    assert LAST_ALERT_INTERNAL in serialized


def test_status_only_change_returns_general_status_to_actual() -> None:
    definition = build_definition()
    serialized = str(definition)

    assert "If_status_only_change" in serialized
    assert "PlanificacionFingerprint" in serialized
    assert "EstadoActividadesFingerprint" in serialized
    assert "StatusExcelDeseado" in serialized
    assert "EstadoSyncExcel" in serialized
    assert "'Actual'" in serialized


def test_flow_has_no_python_or_github_dependency() -> None:
    serialized = str(build_definition()).casefold()

    assert "github" not in serialized
    assert "python" not in serialized
