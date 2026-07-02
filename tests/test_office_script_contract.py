from pathlib import Path


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "power_automate"
    / "office_scripts"
    / "SetGanttStatus.ts"
)


def test_office_script_exposes_three_statuses_and_updates_validation() -> None:
    content = SCRIPT.read_text(encoding="utf-8")
    assert '["Actual", "En Progreso", "Entregar"]' in content
    assert 'source: "Actual,En Progreso,Entregar"' in content
    assert 'getRange("B6")' in content


def test_office_script_protects_entregar_from_automatic_progress_sync() -> None:
    content = SCRIPT.read_text(encoding="utf-8")
    assert 'current === "Entregar" && desired !== "Actual"' in content
    assert 'code: "PROTEGIDO_ENTREGAR"' in content
    assert "allowReplaceEntregar" in content
