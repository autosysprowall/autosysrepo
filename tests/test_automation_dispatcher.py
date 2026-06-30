from __future__ import annotations

import io
import unittest
from datetime import datetime, timedelta, timezone

from openpyxl import Workbook

from src.automation_dispatcher import (
    AutomationService,
    ControlRecord,
    Notification,
    WorkbookMetadata,
    due_tracking_kind,
    extract_workbook_metadata,
    run_system2,
)


NOW = datetime(2026, 7, 1, 12, 0, tzinfo=timezone.utc)


class FakeBackend:
    def __init__(self, metadata: WorkbookMetadata | None = None) -> None:
        self.metadata = metadata or WorkbookMetadata()
        self.patches: list[tuple[str, dict]] = []
        self.grants: list[tuple[str, str]] = []
        self.notifications: list[tuple[str, Notification]] = []
        self.notification_keys: set[tuple[str, str]] = set()
        self.permission_error: Exception | None = None
        self.notification_error: Exception | None = None
        self.metadata_error: Exception | None = None

    def patch_control(self, item_id: str, updates: dict) -> None:
        self.patches.append((item_id, updates))

    def load_metadata(self, record: ControlRecord, *, prefer_gantt: bool = True) -> WorkbookMetadata:
        if self.metadata_error:
            raise self.metadata_error
        return self.metadata

    def grant_edit_access(self, record: ControlRecord, email: str) -> None:
        if self.permission_error:
            raise self.permission_error
        self.grants.append((record.item_id, email))

    def notification_exists(self, item_id: str, kind: str) -> bool:
        return (item_id, kind.casefold()) in self.notification_keys

    def queue_notification(self, record: ControlRecord, notification: Notification) -> None:
        if self.notification_error:
            raise self.notification_error
        key = (record.item_id, notification.kind.casefold())
        if key in self.notification_keys:
            raise AssertionError(f"duplicate notification: {key}")
        self.notification_keys.add(key)
        self.notifications.append((record.item_id, notification))


def record(**changes) -> ControlRecord:
    defaults = {
        "item_id": "10",
        "project_id": "2026-001",
        "project_name": "Proyecto Prueba",
        "budget_link": "https://contoso.sharepoint.com/Presupuesto.xlsx",
        "gantt_link": "https://contoso.sharepoint.com/Proyectos/Proyectos%20Activos/gantt_working.xlsx",
        "engineer_email": "ingeniero@example.com",
        "supervisors_email": "supervisor@example.com",
        "state": "Pendiente de asignación",
    }
    defaults.update(changes)
    return ControlRecord(**defaults)


class AssignmentTests(unittest.TestCase):
    def test_valid_assignment_grants_permission_and_queues_one_email(self) -> None:
        backend = FakeBackend()
        result = AutomationService(backend, NOW).assign(record())

        self.assertTrue(result.permission_granted)
        self.assertEqual("Asignado", result.state)
        self.assertEqual(NOW + timedelta(days=9), result.deadline)
        self.assertEqual([("10", "ingeniero@example.com")], backend.grants)
        self.assertEqual(["AsignacionGantt"], [item.kind for _, item in backend.notifications])
        merged = {key: value for _, patch in backend.patches for key, value in patch.items()}
        self.assertEqual("Asignado", merged["EstadoGantt"])
        self.assertTrue(merged["PermisoGanttOtorgado"])

    def test_missing_email_does_not_share_or_queue(self) -> None:
        backend = FakeBackend(WorkbookMetadata())
        result = AutomationService(backend, NOW).assign(record(engineer_email=""))

        self.assertEqual("Pendiente de asignación", result.state)
        self.assertEqual([], backend.grants)
        self.assertEqual([], backend.notifications)
        self.assertTrue(
            any("vacío o inválido" in patch.get("UltimoErrorTracking", "") for _, patch in backend.patches)
        )

    def test_permission_failure_does_not_mark_permission_or_queue_email(self) -> None:
        backend = FakeBackend()
        backend.permission_error = RuntimeError("Graph 403")
        AutomationService(backend, NOW).assign(record())

        self.assertEqual([], backend.notifications)
        self.assertFalse(
            any(patch.get("PermisoGanttOtorgado") is True for _, patch in backend.patches)
        )
        self.assertTrue(
            any("Graph 403" in patch.get("UltimoErrorTracking", "") for _, patch in backend.patches)
        )

    def test_rerun_does_not_duplicate_permission_or_assignment_email(self) -> None:
        backend = FakeBackend()
        backend.notification_keys.add(("10", "asignaciongantt"))
        current = record(
            state="Asignado",
            permission_granted=True,
            assignment_date=NOW,
            deadline=NOW + timedelta(days=9),
        )
        AutomationService(backend, NOW).assign(current)
        AutomationService(backend, NOW).assign(current)

        self.assertEqual([], backend.grants)
        self.assertEqual([], backend.notifications)

    def test_assignment_preserves_en_progreso_state(self) -> None:
        backend = FakeBackend()
        current = record(
            state="En progreso",
            permission_granted=True,
            assignment_date=NOW - timedelta(days=1),
            deadline=NOW + timedelta(days=8),
        )
        result = AutomationService(backend, NOW).assign(current)
        self.assertEqual("En progreso", result.state)

    def test_existing_engineer_is_assigned_even_if_optional_metadata_fails(self) -> None:
        backend = FakeBackend()
        backend.metadata_error = RuntimeError("Datos unreadable")
        result = AutomationService(backend, NOW).assign(record(supervisors_email=""))
        self.assertEqual("Asignado", result.state)
        self.assertEqual([("10", "ingeniero@example.com")], backend.grants)

    def test_supervisors_are_enriched_when_engineer_already_exists(self) -> None:
        backend = FakeBackend(WorkbookMetadata(supervisors_email="jefe@example.com"))
        result = AutomationService(backend, NOW).assign(record(supervisors_email=""))
        self.assertEqual("jefe@example.com", result.supervisors_email)
        self.assertTrue(
            any(patch.get("SupervisoresEmail") == "jefe@example.com" for _, patch in backend.patches)
        )

    def test_proyectos_terminados_path_is_rejected(self) -> None:
        backend = FakeBackend()
        result = AutomationService(backend, NOW).assign(
            record(gantt_link="https://contoso/Proyectos/Proyectos%20Terminados/gantt.xlsx")
        )
        self.assertEqual("Requiere revisión manual", result.state)
        self.assertEqual([], backend.grants)
        self.assertEqual([], backend.notifications)


class TrackingTests(unittest.TestCase):
    def test_day_3_queues_warning_1_once(self) -> None:
        backend = FakeBackend()
        current = record(
            state="Asignado",
            permission_granted=True,
            assignment_email_sent=True,
            assignment_date=NOW - timedelta(days=3),
        )
        service = AutomationService(backend, NOW)
        service.track(current)
        service.track(current)
        self.assertEqual(["Advertencia1"], [item.kind for _, item in backend.notifications])

    def test_day_6_queues_warning_2_with_supervisor_cc(self) -> None:
        backend = FakeBackend()
        current = record(
            state="En progreso",
            permission_granted=True,
            assignment_email_sent=True,
            warning1_sent=True,
            assignment_date=NOW - timedelta(days=6),
        )
        AutomationService(backend, NOW).track(current)
        notification = backend.notifications[0][1]
        self.assertEqual("Advertencia2", notification.kind)
        self.assertEqual("supervisor@example.com", notification.cc)

    def test_day_9_marks_expired_and_queues_escalation_without_third_warning(self) -> None:
        backend = FakeBackend()
        current = record(
            state="Asignado",
            permission_granted=True,
            assignment_email_sent=True,
            warning1_sent=True,
            warning2_sent=True,
            assignment_date=NOW - timedelta(days=9),
        )
        result = AutomationService(backend, NOW).track(current)
        self.assertEqual("Vencido", result.state)
        self.assertEqual(["Vencimiento"], [item.kind for _, item in backend.notifications])
        self.assertTrue(any(patch.get("EstadoGantt") == "Vencido" for _, patch in backend.patches))
        closed = record(
            state="Vencido",
            assignment_date=NOW - timedelta(days=20),
            expiration_notified=True,
        )
        self.assertEqual("", due_tracking_kind(closed, NOW))

    def test_day_9_does_not_mark_expired_when_queueing_fails(self) -> None:
        backend = FakeBackend()
        backend.notification_error = RuntimeError("SharePoint unavailable")
        current = record(
            state="En progreso",
            permission_granted=True,
            assignment_email_sent=True,
            warning1_sent=True,
            warning2_sent=True,
            assignment_date=NOW - timedelta(days=9),
        )
        with self.assertRaisesRegex(RuntimeError, "SharePoint unavailable"):
            AutomationService(backend, NOW).track(current)
        self.assertFalse(any(patch.get("EstadoGantt") == "Vencido" for _, patch in backend.patches))


class StatusAndMetadataTests(unittest.TestCase):
    def test_review_status_updates_tracking_without_versioning(self) -> None:
        backend = FakeBackend()
        current = record(
            state="En progreso",
            assignment_date=NOW - timedelta(days=4),
        )
        result = AutomationService(backend, NOW).sync_excel_status(
            current,
            WorkbookMetadata(status="En revisión inicial"),
        )
        self.assertEqual("En revisión inicial", result.state)
        merged = {key: value for _, patch in backend.patches for key, value in patch.items()}
        self.assertEqual(4, merged["DiasParaCompletar"])
        self.assertNotIn("Version", json_keys(backend.patches))

    def test_extracts_engineer_supervisors_and_status_from_datos(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        ws.title = "Datos"
        ws["A1"] = "Ingeniero residente"
        ws["B1"] = "Ingeniero@Example.com"
        ws["A2"] = "Supervisores"
        ws["B2"] = "uno@example.com, dos@example.com"
        ws["A3"] = "EstadoGantt"
        ws["B3"] = "En revisión inicial"
        stream = io.BytesIO()
        workbook.save(stream)
        workbook.close()

        metadata = extract_workbook_metadata(stream.getvalue())
        self.assertEqual("ingeniero@example.com", metadata.engineer_email)
        self.assertEqual("uno@example.com;dos@example.com", metadata.supervisors_email)
        self.assertEqual("En revisión inicial", metadata.status)


class DispatcherTests(unittest.TestCase):
    def test_empty_lists_finish_with_zero_errors(self) -> None:
        class EmptyBackend(FakeBackend):
            def process_status_events(self, now: datetime, max_items: int = 20) -> int:
                return 0

            def control_records(self) -> list[ControlRecord]:
                return []

        summary = run_system2(EmptyBackend(), NOW, 20)
        self.assertEqual(0, summary["control_items"])
        self.assertEqual(0, summary["errors"])


def json_keys(patches: list[tuple[str, dict]]) -> str:
    return " ".join(key for _, patch in patches for key in patch)


if __name__ == "__main__":
    unittest.main()
