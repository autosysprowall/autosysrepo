from __future__ import annotations

import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.automation_dispatcher import (  # noqa: E402
    ControlRecord,
    apply_notification_delivery_mode,
    assignment_notification,
    normalize_emails,
    tracking_notification,
)


def main() -> int:
    now = datetime.now(timezone.utc)
    record = ControlRecord(
        item_id="EMAIL-SANITY-CHECK",
        project_id="TEST-000",
        project_name="Prueba de plantillas",
        gantt_link="https://example.invalid/gantt-working.xlsx",
        engineer_email="ingeniero-prueba@example.invalid",
        supervisors_email="supervisor-prueba@example.invalid",
    )
    templates = [
        assignment_notification(record, now, now + timedelta(days=9)),
        tracking_notification(record, "Advertencia1", 3),
        tracking_notification(record, "Advertencia2", 6),
        tracking_notification(record, "Vencimiento", 9),
    ]
    delivered = [apply_notification_delivery_mode(item) for item in templates]
    mode = os.getenv("NOTIFICATION_DELIVERY_MODE", "test").strip().casefold()
    if mode != "live":
        expected = normalize_emails(
            os.getenv(
                "NOTIFICATION_TEST_RECIPIENT",
                "auto.sys@prowallpanama.com",
            )
        )
        for item in delivered:
            if item.to != expected or item.cc:
                raise RuntimeError(
                    f"Modo test inseguro para {item.kind}: To/CC no fueron redirigidos."
                )
            if not item.subject.startswith("[PRUEBA] "):
                raise RuntimeError(
                    f"Modo test inseguro para {item.kind}: falta prefijo de prueba."
                )
    print(
        "Email delivery sanity check passed: "
        f"mode={mode or 'test'} templates={len(delivered)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
