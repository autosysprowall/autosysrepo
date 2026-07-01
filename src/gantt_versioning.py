from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from openpyxl import load_workbook


HEADER_SCAN_LIMIT = 30
COST_TOLERANCE = 0.01


def _normalized(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().casefold()


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(text, pattern).date()
            except ValueError:
                continue
    return None


def _find_gantt_header(ws) -> tuple[int, dict[str, int]]:
    best: tuple[int, dict[str, int]] | None = None
    for row_number in range(1, min(ws.max_row, HEADER_SCAN_LIMIT) + 1):
        headers = {
            _normalized(ws.cell(row_number, column).value): column
            for column in range(1, ws.max_column + 1)
            if _normalized(ws.cell(row_number, column).value)
        }
        if "actividad" in headers and "fecha de fin" in headers:
            if best is None or len(headers) > len(best[1]):
                best = (row_number, headers)
    if best is None:
        raise ValueError("No se pudo detectar el encabezado de la hoja Gantt.")
    return best


def _contractual_end(workbook) -> date | None:
    if "Datos" not in workbook.sheetnames:
        return None
    ws = workbook["Datos"]
    for row in ws.iter_rows():
        for cell in row:
            label = _normalized(cell.value)
            if "fecha" not in label or not any(
                term in label for term in ("final", "fin contractual")
            ):
                continue
            for offset in range(1, 5):
                parsed = _as_date(ws.cell(cell.row, cell.column + offset).value)
                if parsed:
                    return parsed
            parsed = _as_date(ws.cell(cell.row + 1, cell.column).value)
            if parsed:
                return parsed
    return None


def _cost_columns(headers: dict[str, int]) -> dict[str, int]:
    result: dict[str, int] = {}
    for header, column in headers.items():
        if any(term in header for term in ("costo total", "precio total")):
            if not any(term in header for term in ("utilidad", "margen")):
                result[header] = column
    return result


def _numeric(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


@dataclass(frozen=True)
class GanttSnapshot:
    costs: dict[str, float]
    baseline_costs: dict[str, float]
    contractual_end: date | None
    latest_activity_end: date | None

    @property
    def schedule_overrun(self) -> bool:
        return bool(
            self.contractual_end
            and self.latest_activity_end
            and self.latest_activity_end > self.contractual_end
        )


@dataclass(frozen=True)
class VersionDecision:
    version: str
    cost_increase: bool
    schedule_overrun: bool
    reasons: tuple[str, ...]


def snapshot_gantt(content: bytes) -> GanttSnapshot:
    workbook = load_workbook(io.BytesIO(content), data_only=True, read_only=False)
    try:
        if "Gantt" not in workbook.sheetnames:
            raise ValueError("El archivo no contiene la hoja Gantt.")
        ws = workbook["Gantt"]
        header_row, headers = _find_gantt_header(ws)
        cost_columns = _cost_columns(headers)
        costs = {header: 0.0 for header in cost_columns}
        end_column = headers["fecha de fin"]
        latest_end: date | None = None
        for row_number in range(header_row + 1, ws.max_row + 1):
            activity = ws.cell(row_number, headers["actividad"]).value
            if activity in (None, ""):
                continue
            activity_end = _as_date(ws.cell(row_number, end_column).value)
            if activity_end and (latest_end is None or activity_end > latest_end):
                latest_end = activity_end
            for header, column in cost_columns.items():
                value = _numeric(ws.cell(row_number, column).value)
                if value is not None:
                    costs[header] += value
        baseline_costs: dict[str, float] = {}
        if "AutosysVersionBaseline" in workbook.sheetnames:
            baseline_ws = workbook["AutosysVersionBaseline"]
            for row_number in range(3, baseline_ws.max_row + 1):
                header = _normalized(baseline_ws.cell(row_number, 1).value)
                value = _numeric(baseline_ws.cell(row_number, 2).value)
                if header and value is not None:
                    baseline_costs[header] = value
        return GanttSnapshot(
            costs=costs,
            baseline_costs=baseline_costs,
            contractual_end=_contractual_end(workbook),
            latest_activity_end=latest_end,
        )
    finally:
        workbook.close()


def decide_version(working_content: bytes, baseline_v1_content: bytes | None) -> VersionDecision:
    working = snapshot_gantt(working_content)
    cost_increase = False
    reasons: list[str] = []

    baseline_costs: dict[str, float] = {}
    if baseline_v1_content is not None:
        baseline_costs = snapshot_gantt(baseline_v1_content).costs
    elif working.baseline_costs:
        baseline_costs = working.baseline_costs

    if baseline_costs:
        comparable = sorted(set(working.costs) & set(baseline_costs))
        if not comparable:
            raise ValueError(
                "No existen columnas de costo total comparables entre WORKING y v1.0."
            )
        increased = [
            header
            for header in comparable
            if working.costs[header] > baseline_costs[header] + COST_TOLERANCE
        ]
        cost_increase = bool(increased)
        if increased:
            reasons.append(
                "Aumento de costo detectado en: " + ", ".join(increased)
            )

    if working.schedule_overrun:
        reasons.append(
            "Una actividad termina después de la Fecha Final contractual."
        )

    major_change = cost_increase or working.schedule_overrun
    if not reasons:
        reasons.append(
            "Costos sin aumento y actividades dentro de la Fecha Final contractual."
        )
    return VersionDecision(
        version="v2.0" if major_change else "v1.0",
        cost_increase=cost_increase,
        schedule_overrun=working.schedule_overrun,
        reasons=tuple(reasons),
    )
