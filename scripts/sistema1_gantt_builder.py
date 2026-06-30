from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from hashlib import sha1
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from sistema1_llm_planner import request_llm_plan


HEADER_ROW = 10
WEEKDAY_ROW = 11
DATA_START_ROW = 12
START_DATE_COL = 5
END_DATE_COL = 6
DAILY_START_COL = 9
DEFAULT_WINDOW_DAYS = 90
CATEGORY_ORDER = ("PRELIMINARES", "FABRICA", "CAMPO", "ACABADOS")

TOTAL_TERMS = ("subtotal", "total", "gran total", "itbms", "impuesto", "utilidad", "margen", "precio total")
INDIRECT_TERMS = ("alquiler", "combustible", "baño", "banos", "grúa", "grua", "campamento", "administracion")
RESOURCE_TERMS = ("material", "materiales", "acero", "concreto", "clavo", "alambre", "plastico", "plástico")
WORK_TERMS = (
    "fabricacion",
    "fabricación",
    "instalacion",
    "instalación",
    "construccion",
    "construcción",
    "vaciado",
    "colocacion",
    "colocación",
    "montaje",
    "formaleta",
    "pintura",
    "pasteo",
    "suministro",
    "transporte",
)

MONTHS_ES = {
    1: "ENE",
    2: "FEB",
    3: "MAR",
    4: "ABR",
    5: "MAY",
    6: "JUN",
    7: "JUL",
    8: "AGO",
    9: "SEP",
    10: "OCT",
    11: "NOV",
    12: "DIC",
}
WEEKDAYS_ES = ["L", "M", "X", "J", "V", "S", "D"]


@dataclass(frozen=True)
class ProjectIdentity:
    project_id: str
    project_name: str
    folder_name: str
    gantt_file_name: str
    budget_copy_name: str


@dataclass(frozen=True)
class BudgetRow:
    source_row: int
    line_type: str
    item: Any
    code: Any
    description: str
    quantity: Any
    unit: Any
    amount: Any
    requires_review: bool
    review_reason: str


@dataclass(frozen=True)
class BuildResult:
    output_path: Path
    selected_sheet: str
    header_row: int
    rows_written: int
    review_rows: int
    window_start: date
    window_end: date
    notes: list[str]


@dataclass(frozen=True)
class LlmOptions:
    enabled: bool = False
    model: str = ""


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().lower()
    replacements = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ü": "u",
        "ñ": "n",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return re.sub(r"\s+", " ", text)


def display_text(value: Any, max_len: int = 200) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len]


def is_resident_engineer_row(*values: Any) -> bool:
    text = normalize_text(" ".join(display_text(value, 120) for value in values))
    return any(term in text for term in ("ingeniero residente", "residente obra", "ingeniero residente obra"))


def is_month_unit(value: Any) -> bool:
    text = normalize_text(value).strip(". ")
    return text in {"mes", "meses", "mensual"}


def safe_path_name(value: str, max_len: int = 95) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", value).strip(" ._")
    name = re.sub(r"\s+", " ", name)
    if not name:
        name = "Proyecto"
    return name[:max_len].strip(" ._")


def safe_file_stem(value: str, max_len: int = 120) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", value).strip(" ._")
    name = re.sub(r"\s+", " ", name)
    if not name:
        name = "presupuesto"
    return name[:max_len].strip(" ._")


def derive_project_identity(file_name: str) -> ProjectIdentity:
    stem = Path(file_name).stem
    match = re.match(r"(?P<id>\d{4}-\d{3}(?:-[A-Za-z0-9]+)?)\s*[-_ ]\s*(?P<name>.+)", stem)
    if match:
        project_id = match.group("id").upper()
        project_name = safe_path_name(match.group("name"), 80)
    else:
        digest = int(sha1(stem.encode("utf-8")).hexdigest()[:6], 16) % 1000
        project_id = f"PR-{datetime.now().year}-{digest:03d}"
        project_name = safe_path_name(stem, 80)

    folder_name = safe_path_name(f"{project_id}_{project_name}", 120)
    gantt_file_name = f"gannt_{safe_file_stem(stem)}_working.xlsx"
    budget_copy_name = f"{project_id}_presupuesto_aprobado{Path(file_name).suffix or '.xlsx'}"
    return ProjectIdentity(
        project_id=project_id,
        project_name=project_name,
        folder_name=folder_name,
        gantt_file_name=gantt_file_name,
        budget_copy_name=budget_copy_name,
    )


def parse_date_value(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = display_text(value, 40)
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_int_value(value: Any) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    match = re.search(r"\d+", display_text(value, 40).replace(",", ""))
    return int(match.group(0)) if match else None


def find_value_near_label(ws, label_terms: tuple[str, ...]) -> Any:
    for row in ws.iter_rows():
        for cell in row:
            label = normalize_text(cell.value)
            if not label or not all(term in label for term in label_terms):
                continue
            for offset in range(1, 5):
                value = ws.cell(cell.row, cell.column + offset).value
                if display_text(value):
                    return value
            value = ws.cell(cell.row + 1, cell.column).value
            if display_text(value):
                return value
    return None


def read_project_window(wb) -> tuple[date, date, list[str]]:
    notes: list[str] = []
    start = None
    end = None
    duration = None

    if "Datos" in wb.sheetnames:
        ws = wb["Datos"]
        start = parse_date_value(find_value_near_label(ws, ("fecha", "inicio")))
        end = parse_date_value(find_value_near_label(ws, ("fecha", "final")))
        duration = parse_int_value(find_value_near_label(ws, ("duracion",)))

    if start and end and duration:
        if end < start:
            start, end = end, start
            notes.append("Fechas de Datos venian invertidas; se ordenaron.")
        date_duration = max(1, (end - start).days + 1)
        duration_end = start + timedelta(days=duration - 1)
        if duration < date_duration:
            end = duration_end
            notes.append("Datos traia Fecha Final y Duracion no coincidentes; se uso la ventana mas corta.")
    elif start and duration and not end:
        end = start + timedelta(days=duration - 1)
        notes.append("Fecha final calculada desde Fecha de Inicio y Duracion.")
    elif end and duration and not start:
        start = end - timedelta(days=duration - 1)
        notes.append("Fecha de inicio calculada desde Fecha Final y Duracion.")
    elif start and end and end < start:
        start, end = end, start
        notes.append("Fechas de Datos venian invertidas; se ordenaron.")

    if not start:
        today = date.today()
        start = date(today.year, today.month, 1)
        notes.append("No se encontro fecha de inicio; se uso una ventana visual temporal.")
    if not end:
        end = start + timedelta(days=DEFAULT_WINDOW_DAYS - 1)
        notes.append(f"No se encontro fecha final; se uso una ventana visual de {DEFAULT_WINDOW_DAYS} dias.")
    return start, end, notes


def select_budget_sheet(wb) -> str:
    names = wb.sheetnames
    normalized = {normalize_text(name): name for name in names}
    if "presupuesto" in normalized:
        return normalized["presupuesto"]
    general = [name for name in names if "presupuesto" in normalize_text(name) and "general" in normalize_text(name)]
    if general:
        return general[0]
    pres = [name for name in names if "pres" in normalize_text(name)]
    if pres:
        return pres[0]
    for name in names:
        if normalize_text(name) not in {"datos", "destinatarios", "flujo de caja"}:
            return name
    return names[0]


def is_non_empty(value: Any) -> bool:
    return display_text(value) != ""


def detect_header_row(ws) -> int:
    keyword_groups = (
        ("item", "partida", "no"),
        ("cc", "codigo", "c.c"),
        ("actividad", "descripcion", "descripción", "detalle", "concepto", "material"),
        ("cantidad", "cant", "qty"),
        ("unidad", "und", "um"),
        ("total", "subtotal", "monto", "precio", "costo"),
    )
    best_row = 1
    best_score = -1.0
    for row_idx in range(1, min(ws.max_row, 60) + 1):
        values = [ws.cell(row_idx, col_idx).value for col_idx in range(1, ws.max_column + 1)]
        non_empty = [value for value in values if is_non_empty(value)]
        if len(non_empty) < 2:
            continue
        score = len(non_empty)
        for value in non_empty:
            text = normalize_text(value)
            if any(any(term in text for term in group) for group in keyword_groups):
                score += 4
            if isinstance(value, (int, float)):
                score -= 1
        if score > best_score:
            best_score = score
            best_row = row_idx
    return best_row


def find_column(headers: list[str], terms: tuple[str, ...]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        text = normalize_text(header)
        if any(term in text for term in terms):
            return idx
    return None


def find_item_column(headers: list[str]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        text = normalize_text(header).replace(".", "")
        if text in {"item", "partida", "nro", "numero", "no"}:
            return idx
    return find_column(headers, ("item", "partida"))


def find_unit_column(headers: list[str]) -> int | None:
    for idx, header in enumerate(headers, start=1):
        text = normalize_text(header)
        if any(term in text for term in ("costo", "precio", "unitario", "c unidad", "p unitario", "c. unidad")):
            continue
        if any(term in text for term in ("unidad", "und", "um", "u/m")):
            return idx
    return None


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def is_numeric_unit(value: Any) -> bool:
    if is_number(value):
        return True
    text = display_text(value, 40).replace(",", ".")
    return bool(text and re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text))


def clean_unit_value(value: Any) -> Any:
    return None if is_numeric_unit(value) else value


def item_depth(value: Any) -> int:
    text = display_text(value, 40).strip(". ")
    if not text:
        return 0
    return len([part for part in re.split(r"[.\s]+", text) if part])


def classify_row(description: str, item: Any, quantity: Any, unit: Any, amount: Any) -> tuple[str, bool, str]:
    text = normalize_text(description)
    row_text = normalize_text(" ".join(display_text(value, 80) for value in (description, item, quantity, unit, amount)))
    has_quantity = is_non_empty(quantity)
    has_unit = is_non_empty(unit)
    has_amount = is_non_empty(amount)

    if is_month_unit(unit) and text:
        return "ACTIVIDAD_PROBABLE", False, "Unidad mes: se conserva para cronograma de obra."
    if is_resident_engineer_row(description, item, quantity, unit, amount):
        return "NO_CRONOGRAMA_PROBABLE", True, "Ingeniero residente obra es control/administracion, no actividad cronogramable."
    if not text and has_amount:
        return "AMBIGUO", True, "Tiene monto pero no descripcion clara."
    if any(term in row_text for term in TOTAL_TERMS):
        return "NO_CRONOGRAMA_PROBABLE", True, "Subtotal, total, impuesto, margen, precio o resumen financiero."
    if any(term in row_text for term in INDIRECT_TERMS) and not any(term in row_text for term in ("instalacion", "montaje", "construccion")):
        return "NO_CRONOGRAMA_PROBABLE", True, "Indirecto, alquiler, equipo, campamento o administracion; validar antes de cronograma."
    if is_number(unit):
        return "AMBIGUO", True, "La columna Unidad contiene un numero; posible mapeo irregular del presupuesto."
    if item_depth(item) <= 2 and not has_quantity and not has_unit and text:
        return "AGRUPADOR", False, "Agrupador preservado para trazabilidad."
    if any(term in row_text for term in RESOURCE_TERMS) and not any(term in row_text for term in WORK_TERMS):
        return "NO_CRONOGRAMA_PROBABLE", True, "Parece insumo o recurso, no actividad ejecutable."
    if any(term in row_text for term in WORK_TERMS) or (has_quantity and has_unit and text):
        return "ACTIVIDAD_PROBABLE", False, "Actividad o paquete de trabajo probable."
    if text:
        return "AMBIGUO", True, "No hay evidencia suficiente para clasificar como actividad."
    return "AMBIGUO", True, "Fila sin descripcion util."


def value_at(ws, row_idx: int, col_idx: int | None) -> Any:
    if not col_idx:
        return None
    return ws.cell(row_idx, col_idx).value


def fallback_description(ws, row_idx: int, preferred_col: int | None, ignored_cols: set[int]) -> str:
    if preferred_col:
        text = display_text(ws.cell(row_idx, preferred_col).value, 220)
        if text:
            return text
    candidates: list[str] = []
    for col_idx in range(1, ws.max_column + 1):
        if col_idx in ignored_cols:
            continue
        value = ws.cell(row_idx, col_idx).value
        text = display_text(value, 220)
        if text and not is_number(value):
            candidates.append(text)
    return max(candidates, key=len) if candidates else ""


def resolve_budget_columns(headers: list[str]) -> dict[str, int | None]:
    return {
        "item": find_item_column(headers),
        "code": find_column(headers, ("cc", "c.c", "codigo", "centro")),
        "description": find_column(
            headers,
            ("actividad", "descripcion", "detalle", "concepto", "material", "partida"),
        ),
        "quantity": find_column(headers, ("cantidad", "cant", "qty", "volumen", "area")),
        "unit": find_unit_column(headers),
        "amount": find_column(headers, ("total", "subtotal", "monto", "importe", "valor")),
    }


def is_budget_header_row(headers: list[str], columns: dict[str, int | None]) -> bool:
    desc_col = columns["description"]
    unit_col = columns["unit"]
    qty_col = columns["quantity"]
    if not desc_col or not unit_col or not qty_col:
        return False
    desc_header = normalize_text(headers[desc_col - 1])
    unit_header = normalize_text(headers[unit_col - 1])
    qty_header = normalize_text(headers[qty_col - 1])
    return (
        any(term in desc_header for term in ("actividad", "descripcion", "detalle", "concepto", "partida"))
        and any(term in unit_header for term in ("unidad", "und", "u/m"))
        and any(term in qty_header for term in ("cantidad", "cant", "qty"))
    )


def extract_budget_rows(ws, header_row: int) -> list[BudgetRow]:
    headers = [display_text(ws.cell(header_row, col_idx).value, 80) for col_idx in range(1, ws.max_column + 1)]
    item_col = find_item_column(headers)
    code_col = find_column(headers, ("cc", "c.c", "codigo", "código", "centro"))
    desc_col = find_column(headers, ("actividad", "descripcion", "descripción", "detalle", "concepto", "material", "partida"))
    qty_col = find_column(headers, ("cantidad", "cant", "qty", "volumen", "area", "área"))
    unit_col = find_unit_column(headers)
    amount_col = find_column(headers, ("total", "subtotal", "monto", "importe", "valor"))
    ignored_cols = {col for col in (item_col, code_col, qty_col, unit_col, amount_col) if col}
    columns = resolve_budget_columns(headers)

    rows: list[BudgetRow] = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        values = [ws.cell(row_idx, col_idx).value for col_idx in range(1, ws.max_column + 1)]
        if not any(is_non_empty(value) for value in values):
            continue
        row_headers = [display_text(value, 80) for value in values]
        row_columns = resolve_budget_columns(row_headers)
        if is_budget_header_row(row_headers, row_columns):
            columns = row_columns
            continue
        item_col = columns["item"]
        code_col = columns["code"]
        desc_col = columns["description"]
        qty_col = columns["quantity"]
        unit_col = columns["unit"]
        amount_col = columns["amount"]
        ignored_cols = {col for col in (item_col, code_col, qty_col, unit_col, amount_col) if col}
        description = fallback_description(ws, row_idx, desc_col, ignored_cols)
        if not description:
            continue
        item = value_at(ws, row_idx, item_col)
        code = value_at(ws, row_idx, code_col)
        quantity = value_at(ws, row_idx, qty_col)
        raw_unit = value_at(ws, row_idx, unit_col)
        unit = clean_unit_value(raw_unit)
        amount = value_at(ws, row_idx, amount_col)
        line_type, requires_review, reason = classify_row(description, item, quantity, unit, amount)
        if is_numeric_unit(raw_unit):
            requires_review = True
            reason = (
                "El valor numerico detectado en la columna Unidad se omitio; "
                "revisar el mapeo de columnas del presupuesto."
            )
        rows.append(
            BudgetRow(
                source_row=row_idx,
                line_type=line_type,
                item=item,
                code=code,
                description=description,
                quantity=quantity,
                unit=unit,
                amount=amount,
                requires_review=requires_review,
                review_reason=reason,
            )
        )
    return rows


def row_to_llm_candidate(row: BudgetRow) -> dict[str, Any]:
    return {
        "candidate_id": f"R{row.source_row}",
        "source_row": row.source_row,
        "local_kind": row.line_type,
        "item": display_text(row.item, 80),
        "cc": display_text(row.code, 80),
        "description": row.description,
        "quantity": display_text(row.quantity, 80),
        "unit": display_text(row.unit, 80),
        "subtotal": display_text(row.amount, 80),
        "local_reason": row.review_reason,
    }


def apply_llm_plan(
    rows: list[BudgetRow],
    *,
    workbook_name: str,
    sheet_name: str,
    header_row: int,
    options: LlmOptions,
) -> tuple[list[BudgetRow], list[str]]:
    if not options.enabled:
        return rows, ["LLM actividad/cronograma: desactivado; se uso clasificacion local."]
    if not rows:
        return rows, ["LLM actividad/cronograma: sin filas candidatas."]

    candidates = [row_to_llm_candidate(row) for row in rows]
    try:
        plan = request_llm_plan(
            workbook_name=workbook_name,
            sheet_name=sheet_name,
            header_row=header_row,
            candidates=candidates,
            model=options.model or None,
        )
    except Exception as exc:
        return rows, [f"LLM actividad/cronograma fallo; se uso clasificacion local. Error: {exc}"]

    source_by_id = {f"R{row.source_row}": row for row in rows}
    bucket_order = {"PRELIMINARES": 0, "FABRICA": 1, "CAMPO": 2, "ACABADOS": 3, "NO_CRONOGRAMA": 4}
    planned: list[tuple[int, int, int, BudgetRow]] = []
    missing = 0
    for candidate_id, source in source_by_id.items():
        decision = plan.decisions.get(candidate_id)
        if not decision:
            missing += 1
            planned.append(
                (
                    4,
                    999999,
                    source.source_row,
                    BudgetRow(
                        source_row=source.source_row,
                        line_type="NO_CRONOGRAMA",
                        item=source.item,
                        code=source.code,
                        description=source.description,
                        quantity=source.quantity,
                        unit=source.unit,
                        amount=source.amount,
                        requires_review=True,
                        review_reason="LLM no devolvio decision para esta fila; requiere revision.",
                    ),
                )
            )
            continue

        include = decision.include_in_cronograma and decision.planning_bucket != "NO_CRONOGRAMA"
        line_type = decision.planning_bucket if include else "NO_CRONOGRAMA"
        description = decision.actividad_normalizada or source.description
        review = decision.confidence != "ALTA" or not include
        reason = f"LLM {decision.confidence}: {decision.reason or 'Sin razon.'}"
        planned.append(
            (
                bucket_order.get(decision.planning_bucket, 4),
                decision.orden_cronologico,
                source.source_row,
                BudgetRow(
                    source_row=source.source_row,
                    line_type=line_type,
                    item=source.item,
                    code=source.code,
                    description=description,
                    quantity=source.quantity,
                    unit=source.unit,
                    amount=source.amount,
                    requires_review=review,
                    review_reason=reason,
                ),
            )
        )

    notes = [
        "LLM actividad/cronograma: activo.",
        f"LLM filas evaluadas={len(rows)}; decisiones={len(plan.decisions)}; faltantes={missing}.",
        f"LLM uso tokens={plan.usage.get('total_tokens', 0)} en {plan.usage.get('batches', 0)} lote(s).",
    ]
    notes.extend(f"LLM advertencia: {warning}" for warning in plan.warnings[:5])
    return [item[3] for item in sorted(planned, key=lambda item: (item[0], item[1], item[2]))], notes


def is_chronogram_row(row: BudgetRow) -> bool:
    if is_month_unit(row.unit):
        return True
    if is_resident_engineer_row(row.description, row.item, row.code):
        return False
    if row.line_type in CATEGORY_ORDER:
        return True
    if row.line_type in {"NO_CRONOGRAMA", "NO_CRONOGRAMA_PROBABLE"}:
        return False
    if any(term in normalize_text(row.line_type) for term in ("total", "resource", "admin", "indirect")):
        return False
    return row.line_type in {"ACTIVIDAD_PROBABLE", "AGRUPADOR"}


def category_for_row(row: BudgetRow) -> str:
    if is_month_unit(row.unit):
        return "CAMPO"
    if row.line_type in CATEGORY_ORDER:
        return row.line_type
    text = normalize_text(" ".join(display_text(value, 120) for value in (row.description, row.item, row.code)))
    if any(term in text for term in ("pintura", "pasteo", "acabado", "limpieza final", "terminacion")):
        return "ACABADOS"
    if any(term in text for term in ("fabricacion", "produccion", "planta", "formaleta", "molde", "columna", "panel", "pared")):
        return "FABRICA"
    if any(term in text for term in ("instalacion", "montaje", "campo", "obra", "fundacion", "losa", "vaciado", "grua")):
        return "CAMPO"
    return "PRELIMINARES"


def filter_cronogram_rows(rows: list[BudgetRow]) -> tuple[list[BudgetRow], list[str]]:
    kept = [row for row in rows if is_chronogram_row(row)]
    removed = [row for row in rows if not is_chronogram_row(row)]
    notes = [f"Filas fuera de cronograma omitidas del Gantt: {len(removed)}."]
    if removed:
        sample = ", ".join(str(row.source_row) for row in removed[:12])
        notes.append(f"Primeras filas omitidas: {sample}.")
    return kept, notes


def style_base_sheet(ws) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = None
    ws.auto_filter.ref = None
    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 58
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 13
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 42


def write_daily_headers(ws, start: date, end: date) -> int:
    dark_fill = PatternFill("solid", fgColor="1F3864")
    month_fill = PatternFill("solid", fgColor="244B6B")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    headers = [
        "Fuente_Fila",
        "Actividad",
        "Cantidad",
        "Unidad",
        "Fecha_Inicio",
        "Fecha_Fin",
        "Estado_Planificacion",
        "Comentarios",
    ]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(HEADER_ROW, col_idx)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF" if col_idx >= START_DATE_COL else "000000")
        cell.fill = dark_fill if col_idx >= START_DATE_COL else header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    days = max(1, (end - start).days + 1)
    month_start_col = DAILY_START_COL
    current_month = start.month
    current_year = start.year
    for offset in range(days):
        current = start + timedelta(days=offset)
        col_idx = DAILY_START_COL + offset
        ws.column_dimensions[get_column_letter(col_idx)].width = 3.2
        day_cell = ws.cell(HEADER_ROW, col_idx)
        day_cell.value = current
        day_cell.number_format = "d"
        day_cell.font = Font(bold=True, size=8)
        day_cell.alignment = Alignment(horizontal="center", vertical="center")
        day_cell.border = Border(
            left=Side(style="thin", color="D9E2EA"),
            right=Side(style="thin", color="D9E2EA"),
            top=Side(style="thin", color="D9E2EA"),
            bottom=Side(style="thin", color="D9E2EA"),
        )
        weekday = ws.cell(WEEKDAY_ROW, col_idx)
        weekday.value = WEEKDAYS_ES[current.weekday()]
        weekday.font = Font(size=7, color="4B5563")
        weekday.alignment = Alignment(horizontal="center", vertical="center")

        month_changed = current.month != current_month or current.year != current_year
        is_last_day = offset == days - 1
        if month_changed or is_last_day:
            month_end_col = col_idx - 1 if month_changed else col_idx
            if month_end_col >= month_start_col:
                if month_end_col > month_start_col:
                    ws.merge_cells(
                        start_row=9,
                        start_column=month_start_col,
                        end_row=9,
                        end_column=month_end_col,
                    )
                month_cell = ws.cell(9, month_start_col)
                month_cell.value = f"{MONTHS_ES[current_month]} {current_year}"
                month_cell.fill = month_fill
                month_cell.font = Font(bold=True, color="FFFFFF", size=9)
                month_cell.alignment = Alignment(horizontal="center", vertical="center")
            if month_changed:
                month_start_col = col_idx
                current_month = current.month
                current_year = current.year

    return DAILY_START_COL + days - 1


def add_validations(ws, activity_rows: list[int]) -> None:
    if not activity_rows:
        return
    date_validation = DataValidation(
        type="date",
        operator="between",
        formula1="DATE(2020,1,1)",
        formula2="DATE(2035,12,31)",
        allow_blank=True,
    )
    date_validation.error = "Ingrese una fecha valida."
    date_validation.errorTitle = "Fecha invalida"
    ws.add_data_validation(date_validation)
    for row_idx in activity_rows:
        date_validation.add(f"E{row_idx}:F{row_idx}")

    status_validation = DataValidation(
        type="list",
        formula1='"Pendiente,Validado,Requiere Ajuste,Excluido"',
        allow_blank=True,
    )
    ws.add_data_validation(status_validation)
    for row_idx in activity_rows:
        status_validation.add(f"G{row_idx}")


def add_bar_formatting(ws, activity_rows: list[int], calendar_end_col: int) -> None:
    if not activity_rows:
        return
    first_col = get_column_letter(DAILY_START_COL)
    last_col = get_column_letter(calendar_end_col)
    bar_fill = PatternFill(fill_type="solid", start_color="2F80ED", end_color="2F80ED")
    invalid_fill = PatternFill(fill_type="solid", start_color="F8D7DA", end_color="F8D7DA")
    thin = Side(style="thin", color="E3E8EF")

    for row_idx in activity_rows:
        for col_idx in range(DAILY_START_COL, calendar_end_col + 1):
            cell = ws.cell(row_idx, col_idx)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row_range = f"{first_col}{row_idx}:{last_col}{row_idx}"
        formula = (
            f'=AND($E{row_idx}<>"",$F{row_idx}<>"",'
            f'{first_col}$10>=$E{row_idx},'
            f'{first_col}$10<=$F{row_idx})'
        )
        ws.conditional_formatting.add(row_range, FormulaRule(formula=[formula], fill=bar_fill))

        ws.conditional_formatting.add(
            f"E{row_idx}:F{row_idx}",
            FormulaRule(
                formula=[f'=AND($E{row_idx}<>"",$F{row_idx}<>"",$F{row_idx}<$E{row_idx})'],
                fill=invalid_fill,
            ),
        )


def write_budget_rows(ws, rows: list[BudgetRow], calendar_end_col: int) -> tuple[int, list[int]]:
    border = Border(
        left=Side(style="thin", color="B7C9D8"),
        right=Side(style="thin", color="B7C9D8"),
        top=Side(style="thin", color="B7C9D8"),
        bottom=Side(style="thin", color="B7C9D8"),
    )
    section_fill = PatternFill("solid", fgColor="D9EAF7")
    review_fill = PatternFill("solid", fgColor="FFF2CC")
    data_end = DATA_START_ROW - 1
    activity_rows: list[int] = []

    def write_section(category: str, section_row: int) -> None:
        for col_idx in range(1, calendar_end_col + 1):
            cell = ws.cell(section_row, col_idx)
            cell.fill = section_fill
            cell.font = Font(bold=True, color="1F3864")
            cell.border = border
        ws.merge_cells(
            start_row=section_row,
            start_column=1,
            end_row=section_row,
            end_column=calendar_end_col,
        )
        cell = ws.cell(section_row, 1)
        cell.value = category
        cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        ws.row_dimensions[section_row].height = 22

    grouped: dict[str, list[BudgetRow]] = {category: [] for category in CATEGORY_ORDER}
    for row in rows:
        grouped.setdefault(category_for_row(row), []).append(row)

    row_idx = DATA_START_ROW
    for category in CATEGORY_ORDER:
        category_rows = grouped.get(category, [])
        if not category_rows:
            continue

        write_section(category, row_idx)
        data_end = row_idx
        row_idx += 1

        for row in category_rows:
            activity_rows.append(row_idx)
            values = [
                row.source_row,
                row.description,
                row.quantity,
                row.unit,
                None,
                None,
                "Pendiente",
                row.review_reason if row.requires_review else "",
            ]
            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row_idx, col_idx)
                cell.value = value
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=col_idx in {2, 8})
                if row.requires_review:
                    cell.fill = review_fill
            ws.cell(row_idx, 5).number_format = "dd/mm/yyyy"
            ws.cell(row_idx, 6).number_format = "dd/mm/yyyy"
            data_end = row_idx
            row_idx += 1

    extra_categories = sorted(category for category in grouped if category not in CATEGORY_ORDER and grouped[category])
    for category in extra_categories:
        write_section(category, row_idx)
        data_end = row_idx
        row_idx += 1
        for row in grouped[category]:
            activity_rows.append(row_idx)
            values = [row.source_row, row.description, row.quantity, row.unit, None, None, "Pendiente", row.review_reason]
            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row_idx, col_idx)
                cell.value = value
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=col_idx in {2, 8})
            ws.cell(row_idx, 5).number_format = "dd/mm/yyyy"
            ws.cell(row_idx, 6).number_format = "dd/mm/yyyy"
            data_end = row_idx
            row_idx += 1

    if not rows:
        data_end = DATA_START_ROW
        ws.cell(DATA_START_ROW, 2).value = "No se detectaron actividades cronogramables en el presupuesto. Requiere revision manual."
        ws.cell(DATA_START_ROW, 8).value = "El archivo no contiene un bloque presupuestario claro o el LLM excluyo todo como no cronogramable."
    return data_end, activity_rows


def write_metadata_sheet(wb, identity: ProjectIdentity, source_file: str, selected_sheet: str, notes: list[str]) -> None:
    ws = wb.create_sheet("Datos")
    rows = [
        ("ProyectoID", identity.project_id),
        ("NombreProyecto", identity.project_name),
        ("ArchivoPresupuesto", source_file),
        ("HojaPresupuestoUsada", selected_sheet),
        ("Notas", " | ".join(notes)),
    ]
    for row_idx, (label, value) in enumerate(rows, start=1):
        ws.cell(row_idx, 1).value = label
        ws.cell(row_idx, 1).font = Font(bold=True)
        ws.cell(row_idx, 2).value = value
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 90
    ws.freeze_panes = None


def build_gantt_workbook(
    input_path: Path,
    output_path: Path,
    source_file_name: str,
    llm_options: LlmOptions | None = None,
) -> BuildResult:
    source_wb = load_workbook(input_path, data_only=True, read_only=False)
    try:
        selected_sheet = select_budget_sheet(source_wb)
        source_ws = source_wb[selected_sheet]
        header_row = detect_header_row(source_ws)
        rows = extract_budget_rows(source_ws, header_row)
        window_start, window_end, notes = read_project_window(source_wb)
    finally:
        source_wb.close()

    rows, llm_notes = apply_llm_plan(
        rows,
        workbook_name=source_file_name,
        sheet_name=selected_sheet,
        header_row=header_row,
        options=llm_options or LlmOptions(enabled=False),
    )
    notes.extend(llm_notes)
    rows, filter_notes = filter_cronogram_rows(rows)
    notes.extend(filter_notes)

    identity = derive_project_identity(source_file_name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Gantt_Diario"
    style_base_sheet(ws)

    ws["A1"] = "AUTOSYS - GANTT WORKING"
    ws["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws["A2"] = f"Proyecto: {identity.project_id} - {identity.project_name}"
    ws["A3"] = f"Presupuesto fuente: {source_file_name}"
    ws["A4"] = "Las fechas por actividad quedan vacias para que las complete el ingeniero residente."

    calendar_end_col = write_daily_headers(ws, window_start, window_end)
    data_end, activity_rows = write_budget_rows(ws, rows, calendar_end_col)
    add_validations(ws, activity_rows)
    add_bar_formatting(ws, activity_rows, calendar_end_col)
    write_metadata_sheet(wb, identity, source_file_name, selected_sheet, notes)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()

    return BuildResult(
        output_path=output_path,
        selected_sheet=selected_sheet,
        header_row=header_row,
        rows_written=len(rows),
        review_rows=sum(1 for row in rows if row.requires_review),
        window_start=window_start,
        window_end=window_end,
        notes=notes,
    )
