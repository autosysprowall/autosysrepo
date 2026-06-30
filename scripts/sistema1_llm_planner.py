from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Any, Iterable


DEFAULT_MODEL = "gpt-4o-mini"
PLANNING_BUCKETS = ["PRELIMINARES", "FABRICA", "CAMPO", "ACABADOS", "NO_CRONOGRAMA"]
LINE_TYPES = ["SECTION", "ACTIVITY", "MILESTONE", "RESOURCE_ONLY", "TOTAL", "ADMIN_INDIRECT", "UNKNOWN"]

COLUMN_MAPPING_SYSTEM_PROMPT = """
Eres un analista de presupuestos. Tu única tarea es sugerir el índice de columnas
existentes para campos que las reglas determinísticas no pudieron reconocer.

Reglas estrictas:
- Usa solamente índices de columnas recibidos.
- No inventes columnas ni valores.
- Devuelve null cuando no exista evidencia suficiente.
- Una columna no puede representar dos campos.
- "item" contiene códigos jerárquicos como 1.1, 1.1.1 o 2.3.4; no es la actividad.
- "actividad" contiene descripciones textuales como Muro prefabricado, Concreto o Mano de obra.
- "unidad" contiene medidas como m2, m3, kg, ud o global.
- Usa las muestras de contenido para validar el significado del encabezado.
- No cambies ni resumas contenido del presupuesto.
- Devuelve solo JSON válido con el objeto "mapping".
""".strip()


SYSTEM_PROMPT = """
Eres un planificador senior de obra y analista de presupuestos de construccion prefabricada.

El archivo de entrada NO fue creado como cronograma. Fue creado por comercial/presupuesto para costear productos, partidas, materiales, mano de obra, indirectos y totales. Tu trabajo es traducir esa narrativa de presupuesto a una narrativa de cronograma para PMO/residencia, sin perder trazabilidad.

Objetivo:
- Leer principalmente la columna de actividades/descripciones y decidir que filas representan actividades cronogramables.
- Identificar objetos o elementos que seran fabricados en planta/fabrica y que luego necesitaran movilizacion o instalacion en campo.
- Separar, cuando la evidencia lo permita, el ciclo del objeto: diseno/preparacion previa, formaleta/molde, fabricacion en fabrica, suministro/movilizacion a campo, instalacion/ejecucion en campo y acabados.
- Distinguir entre una actividad ejecutable y un costo/recurso que solo explica como se cobra el presupuesto.
- Mantener source_row como la referencia obligatoria a la fila original.

Categorias de cronograma permitidas:
- PRELIMINARES: diseno, planos, aprobaciones, permisos, fianzas, submittals, levantamientos, programacion inicial, preparacion inicial, movilizacion inicial o actividades previas. En barriadas, los planos de casa normal/espejo y planos por tipo de pared deben agruparse idealmente antes del arranque general, no repetirse por cada vivienda salvo que el presupuesto lo separe claramente.
- FABRICA: fabricacion de objetos prefabricados, preparacion de formaletas/moldes, produccion en planta, armado, piezas, columnas, modulos, paneles, paredes, elementos o componentes fabricables. Formaleta es preparacion/molde antes de fabricar, no campo, salvo que el texto diga instalacion de formaleta en obra.
- CAMPO: suministro o llegada operativa a obra de elementos prefabricados, descarga, grua operativa, montaje, instalacion, obra civil, fundaciones, losa sobre tierra, vaciado, excavacion, compactacion, colocacion o ejecucion en sitio. Fundacion/losa/vaciado son senales de que campo empieza a trabajar.
- ACABADOS: pintura, pasteo, cielo raso, limpieza final, terminaciones y cierre estetico/funcional.

Lectura de presupuesto:
- Una fila puede estar en el presupuesto y aun asi quedar fuera del cronograma. No todo costo es actividad.
- Conserva encabezados utiles como SECTION cuando ayudan a leer la secuencia; no conviertas todos los titulos en actividades.
- Si una linea es subtotal, total, impuesto, utilidad, margen, descuento, precio total o resumen financiero, marcala TOTAL/NO_CRONOGRAMA.
- Si budget_structure es activity_based y una linea es insumo puro o material puro, marcala RESOURCE_ONLY/NO_CRONOGRAMA salvo que el texto describa una actividad ejecutable.
- Si budget_structure es material_process, el presupuesto expresa procesos mediante nombres de materiales. En ese modo, Concreto, Acero, Formaleta, Material de produccion, Mano de obra y Transporte son hitos/procesos cronogramables y deben conservarse como filas separadas.
- Si una linea es "ingeniero residente", "ingeniero residente obra", residente, supervision administrativa o personal de control sin accion ejecutable, dejala fuera como ADMIN_INDIRECT/NO_CRONOGRAMA.
- MT produccion o material de produccion normalmente es material/recurso de fabrica; en modo material_process conservalo como proceso de FABRICA.
- MO, Mano de Obra, MO acero, MO concreto, MO instalacion, MO acabados u otras manos de obra pueden ser actividades cronogramables si describen trabajo ejecutable. La unidad m2, m3, kg, ml, hr o similar en MO suele ser base de cobro de los obreros; NO excluyas una MO solo por tener unidad de medicion. Clasifica por la accion: MO acero/concreto de produccion o fabricacion suele ser FABRICA; MO instalacion, vaciado, fundacion o montaje suele ser CAMPO; MO pintura/pasteo/acabados suele ser ACABADOS.
- Si la unidad es mes, meses o mensual, conserva la fila en el cronograma. En presupuestos de obra suele representar permanencia, servicio o actividad sostenida en campo. Clasificala como CAMPO salvo que el texto demuestre claramente que es PRELIMINARES, FABRICA o ACABADOS.
- La unidad m2/m3/kg nunca decide sola. Es evidencia de metrado/base de cobro, no prueba automatica de actividad. Si la descripcion es alquiler, costo, precio, material, subtotal, equipo o indirecto, NO_CRONOGRAMA aunque tenga m2/m3/kg.
- Grua solo es CAMPO si representa una actividad operativa de izaje, montaje, descarga o movilizacion necesaria en obra. Alquiler de grua como costo/equipo sin accion debe ser ADMIN_INDIRECT o RESOURCE_ONLY con NO_CRONOGRAMA.
- Suministro de paredes/paneles/columnas puede significar disponibilidad o entrega a campo. Si dice fabricacion/produccion en planta, clasificalo FABRICA; si dice suministro/descarga/traslado/entrega/montaje en obra, clasificalo CAMPO.
- Plano pared 10cm, planos de casa normal y casa espejo, diseno y programacion son PRELIMINARES y deben ordenarse antes de FABRICA/CAMPO.

Reglas:
- El centro son las actividades cronogramables y los objetos fabricables, no los insumos ni los totales.
- No inventes fechas, duraciones, responsables, dependencias, cantidades ni montos.
- No inventes actividades que no esten trazadas en una fila fuente.
- No conviertas automaticamente todos los materiales en actividades cuando budget_structure sea activity_based. En material_process conserva los procesos materiales reconocidos sin agregar otros.
- En material_process no agrupes, resumas ni renombres Concreto, Acero, Formaleta, Material de produccion, Mano de obra o Transporte. actividad_normalizada debe conservar exactamente la descripcion recibida.
- Si una linea describe un objeto fabricable o prefabricado, clasificala como FABRICA y registra el objeto en la lista objetos_fabricados_planta.
- Si una linea habla de formaleta, formaletas, moldes o preparacion de moldes, usa FABRICA salvo que sea solo alquiler/costo sin actividad.
- Si una linea habla de instalacion, montaje, descarga, transporte a obra o ejecucion en sitio de un objeto fabricado, usa CAMPO y relaciona el objeto si se puede identificar.
- Si un mismo objeto aparece en fabrica y campo, usa el mismo nombre normalizado de objeto para conectar ambas filas.
- Si no hay evidencia suficiente para incluir una fila en el cronograma, excluyela con NO_CRONOGRAMA.
- Usa el orden macro PRELIMINARES -> FABRICA -> CAMPO -> ACABADOS salvo que el presupuesto muestre una secuencia mas especifica. Dentro de PRELIMINARES, planos/diseno/programacion van antes de todo; dentro de FABRICA, formaleta/molde va antes de fabricacion; dentro de CAMPO, fundacion/losa/vaciado y suministro/montaje ordenan el arranque de obra.
- El orden_cronologico debe ser un entero incremental dentro de la logica de proyecto.

Devuelve solo JSON valido.
""".strip()


USER_PROMPT = """
Clasifica y ordena estas lineas de presupuesto para generar un cronograma Gantt desde la perspectiva de planificacion de obra.

Devuelve este JSON:
{
  "objetos_fabricados_planta": [
    {
      "nombre_objeto": "Paneles prefabricados",
      "source_rows": [15, 18],
      "requiere_movilizacion_campo": true,
      "evidencia": "Filas de fabricacion/instalacion trazadas al mismo objeto."
    }
  ],
  "decisions": [
    {
      "candidate_id": "R15",
      "source_row": 15,
      "include_in_cronograma": true,
      "line_type": "ACTIVITY",
      "planning_bucket": "FABRICA",
      "orden_cronologico": 120,
      "actividad_normalizada": "Fabricacion de paneles prefabricados",
      "objeto_fabricado_planta": "Paneles prefabricados",
      "requiere_movilizacion_campo": true,
      "confidence": "ALTA",
      "reason": "Objeto fabricable trazado desde la fila fuente."
    }
  ],
  "warnings": []
}

Campos permitidos:
- line_type: SECTION, ACTIVITY, MILESTONE, RESOURCE_ONLY, TOTAL, ADMIN_INDIRECT, UNKNOWN.
- planning_bucket: PRELIMINARES, FABRICA, CAMPO, ACABADOS, NO_CRONOGRAMA.
- confidence: ALTA, MEDIA, BAJA.

Reglas:
- Debe haber una decision por cada candidate_id recibido.
- No cambies source_row.
- Si include_in_cronograma es false, planning_bucket debe ser NO_CRONOGRAMA.
- Si line_type es TOTAL, RESOURCE_ONLY, ADMIN_INDIRECT o UNKNOWN por falta de evidencia, include_in_cronograma debe ser false salvo que reason explique una excepcion operativa clara.
- Si una fila representa fabricacion en planta, objeto_fabricado_planta debe tener el nombre del objeto.
- Si una fila representa instalacion/montaje/transporte del objeto fabricado, objeto_fabricado_planta debe repetir el mismo nombre del objeto cuando sea identificable.
- Si no hay objeto fabricado asociado, objeto_fabricado_planta debe ser "" y requiere_movilizacion_campo debe ser false.
- objetos_fabricados_planta debe ser una lista consolidada de objetos fabricados en planta detectados en este lote.
- No uses m2/m3/kg/ml/hr como criterio unico para incluir o excluir. En MO puede ser base de pago; en materiales/costos puede ser solo metrado.
- Si la unidad es mes, meses o mensual, include_in_cronograma debe ser true. Usa CAMPO por defecto porque normalmente corresponde a obra, permanencia o servicio mensual.
- Si detectas MO ejecutable, incluyela y clasificala por accion: fabrica para produccion/acero/concreto en planta; campo para instalacion/fundacion/vaciado/montaje; acabados para pintura/pasteo/cielo raso.
- Si detectas MT produccion/material de produccion sin accion ejecutable, dejalo fuera como RESOURCE_ONLY/NO_CRONOGRAMA en activity_based; incluyelo como FABRICA en material_process.
- Si budget_structure es material_process, incluye por separado Concreto, Acero, Formaleta, Material de produccion y Mano de obra en FABRICA, y Transporte en CAMPO. No los combines ni simplifiques.
- Si detectas alquiler de grua/equipo sin accion de izaje/montaje/descarga, dejalo fuera como ADMIN_INDIRECT o RESOURCE_ONLY.
- Si detectas ingeniero residente obra, supervision administrativa o personal indirecto sin actividad ejecutable, dejalo fuera del cronograma.
- Si detectas planos/programacion de barriada o tipos de casa normal/espejo, ordenalos temprano en PRELIMINARES.
- No agregues campos fuera del JSON.
""".strip()


@dataclass(frozen=True)
class LlmDecision:
    candidate_id: str
    source_row: int
    include_in_cronograma: bool
    line_type: str
    planning_bucket: str
    orden_cronologico: int
    actividad_normalizada: str
    confidence: str
    reason: str


@dataclass(frozen=True)
class LlmPlanResult:
    decisions: dict[str, LlmDecision]
    warnings: list[str]
    usage: dict[str, Any]


def normalize_bucket(value: Any) -> str:
    text = str(value or "").strip().upper().replace(" ", "_")
    aliases = {
        "PRELIMINAR": "PRELIMINARES",
        "PRELIMINARES": "PRELIMINARES",
        "PLANTA": "FABRICA",
        "FABRICA": "FABRICA",
        "FABRICACION": "FABRICA",
        "CAMPO": "CAMPO",
        "ACABADO": "ACABADOS",
        "ACABADOS": "ACABADOS",
        "NO_CRONOGRAMA": "NO_CRONOGRAMA",
    }
    return aliases.get(text, "NO_CRONOGRAMA")


def normalize_line_type(value: Any) -> str:
    text = str(value or "").strip().upper()
    return text if text in LINE_TYPES else "UNKNOWN"


def normalize_confidence(value: Any) -> str:
    text = str(value or "").strip().upper()
    return text if text in {"ALTA", "MEDIA", "BAJA"} else "BAJA"


def chunked(items: list[dict[str, Any]], size: int) -> Iterable[list[dict[str, Any]]]:
    for index in range(0, len(items), size):
        yield items[index : index + size]


def parse_plan(raw: dict[str, Any], candidate_ids: set[str]) -> tuple[dict[str, LlmDecision], list[str]]:
    warnings: list[str] = []
    decisions: dict[str, LlmDecision] = {}
    raw_decisions = raw.get("decisions") if isinstance(raw, dict) else None
    if not isinstance(raw_decisions, list):
        raw_decisions = []
        warnings.append("El LLM no devolvio una lista decisions valida.")

    for item in raw_decisions:
        if not isinstance(item, dict):
            continue
        candidate_id = str(item.get("candidate_id") or "").strip()
        if candidate_id not in candidate_ids:
            continue
        include = bool(item.get("include_in_cronograma"))
        bucket = normalize_bucket(item.get("planning_bucket"))
        if not include:
            bucket = "NO_CRONOGRAMA"
        try:
            order = int(item.get("orden_cronologico") or 999999)
        except (TypeError, ValueError):
            order = 999999
        source_row = item.get("source_row")
        try:
            source_row_int = int(source_row)
        except (TypeError, ValueError):
            source_row_int = 0
        decisions[candidate_id] = LlmDecision(
            candidate_id=candidate_id,
            source_row=source_row_int,
            include_in_cronograma=include,
            line_type=normalize_line_type(item.get("line_type")),
            planning_bucket=bucket,
            orden_cronologico=order,
            actividad_normalizada=str(item.get("actividad_normalizada") or "").strip(),
            confidence=normalize_confidence(item.get("confidence")),
            reason=str(item.get("reason") or "").strip(),
        )

    missing = sorted(candidate_ids.difference(decisions))
    if missing:
        warnings.append(f"Faltaron decisiones LLM para {len(missing)} candidatos.")
    raw_warnings = raw.get("warnings") if isinstance(raw, dict) else None
    if isinstance(raw_warnings, list):
        warnings.extend(str(item) for item in raw_warnings if item)
    return decisions, warnings


def request_llm_plan(
    *,
    workbook_name: str,
    sheet_name: str,
    header_row: int,
    candidates: list[dict[str, Any]],
    model: str | None = None,
    chunk_size: int = 18,
    budget_structure: str = "activity_based",
) -> LlmPlanResult:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY no esta configurada.")
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError("Falta instalar openai en el runner.") from exc

    client = OpenAI(api_key=api_key)
    all_decisions: dict[str, LlmDecision] = {}
    all_warnings: list[str] = []
    usages: list[dict[str, Any]] = []
    batches = list(chunked(candidates, max(1, chunk_size)))
    for batch_index, batch in enumerate(batches, start=1):
        payload = {
            "workbook_name": workbook_name,
            "sheet_name": sheet_name,
            "header_row": header_row,
            "planning_buckets": PLANNING_BUCKETS,
            "line_types": LINE_TYPES,
            "budget_structure": budget_structure,
            "batch_index": batch_index,
            "batch_count": len(batches),
            "candidates": batch,
        }
        response = client.chat.completions.create(
            model=model or DEFAULT_MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": USER_PROMPT + "\n\n" + json.dumps(payload, ensure_ascii=False)},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        text = response.choices[0].message.content or "{}"
        raw = json.loads(text)
        candidate_ids = {str(item["candidate_id"]) for item in batch}
        decisions, warnings = parse_plan(raw, candidate_ids)
        all_decisions.update(decisions)
        all_warnings.extend(f"Lote {batch_index}: {warning}" for warning in warnings)
        usage = getattr(response, "usage", None)
        if usage:
            usages.append(usage.model_dump() if hasattr(usage, "model_dump") else dict(usage))

    total_usage = {
        "prompt_tokens": sum(int(item.get("prompt_tokens") or 0) for item in usages),
        "completion_tokens": sum(int(item.get("completion_tokens") or 0) for item in usages),
        "total_tokens": sum(int(item.get("total_tokens") or 0) for item in usages),
        "batches": len(usages),
    }
    return LlmPlanResult(decisions=all_decisions, warnings=all_warnings, usage=total_usage)


def request_llm_column_mapping(
    *,
    workbook_name: str,
    sheet_name: str,
    headers: list[Any],
    missing_fields: tuple[str, ...],
    samples_by_column: dict[int, list[Any]] | None = None,
    model: str | None = None,
) -> dict[str, int | None]:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY no esta configurada.")
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError("Falta instalar openai en el runner.") from exc

    indexed_headers = [
        {"column_index": index, "header": str(value or "").strip()}
        for index, value in enumerate(headers, start=1)
        if str(value or "").strip()
    ]
    payload = {
        "workbook_name": workbook_name,
        "sheet_name": sheet_name,
        "missing_fields": list(missing_fields),
        "allowed_fields": [
            "item",
            "actividad",
            "cc",
            "unidad",
            "cantidad",
            "costo_unitario",
            "costo_total",
        ],
        "headers": indexed_headers,
        "content_samples": {
            str(column): [
                str(value or "").strip()
                for value in values[:8]
                if str(value or "").strip()
            ]
            for column, values in (samples_by_column or {}).items()
        },
        "response_example": {
            "mapping": {
                "item": 3,
                "actividad": 4,
                "cc": None,
                "unidad": 5,
            }
        },
    }
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model or DEFAULT_MODEL,
        messages=[
            {"role": "system", "content": COLUMN_MAPPING_SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        response_format={"type": "json_object"},
        temperature=0,
    )
    raw = json.loads(response.choices[0].message.content or "{}")
    proposed = raw.get("mapping") if isinstance(raw, dict) else {}
    if not isinstance(proposed, dict):
        return {}
    result: dict[str, int | None] = {}
    for field_name in missing_fields:
        value = proposed.get(field_name)
        try:
            result[field_name] = int(value) if value is not None else None
        except (TypeError, ValueError):
            result[field_name] = None
    return result
