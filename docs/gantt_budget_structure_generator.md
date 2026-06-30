# Generador de Gantt basado en presupuesto

## Decisión de diseño

El Gantt WORKING nace ahora de la estructura comercial del presupuesto
aprobado. La planificación LLM anterior se conserva en
`scripts/sistema1_llm_planner.py` para reutilización futura, pero ya no filtra,
resume, renombra ni reordena las filas de la tabla principal.

El cambio corrige el problema del enfoque anterior: una interpretación útil
para planificación no necesariamente conserva las partidas, cantidades,
unidades y costos que el cliente necesita auditar.

## Flujo

1. Se abre una copia local del presupuesto; el original no se modifica.
2. Se selecciona `PRESUPUESTO FLEXIO` como fuente oficial cuando existe.
3. `scripts/budget_extraction.py` detecta encabezado, columnas y bloques.
4. Se extraen actividades y secciones en el mismo orden de la fuente.
5. Se vuelve a abrir el workbook con fórmulas y formato.
6. Se añade o reutiliza la hoja `Gantt`.
7. Se conservan las hojas originales y la hoja `Datos`.
8. Se crea `Assessment`.
9. Se guarda y se reabre el resultado para validarlo.
10. Solo después de validar, Sistema 1 puede subirlo a SharePoint.

La salida es:

```text
{ProyectoID}_gantt_WORKING.xlsx
```

en:

```text
{Proyecto}/gantts/working/
```

## Columnas del Gantt

Cuando existe un CC independiente:

1. `Actividad`
2. `CC`
3. `Fecha de Inicio`
4. `Fecha de Fin`
5. `Estatus`
6. `Unidad`
7. `Cantidad`
8. `Costo Unitario`
9. `Costo Total`

Si el presupuesto no tiene una columna CC independiente, `CC` se omite. No se
usa la misma columna de descripción como CC y no se inventan códigos.

`Fecha de Inicio`, `Fecha de Fin` y `Estatus` quedan vacíos por actividad. El
ingeniero completa las fechas y las reglas condicionales pintan la barra en el
calendario.

## Detección y normalización

Se escanean las primeras 50 filas y se puntúa cada candidata. Para aceptar una
fila se requiere:

- `Actividad` o equivalente;
- al menos dos columnas entre `Unidad`, `Cantidad`, `Costo Unitario` y
  `Costo Total`.

La normalización:

- convierte a minúsculas;
- elimina tildes;
- reemplaza saltos de línea por espacios;
- elimina puntuación y espacios repetidos;
- trata `P.U.`, `PU`, `Precio Unitario` y `Costo Unitario` como equivalentes.

`PRESUPUESTO FLEXIO` tiene prioridad absoluta. Solo se usan hojas
`PRESUPUESTO` u otras variantes como fallback para archivos históricos que no
contienen la hoja oficial de Flexio.

Aliases principales:

| Campo | Ejemplos |
|---|---|
| Actividad | actividad, descripción, concepto, detalle, tarea, partida, rubro, item |
| CC | cc, centro de costo, código, capítulo, familia, id partida |
| Unidad | unidad, und, unid, u/m, um, medida |
| Cantidad | cantidad, cant, qty, q, volumen |
| Costo Unitario | costo unitario, precio unitario, P.U., PU, costo/u |
| Costo Total | costo total, precio total, total, importe, subtotal, monto |

Si faltan columnas y el modo OpenAI está `live`, se puede solicitar un mapping
limitado al LLM. La respuesta solo puede referirse a índices de columnas
existentes. Un índice inválido, duplicado o sin encabezado se descarta. El LLM
no escribe el Excel.

También se detectan encabezados repetidos dentro de una misma hoja. Cada bloque
puede cambiar el orden de columnas sin desplazar Unidad, Cantidad o Costos.

## Estructura reutilizable

`extract_budget_structure()` devuelve `BudgetExtractionResult`, serializable con
`to_dict()`:

```json
{
  "source_file": "",
  "source_sheet": "",
  "header_row": 0,
  "column_mapping": {},
  "rows": [
    {
      "row_number": 0,
      "row_type": "activity",
      "level": 0,
      "actividad": "",
      "cc": "",
      "unidad": "",
      "cantidad": null,
      "costo_unitario": null,
      "costo_total": null,
      "source_cells": {},
      "raw_values": {},
      "warnings": []
    }
  ],
  "assessment": {}
}
```

Esta estructura no depende del writer de Excel y puede reutilizarse para flujo
de caja, análisis comercial, comparación de presupuestos y reportes.

## Tratamiento de filas

- Una fila con actividad y algún valor comercial se conserva como `activity`.
- Una fila con contexto pero sin unidad/cantidad/costos se conserva como
  `section`.
- Las secciones mantienen el orden, se combinan hasta el final del calendario y
  alinean el texto a la derecha.
- Las filas vacías se ignoran.
- Totales, subtotales, ITBMS, utilidad, administración, descuento y resúmenes
  `Costo/<unidad>` o `Precio/<unidad>` no se convierten en actividades.
- Las filas ambiguas se preservan como sección y se registran en `Assessment`.

## Preservación de costos

- Los valores existentes se copian sin sustituirlos.
- Si `Costo Total` falta y existen Cantidad y Costo Unitario, se calcula.
- Si el total fuente difiere más de 2% del cálculo, se conserva el total fuente
  y se registra una advertencia.
- Si una cantidad o costo es una fórmula del presupuesto, el Gantt crea un
  vínculo a la celda fuente dentro del mismo workbook, por ejemplo:

```excel
='PRESUPUESTO FLEXIO'!H13
```

Esto evita perder costos cuando `openpyxl` no encuentra un valor calculado en
la caché del archivo y conserva la relación funcional con el presupuesto.

## Calendario

El calendario usa `Fecha de Inicio` y `Fecha Final` de `Datos`. Cuando ambas
existen, `Fecha Final` tiene prioridad sobre una duración incompatible.

Si falta uno de los extremos, `Duracion` puede expresarse en días, semanas,
meses o años y se usa para calcular únicamente el límite visual del proyecto.
Nunca se inventan fechas por actividad.

## Formato y workbook base

El repositorio no contenía una plantilla binaria neutral. Las plantillas
históricas del prototipo tenían columnas antiguas, paneles congelados o datos de
un proyecto específico. Por esa razón no se incorporó un archivo cliente como
plantilla productiva.

Se conserva el contrato visual del Gantt diario existente:

- filas 9–11 para mes, día y día de semana;
- colores azul oscuro y azul claro;
- secciones combinadas hasta el calendario;
- validación de fechas y estatus;
- formato monetario;
- barras mediante formato condicional;
- anchos de columna razonables;
- sin paneles congelados.

Además, el output parte del workbook del presupuesto, por lo que mantiene sus
hojas, fórmulas, estilos, validaciones y metadatos. Solo `Gantt` y `Assessment`
son áreas controladas por el generador.

## Hoja Assessment

Documenta:

- archivo y hoja fuente;
- fila de encabezado;
- mapping y alias;
- columnas faltantes;
- filas preservadas, ignoradas y ambiguas;
- costos calculados;
- fórmulas vinculadas;
- uso de fallback LLM;
- validación final.

## Validación final

El archivo se reabre con `openpyxl` y se comprueba:

- existencia de `Gantt`;
- encabezados exactos, con la excepción documentada de CC;
- número de actividades;
- fechas por actividad vacías;
- costos presentes cuando la fuente tenía costos;
- tamaño mayor que cero;
- capacidad de volver a guardarse.

Un fallo lanza `GanttReviewRequiredError`. El procesador no sube el Gantt y
actualiza el evento a `RequiereRevision`.

## Casos de RequiereRevision

- no se encuentra Actividad y dos columnas comerciales;
- no se extrae ninguna actividad;
- las fechas del proyecto producen una ventana inválida o imposible para Excel;
- las columnas del Gantt no coinciden;
- desaparecen actividades o costos;
- el workbook no puede reabrirse o guardarse.

## Pruebas

`tests/test_budget_gantt_generator.py` cubre:

- columnas normales y reordenadas;
- aliases `Und`, `Cant`, `P.U.` y `Total`;
- filas antes del encabezado;
- filas vacías, totales y secciones;
- encabezados repetidos con otro orden dentro de la misma hoja;
- total calculado y discrepancia con total fuente;
- fallback LLM limitado a columnas existentes;
- ausencia de CC;
- duración en semanas y prioridad de Fecha Final;
- fórmulas de costos enlazadas;
- formato, colores, validaciones, barras y secciones combinadas;
- rechazo de un presupuesto sin encabezado confiable.

También se ejecutó una regresión local, sin modificar ni subir archivos, sobre
el presupuesto real `2026-066-DOOCOLEGIO ... CAMARAS ELECTRICAS`. Se seleccionó
`PRESUPUESTO FLEXIO`, se detectó el encabezado en la fila 12, se conservaron 19
filas comerciales y sus fórmulas, y la reapertura/validación terminó
correctamente.

No se crean versiones, no se envían correos y no se toca
`Proyectos Terminados`.
