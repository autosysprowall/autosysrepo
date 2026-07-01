# Versionado automático del Gantt

El versionado usa la misma fila del proyecto en
`Control_Gantt_Asignaciones`. No necesita aprobación manual del supervisor.

## Disparador

1. El ingeniero termina el cronograma en el Gantt WORKING.
2. Cambia la celda existente `EstadoGantt` de la hoja `Datos` a
   `En revisión inicial`.
3. El dispatcher lee directamente el Excel en cada ejecución.
4. Si existe un evento `gantt_working_modificado`, también lo procesa como vía
   rápida, pero el evento no es obligatorio.
5. Actualiza la lista, activa internamente `SolicitarVersionado` y crea la
   versión correspondiente.

`SolicitarVersionado` continúa como flag técnico de reintento e idempotencia,
pero el supervisor ya no tiene que editarlo.

El cron interno de GitHub Actions está desactivado. El workflow se ejecuta
manualmente o mediante el dispatcher externo cuando sea activado.

## Decisión v1.0 o v2.0

Python descarga el WORKING y, si existe, la versión `v1.0`:

- si los costos totales comparables se mantienen o disminuyen, queda en
  `v1.0`;
- si aumenta al menos una columna comparable de costo total o precio total,
  queda en `v2.0`;
- si la fecha final de alguna actividad supera la Fecha Final contractual,
  queda en `v2.0`, aunque los costos no aumenten.

Cada Gantt nuevo contiene la hoja oculta `AutosysVersionBaseline`, creada con
los totales del presupuesto original. En el primer versionado se compara contra
esa línea base; posteriormente se compara contra `v1.0`.

Un Gantt antiguo que no tenga esta hoja y tampoco tenga una `v1.0` previa se
rechaza con una instrucción de regenerar el WORKING. No se asume una versión
sin comparación financiera verificable.

La comparación considera todas las columnas cuyo encabezado representa
`Costo Total` o `Precio Total`. `Costo Unitario`, `Utilidad` y `Margen` no se
suman como costo del proyecto.

## Archivos

El WORKING permanece en `Proyectos Activos/.../gantts/working/`. La versión se
guarda en `Proyectos Activos/.../gantts/versionados/` como:

- `{ProyectoID}_gantt_v1.0.xlsx`;
- `{ProyectoID}_gantt_v2.0.xlsx`.

El WORKING no se mueve ni se renombra. `Proyectos Terminados` está rechazado.

## Campos actualizados

- `SolicitarVersionado = No` después del éxito;
- `VersionActual = v1.0` o `v2.0`;
- `GanttVersionLink`;
- `GanttVersionIdentifier`;
- `FechaUltimoVersionado`;
- `FechaAprobacion`;
- `EstadoGantt = Aprobado / Versionado`;
- `MotivoUltimoVersionado`, si esa columna opcional existe;
- `UltimoErrorVersionado` vacío.

Ante un fallo, el proyecto no se aprueba, el flag técnico permanece activo,
`VersionadoIntentos` aumenta y la causa queda en `UltimoErrorVersionado`.

No existe todavía historial con `v1.1`, `v1.2`, `v2.1` u otros niveles.
