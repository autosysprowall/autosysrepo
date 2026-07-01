# Versionado del Gantt

El versionado usa la misma fila del proyecto en
`Control_Gantt_Asignaciones`. No crea otra lista ni otro workflow.

## Disparador humano

Cuando el supervisor aprueba el Gantt WORKING:

1. establece `EstadoGantt = En revisión inicial`;
2. cambia `SolicitarVersionado = Sí`;
3. ejecuta manualmente `SharePoint Automation Dispatcher`, o espera al
   dispatcher externo de 15 minutos cuando sea activado.

El cron interno de GitHub Actions está desactivado. El workflow conserva
únicamente `workflow_dispatch`.

## Decisión v1.0 o v2.0

Python descarga el WORKING y, si existe, la versión `v1.0`:

- la primera aprobación queda en `v1.0`, salvo que ya exista una actividad
  posterior a la Fecha Final contractual;
- si los costos totales comparables se mantienen o disminuyen, queda en
  `v1.0`;
- si aumenta al menos una columna comparable de costo total o precio total,
  queda en `v2.0`;
- si la fecha final de alguna actividad supera la Fecha Final contractual,
  queda en `v2.0`, aunque los costos no aumenten.

Cada Gantt nuevo contiene la hoja oculta `AutosysVersionBaseline`, creada con
los totales del presupuesto original. En la primera aprobación se compara
contra esa línea base; en aprobaciones posteriores se compara contra `v1.0`.
La hoja no contiene credenciales ni lógica de correo.

Un Gantt antiguo que no tenga esta hoja y tampoco tenga una `v1.0` previa se
rechaza con una instrucción de regenerar el WORKING. No se asume `v1.0` sin una
comparación financiera verificable.

La comparación considera todas las columnas monetarias conservadas por el
generador cuyos encabezados representan `Costo Total` o `Precio Total`.
`Costo Unitario`, `Utilidad` y `Margen` no se suman como costo del proyecto.
Si existe una `v1.0` pero no hay ninguna columna comparable, el proceso falla
con un error accionable y no aprueba el proyecto.

## Archivos

El WORKING permanece en `Proyectos Activos/.../gantts/working/`. La versión se
guarda en `Proyectos Activos/.../gantts/versionados/` como:

- `{ProyectoID}_gantt_v1.0.xlsx`;
- `{ProyectoID}_gantt_v2.0.xlsx`.

El WORKING no se mueve ni se renombra. `Proyectos Terminados` está rechazado
explícitamente. Al aprobar nuevamente el mismo nivel, el archivo de ese nivel
se actualiza con el WORKING aprobado.

## Campos actualizados

- `SolicitarVersionado = No`;
- `VersionActual = v1.0` o `v2.0`;
- `GanttVersionLink`;
- `GanttVersionIdentifier`;
- `FechaUltimoVersionado`;
- `FechaAprobacion`;
- `EstadoGantt = Aprobado / Versionado`;
- `MotivoUltimoVersionado`;
- `UltimoErrorVersionado` vacío.

Ante un fallo no se aprueba el proyecto, la solicitud permanece activa,
`VersionadoIntentos` aumenta y la causa queda en `UltimoErrorVersionado`.

## Alcance pendiente

No existe todavía un historial con `v1.1`, `v1.2`, `v2.1` u otros niveles.
Tampoco se envía un correo de versionado.
