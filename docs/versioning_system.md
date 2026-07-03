# Versionado automático del Gantt

El versionado usa la misma fila del proyecto en
`Control_Gantt_Asignaciones`. No necesita aprobación manual del supervisor.

## Disparador

1. El ingeniero termina el cronograma en el Gantt WORKING.
2. Cambia el selector visible `Gantt!B6`, rotulado
   `Estado general del Gantt`, a `Entregar`.
3. El dispatcher lee directamente el Excel en cada ejecución.
4. Si existe un evento `gantt_working_modificado`, también lo procesa como vía
   rápida, pero el evento no es obligatorio.
5. Actualiza la lista, activa internamente `SolicitarVersionado` y crea la
   versión correspondiente.

`SolicitarVersionado` continúa como flag técnico de reintento e idempotencia,
pero el supervisor ya no tiene que editarlo.

El cron interno de GitHub Actions está desactivado. El workflow se ejecuta
manualmente o mediante el dispatcher externo cuando sea activado.

## Versiones acumulativas

Python enumera todas las versiones existentes del proyecto y ordena
numéricamente sus nombres. La primera decisión parte de la línea base del
presupuesto; las siguientes parten de la última versión aprobada.

- Sin salto mayor: conserva el entero e incrementa el decimal:
  `v1.0 → v1.1 → v1.2` o `v2.0 → v2.1`.
- Con salto mayor: incrementa el entero y reinicia el decimal:
  `v1.2 → v2.0` o `v2.3 → v3.0`.
- Una versión existente nunca se sobrescribe.

El salto mayor usa una condición OR:

1. el nuevo `Costo Total` supera el de la versión inmediatamente anterior; o
2. la fecha final más tardía del WORKING supera la fecha más tardía ya
   aprobada.

Si cualquiera se cumple, avanza el entero. Si ninguna se cumple, avanza solo
el decimal.

Cada Gantt nuevo contiene la hoja oculta `AutosysVersionBaseline`, creada con
el total del presupuesto original. En el primer versionado el costo se compara
contra esa línea base; posteriormente se compara contra la versión
inmediatamente anterior.

La `Fecha Final contractual` también se conserva en esa hoja oculta y es la
fuente inicial de cronología. Para revisiones posteriores, el límite es el
máximo entre esa fecha contractual y todas las fechas finales ya aprobadas.
Por eso un atraso aceptado no genera otro salto mayor hasta que una nueva
entrega supere esa fecha más atrasada.

Un Gantt antiguo que no tenga esta hoja y tampoco tenga una versión previa se
rechaza con una instrucción de regenerar el WORKING. No se asume una versión
sin comparación financiera verificable.

La comparación suma únicamente la columna canónica `Costo Total`.
`Costo Unitario`, `Precio Total`, `Precio Unitario`, `Utilidad` y `Margen` no
se suman como costo del proyecto.

## Archivos

El WORKING permanece en `Proyectos Activos/.../gantts/working/`. La versión se
guarda en `Proyectos Activos/.../gantts/versionados/` como:

- `{ProyectoID}_gantt_v1.0.xlsx`;
- `{ProyectoID}_gantt_v1.1.xlsx`;
- `{ProyectoID}_gantt_v2.0.xlsx`;
- y así sucesivamente.

El WORKING no se mueve ni se renombra. `Proyectos Terminados` está rechazado.

## Campos actualizados

- `SolicitarVersionado = No` después del éxito;
- `VersionActual =` la versión acumulativa creada;
- `GanttVersionLink`;
- `GanttVersionIdentifier`;
- `FechaUltimoVersionado`;
- `FechaAprobacion`;
- `EstadoGantt = Actual`;
- `StatusExcelDeseado = Actual`;
- `EstadoSyncExcel = Pendiente`, para que Power Automate actualice `B6`;
- cierre y acumulación del tiempo que permaneció `En Progreso`;
- `MotivoUltimoVersionado`, si esa columna opcional existe;
- `UltimoErrorVersionado` vacío.

Ante un fallo, el proyecto no se aprueba, el flag técnico permanece activo,
`VersionadoIntentos` aumenta y la causa queda en `UltimoErrorVersionado`.

Para reevaluar explícitamente un control ya versionado después de corregir una
regla, la ejecución manual admite `force_version_recheck=true`. Esta opción
requiere `control_item_id` y que `Gantt!B6` permanezca en
`Entregar`; nunca se usa en ejecuciones automáticas.

## Estados visibles

- `En Progreso`: el Gantt está siendo trabajado y su contador permanece
  activo.
- `Entregar`: solicita validación y versionado. Si el proceso falla, permanece
  en este estado para reintentar.
- `Actual`: el WORKING coincide con la versión vigente.

La lista de SharePoint es la fuente de verdad. Power Automate sincroniza
`Gantt!B6` mediante el Office Script `SetGanttStatus`; un bloqueo temporal del
Excel no invalida el versionado.

La columna `Estatus` de las actividades tampoco forma parte del fingerprint de
planificación. Cambiar una actividad entre `Pendiente`, `En Progreso` y
`Completada` no crea una versión. El cambio solo alimenta el monitor de atrasos
de Power Automate y el estado general vuelve a `Actual` cuando la planificación
no cambió.

Cada archivo dentro de `gantts/versionados` forma el historial inmutable del
proyecto.
