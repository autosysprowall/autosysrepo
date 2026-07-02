# PA_S2_SincronizarEstadoGanttExcel

Este flujo sincroniza exclusivamente `Gantt!B6`. La lista
`Control_Gantt_Asignaciones` es la fuente de verdad.

## Despliegue

- Environment:
  `Default-ed7d4cfd-f42f-48ee-85ae-2e2be3539cd4`
- Flow ID: `f340a819-d200-4ecd-995b-57bc86467730`
- Estado: activo.
- Script de despliegue:
  `power_automate/scripts/deploy_status_sync_flow.py`

## Requisitos

- Conector estándar SharePoint.
- Conector estándar Excel Online (Business).
- Office Script `SetGanttStatus`, basado en
  `power_automate/office_scripts/SetGanttStatus.ts`.
- Conexiones ejecutadas por `auto.sys@prowallpanama.com`.

## Trigger

SharePoint: **When an item is created or modified**.

- Sitio: `https://sciprowall.sharepoint.com/sites/PROYECTOSPROWALL`
- Lista: `Control_Gantt_Asignaciones`
- Concurrency Control: `1`
- Trigger condition:

```text
@and(
  equals(triggerBody()?['EstadoSyncExcel'],'Pendiente'),
  not(empty(triggerBody()?['GanttWorkingIdentifier'])),
  or(
    equals(triggerBody()?['StatusExcelDeseado'],'Actual'),
    equals(triggerBody()?['StatusExcelDeseado'],'En Progreso')
  )
)
```

`Entregar` nunca es escrito automáticamente. Solo el usuario puede
seleccionarlo.

## Scope principal

1. **Update item**:
   - `EstadoSyncExcel = Procesando`
   - preservar todos los campos obligatorios.
2. **Run script** de Excel Online (Business):
   - archivo: `GanttWorkingIdentifier`;
   - script: `SetGanttStatus`;
   - `desiredStatus = StatusExcelDeseado`;
   - `expectedCurrentStatus = StatusExcel`;
   - `allowReplaceEntregar`:

```text
@and(
  equals(triggerBody()?['StatusExcelDeseado'],'Actual'),
  equals(triggerBody()?['EstadoGantt'],'Actual'),
  not(empty(triggerBody()?['VersionActual']))
)
```

Antes de cambiar el valor, el script reemplaza la validación de `B6` por
`Actual,En Progreso,Entregar`. Esto migra los WORKING existentes sin
reconstruir el archivo.

3. Evaluar `result.code`.

### `ACTUALIZADO` o `YA_SINCRONIZADO`

1. Esperar 30 segundos para que Excel/SharePoint publique la metadata.
2. SharePoint: **Get file metadata** usando `GanttWorkingIdentifier`.
3. **Update item**:
   - `StatusExcel = result.finalStatus`
   - `EstadoSyncExcel = Sincronizado`
   - `FechaUltimoSyncExcel = utcNow()`
   - `UltimoETagAutomatizacion = ETag`
   - `GanttWorkingETag = ETag`
   - `FechaUltimaModificacionGantt = Last modified`
   - `UltimoErrorSyncExcel` vacío
   - `ProximoIntentoSyncExcel` vacío

### `PROTEGIDO_ENTREGAR`, `REQUIERE_VERSION_EXITOSA` o `CONFLICTO_ESTADO`

No modificar la celda. Actualizar:

- `StatusExcel = result.finalStatus`
- `EstadoSyncExcel = Omitido`
- `FechaUltimoSyncExcel = utcNow()`
- `UltimoErrorSyncExcel = result.code`

El dispatcher leerá `Entregar` y procesará el versionado.

## Scope de fallo

Configurar **Run after** para `Failed`, `Timed out` y `Skipped` del Scope
principal.

Actualizar el mismo item:

- `EstadoSyncExcel = Error`
- `IntentosSyncExcel = add(coalesce(IntentosSyncExcel,0),1)`
- `UltimoErrorSyncExcel = result('Scope_Principal')`
- `ProximoIntentoSyncExcel = addMinutes(utcNow(),15)`

No modificar:

- `EstadoGantt`
- `StatusExcelDeseado`
- `SolicitarVersionado`
- vínculos o identificadores
- flags de correo

El dispatcher cambia nuevamente `Error` a `Pendiente` cuando llega
`ProximoIntentoSyncExcel`, hasta un máximo de cinco intentos.

## Protección contra bucles

La escritura de `B6` genera una nueva modificación del archivo. El flujo guarda
el nuevo eTag en `UltimoETagAutomatizacion` y `GanttWorkingETag`. El dispatcher
ignora ese eTag y no cambia `Actual` a `En Progreso`.

## Resultado esperado

| Estado lista | B6 observado | Acción |
|---|---|---|
| `Actual` | `Entregar` | Solo reemplazar por `Actual` después de versionar. |
| `Actual` | `Actual` | Sin cambios. |
| `En Progreso` | `Actual` | Cambiar B6 a `En Progreso`. |
| Cualquiera | `Entregar` | No sobrescribir; el dispatcher versiona. |
