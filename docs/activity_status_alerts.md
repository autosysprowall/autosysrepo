# Alertas por estatus de actividades

Esta funcionalidad vive en la rama `feature/activity-status-alerts`.

El Gantt permite `Pendiente`, `En Progreso` y `Completada`. Una actividad queda
atrasada cuando:

- está `Pendiente` y la fecha actual supera `Fecha de Inicio`; o
- está `En Progreso` y la fecha actual supera `Fecha de Fin`.

`Completada` nunca se marca atrasada.

## Flujo

`PA_S3_MonitorearEstadosActividades` corre cada 15 minutos con conectores
estándar de SharePoint, Excel Online Business y Outlook. Ejecuta
`GetGanttActivityStatus`, actualiza `Control_Gantt_Asignaciones` y envía un
correo consolidado solamente cuando cambia el fingerprint de alertas.

El fingerprint de planificación excluye `Estatus`. Si solo cambia esta
columna, el flujo conserva el estado general en `Actual` y solicita que
`PA_S2_SincronizarEstadoGanttExcel` refleje `Actual` en `Gantt!B6`.

El flujo se despliega detenido primero. Al activarlo, la primera ejecución
establece la línea base de fingerprints. A partir de ahí, solamente una
combinación nueva de actividades atrasadas produce otro correo.

## Columnas de SharePoint

El flujo usa:

- `PlanificacionFingerprint`
- `EstadoActividadesFingerprint`
- `AlertasActividadesFingerprint`
- `ActividadesAtrasadas`
- `ResumenActividadesAtrasadas`
- `FechaLecturaActividades`
- `UltimaAlertaActividadesFingerprint`
- `FechaUltimoCorreoActividades`
- `UltimoErrorActividades`

`UltimaAlertaActividadesFingerprint` tiene como nombre interno de SharePoint
`UltimaAlertaActividadesFingerpri`.

## Versionado

Python calcula un fingerprint de planificación que excluye `Estatus`. Si un
WORKING se entrega y la única diferencia frente a la versión anterior son los
estados de actividades, no crea otra versión y devuelve el estado general a
`Actual`.
