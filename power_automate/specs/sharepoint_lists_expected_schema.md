# Esquema esperado de listas SharePoint

Las listas ya existen manualmente. Este documento define los nombres y valores que los flujos y Python esperan.

## Cola_Automatizacion_Proyectos

Funcion: cola tecnica de eventos que Python/GitHub Actions procesa cada hora.

| Columna | Tipo recomendado | Requerida | Uso |
|---|---|---:|---|
| `Title` | Single line of text | Si | Nombre visible del evento. |
| `EventType` | Choice o Single line of text | Si | Tipo de evento. |
| `Estado` | Choice | Si | Estado de procesamiento. |
| `FileName` | Single line of text | Si | Nombre del archivo con extension. |
| `FileIdentifier` | Single line of text | Si | Identificador SharePoint del archivo. |
| `FolderPath` | Single line of text | Si | Ruta de carpeta fuente. |
| `FileLink` | Hyperlink o Single line of text | Si | Link al archivo. |
| `CreatedByEmail` | Single line of text | No | Correo de quien subio el archivo. |
| `CreatedTime` | Date and time | No | Fecha de creacion del archivo. |
| `Intentos` | Number | Si | Numero de intentos Python. |
| `UltimoError` | Multiple lines of text | No | Error mas reciente. |
| `FechaProcesado` | Date and time | No | Fecha en que Python termino. |
| `ProyectoID` | Single line of text | No | ID detectado/generado por Python. |
| `Notas` | Multiple lines of text | No | Observaciones tecnicas. |

### Valores permitidos de EventType

- `presupuesto_aprobado`
- `gantt_working_modificado`

### Valores permitidos de Estado

- `Pendiente`
- `Procesando`
- `Procesado`
- `Error`
- `RequiereRevision`

## Control_Gantt_Asignaciones

Funcion: seguimiento humano del Gantt WORKING generado por Python.

| Columna | Tipo recomendado | Requerida | Uso |
|---|---|---:|---|
| `Title` | Single line of text | Si | Nombre visible del registro. |
| `ProyectoID` | Single line of text | Si | ID del proyecto. |
| `NombreProyecto` | Single line of text | Si | Nombre legible del proyecto. |
| `PresupuestoLink` | Hyperlink o Single line of text | Si | Link al presupuesto aprobado. |
| `GanttWorkingLink` | Hyperlink o Single line of text | Si | Link al Gantt de trabajo. |
| `IngenieroAsignado` | Person o Single line of text | No | Nombre del ingeniero. |
| `IngenieroEmail` | Single line of text | No | Correo operativo para notificaciones. |
| `EstadoGantt` | Choice | Si | Estado humano del Gantt. |
| `FechaGanttGenerado` | Date and time | No | Fecha de generacion por Python. |
| `FechaAsignacion` | Date and time | No | Fecha de asignacion al ingeniero. |
| `FechaLimite` | Date and time | No | Fecha limite del trabajo inicial. |
| `FechaInicioReal` | Date and time | No | Fecha real de inicio del proyecto/actividad si aplica. |
| `FechaEnvioRevision` | Date and time | No | Fecha en que se envio a revision inicial. |
| `FechaAprobacion` | Date and time | No | Fecha de aprobacion. |
| `DiasParaCompletar` | Number | No | Dias entre asignacion y envio a revision. |
| `Advertencia1Enviada` | Yes/No | Si | Control de advertencia dia 3. |
| `Advertencia2Enviada` | Yes/No | Si | Control de advertencia dia 6. |
| `Notas` | Multiple lines of text | No | Observaciones humanas/tecnicas. |
| `GanttWorkingIdentifier` | Single line of text | No | Identificador estable del Gantt. |
| `PresupuestoIdentifier` | Single line of text | No | Identificador estable del presupuesto. |
| `SupervisoresEmail` | Multiple lines of text | No | Correos separados por punto y coma. |
| `PermisoGanttOtorgado` | Yes/No | Si | Evita repetir el permiso. |
| `FechaPermisoOtorgado` | Date and time | No | Fecha del permiso de edicion. |
| `CorreoAsignacionEnviado` | Yes/No | Si | Confirmacion de Power Automate. |
| `FechaCorreoAsignacion` | Date and time | No | Fecha confirmada del correo. |
| `FechaAdvertencia1` | Date and time | No | Fecha confirmada del primer aviso. |
| `FechaAdvertencia2` | Date and time | No | Fecha confirmada del segundo aviso. |
| `VencimientoNotificado` | Yes/No | Si | Evita repetir escalamiento. |
| `FechaVencimientoNotificado` | Date and time | No | Fecha confirmada del escalamiento. |
| `UltimoTrackingRun` | Date and time | No | Ultima revision del dispatcher. |
| `TrackingIntentos` | Number | Si | Intentos fallidos de asignacion/tracking. |
| `UltimoErrorTracking` | Multiple lines of text | No | Ultimo error accionable. |
| `StatusExcel` | Single line of text | No | Ultimo status leido de Datos. |
| `FechaLecturaStatusExcel` | Date and time | No | Momento de lectura del status. |
| `StatusExcelDeseado` | Single line of text | No | Estado que Power Automate debe reflejar en `Gantt!B6`. |
| `EstadoSyncExcel` | Single line of text | No | `Sincronizado`, `Pendiente`, `Procesando`, `Omitido` o `Error`. |
| `IntentosSyncExcel` | Number | No | Reintentos de escritura de `B6`. |
| `ProximoIntentoSyncExcel` | Date and time | No | Próximo reintento permitido. |
| `UltimoErrorSyncExcel` | Multiple lines of text | No | Error del conector Excel/Office Script. |
| `FechaUltimoSyncExcel` | Date and time | No | Última sincronización exitosa. |
| `FechaInicioEnProgreso` | Date and time | No | Inicio del ciclo actual de trabajo. |
| `MinutosEnProgresoActual` | Number | No | Duración del ciclo abierto. |
| `MinutosEnProgresoAcumulados` | Number | No | Duración acumulada de ciclos cerrados. |
| `FechaEntregaSolicitada` | Date and time | No | Momento en que se leyó `Entregar`. |
| `GanttWorkingETag` | Single line of text | No | Último contenido conocido por el dispatcher. |
| `UltimoETagAutomatizacion` | Single line of text | No | eTag producido por Power Automate; evita bucles. |
| `FechaUltimaModificacionGantt` | Date and time | No | Última modificación observada en SharePoint. |

### Valores permitidos de EstadoGantt

- `Pendiente de asignación`
- `En Progreso`
- `Entregar`
- `Actual`
- `Vencido`
- `Requiere revisión manual`

Durante la transición se siguen leyendo `Asignado`, `En progreso`,
`En revisión inicial` y `Aprobado / Versionado` para compatibilidad, pero no se
escriben en registros nuevos.

## Cola_Notificaciones_Gantt

Funcion: cola idempotente para que Power Automate envie correos con Outlook.

| Columna | Tipo recomendado | Uso |
|---|---|---|
| `Title` | Single line of text | Nombre visible. |
| `ProyectoID` | Single line of text | Proyecto relacionado. |
| `TipoNotificacion` | Choice | Asignacion, avisos o vencimiento. |
| `EstadoNotificacion` | Choice | `Pendiente`, `Enviado` o `Error`. |
| `To` / `Cc` | Multiple lines of text | Destinatarios. |
| `Subject` | Single line of text | Asunto. |
| `Body` | Multiple lines of text | Cuerpo. |
| `RelatedControlItemID` | Single line of text | ID del control relacionado. |
| `Intentos` | Number | Intentos del flujo. |
| `UltimoError` | Multiple lines of text | Error de Outlook/SharePoint. |
| `FechaCreacion` / `FechaEnvio` | Date and time | Auditoria. |
| `Notas` | Multiple lines of text | Observaciones. |

## Reglas generales

- No usar Power BI.
- No usar HTTP premium hacia GitHub.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.
- Power Automate registra y notifica; Python procesa.
- GitHub Actions ejecuta el dispatcher una vez por hora.
