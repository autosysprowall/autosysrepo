# Esquema esperado de listas SharePoint

Las listas ya existen manualmente. Este documento define los nombres y valores que los flujos y Python esperan.

## Cola_Automatizacion_Proyectos

Funcion: cola tecnica de eventos que Python/GitHub Actions procesara cada 15 minutos.

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

### Valores permitidos de EstadoGantt

- `Pendiente de asignación`
- `Asignado`
- `En progreso`
- `En revisión inicial`
- `Aprobado / Versionado`
- `Vencido`
- `Requiere revisión manual`

## Reglas generales

- No usar Power BI.
- No usar HTTP premium hacia GitHub.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.
- Power Automate registra y notifica; Python procesa.
- GitHub Actions revisa la cola cada 15 minutos.
