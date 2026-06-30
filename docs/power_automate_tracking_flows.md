# Power Automate para Sistema 2

Estos flujos usan únicamente SharePoint y Office 365 Outlook, conectores
estándar. No usan HTTP premium, no llaman GitHub y no procesan
`Proyectos Terminados`.

No se generó un paquete importable porque este entorno no tiene una conexión
autenticada a Power Platform. Los pasos siguientes son la definición exacta
para construirlos en la interfaz de Power Automate.

## Lista `Cola_Notificaciones_Gantt`

El dispatcher puede crear la lista y estas columnas en una sola operación. Si
la lista ya existe pero le faltan columnas, agregarlas manualmente o conceder
`Sites.Manage.All` a la aplicación de Entra:

| Columna | Tipo |
|---|---|
| `Title` | Texto |
| `ProyectoID` | Texto |
| `TipoNotificacion` | Choice: `AsignacionGantt`, `Advertencia1`, `Advertencia2`, `Vencimiento` |
| `EstadoNotificacion` | Choice: `Pendiente`, `Enviado`, `Error` |
| `To` | Texto multilínea |
| `Cc` | Texto multilínea |
| `Subject` | Texto |
| `Body` | Texto multilínea |
| `RelatedControlItemID` | Texto |
| `Intentos` | Número |
| `UltimoError` | Texto multilínea |
| `FechaCreacion` | Fecha y hora |
| `FechaEnvio` | Fecha y hora |
| `Notas` | Texto multilínea |

## Flujo 1: `PA_S2_EnviarNotificacionesGantt`

1. Crear un **Automated cloud flow**.
2. Trigger: SharePoint, **When an item is created or modified**.
3. Seleccionar el sitio y `Cola_Notificaciones_Gantt`.
4. En Settings del trigger, activar Concurrency Control y usar grado `1`.
5. Agregar una Condition con AND:
   - `EstadoNotificacion` es `Pendiente`;
   - `To` no está vacío;
   - `Subject` no está vacío;
   - `Body` no está vacío.
6. Rama No: terminar con estado `Succeeded`.
7. Rama Sí: Office 365 Outlook, **Send an email (V2)**:
   - To: `To`;
   - Cc: `Cc`;
   - Subject: `Subject`;
   - Body: `Body`.
8. Después del envío, SharePoint **Update item** sobre la notificación:
   - `EstadoNotificacion = Enviado`;
   - `FechaEnvio = utcNow()`;
   - `UltimoError` vacío;
   - preservar el resto de campos requeridos.
9. SharePoint **Get item** en `Control_Gantt_Asignaciones` usando
   `int(RelatedControlItemID)`.
10. Agregar Switch por `TipoNotificacion` y actualizar el mismo control:
    - `AsignacionGantt`: `CorreoAsignacionEnviado = Yes`,
      `FechaCorreoAsignacion = utcNow()`;
    - `Advertencia1`: `Advertencia1Enviada = Yes`,
      `FechaAdvertencia1 = utcNow()`;
    - `Advertencia2`: `Advertencia2Enviada = Yes`,
      `FechaAdvertencia2 = utcNow()`;
    - `Vencimiento`: `VencimientoNotificado = Yes`,
      `FechaVencimientoNotificado = utcNow()`. Mantener `Vencido` si Python ya
      lo estableció; solo cambiarlo si aún está `Asignado` o `En progreso`.

### Manejo del fallo

Crear un Scope `Fallo` configurado con **Run after** para timeout y error del
scope de envío. En él, actualizar el mismo item:

- `EstadoNotificacion = Error`;
- `Intentos = add(coalesce(triggerBody()?['Intentos'], 0), 1)`;
- `UltimoError` con el mensaje disponible de `result('Scope_Envio')`;
- no modificar flags de envío en `Control_Gantt_Asignaciones`.

Para reintentar, corregir la causa y cambiar ese mismo item de `Error` a
`Pendiente`. No crear otro item.

## Flujo 2: `PA_S2_GanttWorkingModificado_A_Cola`

1. Crear un **Automated cloud flow**.
2. Trigger: SharePoint, **When a file is created or modified (properties only)**.
3. Seleccionar el sitio y la biblioteca de `/Proyectos/`.
4. Agregar una Condition AND usando ruta y nombre:
   - la ruta contiene `/Proyectos Activos/`;
   - la ruta contiene `/gantts/`;
   - la ruta no contiene `/Proyectos Terminados/`;
   - el nombre no comienza con `~$`;
   - el nombre comienza con `gannt_`;
   - el nombre termina con `_working.xlsx`.
5. Rama Sí: SharePoint **Create item** en
   `Cola_Automatizacion_Proyectos`:
   - `Title = Gantt WORKING modificado - <File name with extension>`;
   - `EventType = gantt_working_modificado`;
   - `Estado = Pendiente`;
   - `FileName = File name with extension`;
   - `FileIdentifier = Identifier`;
   - `FolderPath = Folder path`;
   - `FileLink = Link to item`;
   - `Intentos = 0`;
   - `Notas = Detectado por Power Automate`.
6. Rama No: terminar sin acciones.

Este flujo registra el evento. No abre Excel, no crea versiones y no mueve el
Gantt. Los patrones reflejan la estructura real existente del generador; no se
crea una carpeta `working` adicional.

## Evitar duplicados con flujos antiguos

Al activar Sistema 2, mantener apagados:

- `PA_S1_NotificarGanttAsignado`;
- `PA_S1_AdvertenciasGantt`;
- `PA_S1_GanttEnRevision`.

Esos documentos quedan como referencia histórica. Si se ejecutan en paralelo,
pueden duplicar correos o competir por flags. En Sistema 2, la única fuente de
correos es `Cola_Notificaciones_Gantt`.

## Prueba de aceptación

1. Crear una notificación de prueba `Pendiente` con un destinatario controlado.
2. Confirmar que llega un solo correo.
3. Confirmar `EstadoNotificacion = Enviado` y el flag correspondiente en
   `Control_Gantt_Asignaciones`.
4. Volver a guardar el item sin cambiarlo a `Pendiente`: no debe reenviar.
5. Modificar un `gannt_*_working.xlsx` de prueba en Proyectos Activos y confirmar
   un evento `gantt_working_modificado`.
6. Verificar que ningún flujo contiene una acción sobre Proyectos Terminados.
