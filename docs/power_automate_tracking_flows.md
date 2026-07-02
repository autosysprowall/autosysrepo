# Power Automate para Sistema 2

Estos flujos usan únicamente SharePoint y Office 365 Outlook, conectores
estándar. No usan HTTP premium, no llaman GitHub y no procesan
`Proyectos Terminados`.

Estado al 30 de junio de 2026:

- `Cola_Notificaciones_Gantt` y sus columnas ya existen en SharePoint.
- `PA_S2_EnviarNotificacionesGantt` está creado, activo y validado en modo de
  prueba.
- `PA_S2_GanttWorkingModificado_A_Cola` todavía está pendiente de crear.
- `PA_S2_SincronizarEstadoGanttExcel` está definido en
  `power_automate/flows/PA_S2_SincronizarEstadoGanttExcel.md` y usa el Office
  Script `power_automate/office_scripts/SetGanttStatus.ts`.
- El Switch que confirma flags en `Control_Gantt_Asignaciones` y el Scope de
  fallo todavía están pendientes antes de activar destinatarios reales.
- `PA_S1_DevolverPresupuestoInvalido` está creado y guardado sin errores, pero
  permanece apagado. Faltan los destinatarios aprobados y el texto final de
  Contabilidad/Comercial. Su contrato, adjunto y orden seguro de eliminación
  están en `docs/budget_return_flow.md`.

## Lista `Cola_Notificaciones_Gantt`

El dispatcher puede crear la lista y estas columnas en una sola operación. Si
la lista ya existe pero le faltan columnas, agregarlas manualmente o conceder
`Sites.Manage.All` a la aplicación de Entra:

Ruta para el permiso: **Microsoft Entra admin center > App registrations >
aplicación de Autosys > API permissions > Add a permission > Microsoft Graph >
Application permissions > Sites.Manage.All > Grant admin consent**. No requiere
un conector premium de Power Automate.

Alternativa sin ampliar el permiso de la app: en SharePoint, **Site contents >
New > List > Blank list**, crear `Cola_Notificaciones_Gantt` y luego agregar las
columnas de la tabla siguiente.

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

Este flujo consume los mensajes ya redactados por Python para:

- asignación;
- Advertencia 1;
- Advertencia 2;
- vencimiento.

La cola entrega `TipoNotificacion`, `Subject`, `Body`, `To` y `Cc`. Power
Automate no debe volver a redactar ni reemplazar esos campos. El modo `test`
se aplica antes de crear el item, por lo que el flujo puede usar siempre los
valores dinámicos. No activar `live` hasta validar las cuatro ramas.

### Configuración activa de prueba

El flujo activo usa el trigger **When an item is created or modified** sobre
`Cola_Notificaciones_Gantt`, concurrencia `1` y esta condición de trigger:

```text
@equals(triggerBody()?['EstadoNotificacion'],'Pendiente')
```

Durante la validación, Python escribe `To = auto.sys@prowallpanama.com`, elimina
el CC real y agrega `[PRUEBA]` al asunto. **Send an email (V2)** debe usar el
campo dinámico `To`; no necesita un destinatario fijo. Después
del envío exitoso, **Update item** preserva los campos, establece
`EstadoNotificacion = Enviado` y `FechaEnvio = utcNow()`.

La prueba controlada del 30 de junio terminó `Succeeded`: envío en 0.9 segundos
y actualización del item en 0.7 segundos. No se usaron destinatarios reales.
No sustituir el destinatario fijo por `To` hasta obtener aprobación del
supervisor y terminar el Switch y el manejo de fallos descritos abajo.

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

El flujo está desplegado y activo:

- Environment:
  `Default-ed7d4cfd-f42f-48ee-85ae-2e2be3539cd4`
- Flow ID: `aef2a007-ee4f-4ac5-8a82-ce94b374de1f`
- Script de lectura: `GetGanttStatus`
- Script de despliegue:
  `power_automate/scripts/deploy_gantt_modified_flow.py`

1. Trigger: SharePoint, **When a file is created or modified (properties only)**.
2. Seleccionar el sitio y la biblioteca de `/Proyectos/`.
3. Aplicar una condición AND usando ruta y nombre:
   - la ruta contiene `/Proyectos Activos/`;
   - la ruta contiene `/gantts/working/`;
   - la ruta no contiene `/Proyectos Terminados/`;
   - el nombre no comienza con `~$`;
   - el nombre termina con `_gantt_WORKING.xlsx`.
4. Resolver una única fila de `Control_Gantt_Asignaciones` mediante el
   `ProyectoID` contenido en el nombre.
5. Ejecutar `GetGanttStatus` y leer `Gantt!B6`.
6. Si el cambio fue producido por `SetGanttStatus`, ignorarlo para evitar
   bucles.
7. Si el estado es `Actual` o `En Progreso` y la lista estaba en `Actual`:
   - cambiar `EstadoGantt` a `En Progreso`;
   - iniciar `FechaInicioEnProgreso` solo si estaba vacía;
   - solicitar `B6 = En Progreso` mediante `EstadoSyncExcel = Pendiente`.
8. Si el estado es `Entregar`:
   - cambiar el control a `Entregar`;
   - activar `SolicitarVersionado`;
   - crear, si no existe otro pendiente, un item en
     `Cola_Automatizacion_Proyectos`:
   - `Title = <File name with extension>`;
   - `EventType = gantt_working_modificado`;
   - `Estado = Pendiente`;
   - `FileName = File name with extension`;
   - `FileIdentifier = Identifier`;
   - `FolderPath = Folder path`;
   - `FileLink = Link to item`;
   - `Intentos = 0`;
   - `Notas = Detectado por Power Automate`.
Este flujo lee solamente la celda de estado mediante Office Scripts. No
interpreta costos o fechas, no crea versiones y no mueve el Gantt. La
validación financiera y cronológica permanece exclusivamente en Python.

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
