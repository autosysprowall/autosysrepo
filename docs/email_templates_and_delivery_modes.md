# Plantillas de correo y modos de entrega

Las plantillas provienen de `correos.txt`. Los textos entre corchetes son
valores dinámicos; no deben enviarse literalmente.

La conexión de Outlook envía desde `auto.sys@prowallpanama.com`. Los valores
`[Ingeniero extraído del Gantt]` y `[Persona subiendo el archivo]` se usan como
destinatarios `To`, no como remitentes `From`, porque Outlook no puede suplantar
esas cuentas sin permisos **Send As**.

## Modos

### `test`

- `NOTIFICATION_DELIVERY_MODE=test`;
- `NOTIFICATION_TEST_RECIPIENT=auto.sys@prowallpanama.com`;
- todas las notificaciones de Gantt se redirigen exclusivamente a autosys;
- se elimina el CC real;
- el asunto recibe `[PRUEBA]`;
- el cuerpo muestra `Destinatario real previsto` y `CC real previsto` para
  auditar los correos extraídos sin contactar a esas personas.

Si un presupuesto aceptado no contiene un correo válido del ingeniero, `test`
encola una vista previa para autosys pero no concede permisos ni marca el
proyecto como asignado. En `live` no se envía nada hasta completar
`IngenieroEmail`. Antes de pasar a `live`, eliminar las vistas previas de prueba
para que no bloqueen el envío real por idempotencia.

### `live`

- `NOTIFICATION_DELIVERY_MODE=live`;
- asignación y advertencias van al ingeniero;
- los supervisores y correos de escalamiento se agregan a CC;
- Power Automate debe usar los campos dinámicos `To`, `Cc`, `Subject` y `Body`
  del item, sin un destinatario fijo adicional.

Cambiar el modo solo en la variable de repositorio
`NOTIFICATION_DELIVERY_MODE`. No cambiar el YAML.

## Asignación

- To: `IngenieroEmail`;
- CC: `SupervisoresEmail`;
- asunto: `Asignación Cronograma Proyecto {NombreProyecto}`;
- cuerpo: texto final de `correos.txt`, nombre del proyecto y
  `GanttWorkingLink`;
- adjunto: `Guia_AutoSys_Ingenieros_Residentes_Planta.pdf`.

El enlace se envía como referencia al archivo, pero la automatización no
concede edición al ingeniero. El permiso `write` del WORKING se concede
únicamente a `auto.sys@prowallpanama.com`.

La guía está incorporada como adjunto en `PA_S2_EnviarNotificacionesGantt`.
`ENGINEER_GUIDE_URL` puede conservar un enlace adicional, pero no sustituye el
PDF adjunto.

## Advertencia 1 y 2

- To: `IngenieroEmail`;
- CC: `SupervisoresEmail`, `jaime.madrid@prowallpanama.com` y
  `enrique.correa@prowallpanama.com`;
- asunto: `Advertencia Cronograma Proyecto {NombreProyecto}`;
- el cuerpo usa 3 o 6 días según corresponda.

Los correos fijos pueden reemplazarse mediante `GANTT_ESCALATION_CC`.

## Presupuesto no válido

Este correo lo envía `PA_S1_DevolverPresupuestoInvalido`:

- To live: `CreatedByEmail`;
- CC live: `jaime.madrid@prowallpanama.com`;
- To test: `auto.sys@prowallpanama.com`;
- CC test: vacío;
- asunto live: `Presupuesto No Válido Proyecto {NombreProyecto}`;
- asunto test: `[PRUEBA] Presupuesto No Válido Proyecto {NombreProyecto}`;
- adjunta el presupuesto original;
- adjunta `Guia_AutoSys_Comercial_Presupuesto.pdf`;
- incluye `FileLink` y `UltimoError`;
- en prueba muestra `CreatedByEmail` y el CC real previsto dentro del cuerpo,
  pero envía exclusivamente a autosys.

En Power Automate deben existir dos variables:

- `DeliveryMode`, con `test` o `live`;
- `TestRecipient`, con `auto.sys@prowallpanama.com`.

Cuando exista, la URL de la guía puede guardarse en una variable de entorno de
la solución llamada `BUDGET_GUIDE_URL`. Mientras esté vacía, el flujo no debe
mostrar texto de guía ni un vínculo de reemplazo.

Expresión para **To**:

```text
if(
  equals(variables('DeliveryMode'),'test'),
  variables('TestRecipient'),
  triggerBody()?['CreatedByEmail']
)
```

Expresión para **Cc**:

```text
if(
  equals(variables('DeliveryMode'),'test'),
  '',
  'jaime.madrid@prowallpanama.com'
)
```

Antes de activar `live`, validar que `CreatedByEmail` contenga un correo. Si
está vacío, el flujo debe fallar sin borrar el archivo ni el item de la cola.
