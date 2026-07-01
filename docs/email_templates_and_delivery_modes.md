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
- el cuerpo muestra los destinatarios reales previstos para auditoría.

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
- cuerpo: texto de asignación, fechas, `GanttWorkingLink` y guía de ingenieros.

La guía se configura con la variable `ENGINEER_GUIDE_URL`. Mientras esté
vacía, se omite completamente del correo.

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
- incluye `FileLink` y `UltimoError`;
- la guía de presupuesto se omite mientras no exista.

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
