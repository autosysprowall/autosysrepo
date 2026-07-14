# Prueba integral en modo laboratorio

Esta guía describe la rama `test/full-system-lab`. Su propósito es probar el
flujo completo sin enviar correos a ingenieros, supervisores ni CCs ejecutivos.

## Seguridad de correos

En esta rama, el workflow `SharePoint Automation Dispatcher` fuerza:

- `NOTIFICATION_DELIVERY_MODE=test`
- `NOTIFICATION_TEST_RECIPIENT=auto.sys@prowallpanama.com`
- `GANTT_ESCALATION_CC=""`

Por lo tanto, las notificaciones creadas por Python se envían solo a autosys.
El cuerpo conserva el destinatario real previsto y el CC real previsto para
auditoría.

El flujo `PA_S3_MonitorearEstadosActividades` también está generado en modo
prueba:

- To: `auto.sys@prowallpanama.com`
- Cc: vacío
- Subject: prefijo `[PRUEBA]`
- Body: incluye ingeniero y supervisores reales previstos.

La devolución de presupuestos inválidos también queda en modo prueba:

- To: `auto.sys@prowallpanama.com`
- Cc: vacío
- Subject: prefijo `[PRUEBA]`
- Body: incluye el uploader real previsto y el CC real previsto.

## Tiempos acelerados

El workflow fija estos tiempos:

| Evento | Producción | Laboratorio |
|---|---:|---:|
| Fecha límite del Gantt | 9 días | 30 minutos |
| Advertencia 1 | día 3 | minuto 10 |
| Advertencia 2 | día 6 | minuto 20 |
| Vencimiento | día 9 | minuto 30 |

Variables:

```text
GANTT_ASSIGNMENT_DEADLINE_MINUTES=30
GANTT_WARNING1_MINUTES=10
GANTT_WARNING2_MINUTES=20
GANTT_EXPIRATION_MINUTES=30
```

Si estas variables no están presentes, Python conserva la lógica productiva de
3, 6 y 9 días.

## Qué no hace esta rama

- No prende Power Automate.
- No habilita GitHub Actions.
- No cambia producción automáticamente.
- No elimina listas ni archivos por sí sola.

Para ejecutar la prueba real hay que habilitar manualmente los flujos y el
workflow, apuntando a `test/full-system-lab`.

## Trigger de presupuestos movidos

Para el laboratorio, el registro automático de presupuestos debe usar el flow:

```text
PA_S1_RegistrarPresupuestoMovido_A_Cola
```

Este flow usa SharePoint **When a file is created or modified (properties
only)** sobre:

```text
/Documentos compartidos/Proyectos/Presupuestos Aprobados
```

Esto permite detectar archivos que fueron movidos a la carpeta, no solo archivos
creados originalmente ahí. El flow anterior `PA_S1_RegistrarPresupuestoAprobado`
usa `When a file is created` y puede no dispararse cuando SharePoint registra el
movimiento como modificación. Para evitar duplicados, el flow nuevo revisa si ya
existe un item `Pendiente` o `Procesando` con el mismo `FileID` antes de crear
otro item en `Cola_Automatizacion_Proyectos`.
