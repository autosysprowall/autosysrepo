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

## Detección de presupuestos aprobados

En esta rama, el registro inicial ya no depende del trigger de Power Automate.
Antes de ejecutar el dispatcher principal, GitHub Actions escanea:

```text
Proyectos/Presupuestos Aprobados
```

El paso `Enqueue missing approved budgets` crea items pendientes en
`Cola_Automatizacion_Proyectos` solo para presupuestos `.xlsx` que:

- no estén ya en la cola;
- no tengan ya un item en `Control_Gantt_Asignaciones` con el mismo
  `ProyectoID`;
- no sean archivos temporales `~$`.

Esto evita el bottleneck observado con SharePoint/Power Automate, donde un
archivo movido a la carpeta puede no disparar inmediatamente el trigger o puede
no dispararlo del todo según cómo SharePoint registre el movimiento.

Los flows de Power Automate `PA_S1_RegistrarPresupuestoAprobado` y
`PA_S1_RegistrarPresupuestoMovido_A_Cola` quedan como fallback opcional, pero
no son la fuente principal para la prueba integral. Si alguno crea un item antes
de GitHub, el escaneo de GitHub detecta que ya existe y no duplica el evento.

Para probar desde cero:

1. Vaciar estado si hace falta con `clear_all_queues=true`.
2. Mover/subir presupuestos a:

```text
/Documentos compartidos/Proyectos/Presupuestos Aprobados
```

3. Ejecutar `SharePoint Automation Dispatcher` en la rama
   `test/full-system-lab` sin marcar casillas.

El primer paso del workflow encolará los presupuestos faltantes y luego el
dispatcher procesará la cola en la misma corrida.
