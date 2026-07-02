# Sistema 2: asignación y tracking del Gantt WORKING

## Alcance

Sistema 2 comienza después de que Sistema 1 genera el Gantt WORKING. Comparte
solamente ese archivo con el ingeniero, encola el correo de asignación, controla
los días 3, 6 y 9, sincroniza el estado `En revisión inicial` y crea
automáticamente la versión correspondiente.

No mueve ni renombra el WORKING, no modifica Gantts cerrados y rechaza cualquier
ruta que contenga `Proyectos Terminados`.

## Arquitectura

```text
Sistema 1 genera Gantt WORKING y actualiza Control_Gantt_Asignaciones
                              |
                              v
GitHub Actions -> src/automation_dispatcher.py
  |-- completa correos desde la hoja Datos
  |-- concede permiso de edición al archivo
  |-- calcula asignación y fecha límite
  |-- crea notificaciones idempotentes
  `-- procesa eventos gantt_working_modificado
                              |
                              v
Cola_Notificaciones_Gantt -> PA_S2_EnviarNotificacionesGantt -> Outlook
```

GitHub Actions ejecuta un único dispatcher. Power Automate no llama a GitHub:
registra cambios del archivo en la cola técnica y envía los correos de la cola
de notificaciones.

## Listas y columnas

Se reutilizan las columnas existentes de `Control_Gantt_Asignaciones` y el
dispatcher crea, si faltan:

| Columna | Uso |
|---|---|
| `GanttWorkingIdentifier` | Identificador estable del Gantt. |
| `PresupuestoIdentifier` | Identificador estable del presupuesto. |
| `SupervisoresEmail` | Correos normalizados con `;`. |
| `PermisoGanttOtorgado` / `FechaPermisoOtorgado` | Idempotencia del permiso. |
| `CorreoAsignacionEnviado` / `FechaCorreoAsignacion` | Confirmación de Power Automate. |
| `FechaAdvertencia1` / `FechaAdvertencia2` | Fechas confirmadas de aviso. |
| `VencimientoNotificado` / `FechaVencimientoNotificado` | Confirmación del escalamiento. |
| `UltimoTrackingRun` / `TrackingIntentos` / `UltimoErrorTracking` | Diagnóstico. |
| `StatusExcel` / `FechaLecturaStatusExcel` | Último status leído del libro. |

La lista `Cola_Notificaciones_Gantt` separa la decisión de negocio del envío
Outlook. Sus columnas están detalladas en
[`power_automate_tracking_flows.md`](power_automate_tracking_flows.md).

El dispatcher crea `Cola_Notificaciones_Gantt` con todas sus columnas en una
sola operación mediante Microsoft Graph. La aplicación de Entra requiere
`Sites.ReadWrite.All` y `Files.ReadWrite.All`.

Agregar columnas individualmente a una lista ya existente requiere además
`Sites.Manage.All`. Si no está concedido, el dispatcher muestra advertencias y
continúa usando las columnas disponibles. Para auditoría completa, conceder ese
permiso con consentimiento de administrador o crear manualmente las columnas
recomendadas y ejecutar de nuevo.

## Correos desde `Datos`

Primero se usa `IngenieroEmail` de la lista. Si está vacío o no es válido, se
abre primero el presupuesto asociado y luego, como respaldo, el Gantt.

En la hoja `Datos` se buscan etiquetas normalizadas:

- ingeniero: `Ingeniero residente`, `IngenieroEmail` o equivalentes;
- supervisores: `Supervisores`, `SupervisoresEmail` o equivalentes.

Los correos se validan y se normalizan separados por `;`. Si no hay un correo de
ingeniero válido, no se comparte el archivo y no se encola correo; el registro
queda `Pendiente de asignación` con un error accionable.

## Permiso y asignación

Microsoft Graph resuelve el Gantt por `GanttWorkingIdentifier` o por
`GanttWorkingLink` y concede permiso `write` solamente al archivo mediante
`invite`. No comparte la carpeta.

Después del permiso:

- `FechaAsignacion` se conserva si ya existía, o se establece al momento actual;
- `FechaLimite = FechaAsignacion + 9 días`, si estaba vacía;
- el estado pasa a `Asignado`, salvo que ya estuviera `En progreso`;
- se crea una única notificación `AsignacionGantt`.

El correo instruye trabajar en el mismo archivo y marcar
`En revisión inicial` al terminar.

## Tracking

Los días se calculan por fecha calendario UTC desde `FechaAsignacion`.

| Momento | Acción única |
|---|---|
| Día 3 o después | Encola `Advertencia1` al ingeniero. |
| Día 6 o después | Encola `Advertencia2` al ingeniero, con supervisores en copia. |
| Día 9 o después | Encola `Vencimiento`; luego marca `Vencido`. |

Solo existen dos advertencias. Si una ejecución se retrasó, encola la acción
vigente de mayor prioridad: no envía advertencias atrasadas después del
vencimiento. Cada tipo solo puede existir una vez por registro.

La idempotencia combina:

- flags confirmados por Power Automate;
- búsqueda de una notificación ya existente por
  `RelatedControlItemID + TipoNotificacion`;
- flags de permiso y estados cerrados.

Una notificación existente en `Pendiente`, `Enviado` o `Error` impide crear un
duplicado. Los errores se corrigen y reintentan sobre el mismo item.

### Modo de entrega de pruebas

Mientras `NOTIFICATION_DELIVERY_MODE=test`, todas las notificaciones se
redireccionan a `NOTIFICATION_TEST_RECIPIENT` y el asunto recibe el prefijo
`[PRUEBA]`. El cuerpo conserva los destinatarios reales previstos solamente como
auditoría. La configuración inicial usa `auto.sys@prowallpanama.com`.

No cambiar `NOTIFICATION_DELIVERY_MODE` a `live` hasta la aprobación del
supervisor.

Las plantillas finales, destinatarios y variables de guía están documentados
en `docs/email_templates_and_delivery_modes.md`.

Además de la redirección en Python, el flujo activo de Power Automate mantiene
su campo `To` fijado a `auto.sys@prowallpanama.com`. Esta segunda barrera se
mantendrá durante la validación con el supervisor.

## Status del Excel

Los Gantts nuevos muestran el estado general en la hoja `Gantt`:

```text
Gantt!A6 | Estado general del Gantt
Gantt!B6 | En progreso
```

`Gantt!B6` tiene una lista con `En progreso` y `En revisión inicial`. El
dispatcher prioriza esta celda visible y conserva `EstadoGantt` en `Datos`
únicamente como compatibilidad para libros anteriores.

Power Automate registra un evento `gantt_working_modificado`; Python descarga el
libro y lee el estado. Si encuentra `En revisión inicial`, actualiza:

- `EstadoGantt = En revisión inicial`;
- `FechaEnvioRevision`, solo si estaba vacía;
- `DiasParaCompletar`;
- `StatusExcel` y `FechaLecturaStatusExcel`.

Los Gantts antiguos que solo tengan `EstadoGantt` en `Datos` continúan siendo
compatibles.

## Dispatcher y horario

El workflow principal es
`.github/workflows/automation-dispatcher.yml`. Ejecuta:

```bash
python src/automation_dispatcher.py
```

Primero procesa Sistema 1 y después Sistema 2. Si Sistema 1 falla, Sistema 2 se
ejecuta igualmente y el job termina con error para conservar visibilidad.

En una ejecución manual, `control_item_id` limita Sistema 2 a un único item de
`Control_Gantt_Asignaciones`. Este modo aislado tampoco consume eventos globales
de modificación de Gantt. Activar además `skip_system1` evita consumir la cola
de presupuestos durante las primeras pruebas controladas.

El cron interno de GitHub Actions está desactivado. El workflow conserva
`workflow_dispatch` y puede ser llamado por el dispatcher externo de 15
minutos cuando este sea activado. El fallback diario documentado, no activado, es
`17 12 * * *`.

## Errores operativos

- Correo vacío: completar `IngenieroEmail` o la hoja `Datos` y reejecutar.
- Permiso fallido: revisar link/identifier y permisos Graph; el flag permanece
  en falso.
- Link roto: el item pasa a `Requiere revisión manual`.
- Cola vacía: termina correctamente.
- Envío Outlook fallido: Power Automate deja la notificación en `Error`; no
  marca el flag de enviado.

## Pendiente para una fase posterior

El MVP usa `v1.0` para costos iguales o menores y `v2.0` para aumentos de costo
o actividades posteriores a la fecha final contractual. Las revisiones `v1.1`,
`v1.2`, `v2.1`, el
historial formal y cualquier correo de aprobación permanecen fuera de Sistema
2. Véase `docs/versioning_system.md`.
