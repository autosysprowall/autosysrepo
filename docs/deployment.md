# Deployment de AutoSys

## Rama operativa

La rama consolidada y predeterminada del repositorio es:

```text
feature/power-automate-sistema1-flows
```

Esta rama contiene Sistema 1, generación y validación del Gantt, asignación,
tracking, sincronización de estado, versionado acumulativo y devolución de
presupuestos inválidos. No es necesario crear una rama `main` adicional:
cambiar el nombre obligaría a actualizar el dispatcher externo sin aportar una
separación nueva.

## Correos

El deployment usa:

```text
NOTIFICATION_DELIVERY_MODE=live
```

La asignación, advertencias y vencimiento usan los correos extraídos de la hoja
`Datos`. La devolución usa `CreatedByEmail`. Los flujos de Outlook adjuntan las
guías vigentes desde `power_automate/assets/`.

El permiso de edición que concede Python sobre el Gantt WORKING se dirige
únicamente a `auto.sys@prowallpanama.com`. Los correos del ingeniero no se usan
para conceder acceso.

## Reinicio integral para una muestra nueva

En GitHub:

1. Abrir **Actions**.
2. Seleccionar **SharePoint Automation Dispatcher**.
3. Pulsar **Run workflow**.
4. Seleccionar `feature/power-automate-sistema1-flows`.
5. Usar:
   - `top`: `50`;
   - `max_items`: `5` o más que la cantidad de presupuestos;
   - `reset_system1_data`: `true`;
   - `reset_related_tracking`: `true`;
   - `reset_queue_only`: `false`;
   - `skip_system1`: `false`;
   - `skip_system2`: `false`.
6. Ejecutar el workflow una sola vez.

`reset_system1_data=true` vacía la cola técnica y el contenido procesable de
`Proyectos Activos`, y vuelve a encolar los presupuestos aprobados. No toca
`Proyectos Terminados`.

Si solo se necesita volver a crear eventos sin borrar los proyectos activos,
usar `reset_queue_only=true` y dejar `reset_system1_data=false`; esta opción no
es apropiada para una prueba integral desde cero.

Si `Presupuestos Aprobados` está vacío y se quiere preparar una carga limpia,
usar `clear_queue_only=true`. Esta opción elimina los items existentes de la
cola sin reencolar archivos y no toca `Proyectos Activos`.

## Dispatcher externo

El cuerpo que invoca GitHub debe contener:

```json
{
  "ref": "feature/power-automate-sistema1-flows"
}
```

El cron interno de GitHub permanece desactivado. La periodicidad depende del
dispatcher externo configurado cada 15 minutos.
