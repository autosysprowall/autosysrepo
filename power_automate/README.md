# Power Automate - Sistema 1

Esta carpeta documenta Power Automate para Autosys. Sistema 1 registra
presupuestos aprobados. Sistema 2 registra modificaciones del Gantt WORKING,
sincroniza su estado general y envía las notificaciones decididas por Python.
Sistema 3 supervisa los estados de las actividades dentro del Gantt.

Power Automate no genera Gantts, no llama a GitHub y no usa Power BI. Su rol es registrar eventos, notificar personas y actualizar estados humanos.

## Alcance

Los sistemas cubren:

- deteccion de presupuestos nuevos en `/Proyectos/presupuestos aprobados/`;
- registro de eventos en `Cola_Automatizacion_Proyectos`;
- cola idempotente de notificaciones;
- correo de asignacion, advertencias dia 3 y dia 6;
- vencimiento/escalamiento dia 9;
- registro de modificaciones para leer el status del Excel.
- alertas por actividades `Pendiente` o `En Progreso` fuera de fecha.

Fuera de alcance:

- Power BI;
- Power Automate llamando a GitHub;
- conectores HTTP premium hacia GitHub;
- mover, borrar o reorganizar archivos;
- cualquier operacion dentro de `/Proyectos/PROYECTOS TERMINADOS/`.

## Estructura documental esperada

Raiz documental:

```text
/Proyectos/
├── presupuestos aprobados/
├── Proyectos Activos/
├── PROYECTOS TERMINADOS/
└── 99_Legacy/
```

La carpeta `/Proyectos/PROYECTOS TERMINADOS/` es intocable. Ningun flujo debe leer, mover, editar, copiar, crear ni borrar nada dentro de esa carpeta.

## Listas requeridas

Las listas ya existen manualmente en SharePoint:

- `Cola_Automatizacion_Proyectos`
- `Control_Gantt_Asignaciones`

El esquema esperado esta documentado en [specs/sharepoint_lists_expected_schema.md](specs/sharepoint_lists_expected_schema.md).

## Flujos definidos

| Flujo | Tipo | Proposito |
|---|---|---|
| `PA_S1_RegistrarPresupuestoAprobado` | Automated cloud flow | Registra un presupuesto `.xlsx` nuevo en la cola tecnica. |
| `PA_S1_NotificarGanttAsignado` | Automated cloud flow | Notifica al ingeniero cuando un Gantt WORKING queda asignado. |
| `PA_S1_AdvertenciasGantt` | Scheduled cloud flow | Envia advertencias dia 3 y 6, y vence/escalona dia 9. |
| `PA_S1_GanttEnRevision` | Automated cloud flow | Registra fecha de envio a revision inicial y dias usados. |
| `PA_S2_EnviarNotificacionesGantt` | Automated cloud flow | Envia items pendientes de la cola de notificaciones y confirma flags. |
| `PA_S2_GanttWorkingModificado_A_Cola` | Automated cloud flow | Registra modificaciones del Gantt para lectura por Python. |
| `PA_S2_SincronizarEstadoGanttExcel` | Automated cloud flow | Sincroniza `Gantt!B6` mediante Office Scripts sin reemplazar el workbook completo. |
| `PA_S3_MonitorearEstadosActividades` | Scheduled cloud flow | Lee cada 15 minutos los estados por actividad, registra atrasos y envía una alerta idempotente. |

Documentacion nodo por nodo:

- [flows/PA_S1_RegistrarPresupuestoAprobado.md](flows/PA_S1_RegistrarPresupuestoAprobado.md)
- [flows/PA_S1_NotificarGanttAsignado.md](flows/PA_S1_NotificarGanttAsignado.md)
- [flows/PA_S1_AdvertenciasGantt.md](flows/PA_S1_AdvertenciasGantt.md)
- [flows/PA_S1_GanttEnRevision.md](flows/PA_S1_GanttEnRevision.md)
- [Definicion completa de los flujos de Sistema 2](../docs/power_automate_tracking_flows.md)
- [Sincronizacion segura de Gantt!B6](flows/PA_S2_SincronizarEstadoGanttExcel.md)
- [Alertas por estado de actividades](../docs/activity_status_alerts.md)
- Office Scripts:
  `office_scripts/GetGanttStatus.ts` y
  `office_scripts/SetGanttStatus.ts`,
  `office_scripts/GetGanttActivityStatus.ts`.

Los tres flujos S1 de correo/tracking quedan como referencia historica y no
deben activarse junto con Sistema 2, porque duplicarian correos o estados.

Contenido dinamico y expresiones:

- [specs/flow_variables_and_dynamic_content.md](specs/flow_variables_and_dynamic_content.md)

## GitHub Actions

Power Automate solo registra eventos en `Cola_Automatizacion_Proyectos`.

GitHub Actions conserva `workflow_dispatch`. El cron interno está desactivado;
el dispatcher externo lo invoca cada 15 minutos usando la rama
`feature/power-automate-sistema1-flows`.

## Conectores

Conectores estandar esperados:

- SharePoint
- Office 365 Outlook

No usar:

- HTTP premium hacia GitHub;
- Power BI;
- conectores premium si no son estrictamente necesarios.

## Importacion o construccion manual

No se genero paquete importable desde esta maquina porque Power Platform CLI (`pac`) no esta instalado ni hay una conexion validada al entorno Power Platform. La carpeta [solution/](solution/) queda preparada para colocar un export real cuando se cree desde Power Automate o Power Platform CLI.

Mientras tanto, cada flujo puede construirse manualmente siguiendo los documentos nodo por nodo en `flows/`.

Recomendacion para construccion:

1. Crear los flujos apagados.
2. Validar sitio, biblioteca y lista en ambiente controlado.
3. Probar con un archivo `.xlsx` de prueba fuera de produccion operativa.
4. Activar solo cuando Python/GitHub Actions ya pueda procesar la cola.
5. Mantener apagados los flujos S1 antiguos de correo, advertencias y revision.

## Seguridad

- No activar flujos automaticamente desde este repo.
- No ejecutar contra produccion durante la construccion.
- No borrar archivos.
- No mover archivos.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.
- No hardcodear correos reales salvo configuracion formal aprobada.
- Registrar errores en listas, no ocultarlos.
