# Power Automate - Sistema 1

Esta carpeta documenta la arquitectura inicial de Power Automate para el Sistema 1 de Autosys: registrar presupuestos aprobados, dar seguimiento humano al Gantt WORKING y dejar que GitHub Actions/Python procese la cola cada 15 minutos.

Power Automate no genera Gantts, no llama a GitHub y no usa Power BI. Su rol es registrar eventos, notificar personas y actualizar estados humanos.

## Alcance

Sistema 1 cubre:

- deteccion de presupuestos nuevos en `/Proyectos/presupuestos aprobados/`;
- registro de eventos en `Cola_Automatizacion_Proyectos`;
- notificacion cuando un Gantt queda asignado a un ingeniero;
- advertencias dia 3 y dia 6;
- vencimiento/escalamiento dia 9;
- registro de envio a revision inicial.

Fuera de alcance:

- Power BI;
- versionado oficial del Gantt;
- Power Automate llamando a GitHub;
- conectores HTTP premium hacia GitHub;
- mover, borrar o reorganizar archivos;
- cualquier operacion dentro de `/Proyectos/PROYECTOS TERMINADOS/`.

## Estructura documental esperada

Raiz documental:

```text
/Proyectos/
├── presupuestos aprobados/
├── 02_Activos/
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

Documentacion nodo por nodo:

- [flows/PA_S1_RegistrarPresupuestoAprobado.md](flows/PA_S1_RegistrarPresupuestoAprobado.md)
- [flows/PA_S1_NotificarGanttAsignado.md](flows/PA_S1_NotificarGanttAsignado.md)
- [flows/PA_S1_AdvertenciasGantt.md](flows/PA_S1_AdvertenciasGantt.md)
- [flows/PA_S1_GanttEnRevision.md](flows/PA_S1_GanttEnRevision.md)

Contenido dinamico y expresiones:

- [specs/flow_variables_and_dynamic_content.md](specs/flow_variables_and_dynamic_content.md)

## GitHub Actions

Power Automate solo registra eventos en `Cola_Automatizacion_Proyectos`.

GitHub Actions debe correr por horario cada 15 minutos y Python debe leer los items `Pendiente`.

Ejemplo:

```yaml
on:
  schedule:
    - cron: "7,22,37,52 * * * *"
  workflow_dispatch:
```

El schedule puede retrasarse algunos minutos por infraestructura de GitHub. Es aceptable.

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

## Seguridad

- No activar flujos automaticamente desde este repo.
- No ejecutar contra produccion durante la construccion.
- No borrar archivos.
- No mover archivos.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.
- No hardcodear correos reales salvo configuracion formal aprobada.
- Registrar errores en listas, no ocultarlos.
