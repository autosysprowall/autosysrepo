# PA_S1_NotificarGanttAsignado

## Objetivo

Notificar al ingeniero cuando Python haya creado un registro en `Control_Gantt_Asignaciones` y el Gantt WORKING este listo para trabajo humano.

## Tipo

Automated cloud flow.

## Trigger

Conector: SharePoint  
Accion: `When an item is created`

Configuracion:

| Campo | Valor |
|---|---|
| Site Address | Sitio SharePoint del proyecto |
| List Name | `Control_Gantt_Asignaciones` |

## Condicion principal

Continuar solo si:

- `EstadoGantt` = `Asignado`
- `IngenieroEmail` no esta vacio
- `GanttWorkingLink` no esta vacio

Expresion recomendada en modo avanzado:

```text
@and(
  equals(coalesce(triggerOutputs()?['body/EstadoGantt/Value'], triggerOutputs()?['body/EstadoGantt']), 'Asignado'),
  not(empty(triggerOutputs()?['body/IngenieroEmail'])),
  not(empty(triggerOutputs()?['body/GanttWorkingLink']))
)
```

## Validacion simple de correo

Agregar segunda condicion dentro de `If yes`:

```text
@contains(triggerOutputs()?['body/IngenieroEmail'], '@')
```

## Rama correo valido

Accion: Office 365 Outlook - `Send an email (V2)`

| Campo | Valor |
|---|---|
| To | `IngenieroEmail` |
| Subject | `Gantt asignado - {ProyectoID} {NombreProyecto}` |
| Importance | Normal |

Cuerpo sugerido:

```text
Se le ha asignado el Gantt del proyecto {ProyectoID} - {NombreProyecto}.

Archivo de trabajo: {GanttWorkingLink}

Fecha de asignacion: {FechaAsignacion}
Fecha limite: {FechaLimite}

Cuando termine, actualice el estado del registro a "En revisión inicial".
```

## Rama correo invalido o vacio

Accion: SharePoint - `Update item`

Actualizar el mismo item:

| Campo | Valor |
|---|---|
| `EstadoGantt` | `Requiere revisión manual` |
| `Notas` | concatenar nota existente + `No se envio correo: IngenieroEmail vacio o invalido.` |

Si se prefiere no cambiar estado, al menos actualizar `Notas`.

## Validaciones de seguridad

- No modificar archivos Excel.
- No crear versiones.
- No llamar GitHub.
- No usar HTTP.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.

## Resultado esperado

El ingeniero recibe un correo con el link al Gantt WORKING y la fecha limite. El registro permanece controlado desde `Control_Gantt_Asignaciones`.
