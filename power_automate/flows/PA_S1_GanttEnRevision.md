# PA_S1_GanttEnRevision

## Objetivo

Registrar automaticamente cuando un ingeniero cambia `EstadoGantt` a `En revisión inicial`.

Este flujo no crea versiones oficiales. El versionado `v1.0` pertenece a una fase posterior.

## Tipo

Automated cloud flow.

## Trigger

Conector: SharePoint  
Accion: `When an item is created or modified`

Configuracion:

| Campo | Valor |
|---|---|
| Site Address | Sitio SharePoint del proyecto |
| List Name | `Control_Gantt_Asignaciones` |

## Condicion principal

Continuar solo si:

- `EstadoGantt` = `En revisión inicial`
- `FechaEnvioRevision` esta vacia

Expresion recomendada:

```text
@and(
  equals(coalesce(triggerOutputs()?['body/EstadoGantt/Value'], triggerOutputs()?['body/EstadoGantt']), 'En revisión inicial'),
  empty(triggerOutputs()?['body/FechaEnvioRevision'])
)
```

Esta condicion evita loop infinito: despues de actualizar `FechaEnvioRevision`, el flujo se dispara otra vez pero ya no cumple la condicion.

## Calcular fecha de revision

Compose: `FechaRevisionActual`

```text
utcNow()
```

## Calcular dias para completar

Si `FechaAsignacion` existe:

Compose: `DiasParaCompletar`

```text
int(div(sub(ticks(outputs('FechaRevisionActual')), ticks(triggerOutputs()?['body/FechaAsignacion'])), 864000000000))
```

Si `FechaAsignacion` esta vacia, no calcular dias y agregar nota de revision manual.

## Actualizar item

Accion: SharePoint - `Update item`

Actualizar el mismo registro:

| Campo | Valor |
|---|---|
| `FechaEnvioRevision` | `FechaRevisionActual` |
| `DiasParaCompletar` | salida de `DiasParaCompletar` |
| `Notas` | conservar nota anterior + `Enviado a revision inicial registrado por Power Automate.` |

No cambiar `EstadoGantt`; debe permanecer `En revisión inicial`.

## Notificar supervisor

Solo si existe campo de supervisor o configuracion formal aprobada.

Accion: Office 365 Outlook - `Send an email (V2)`

Asunto:

```text
Gantt enviado a revision - {ProyectoID} {NombreProyecto}
```

Cuerpo:

```text
El Gantt del proyecto {ProyectoID} - {NombreProyecto} fue marcado como "En revisión inicial".

Archivo de trabajo: {GanttWorkingLink}
Fecha de envio a revision: {FechaEnvioRevision}
Dias para completar: {DiasParaCompletar}
```

## Validaciones de seguridad

- No crear version oficial.
- No mover archivos.
- No editar Excel.
- No llamar GitHub.
- No usar Power BI.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.
