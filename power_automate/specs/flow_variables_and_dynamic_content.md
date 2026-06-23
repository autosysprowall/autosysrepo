# Variables, contenido dinamico y expresiones

Los nombres internos exactos pueden variar segun idioma del tenant y tipo de columna. Este documento define la intencion y expresiones base.

## Rutas relevantes

| Concepto | Valor |
|---|---|
| Raiz documental | `/Proyectos/` |
| Inbox presupuestos | `/Proyectos/presupuestos aprobados/` |
| Activos | `/Proyectos/02_Activos/` |
| Terminados intocable | `/Proyectos/PROYECTOS TERMINADOS/` |
| Legacy | `/Proyectos/99_Legacy/` |

## PA_S1_RegistrarPresupuestoAprobado

Trigger: SharePoint `When a file is created (properties only)`.

Contenido dinamico esperado:

| Dato requerido | Contenido dinamico comun |
|---|---|
| Nombre con extension | `Filename with extension` |
| Identificador | `Identifier` |
| Ruta carpeta | `Folder path`, `Path` o `{Path}` |
| Link archivo | `Link to item` |
| Creador email | `Created by Email`, `Author Email` |
| Fecha creacion | `Created` |

Condicion:

```text
@and(
  endsWith(toLower(triggerOutputs()?['body/{FilenameWithExtension}']), '.xlsx'),
  not(startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], '~$')),
  contains(triggerOutputs()?['body/{Path}'], '/Proyectos/presupuestos aprobados/')
)
```

Si `{Path}` no existe, reemplazar por el token interno que Power Automate muestre para la ruta de carpeta.

## PA_S1_NotificarGanttAsignado

Condicion:

```text
@and(
  equals(coalesce(triggerOutputs()?['body/EstadoGantt/Value'], triggerOutputs()?['body/EstadoGantt']), 'Asignado'),
  not(empty(triggerOutputs()?['body/IngenieroEmail'])),
  not(empty(triggerOutputs()?['body/GanttWorkingLink']))
)
```

Validacion minima de correo:

```text
@contains(triggerOutputs()?['body/IngenieroEmail'], '@')
```

Asunto:

```text
concat('Gantt asignado - ', triggerOutputs()?['body/ProyectoID'], ' ', triggerOutputs()?['body/NombreProyecto'])
```

## PA_S1_AdvertenciasGantt

Filter Query recomendado:

```text
(EstadoGantt eq 'Asignado' or EstadoGantt eq 'En progreso') and FechaAsignacion ne null and FechaLimite ne null
```

Dias desde asignacion:

```text
int(div(sub(ticks(utcNow()), ticks(items('Apply_to_each')?['FechaAsignacion'])), 864000000000))
```

Vencimiento:

```text
@greaterOrEquals(outputs('DiasDesdeAsignacion'), 9)
```

Advertencia 2:

```text
@and(
  greaterOrEquals(outputs('DiasDesdeAsignacion'), 6),
  equals(items('Apply_to_each')?['Advertencia2Enviada'], false)
)
```

Advertencia 1:

```text
@and(
  greaterOrEquals(outputs('DiasDesdeAsignacion'), 3),
  equals(items('Apply_to_each')?['Advertencia1Enviada'], false)
)
```

## PA_S1_GanttEnRevision

Condicion:

```text
@and(
  equals(coalesce(triggerOutputs()?['body/EstadoGantt/Value'], triggerOutputs()?['body/EstadoGantt']), 'En revisión inicial'),
  empty(triggerOutputs()?['body/FechaEnvioRevision'])
)
```

Fecha actual:

```text
utcNow()
```

Dias para completar:

```text
int(div(sub(ticks(outputs('FechaRevisionActual')), ticks(triggerOutputs()?['body/FechaAsignacion'])), 864000000000))
```

## GitHub Actions

Power Automate no dispara GitHub.

Ejemplo de schedule:

```yaml
on:
  schedule:
    - cron: "7,22,37,52 * * * *"
  workflow_dispatch:
```

## Campos Choice

Power Automate puede exponer Choice columns como:

- `EstadoGantt`
- `EstadoGantt Value`
- `EstadoGantt/Value`

Por eso las expresiones usan:

```text
coalesce(triggerOutputs()?['body/Campo/Value'], triggerOutputs()?['body/Campo'])
```

Ajustar el nombre interno si Power Automate muestra otro token.
