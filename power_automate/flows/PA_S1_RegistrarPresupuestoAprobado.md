# PA_S1_RegistrarPresupuestoAprobado

## Objetivo

Registrar en `Cola_Automatizacion_Proyectos` cada presupuesto aprobado nuevo que entre a:

```text
/Proyectos/presupuestos aprobados/
```

Este flujo no llama a GitHub. GitHub Actions procesara la cola en su siguiente ejecucion programada de 15 minutos.

## Tipo

Automated cloud flow.

## Trigger

Conector: SharePoint  
Accion: `When a file is created (properties only)`

Configuracion:

| Campo | Valor |
|---|---|
| Site Address | Sitio SharePoint donde vive `/Proyectos/` |
| Library Name | Biblioteca de documentos que contiene `Proyectos` |
| Folder | `/Proyectos/presupuestos aprobados/` |

## Condicion principal

Agregar un bloque `Condition` despues del trigger.

Expresion recomendada en modo avanzado:

```text
@and(
  endsWith(toLower(triggerOutputs()?['body/{FilenameWithExtension}']), '.xlsx'),
  not(startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], '~$')),
  contains(triggerOutputs()?['body/{Path}'], '/Proyectos/presupuestos aprobados/')
)
```

Si el tenant expone `Folder path` con otro nombre interno, usar el contenido dinamico equivalente. Ver [../specs/flow_variables_and_dynamic_content.md](../specs/flow_variables_and_dynamic_content.md).

## Rama If yes

Accion: SharePoint - `Create item`

Lista destino: `Cola_Automatizacion_Proyectos`

Mapeo de campos:

| Campo lista | Valor |
|---|---|
| `Title` | `Filename with extension` |
| `EventType` | `presupuesto_aprobado` |
| `Estado` | `Pendiente` |
| `FileName` | `Filename with extension` |
| `FileIdentifier` | `Identifier` |
| `FolderPath` | `Folder path` o `Path` |
| `FileLink` | `Link to item` |
| `CreatedByEmail` | `Created by Email` |
| `CreatedTime` | `Created` |
| `Intentos` | `0` |
| `UltimoError` | vacio |
| `FechaProcesado` | vacio |
| `ProyectoID` | vacio |
| `Notas` | `Registrado por Power Automate` |

## Rama If no

No hacer nada.

Opcional para pruebas: agregar `Compose` con mensaje `Archivo ignorado por no cumplir condiciones`. Eliminarlo o dejarlo deshabilitado antes de produccion.

## Validaciones de seguridad

- No incluir acciones `Delete file`.
- No incluir acciones `Move file`.
- No incluir acciones dentro de `/Proyectos/PROYECTOS TERMINADOS/`.
- No llamar a GitHub.
- No usar HTTP.

## Resultado esperado

Un item queda en `Cola_Automatizacion_Proyectos` con:

```text
EventType = presupuesto_aprobado
Estado = Pendiente
Intentos = 0
```

Python/GitHub Actions tomara ese item posteriormente.
