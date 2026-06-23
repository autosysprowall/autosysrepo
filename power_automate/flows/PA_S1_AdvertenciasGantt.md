# PA_S1_AdvertenciasGantt

## Objetivo

Enviar dos advertencias maximas y luego marcar vencimiento si el Gantt asignado no avanza.

Regla actual:

| Dia desde asignacion | Accion |
|---:|---|
| 0 | Gantt asignado |
| 3 | Advertencia 1 |
| 6 | Advertencia 2 |
| 9 | Vencimiento / escalamiento |

No existe tercera advertencia.

## Tipo

Scheduled cloud flow.

## Trigger

Conector: Schedule  
Accion: `Recurrence`

Configuracion recomendada:

| Campo | Valor |
|---|---|
| Frequency | Day |
| Interval | 1 |
| Time zone | America/Panama |
| Start time | Hora operativa aprobada por el equipo |

## Obtener items

Accion: SharePoint - `Get items`

Lista: `Control_Gantt_Asignaciones`

Filter Query recomendado:

```text
(EstadoGantt eq 'Asignado' or EstadoGantt eq 'En progreso') and FechaAsignacion ne null and FechaLimite ne null
```

Si `EstadoGantt` es columna Choice y el conector no acepta el filtro directo, traer items con filtro mas simple y aplicar las condiciones dentro del flujo.

## Apply to each

Recorrer cada item devuelto por `Get items`.

Calcular dias desde asignacion con `Compose`:

Nombre: `DiasDesdeAsignacion`

Expresion:

```text
int(div(sub(ticks(utcNow()), ticks(items('Apply_to_each')?['FechaAsignacion'])), 864000000000))
```

## Orden de evaluacion

Evaluar en este orden para evitar enviar advertencias a un item ya vencido:

1. Vencimiento dia 9.
2. Advertencia 2 dia 6.
3. Advertencia 1 dia 3.

## Vencimiento

Condicion:

```text
@greaterOrEquals(outputs('DiasDesdeAsignacion'), 9)
```

Acciones en `If yes`:

1. SharePoint - `Update item`
   - `EstadoGantt` = `Vencido`
   - conservar otros campos
   - agregar nota: `Vencido automaticamente por PA_S1_AdvertenciasGantt.`
2. Office 365 Outlook - `Send an email (V2)`
   - To: `IngenieroEmail`
   - CC: supervisor solo si existe campo/configuracion formal.
   - Subject: `Gantt vencido - {ProyectoID} {NombreProyecto}`
   - Body:

```text
El Gantt del proyecto {ProyectoID} - {NombreProyecto} llego al dia 9 sin pasar a revision.

Archivo de trabajo: {GanttWorkingLink}
Estado actualizado: Vencido
```

## Advertencia 2

Solo evaluar si no entro en vencimiento.

Condicion:

```text
@and(
  greaterOrEquals(outputs('DiasDesdeAsignacion'), 6),
  equals(items('Apply_to_each')?['Advertencia2Enviada'], false)
)
```

Acciones:

1. Enviar correo al ingeniero.
2. Actualizar item:
   - `Advertencia2Enviada` = `true`
   - conservar `EstadoGantt`.

Asunto:

```text
Advertencia 2 - Gantt pendiente - {ProyectoID} {NombreProyecto}
```

## Advertencia 1

Solo evaluar si no entro en vencimiento ni advertencia 2.

Condicion:

```text
@and(
  greaterOrEquals(outputs('DiasDesdeAsignacion'), 3),
  equals(items('Apply_to_each')?['Advertencia1Enviada'], false)
)
```

Acciones:

1. Enviar correo al ingeniero.
2. Actualizar item:
   - `Advertencia1Enviada` = `true`
   - conservar `EstadoGantt`.

Asunto:

```text
Advertencia 1 - Gantt pendiente - {ProyectoID} {NombreProyecto}
```

## Consideraciones

- No enviar correos si `IngenieroEmail` esta vacio.
- Si el correo esta vacio, marcar `EstadoGantt = Requiere revisión manual` o agregar nota.
- No crear tercera advertencia.
- No modificar archivos.
- No versionar Gantts.
- No tocar `/Proyectos/PROYECTOS TERMINADOS/`.
