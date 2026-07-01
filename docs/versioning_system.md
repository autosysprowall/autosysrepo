# Versionado inicial del Gantt

El versionado usa la misma fila del proyecto en
`Control_Gantt_Asignaciones`. No crea otra lista, otro workflow ni un historial
de revisiones.

## Disparador humano

El supervisor revisa el Gantt WORKING y, cuando está conforme:

1. confirma que `EstadoGantt` sea `En revisión inicial`;
2. cambia `SolicitarVersionado` a `Sí`.

El dispatcher horario detecta la solicitud. No se versionan proyectos en
`Asignado`, `En progreso`, `Vencido` o `Requiere revisión manual`.

## Resultado

Python:

1. resuelve el archivo indicado por `GanttWorkingIdentifier` o
   `GanttWorkingLink`;
2. confirma que esté en `Proyectos Activos/.../gantts/working`;
3. rechaza cualquier ruta de `Proyectos Terminados`;
4. crea, si hace falta, la carpeta hermana `gantts/versionados`;
5. copia el WORKING sin modificarlo como
   `{ProyectoID}_gantt_v1.0.xlsx`;
6. actualiza la misma fila de control.

Campos finales:

- `SolicitarVersionado = No`;
- `VersionActual = v1.0`;
- `GanttVersionLink`;
- `GanttVersionIdentifier`;
- `FechaUltimoVersionado`;
- `FechaAprobacion`;
- `EstadoGantt = Aprobado / Versionado`;
- `UltimoErrorVersionado` vacío.

Los nombres internos históricos `VersionadoActual`, `GanntVersionLink` y
`GanntVersionIdentifier` están soportados mediante aliases.

## Idempotencia

- Si `v1.0` ya existe en SharePoint, se reutiliza; nunca se sobrescribe.
- Si la lista ya contiene `v1.0` y su identificador o enlace, solo se recuperan
  los flags finales.
- Una ejecución repetida no crea una segunda versión.
- El WORKING no se mueve, renombra ni sobrescribe.

## Errores

Ante un fallo:

- no se establece `Aprobado / Versionado`;
- `SolicitarVersionado` permanece activo para reintento;
- se incrementa `VersionadoIntentos`;
- la causa se guarda en `UltimoErrorVersionado`;
- GitHub Actions termina con error accionable.

## Alcance

Esta implementación cubre únicamente la versión inicial `v1.0`. Las revisiones
`v1.1`, `v1.2`, `v2.0` y un historial formal quedan fuera de este MVP.
