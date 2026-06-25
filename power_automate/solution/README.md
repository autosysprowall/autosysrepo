# Solucion Power Automate

## Estado actual

El flujo principal del Sistema 1 fue creado en Dataverse/Power Automate como cloud flow en borrador:

- Nombre: `PA_S1_RegistrarPresupuestoAprobado`
- Workflow ID: `02b348f0-9c70-f111-ab0f-000d3a34454c`
- Ambiente: `Prowall (default)`
- Estado: `Draft/Off`
- Conector usado: SharePoint (`shared_sharepointonline`)

El flujo no pudo activarse automaticamente porque el ambiente no tiene una conexion SharePoint asociada al flow.

Error de activacion recibido:

```text
FlowMissingConnection: Falta una conexión para la API "shared_sharepointonline" en el flujo.
```

## Definicion local generada

Los archivos generados quedan en:

```text
power_automate/solution/generated/
├── PA_S1_RegistrarPresupuestoAprobado.clientdata.json
└── PA_S1_RegistrarPresupuestoAprobado.definition.json
```

El script usado para crear/actualizar el workflow es:

```text
power_automate/scripts/upload_registrar_presupuesto_flow.py
```

## Comando usado

```bash
python power_automate/scripts/upload_registrar_presupuesto_flow.py --apply
```

## Pendientes para que funcione en produccion

1. Abrir el flujo en Power Automate.
2. Configurar o reparar la conexion SharePoint para `shared_sharepointonline`.
3. En la lista `Cola_Automatizacion_Proyectos`, actualizar la columna `EventType` para que permita el valor `presupuesto_aprobado`.
4. Guardar el flujo.
5. Activar el flujo.
6. Probar subiendo un `.xlsx` nuevo a:

```text
/Proyectos/Presupuestos Aprobados/
```

## Nota de esquema real detectado

La lista `Cola_Automatizacion_Proyectos` existe, pero sus nombres internos no son exactamente los del documento inicial:

- `FileName` tiene nombre interno `Filename`.
- `FileIdentifier` tiene nombre interno `FileID`.
- `EventType` es Choice y actualmente solo permite `Opción 1`.

La definicion del flujo ya usa los nombres internos reales `Filename` y `FileID`.
