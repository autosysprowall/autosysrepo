# Solucion Power Automate

## Estado actual

El flujo principal del Sistema 1 está desplegado y activo en Power Automate:

- Nombre: `PA_S1_RegistrarPresupuestoAprobado`
- Workflow ID: `02b348f0-9c70-f111-ab0f-000d3a34454c`
- Ambiente: `Prowall (default)`
- Estado verificado: `Started`
- Conector usado: SharePoint (`shared_sharepointonline`)

La conexión SharePoint ya está asociada. La sección siguiente se conserva como
referencia para regenerar la definición, no como un pendiente de deployment.

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

## Prueba de producción

Probar subiendo un `.xlsx` nuevo a:

```text
/Proyectos/Presupuestos Aprobados/
```

## Nota de esquema real detectado

La lista `Cola_Automatizacion_Proyectos` existe, pero sus nombres internos no son exactamente los del documento inicial:

- `FileName` tiene nombre interno `Filename`.
- `FileIdentifier` tiene nombre interno `FileID`.
- `EventType` es Choice y actualmente solo permite `Opción 1`.

La definicion del flujo ya usa los nombres internos reales `Filename` y `FileID`.
