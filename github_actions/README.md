# GitHub Actions - Sistema 1

## Workflow

Archivo:

```text
.github/workflows/sistema1-poll-queue.yml
```

Funcion:

- Corre cada 15 minutos con cron `7,22,37,52 * * * *`.
- Tambien puede ejecutarse manualmente con `workflow_dispatch`.
- Lee la lista SharePoint `Cola_Automatizacion_Proyectos`.
- Procesa items `Pendiente` con `EventType = presupuesto_aprobado`.
- Marca el item como `Procesando`, incrementa `Intentos`, descarga el presupuesto y genera un Gantt WORKING.
- Crea la carpeta del proyecto bajo `/Proyectos/Proyectos Activos/`.
- Copia el presupuesto aprobado directamente dentro de la carpeta del proyecto.
- Sube el Gantt a la subcarpeta `gantts/`, sin crear carpeta `working`.
- Actualiza la cola como `Procesado` o `Error`.
- Si la lista `Control_Gantt_Asignaciones` esta disponible, crea un item de seguimiento con estado `Pendiente de asignación`.

## Variables configuradas

Estas variables quedan configuradas en GitHub:

```text
MS_GRAPH_CLIENT_ID
MS_GRAPH_TENANT_ID
SP_SITE_HOSTNAME
SP_SITE_PATH
SP_QUEUE_LIST_NAME
SP_QUEUE_LIST_ID
SP_CONTROL_LIST_NAME
SP_CONTROL_LIST_ID
SP_ACTIVE_PROJECTS_ROOT
```

`SP_CONTROL_LIST_NAME`, `SP_CONTROL_LIST_ID` y `SP_ACTIVE_PROJECTS_ROOT` son opcionales. Si no se configuran, el worker usa:

```text
Control_Gantt_Asignaciones
(resuelve la lista por nombre)
Proyectos/Proyectos Activos
```

## Secret requerido

GitHub Actions necesita autenticacion no interactiva. El login local por device code no funciona en runners.

Debe existir este secret en GitHub:

```text
MS_GRAPH_CLIENT_SECRET
```

Debe venir de Microsoft Entra ID para la app:

```text
Autosys Local Graph Prototype
```

Permisos requeridos esperados para la app:

```text
Sites.ReadWrite.All
Files.ReadWrite.All
```

Si la app solo va a leer la cola al inicio, `Sites.Read.All` podria bastar, pero el Sistema 1 completo necesitara escribir estados, descargar presupuestos, crear carpetas/subir Gantts y crear registros en listas.

## Comandos utiles

Configurar el secret:

```bash
gh secret set MS_GRAPH_CLIENT_SECRET
```

Ejecutar manualmente:

```bash
gh workflow run sistema1-poll-queue.yml -f top=50 -f max_items=5
```

Ver ultimo run:

```bash
gh run list --workflow sistema1-poll-queue.yml --limit 5
gh run view <run-id> --log
```

## Comportamiento del Gantt generado

El worker genera un archivo `.xlsx` con hoja `Gantt_Diario`.

Columnas principales:

- `ID_Gantt`
- `Fuente_Fila`
- `Tipo_Linea_Autosys`
- `Item`
- `CC`
- `Actividad`
- `Cantidad`
- `Unidad`
- `Monto_Referencia`
- `Fecha_Inicio`
- `Fecha_Fin`
- `Estado_Planificacion`
- `Requiere_Revision`
- `Motivo_Revision`

Las fechas por actividad quedan vacias. El ingeniero residente las completa. El cronograma diario se pinta automaticamente con barras azules cuando `Fecha_Inicio` y `Fecha_Fin` intersectan los dias del calendario.

El builder no usa OpenAI todavia en GitHub Actions. La clasificacion inicial es conservadora y marca para revision filas ambiguas, indirectos, totales o filas con columnas de unidad sospechosas.
