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
- Reporta cuantos items `Pendiente` existen.
- No procesa ni modifica items todavia; este primer paso valida conectividad GitHub -> Microsoft Graph -> SharePoint.

## Variables configuradas

Estas variables quedan configuradas en GitHub:

```text
MS_GRAPH_CLIENT_ID
MS_GRAPH_TENANT_ID
SP_SITE_HOSTNAME
SP_SITE_PATH
SP_QUEUE_LIST_NAME
SP_QUEUE_LIST_ID
```

## Secret pendiente

GitHub Actions necesita autenticacion no interactiva. El login local por device code no funciona en runners.

Falta crear este secret en GitHub:

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
gh workflow run sistema1-poll-queue.yml
```

Ver ultimo run:

```bash
gh run list --workflow sistema1-poll-queue.yml --limit 5
gh run view <run-id> --log
```
