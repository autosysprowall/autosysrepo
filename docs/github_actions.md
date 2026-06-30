# GitHub Actions: procesador de cola de SharePoint

## Workflow principal

- Nombre en Actions: `Process SharePoint Queue`
- Archivo: `.github/workflows/sistema1-poll-queue.yml`
- Rama predeterminada del repositorio: `feature/power-automate-sistema1-flows`
- Python: `3.11`

El workflow hace checkout, instala `requirements.txt`, ejecuta un sanity check del
entrypoint y procesa los eventos `Pendiente` con
`EventType = presupuesto_aprobado` de `Cola_Automatizacion_Proyectos`.

El destino de proyectos activos queda limitado a `/Proyectos/02_Activos/`. El
procesador rechaza una configuración que apunte a
`/Proyectos/PROYECTOS TERMINADOS/`; esa carpeta no forma parte del procesamiento.

## Ejecución automática y manual

GitHub Actions evalúa la cola cada 15 minutos con:

```yaml
schedule:
  - cron: "7,22,37,52 * * * *"
```

GitHub puede demorar algunos minutos una ejecución programada. Power Automate no
dispara este workflow directamente: solamente registra eventos en SharePoint.

Para probarlo manualmente:

1. Abrir el repositorio en GitHub.
2. Ir a **Actions**.
3. Seleccionar **Process SharePoint Queue**.
4. Elegir **Run workflow**.
5. Seleccionar la rama que contiene el workflow y pulsar **Run workflow**.

También puede ejecutarse con GitHub CLI:

```bash
gh workflow run sistema1-poll-queue.yml --ref feature/power-automate-sistema1-flows -f top=50 -f max_items=5
```

## Secrets requeridos

Configurar en **Settings > Secrets and variables > Actions**:

```text
MS_TENANT_ID
MS_CLIENT_ID
MS_CLIENT_SECRET
OPENAI_API_KEY
```

El workflow nunca imprime sus valores. Durante la transición acepta el secret
heredado `MS_GRAPH_CLIENT_SECRET` como respaldo de `MS_CLIENT_SECRET`; el nombre
canónico nuevo debe configurarse y el heredado puede retirarse después.

Las variables no secretas de SharePoint (`SP_SITE_HOSTNAME`, `SP_SITE_PATH`,
`SP_QUEUE_LIST_NAME` y los IDs opcionales de listas) siguen configurándose como
GitHub Actions Variables.

## Comando ejecutado

El entrypoint real del repositorio es:

```bash
python scripts/sistema1_poll_queue.py --top 50 --max-items 5 --process
```

No se creó un entrypoint paralelo ni se cambió la lógica del generador de Gantts.

## Resultados y errores comunes

- **Cola vacía:** termina exitosamente y registra
  `No pending events found. No hay eventos pendientes para procesar.`
- **Missing required GitHub Secrets:** falta al menos uno de los cuatro secrets.
  Configurarlo y volver a ejecutar manualmente.
- **No se pudo obtener token Microsoft Graph:** revisar tenant, client ID, client
  secret, permisos de aplicación y consentimiento de administrador en Entra ID.
- **Graph GET/PATCH/POST/PUT:** revisar el código HTTP mostrado, permisos
  `Sites.ReadWrite.All` y `Files.ReadWrite.All`, sitio y nombres/IDs de listas.
- **Error al instalar dependencias:** revisar la salida de `pip` y que las
  versiones de `requirements.txt` estén disponibles para Python 3.11.
- **Error de importación:** el sanity check falla antes de autenticarse y muestra
  el módulo que falta.

Un fallo de autenticación ocurre antes de leer o actualizar elementos. Un fallo
por elemento se guarda como `Error`; no se marca ese elemento como `Procesado`.

Para consultar ejecuciones:

```bash
gh run list --workflow sistema1-poll-queue.yml --limit 10
gh run view <run-id> --log
```
