# GitHub Actions: dispatcher de automatización SharePoint

## Workflow principal

- Nombre en Actions: `SharePoint Automation Dispatcher`
- Archivo: `.github/workflows/automation-dispatcher.yml`
- Python: `3.11`

El workflow hace checkout, instala `requirements.txt`, ejecuta un sanity check del
entrypoint y ejecuta en orden:

1. Sistema 1 para eventos `presupuesto_aprobado`.
2. Sistema 2 para asignación, tracking y eventos
   `gantt_working_modificado`.

Un error de Sistema 1 no impide que Sistema 2 haga su revisión, aunque el job
termina con error para que la falla siga siendo visible.

El destino de proyectos activos queda limitado a `/Proyectos/Proyectos Activos/`. El
procesador rechaza una configuración que apunte a
`/Proyectos/PROYECTOS TERMINADOS/`; esa carpeta no forma parte del procesamiento.

## Ejecución automática y manual

El cron interno de GitHub Actions está desactivado. El workflow conserva:

```yaml
workflow_dispatch:
```

El dispatcher externo previsto para invocarlo cada 15 minutos permanece
apagado durante la validación. Power Automate solamente registra eventos en
SharePoint.

Variables de correo:

- `NOTIFICATION_DELIVERY_MODE`: `test` o `live`;
- `NOTIFICATION_TEST_RECIPIENT`: destinatario único para pruebas;
- `ENGINEER_GUIDE_URL`: vínculo a la guía PDF de ingenieros y planta;
- `GANTT_ESCALATION_CC`: correos adicionales de escalamiento separados por
  punto y coma.

Para probarlo manualmente:

1. Abrir el repositorio en GitHub.
2. Ir a **Actions**.
3. Seleccionar **SharePoint Automation Dispatcher**.
4. Elegir **Run workflow**.
5. Seleccionar la rama que contiene el workflow y pulsar **Run workflow**.

También puede ejecutarse con GitHub CLI:

```bash
gh workflow run automation-dispatcher.yml --ref feature/power-automate-sistema1-flows -f top=50 -f max_items=5
```

Para reprocesar manualmente un item específico que ya esté marcado como
`Procesado`, agregar `-f item_id=<ID>`. Esta opción no se usa en el schedule.

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
`SP_QUEUE_LIST_NAME`, `SP_CONTROL_LIST_NAME` y los IDs opcionales de listas)
siguen configurándose como GitHub Actions Variables.

## Comando ejecutado

El entrypoint real del repositorio es:

```bash
python src/automation_dispatcher.py --top 50 --max-items 5
```

El dispatcher reutiliza el entrypoint funcional de Sistema 1 y añade Sistema 2.
No reescribe la lógica del generador.

## Resultados y errores comunes

- **Colas/listas sin trabajo:** termina exitosamente con contadores en cero.
- **Missing required GitHub Secrets:** falta al menos uno de los cuatro secrets.
  Configurarlo y volver a ejecutar manualmente.
- **No se pudo obtener token Microsoft Graph:** revisar tenant, client ID, client
  secret, permisos de aplicación y consentimiento de administrador en Entra ID.
- **Graph GET/PATCH/POST/PUT:** revisar el código HTTP mostrado, permisos
  `Sites.ReadWrite.All` y `Files.ReadWrite.All`, sitio y nombres/IDs de listas.
- **WARNING al crear columnas:** las operaciones normales continúan, pero
  ampliar una lista existente requiere `Sites.Manage.All` o crear esas columnas
  manualmente.
- **Error al instalar dependencias:** revisar la salida de `pip` y que las
  versiones de `requirements.txt` estén disponibles para Python 3.11.
- **Error de importación:** el sanity check falla antes de autenticarse y muestra
  el módulo que falta.

Un fallo de autenticación ocurre antes de leer o actualizar elementos. Un fallo
por elemento se guarda como `Error`; no se marca ese elemento como `Procesado`.

Para consultar ejecuciones:

```bash
gh run list --workflow automation-dispatcher.yml --limit 10
gh run view <run-id> --log
```

La ejecución manual permite validar el dispatcher sin depender de un
scheduler. Si se vuelve a activar un cron de GitHub, debe recordarse que no
tiene garantía de hora exacta.
