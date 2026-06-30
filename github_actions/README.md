# GitHub Actions

El workflow principal del Sistema 1 es `Process SharePoint Queue`, definido en:

```text
.github/workflows/sistema1-poll-queue.yml
```

La guía de configuración, ejecución manual, schedule, secrets y diagnóstico está
en [docs/github_actions.md](../docs/github_actions.md).

Power Automate registra eventos en SharePoint y no invoca GitHub Actions
directamente. GitHub Actions revisa la cola mediante su schedule de 15 minutos.
