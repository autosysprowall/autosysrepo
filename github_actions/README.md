# GitHub Actions

El único workflow principal es `SharePoint Automation Dispatcher`, definido en:

```text
.github/workflows/automation-dispatcher.yml
```

La guía de configuración, ejecución manual, schedule, secrets y diagnóstico está
en [docs/github_actions.md](../docs/github_actions.md).

Power Automate registra eventos en SharePoint y envía notificaciones, pero no
invoca GitHub Actions directamente. GitHub Actions ejecuta Sistema 1 y Sistema 2
mediante su schedule por hora (`17 * * * *`) o manualmente.
