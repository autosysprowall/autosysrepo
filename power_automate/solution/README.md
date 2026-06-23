# Solucion Power Automate

No se genero un paquete importable en esta ejecucion porque, aunque Power Platform CLI (`pac`) ya esta instalado localmente, no hay un perfil autenticado contra un environment de Power Platform.

Esta carpeta queda reservada para colocar una solucion exportada desde Power Automate o Power Platform CLI cuando se construyan los flujos en el tenant.

## Estado actual

- Paquete importable: no generado.
- CLI: `pac` instalado localmente.
- Version validada: `2.8.1`.
- Motivo pendiente: falta ejecutar `pac auth create` y conectar con el environment correcto.
- Alternativa entregada: especificacion manual exacta nodo por nodo en `../flows/`.

## Comandos esperados si se usa Power Platform CLI despues

```bash
pac auth create
pac solution init --publisher-name Autosys --publisher-prefix autosys
pac solution add-reference --path <ruta-del-flujo-o-solucion>
pac solution pack --zipfile Autosys_Sistema1_PowerAutomate.zip --folder .
```

Los comandos finales pueden variar segun si los flujos se crean primero en una solucion dentro del tenant y luego se exportan.
