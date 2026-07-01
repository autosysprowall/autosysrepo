# Devolución de presupuestos inválidos

Este circuito se ejecuta antes de generar un Gantt. No asigna ingenieros, no
crea versiones y no toca `Proyectos Terminados`.

## Compuerta de calidad

El presupuesto se acepta únicamente cuando:

- contiene un bloque identificado como `Presupuesto General`;
- existe una sola tabla comercial confiable;
- `Actividad` contiene principalmente texto descriptivo;
- `Cantidad` contiene principalmente números;
- `Unidad` contiene unidades reconocibles;
- `Ítem`, cuando existe, contiene códigos jerárquicos como `1.1` o `1.1.1`;
- las columnas de costo contienen valores numéricos y no ocupan el rol de otra
  columna.

Las reglas determinísticas validan encabezado y contenido. El LLM solo puede
proponer un mapping entre columnas existentes; la validación de contenido debe
aprobarlo antes de usarlo.

## Causas de rechazo

| Código | Significado |
|---|---|
| `SIN_PRESUPUESTO_GENERAL` | El libro parece dividido o incompleto. |
| `ENCABEZADOS_REPETIDOS` | Hay más de una tabla o los encabezados están desordenados. |
| `ENCABEZADOS_DESORDENADOS` | Una columna recibió dos roles incompatibles. |
| `CONTENIDO_DE_COLUMNAS_INVALIDO` | Los valores no corresponden al encabezado detectado. |
| `COLUMNAS_DE_COSTO_AUSENTES` | No existe una columna confiable de costo o precio. |
| `CONTENIDO_DE_COSTOS_INVALIDO` | Una columna financiera contiene datos incompatibles. |
| `ARCHIVO_ROTO_CAUSA_DESCONOCIDA` | Excel no puede abrir el archivo y no existe una causa estructural más específica. |

Un rechazo no produce ni sube ningún Gantt. El item original de
`Cola_Automatizacion_Proyectos` cambia a:

- `EventType` conserva `presupuesto_aprobado`;
- `Estado = RequiereRevision`;
- `UltimoError = <código y explicación>`;
- `Notas` comenzando por `DEVOLVER_PRESUPUESTO`.

El item conserva `FileIdentifier`, `FolderPath`, `FileLink` y `FileName` para
que Power Automate pueda obtener el archivo original.

## Flujo de Power Automate

Nombre: `PA_S1_DevolverPresupuestoInvalido`.

El flujo fue creado el 30 de junio de 2026 en el ambiente `Prowall (default)`,
con ID `56a4ec6a-5fd0-4c74-b9cb-7c65b0375928`. Quedó apagado deliberadamente.
Debe permanecer en modo `test` hasta validar el texto final. La guía PDF se
omite mientras no exista.

La conexión Outlook usada es `auto.sys@prowallpanama.com`; el remitente efectivo
será esa cuenta mientras no se configure `From (Send as)`.

1. Trigger SharePoint **When an item is created or modified** sobre
   `Cola_Automatizacion_Proyectos`.
2. Activar concurrencia y establecer grado `1`.
3. Condición de trigger:

   ```text
   @and(
     equals(triggerBody()?['Estado'],'RequiereRevision'),
     startsWith(triggerBody()?['Notas'],'DEVOLVER_PRESUPUESTO')
   )
   ```

4. **Get file content** usando `FileIdentifier`. El flujo debe fallar antes de
   enviar si no puede obtener el contenido.
5. **Send an email (V2)**:
   - To live: `CreatedByEmail`;
   - CC live: `jaime.madrid@prowallpanama.com`;
   - To test: `auto.sys@prowallpanama.com`;
   - CC test: vacío;
   - Subject: `Presupuesto No Válido Proyecto <NombreProyecto>`;
   - Body: texto aprobado, `UltimoError`, nombre, enlace y guía de presupuesto;
   - Attachment Name: `FileName`;
   - Attachment Content: salida de **Get file content**.
6. Después de un envío exitoso, ejecutar **Update item** antes de borrar nada:
   - `Estado = Procesado`;
   - `FechaProcesado = utcNow()`;
   - `Notas = DEVOLUCION_ENVIADA`.
7. Eliminar el archivo fuente usando el mismo `FileIdentifier`.
8. Eliminar el item de `Cola_Automatizacion_Proyectos`.

La definición guardada contiene estas cinco acciones en ese orden:

1. `Get file content`;
2. `Send an email (V2)` con el presupuesto adjunto;
3. `Update item`;
4. `Delete file`;
5. `Delete item`.

El diseñador validó la definición con cero errores. Muestra una advertencia
esperada de posible ciclo porque el flujo actualiza la misma lista que lo
dispara; la condición `DEVOLVER_PRESUPUESTO` y el cambio a `Procesado` cortan
ese ciclo.

El orden es obligatorio. Marcar el item como `Procesado` antes de eliminar el
archivo impide un segundo correo si la eliminación final del item falla. Si el
correo falla, no se elimina ni el archivo ni el item y se conserva la causa
para reintento.

## Política de la cola

Los eventos exitosos normales permanecen con estado terminal para conservar
idempotencia. Los eventos rechazados pueden eliminarse únicamente después de:

1. adjuntar y enviar el presupuesto;
2. marcar el item como `Procesado`;
3. retirar el archivo de `Presupuestos Aprobados`.

Así la cola no contiene un rechazo finalizado, pero tampoco pierde el control
de duplicados durante el proceso.

Las expresiones exactas para alternar `test/live` están en
`docs/email_templates_and_delivery_modes.md`.
