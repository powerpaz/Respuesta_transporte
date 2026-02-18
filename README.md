# Respuesta

Módulo web (cliente, sin servidor) para **validar modelamientos de transporte escolar** y **generar una respuesta en Word (DOCX)**.

## Uso
1. Publica la carpeta `Respuesta/` en GitHub Pages.
2. Abre el `index.html`.
3. Carga: Memorando (PDF) + Modelamiento (Excel .xlsx/.xls/.xlsm) + (opcional) referencia de beneficiarios + (opcional) plantilla DOCX oficial.
4. Ejecuta **Validar** y luego **Generar respuesta (Word)**.

## Plantilla DOCX
La plantilla debe contener placeholders:
- `{{memo_nro}}`, `{{memo_fecha}}`, `{{para}}`, `{{de}}`, `{{asunto}}`
- `{{resultado_general}}`, `{{resumen_texto}}`, `{{conclusion}}`, `{{firma}}`

La plantilla incluida está en `templates/plantilla_respuesta.docx`.
