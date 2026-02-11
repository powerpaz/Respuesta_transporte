# Respuesta (Transporte Escolar) — GitHub Pages

Módulo cliente (sin backend) para:
- Leer un Memorando (PDF) y extraer datos básicos
- Validar un Excel de modelamiento (estructura + beneficiarios + reglas básicas)
- Generar una respuesta editable en Word (DOCX) usando una plantilla

## Uso
1. Abrir `index.html` en el navegador (o subir el repo a GitHub Pages)
2. Cargar:
   - Memorando (PDF)
   - Modelamiento (Excel)
   - (Opcional) Referencia de beneficiarios (Excel con columnas AMIE, INICIAL, EGB, BACHILLERATO)
   - (Opcional) Plantilla DOCX con placeholders

## Plantillas
Incluye:
- `templates/plantilla_respuesta.docx` (simple, lista para docxtemplater)
- `templates/plantilla_oficial_usuario.docx` (la tuya, por si deseas adaptarla a placeholders)

## Nota
Todo corre en local (no se suben archivos a ningún servidor).
