# CHG Ascensores — Informe Membretado

Automatización que transforma el PDF crudo exportado desde MaintainX en un PDF con el membrete oficial de CHG Ascensores.

```
PDF crudo MaintainX  →  n8n  →  PDF membretado CHG
```

## Arquitectura

| Componente | Descripción |
|---|---|
| `src/maintainx_parser.py` | Parser PyMuPDF: extrae campos, fotos y firma del PDF crudo |
| `src/render_html.py` | Renderizado Jinja2: genera el HTML final con el membrete embebido |
| `template/informe.html` | Plantilla HTML/CSS genérica (secciones, fotos y firma dinámicos) |
| `template/assets/membrete.png` | PNG del membrete CHG (extraído del DOCX de referencia) |
| `workflows/chg-informe-membretado.json` | Workflow n8n completo, listo para importar |

## Requisitos del servidor n8n

```bash
pip install pymupdf jinja2
```

Gotenberg (render HTML → PDF):
```bash
docker run -d --name gotenberg -p 3000:3000 gotenberg/gotenberg:8
```

## Despliegue

1. Copiar `src/`, `template/` al servidor n8n en `/opt/n8n-resources/chg-informe-membretado/`
2. Importar `workflows/chg-informe-membretado.json` en n8n
3. Activar el workflow

Rutas configurables vía variables de entorno en el servidor n8n:
```
CHG_TEMPLATE_PATH=/opt/n8n-resources/chg-informe-membretado/template/informe.html
CHG_MEMBRETE_PATH=/opt/n8n-resources/chg-informe-membretado/template/assets/membrete.png
```

Si Gotenberg corre en host (no Docker network), cambiar la URL en el nodo **Generar PDF con Gotenberg**:
```
http://localhost:3000/forms/chromium/convert/html
```

## Intake (trigger)

El workflow usa un **Webhook** como trigger. CHG envía el PDF mediante `POST multipart/form-data`:

```bash
curl -X POST https://wn8nw.omegaconsultingai.com/webhook/chg-informe \
  -F "file=@orden_488.pdf" \
  --output informe_membretado.pdf
```

La respuesta es el PDF membretado como binario (`Content-Type: application/pdf`).

## Flujo de nodos

```
Recibir PDF via Webhook
  → Parsear PDF MaintainX          (Python / PyMuPDF)
    → Validar Parseo               (IF: validacion_ok == true)
        TRUE  → Renderizar HTML Membretado   (Python / Jinja2)
                  → Generar PDF con Gotenberg (HTTP Request)
                      → Responder con PDF
        FALSE → Preparar Mensaje de Error
                  → Responder con Error (HTTP 422)
```

## Validación del parser

El parser extrae el contador `X / Y Campos completados` del PDF y lo compara contra los campos parseados. Si no coinciden, o si aparece un valor fuera del conjunto conocido (`Bueno`, `Malo`, `No Aplica`, `condición estable`), el nodo **Validar Parseo** enruta a la rama de error.

## Riesgo conocido

Si MaintainX rediseña su plantilla de PDF, el parser puede desalinearse. El sistema detectará el problema (validación fallida) y notificará en lugar de generar un PDF silenciosamente incorrecto. Ajustar `NOISE_Y_TOP` / `NOISE_Y_BOTTOM` en `src/maintainx_parser.py` si MaintainX cambia el encabezado o pie.
