"""
Renderiza la plantilla Jinja2 con el JSON parseado.
Se importa desde el nodo Code (Python) del workflow n8n.

Requisitos: pip install jinja2
"""
import base64
from jinja2 import Environment, BaseLoader


def render_html(data, template_path, membrete_path):
    """
    Retorna el HTML renderizado como string base64 (listo para enviar a Gotenberg).
    """
    with open(membrete_path, "rb") as f:
        membrete_b64 = base64.b64encode(f.read()).decode("utf-8")

    with open(template_path, "r", encoding="utf-8") as f:
        template_str = f.read()

    env      = Environment(loader=BaseLoader())
    template = env.from_string(template_str)
    html_out = template.render(data=data, membrete_base64=membrete_b64)

    return base64.b64encode(html_out.encode("utf-8")).decode("utf-8")
