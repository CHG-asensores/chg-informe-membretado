"""
Servidor web para generar informes membretados CHG Ascensores.
Reemplaza el workflow n8n — corre localmente sin Docker.

Uso:
    pip install flask pymupdf jinja2 playwright
    python -m playwright install chromium
    python app.py
"""
import base64
import io
import os
import sys
from pathlib import Path

import fitz  # PyMuPDF para overlay del membrete
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime
from flask import Flask, jsonify, render_template_string, request, send_file
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent / "src"))
from maintainx_parser import parse_pdf
from render_html import render_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "template" / "informe.html"
MEMBRETE_PATH = BASE_DIR / "template" / "assets" / "membrete.png"


def html_to_pdf(html_bytes: bytes) -> bytes:
    """Renderiza HTML a PDF usando Chromium headless."""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.set_content(html_bytes.decode("utf-8"), wait_until="networkidle")
        pdf = page.pdf(
            format="A4",
            margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
            print_background=True,
            prefer_css_page_size=True,
        )
        browser.close()
    return pdf


def apply_letterhead(pdf_bytes: bytes, membrete_path: str) -> bytes:
    """Sobrepone el membrete como fondo en cada página del PDF.

    insert_image con overlay=False lo coloca DETRÁS del contenido existente,
    así el texto y elementos del informe quedan por encima del membrete.
    Garantiza que el membrete se repita idéntico en TODAS las páginas.
    """
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    with open(membrete_path, "rb") as f:
        membrete_bytes = f.read()
    for page in src:
        page.insert_image(page.rect, stream=membrete_bytes, overlay=False)
        # Texto "Generado para CHG Ascensores" en bottom-left de cada página
        page.insert_text(
            (40, page.rect.height - 18),
            "Generado para CHG Ascensores",
            fontsize=7,
            color=(0.42, 0.45, 0.50),
        )
    out = src.tobytes()
    src.close()
    return out


def generar_docx_con_observaciones(data, observaciones_extra):
    doc = DocxDocument()
    seccion = doc.sections[0]
    seccion.page_height = Cm(29.7)
    seccion.page_width = Cm(21.0)
    seccion.top_margin = Cm(3.2)
    seccion.bottom_margin = Cm(2.8)
    seccion.left_margin = Cm(1.4)
    seccion.right_margin = Cm(1.4)
    AZUL = RGBColor(0x1A, 0x1A, 0x2E)
    GRIS = RGBColor(0x6B, 0x72, 0x80)
    VERDE = RGBColor(0x16, 0xA3, 0x4A)
    ROJO = RGBColor(0xC0, 0x39, 0x2B)
    LOGO = BASE_DIR / "template" / "assets" / "membrete.png"
    hdr = doc.sections[0].header
    ph = hdr.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO.exists():
        ph.add_run().add_picture(str(LOGO), width=Cm(8))
    ftr = doc.sections[0].footer
    pf = ftr.paragraphs[0]
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = pf.add_run("CHG Ascensores | comercial@chgascensor.com | (01) 627-9422 | Av. La Encalada N110 - Surco")
    rf.font.size = Pt(7)
    rf.font.color.rgb = GRIS
    meta = data.get("meta", {})
    secciones = data.get("secciones", [])
    fotos = data.get("fotos", [])
    obs = data.get("observaciones", {})
    firma = data.get("firma", {})
    pt = doc.add_paragraph()
    pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = pt.add_run(meta.get("titulo", "INFORME DE MANTENIMIENTO"))
    rt.bold = True
    rt.font.size = Pt(14)
    rt.font.color.rgb = AZUL
    doc.add_paragraph()
    tabla = doc.add_table(rows=0, cols=4)
    tabla.style = "Table Grid"
    def add_meta_row(t, l1, v1, l2, v2):
        fila = t.add_row()
        for ci, txt, bold in [(0,l1,True),(1,v1,False),(2,l2,True),(3,v2,False)]:
            p = fila.cells[ci].paragraphs[0]
            r = p.add_run(str(txt) if txt else "-")
            r.bold = bold
            r.font.size = Pt(9)
            if bold:
                r.font.color.rgb = GRIS
    add_meta_row(tabla, "ESTADO", meta.get("estado","-"), "FECHA VENCIMIENTO", meta.get("fecha_vencimiento","-"))
    add_meta_row(tabla, "TIEMPO EST.", meta.get("tiempo_estimado","-"), "TIPO TRABAJO", meta.get("tipo_trabajo","-"))
    add_meta_row(tabla, "ASIGNADOS", ", ".join(meta.get("asignados",[])) or "-", "CATEGORIAS", ", ".join(meta.get("categorias",[])) or "-")
    add_meta_row(tabla, "UBICACION", meta.get("ubicacion","-"), "ACTIVO", meta.get("activo","-"))
    add_meta_row(tabla, "INGRESO", meta.get("fecha_hora_ingreso","-"), "CAMPOS", meta.get("campos_completados","-"))
    doc.add_paragraph()
    if meta.get("procedimiento"):
        pp = doc.add_paragraph(meta["procedimiento"])
        pp.runs[0].bold = True
        pp.runs[0].font.size = Pt(11)
        pp.runs[0].font.color.rgb = AZUL
        doc.add_paragraph()
    for sec in secciones:
        ps = doc.add_paragraph(sec["nombre"])
        ps.runs[0].bold = True
        ps.runs[0].font.size = Pt(10)
        ps.runs[0].font.color.rgb = AZUL
        ts = doc.add_table(rows=0, cols=2)
        ts.style = "Table Grid"
        for campo in sec.get("campos", []):
            e = campo.get("etiqueta", "")
            v = campo.get("valor", "-")
            fila = ts.add_row()
            fila.cells[0].width = Cm(9)
            fila.cells[1].width = Cm(7)
            fila.cells[0].paragraphs[0].add_run(e).font.size = Pt(9)
            rv = fila.cells[1].paragraphs[0].add_run(str(v))
            rv.font.size = Pt(9)
            if v == "Malo":
                rv.font.color.rgb = ROJO
            elif v == "Bueno":
                rv.font.color.rgb = VERDE
        doc.add_paragraph()
    if obs.get("evaluacion_final") or obs.get("para_cliente"):
        po = doc.add_paragraph("OBSERVACIONES Y RECOMENDACIONES")
        po.runs[0].bold = True
        po.runs[0].font.size = Pt(11)
        po.runs[0].font.color.rgb = AZUL
        to = doc.add_table(rows=0, cols=2)
        to.style = "Table Grid"
        if obs.get("evaluacion_final"):
            fila = to.add_row()
            fila.cells[0].paragraphs[0].add_run("Evaluacion final:").font.size = Pt(9)
            fila.cells[1].paragraphs[0].add_run(obs["evaluacion_final"]).font.size = Pt(9)
        if obs.get("para_cliente"):
            fila = to.add_row()
            fila.cells[0].paragraphs[0].add_run("Para cliente:").font.size = Pt(9)
            fila.cells[1].paragraphs[0].add_run(obs["para_cliente"]).font.size = Pt(9)
        if obs.get("para_chg") and obs["para_chg"] != "-":
            fila = to.add_row()
            fila.cells[0].paragraphs[0].add_run("Para CHG:").font.size = Pt(9)
            fila.cells[1].paragraphs[0].add_run(obs["para_chg"]).font.size = Pt(9)
        doc.add_paragraph()
    if firma.get("texto"):
        pfi = doc.add_paragraph("FIRMA DEL CLIENTE")
        pfi.runs[0].bold = True
        pfi.runs[0].font.size = Pt(10)
        pfi.runs[0].font.color.rgb = AZUL
        if firma.get("data_base64"):
            import base64, io as _io
            img_bytes = base64.b64decode(firma["data_base64"])
            doc.add_picture(_io.BytesIO(img_bytes), width=Cm(5))
        pft = doc.add_paragraph(firma["texto"])
        pft.runs[0].font.size = Pt(9)
        pft.runs[0].italic = True
        if meta.get("fecha_hora_salida"):
            ps2 = doc.add_paragraph("Fecha y Hora de salida: " + meta["fecha_hora_salida"])
            ps2.runs[0].font.size = Pt(9)
        doc.add_paragraph()
    if observaciones_extra:
        poe = doc.add_paragraph("OBSERVACIONES ADICIONALES")
        poe.runs[0].bold = True
        poe.runs[0].font.size = Pt(11)
        poe.runs[0].font.color.rgb = ROJO
        pob = doc.add_paragraph(observaciones_extra)
        pob.runs[0].font.size = Pt(10)
        pob.runs[0].italic = True
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

@app.route("/")
def index():
    return render_template_string(UI)


@app.route("/generate", methods=["POST"])
def generate():
    if "pdf" not in request.files:
        return jsonify({"error": "No se recibió ningún archivo PDF."}), 400

    pdf_bytes = request.files["pdf"].read()
    observaciones = request.form.get("observaciones", "").strip()
    if not pdf_bytes:
        return jsonify({"error": "El archivo está vacío."}), 400

    # 1. Parsear
    try:
        data = parse_pdf(pdf_bytes)
    except Exception as exc:
        return jsonify({"error": f"Error al parsear el PDF: {exc}"}), 422

    # 2. Renderizar HTML
    try:
        html_b64 = render_html(data, str(TEMPLATE_PATH), str(MEMBRETE_PATH))
        html_bytes = base64.b64decode(html_b64)
    except Exception as exc:
        return jsonify({"error": f"Error al renderizar el HTML: {exc}"}), 500

    # 3. Convertir a PDF vía Playwright (Chromium) — contenido sin membrete
    try:
        pdf_raw = html_to_pdf(html_bytes)
    except Exception as exc:
        return jsonify({"error": f"Error al generar PDF con Playwright: {exc}"}), 500

    # 4. Aplicar membrete como overlay en cada página (PyMuPDF)
    try:
        pdf_out = apply_letterhead(pdf_raw, str(MEMBRETE_PATH))
    except Exception as exc:
        return jsonify({"error": f"Error al aplicar membrete: {exc}"}), 500

    # 4. Guardar PDF temporalmente y servir URL de descarga
    numero   = data.get("meta", {}).get("numero_orden", "informe")
    filename = f"informe-{numero}.pdf"

    # Guardamos en memoria con un ID de sesión simple
    import hashlib, time
    file_id = hashlib.md5(f"{numero}{time.time()}".encode()).hexdigest()[:12]
    _pdf_store[file_id] = (filename, pdf_out)

    val = data.get("validacion", {})
    warnings = val.get("valores_fuera_de_conjunto", [])

    return jsonify({
        "ok":           True,
        "warning":      bool(warnings),
        "meta":         data.get("meta", {}),
        "warnings":     warnings,
        "filename":     filename,
        "download_url": f"/download/{file_id}",
    })


_pdf_store: dict = {}
@app.route('/generate-word', methods=['POST'])
def generate_word():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No se recibio PDF'}), 400
    pdf_bytes = request.files['pdf'].read()
    observaciones = request.form.get('observaciones', '').strip()
    if not pdf_bytes:
        return jsonify({'error': 'Archivo vacio'}), 400
    try:
        data = parse_pdf(pdf_bytes)
    except Exception as exc:
        return jsonify({'error': f'Error al parsear: {exc}'}), 422
    try:
        docx_bytes = generar_docx_con_observaciones(data, observaciones)
    except Exception as exc:
        return jsonify({'error': f'Error al generar Word: {exc}'}), 500
    numero = data.get('meta', {}).get('numero_orden', 'informe')
    return send_file(io.BytesIO(docx_bytes), mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', as_attachment=True, download_name=f'informe-{numero}-observaciones.docx')



@app.route("/download/<file_id>")
def download(file_id):
    if file_id not in _pdf_store:
        return "Archivo no encontrado o expirado.", 404
    filename, pdf_bytes = _pdf_store[file_id]
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("\n  CHG Informe Membretado")
    print("  -------------------------------------")
    print(f"  Servidor:  http://localhost:{port}")
    print("  Motor PDF: Playwright (Chromium)")
    print("  -------------------------------------\n")
    app.run(host="0.0.0.0", port=port, debug=True)
