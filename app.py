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


import io
import base64
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """
    Función auxiliar para establecer los bordes de una celda específica.
    kwargs acepta top, bottom, start, end con diccionarios de propiedades (ej. {'val': 'single', 'sz': '4', 'color': 'auto'})
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key, val in edge_data.items():
                element.set(qn('w:{}'.format(key)), str(val))

def generar_docx_con_observaciones(data, observaciones_extra):
    doc = Document()
    seccion = doc.sections[0]
    
    # Configuración de página A4 y márgenes
    seccion.page_height = Cm(29.7)
    seccion.page_width = Cm(21.0)
    seccion.top_margin = Cm(2.5)
    seccion.bottom_margin = Cm(2.5)
    seccion.left_margin = Cm(1.5)
    seccion.right_margin = Cm(1.5)
    
    # Colores corporativos (basados en el PDF)
    AZUL_OSCURO = RGBColor(0x1B, 0x36, 0x5D) # Azul institucional
    GRIS_CLARO = RGBColor(0x6B, 0x72, 0x80)
    NEGRO = RGBColor(0x00, 0x00, 0x00)

    # 1. ENCABEZADO Y PIE DE PÁGINA
    LOGO = BASE_DIR / "template" / "assets" / "membrete.png"

    hdr = seccion.header
    ph = hdr.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if LOGO.exists():
        ph.add_run().add_picture(str(LOGO), width=Cm(4))
    else:
        rh = ph.add_run("CHG ASCENSORES")
        rh.bold = True
        rh.font.size = Pt(14)
        rh.font.color.rgb = AZUL_OSCURO

    ftr = seccion.footer
    pf = ftr.paragraphs[0]
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rf = pf.add_run("Generado para CHG Ascensores\n947234073 | (01) 627-9422 | comercial@chgascensor.com\nAv. La Encalada N°110-Surco")
    rf.font.size = Pt(8)
    rf.font.color.rgb = GRIS_CLARO

    # Extracción de datos
    meta = data.get("meta", {})
    secciones = data.get("secciones", [])
    fotos = data.get("fotos", [])
    obs = data.get("observaciones", {})
    firma = data.get("firma", {})

    # 2. TABLA DE METADATOS (Diseño idéntico al PDF)
    tabla_meta = doc.add_table(rows=4, cols=4)
    tabla_meta.autofit = False
    
    # Anchos de columna proporcionales
    col_widths = [Cm(4), Cm(5.5), Cm(4.5), Cm(5.5)]
    for row in tabla_meta.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    def fill_cell(cell, titulo, valor, es_hecho=False):
        p = cell.paragraphs[0]
        p.style.font.size = Pt(8)
        run_titulo = p.add_run(f"{titulo}\n")
        run_titulo.bold = True
        run_titulo.font.size = Pt(8)
        run_titulo.font.color.rgb = GRIS_CLARO
        
        if es_hecho and valor.lower() == "hecho":
            run_valor = p.add_run("☑ Hecho")
        else:
            run_valor = p.add_run(str(valor) if valor else "-")
        run_valor.font.size = Pt(10)
        run_valor.font.color.rgb = NEGRO
        run_valor.bold = True
        
        set_cell_border(cell, bottom={"val": "single", "sz": "4", "color": "D3D3D3"})

    fill_cell(tabla_meta.cell(0,0), "ESTADO", meta.get("estado","-"), es_hecho=True)
    fill_cell(tabla_meta.cell(0,1), "FECHA DE VENCIMIENTO", meta.get("fecha_vencimiento","-"))
    fill_cell(tabla_meta.cell(0,2), "TIEMPO ESTIMADO", meta.get("tiempo_estimado","-"))
    fill_cell(tabla_meta.cell(0,3), "TIPO DE TRABAJO", meta.get("tipo_trabajo","-"))
    
    fill_cell(tabla_meta.cell(1,0), "ASIGNADOS", ", ".join(meta.get("asignados",[])) or "-")
    fill_cell(tabla_meta.cell(1,1), "CATEGORÍAS", ", ".join(meta.get("categorias",[])) or "-")
    fill_cell(tabla_meta.cell(1,2), "UBICACIÓN", meta.get("ubicacion","-"))
    fill_cell(tabla_meta.cell(1,3), "ACTIVO", f"{meta.get('activo','-')}")

    # Fusionar celdas para el procedimiento si existe (como en el PDF)
    cell_proc_title = tabla_meta.cell(2,0)
    cell_proc_title.merge(tabla_meta.cell(2,3))
    fill_cell(cell_proc_title, "PROCEDIMIENTO", meta.get("procedimiento", "Ascensor") + "\n☑ INFORME DE MANTENIMIENTO PREVENTIVO")
    
    doc.add_paragraph() # Espacio

    # 3. SECCIONES Y CHECKLISTS (Con casillas de verificación ☑)
    for sec in secciones:
        ps = doc.add_paragraph()
        run_sec = ps.add_run(sec["nombre"].upper())
        run_sec.bold = True
        run_sec.font.size = Pt(11)
        run_sec.font.color.rgb = AZUL_OSCURO
        
        # Crear tabla invisible para alinear Etiqueta y ☑ Valor
        ts = doc.add_table(rows=0, cols=2)
        ts.autofit = False
        for campo in sec.get("campos", []):
            e = campo.get("etiqueta", "")
            v = campo.get("valor", "-")
            
            fila = ts.add_row()
            fila.cells[0].width = Cm(11)
            fila.cells[1].width = Cm(8)
            
            p_etiqueta = fila.cells[0].paragraphs[0]
            p_etiqueta.add_run(f"{e}:").font.size = Pt(9.5)
            
            p_valor = fila.cells[1].paragraphs[0]
            rv = p_valor.add_run(f"☑ {v}")
            rv.font.size = Pt(9.5)
            rv.bold = True
            
            # Colores semánticos sutiles pero legibles
            if v == "Malo":
                rv.font.color.rgb = RGBColor(0xD3, 0x2F, 0x2F) # Rojo
            elif v == "Bueno":
                rv.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32) # Verde oscuro
            else:
                rv.font.color.rgb = GRIS_CLARO # No Aplica
                
            # Línea sutil divisoria
            set_cell_border(fila.cells[0], bottom={"val": "single", "sz": "2", "color": "E5E7EB"})
            set_cell_border(fila.cells[1], bottom={"val": "single", "sz": "2", "color": "E5E7EB"})

        doc.add_paragraph() # Espacio entre secciones

    # 4. OBSERVACIONES Y RECOMENDACIONES
    po = doc.add_paragraph()
    run_obs = po.add_run("Observaciones y Recomendaciones")
    run_obs.bold = True
    run_obs.font.size = Pt(12)
    run_obs.font.color.rgb = AZUL_OSCURO

    # Unir observaciones del JSON base y las extra
    todas_observaciones = []
    if obs.get("para_cliente"):
        todas_observaciones.append(obs["para_cliente"])
    if obs.get("para_chg") and obs["para_chg"] != "-":
        todas_observaciones.append(obs["para_chg"])
        
    if isinstance(observaciones_extra, list):
        todas_observaciones.extend(observaciones_extra)
    elif isinstance(observaciones_extra, str) and observaciones_extra:
        todas_observaciones.append(observaciones_extra)

    for observacion in todas_observaciones:
        p_item = doc.add_paragraph(style='List Bullet')
        r_item = p_item.add_run(observacion)
        r_item.font.size = Pt(10)
        
    if obs.get("evaluacion_final"):
        doc.add_paragraph()
        p_eval = doc.add_paragraph()
        r_eval_t = p_eval.add_run("Evaluación final: ")
        r_eval_t.bold = True
        r_eval_t.font.size = Pt(10)
        r_eval_v = p_eval.add_run(obs["evaluacion_final"])
        r_eval_v.font.size = Pt(10)

    # 5. FIRMA
    doc.add_paragraph()
    pfi = doc.add_paragraph()
    run_firma = pfi.add_run("Firma del cliente:")
    run_firma.bold = True
    run_firma.font.size = Pt(11)
    run_firma.font.color.rgb = AZUL_OSCURO

    if firma.get("data_base64"):
        img_bytes = base64.b64decode(firma["data_base64"])
        doc.add_picture(io.BytesIO(img_bytes), width=Cm(6))
        
    if firma.get("texto"):
        pft = doc.add_paragraph(firma["texto"])
        pft.runs[0].font.size = Pt(9)
        pft.runs[0].font.color.rgb = GRIS_CLARO

    if meta.get("fecha_hora_salida"):
        ps2 = doc.add_paragraph(f"Fecha y Hora de salida: {meta['fecha_hora_salida']}")
        ps2.runs[0].font.size = Pt(9)
        ps2.runs[0].font.color.rgb = GRIS_CLARO

    # 6. FOTOGRAFÍAS (Forzamos salto de página si hay fotos para que se vean como anexos)
    if fotos:
        doc.add_page_break()
        for foto in fotos:
            p_foto_title = doc.add_paragraph()
            r_ft_title = p_foto_title.add_run(foto.get("titulo", "Fotografía:"))
            r_ft_title.bold = True
            r_ft_title.font.size = Pt(11)
            r_ft_title.font.color.rgb = AZUL_OSCURO
            
            if foto.get("data_base64"):
                try:
                    img_bytes = base64.b64decode(foto["data_base64"])
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # Mantenemos las fotos grandes como en el PDF
                    p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(15))
                except Exception as e:
                    doc.add_paragraph(f"[Error al cargar imagen: {str(e)}]")
                    
            if foto.get("descripcion"):
                p_desc = doc.add_paragraph(foto["descripcion"])
                p_desc.runs[0].font.size = Pt(9)
                p_desc.runs[0].font.color.rgb = GRIS_CLARO
            
            doc.add_paragraph() # Espacio extra entre fotos

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
