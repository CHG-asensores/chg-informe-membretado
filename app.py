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

import io
import base64
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """
    Permite manipular los bordes de una celda específica de Word para dar el aspecto de "formulario web".
    kwargs acepta: top, bottom, start, end, left, right.
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

def remove_all_borders(table):
    """Elimina todos los bordes de una tabla para que sirva solo para maquetación."""
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, 
                            top={"val": "nil"}, bottom={"val": "nil"}, 
                            left={"val": "nil"}, right={"val": "nil"})

def generar_docx_con_observaciones(data, observaciones_extra):
    doc = Document()
    
    # Configuración de márgenes y tamaño (A4)
    seccion = doc.sections[0]
    seccion.page_height = Cm(29.7)
    seccion.page_width = Cm(21.0)
    seccion.top_margin = Cm(2.0)
    seccion.bottom_margin = Cm(2.0)
    seccion.left_margin = Cm(1.5)
    seccion.right_margin = Cm(1.5)
    
    # Colores exactos
    AZUL_TITULO = RGBColor(0x1B, 0x36, 0x5D)
    GRIS_ETIQUETA = RGBColor(0x6B, 0x72, 0x80)
    NEGRO_VALOR = RGBColor(0x11, 0x18, 0x27)
    VERDE_OK = RGBColor(0x16, 0xA3, 0x4A)
    ROJO_MAL = RGBColor(0xDC, 0x26, 0x26)
    GRIS_LINEA = {"val": "single", "sz": "2", "color": "E5E7EB"} # Borde inferior sutil

    # Estilos por defecto
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial' # Lo más parecido al PDF estándar
    font.size = Pt(10)

    # Encabezado Membretado
    LOGO = BASE_DIR / "template" / "assets" / "membrete.png"

    hdr = seccion.header
    ph = hdr.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if LOGO.exists():
        ph.add_run().add_picture(str(LOGO), width=Cm(4))

    # Extracción de datos
    meta = data.get("meta", {})
    secciones = data.get("secciones", [])
    fotos = data.get("fotos", [])
    obs = data.get("observaciones", {})
    firma = data.get("firma", {})

    # Título principal (Ej: M. Preventivo EDIF. SAN BORJA NORTE #488)
    p_titulo = doc.add_paragraph()
    p_titulo.paragraph_format.space_after = Pt(2)
    r_titulo = p_titulo.add_run(meta.get("titulo", "INFORME DE MANTENIMIENTO"))
    r_titulo.bold = True
    r_titulo.font.size = Pt(14)
    r_titulo.font.color.rgb = AZUL_TITULO

    doc.add_paragraph() # Espacio

    # --- TABLA DE METADATOS (Estilo MaintainX) ---
    tabla_meta = doc.add_table(rows=4, cols=4)
    tabla_meta.autofit = False
    
    # Anchos para simular el formulario
    for row in tabla_meta.rows:
        row.cells[0].width = Cm(4.0)
        row.cells[1].width = Cm(5.0)
        row.cells[2].width = Cm(4.5)
        row.cells[3].width = Cm(4.5)

    def fill_meta_cell(cell, etiqueta, valor, is_checkbox=False):
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.0
        
        # Etiqueta gris pequeña
        r_etiq = p.add_run(f"{etiqueta.upper()}\n")
        r_etiq.bold = True
        r_etiq.font.size = Pt(7.5)
        r_etiq.font.color.rgb = GRIS_ETIQUETA
        
        # Valor negro más grande
        val_str = str(valor) if valor else "-"
        if is_checkbox and val_str.lower() != "-":
            r_val = p.add_run(f"☑ {val_str}")
        else:
            r_val = p.add_run(val_str)
        r_val.font.size = Pt(9.5)
        r_val.font.color.rgb = NEGRO_VALOR
        
        # Solo borde inferior para separar
        set_cell_border(cell, top={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"}, bottom=GRIS_LINEA)

    fill_meta_cell(tabla_meta.cell(0,0), "ESTADO", meta.get("estado","-"), True)
    fill_meta_cell(tabla_meta.cell(0,1), "FECHA DE VENCIMIENTO", meta.get("fecha_vencimiento","-"))
    fill_meta_cell(tabla_meta.cell(0,2), "TIEMPO ESTIMADO", meta.get("tiempo_estimado","-"))
    fill_meta_cell(tabla_meta.cell(0,3), "TIPO DE TRABAJO", meta.get("tipo_trabajo","-"))
    
    fill_meta_cell(tabla_meta.cell(1,0), "ASIGNADOS", ", ".join(meta.get("asignados",[])) or "-")
    fill_meta_cell(tabla_meta.cell(1,1), "CATEGORÍAS", ", ".join(meta.get("categorias",[])) or "-")
    
    # Merge para ubicación y activo
    cell_ubi = tabla_meta.cell(2,0)
    cell_ubi.merge(tabla_meta.cell(2,1))
    fill_meta_cell(cell_ubi, "UBICACIÓN", meta.get("ubicacion","-"))
    
    cell_act = tabla_meta.cell(2,2)
    cell_act.merge(tabla_meta.cell(2,3))
    fill_meta_cell(cell_act, "ACTIVO", meta.get("activo","-"))

    # Merge para procedimiento
    cell_proc = tabla_meta.cell(3,0)
    cell_proc.merge(tabla_meta.cell(3,3))
    proc_val = meta.get("procedimiento", "-")
    if proc_val != "-":
        proc_val += "\n☑ INFORME DE MANTENIMIENTO PREVENTIVO"
    fill_meta_cell(cell_proc, "PROCEDIMIENTO", proc_val)

    doc.add_paragraph()

    # --- SECCIONES Y CHECKLISTS ---
    for sec in secciones:
        ps = doc.add_paragraph()
        ps.paragraph_format.space_before = Pt(12)
        ps.paragraph_format.space_after = Pt(4)
        run_sec = ps.add_run(sec["nombre"])
        run_sec.bold = True
        run_sec.font.size = Pt(11)
        run_sec.font.color.rgb = AZUL_TITULO
        
        # Tabla sin bordes para alinear Etiquetas y Valores
        ts = doc.add_table(rows=0, cols=2)
        remove_all_borders(ts)
        
        for campo in sec.get("campos", []):
            e = campo.get("etiqueta", "")
            v = campo.get("valor", "-")
            
            fila = ts.add_row()
            fila.cells[0].width = Cm(10)
            fila.cells[1].width = Cm(8)
            
            p_etiq = fila.cells[0].paragraphs[0]
            p_etiq.paragraph_format.space_after = Pt(2)
            r_e = p_etiq.add_run(f"{e}:")
            r_e.font.size = Pt(9.5)
            r_e.font.color.rgb = NEGRO_VALOR
            
            p_val = fila.cells[1].paragraphs[0]
            p_val.paragraph_format.space_after = Pt(2)
            
            r_chk = p_val.add_run("☑ ")
            r_chk.font.size = Pt(10)
            r_v = p_val.add_run(str(v))
            r_v.font.size = Pt(9.5)
            r_v.bold = True
            
            # Asignar color al texto del valor
            if v == "Bueno":
                r_v.font.color.rgb = VERDE_OK
                r_chk.font.color.rgb = VERDE_OK
            elif v == "Malo":
                r_v.font.color.rgb = ROJO_MAL
                r_chk.font.color.rgb = ROJO_MAL
            else:
                r_v.font.color.rgb = GRIS_ETIQUETA
                r_chk.font.color.rgb = GRIS_ETIQUETA
                
            # Línea divisoria inferior finita en la tabla
            set_cell_border(fila.cells[0], bottom=GRIS_LINEA)
            set_cell_border(fila.cells[1], bottom=GRIS_LINEA)

    # --- OBSERVACIONES Y RECOMENDACIONES ---
    p_obs_title = doc.add_paragraph()
    p_obs_title.paragraph_format.space_before = Pt(18)
    r_obs_title = p_obs_title.add_run("Observaciones y Recomendaciones")
    r_obs_title.bold = True
    r_obs_title.font.size = Pt(12)
    r_obs_title.font.color.rgb = AZUL_TITULO

    # Juntar todas las observaciones
    todas_observaciones = []
    if obs.get("para_cliente"):
        todas_observaciones.append(obs["para_cliente"])
    if obs.get("para_chg") and obs["para_chg"] != "-":
        todas_observaciones.append(f"Para CHG: {obs['para_chg']}")
        
    if isinstance(observaciones_extra, list):
        todas_observaciones.extend(observaciones_extra)
    elif isinstance(observaciones_extra, str) and observaciones_extra:
        todas_observaciones.append(observaciones_extra)

    for observacion in todas_observaciones:
        p_item = doc.add_paragraph(observacion) # Sin viñeta para que se vea como en el PDF
        p_item.paragraph_format.left_indent = Cm(0.5)
        p_item.runs[0].font.size = Pt(9.5)
        p_item.runs[0].font.color.rgb = NEGRO_VALOR
        
    if obs.get("evaluacion_final"):
        p_eval = doc.add_paragraph()
        p_eval.paragraph_format.space_before = Pt(8)
        r_et_ev = p_eval.add_run("Evaluación final: ")
        r_et_ev.bold = True
        r_et_ev.font.color.rgb = GRIS_ETIQUETA
        r_ev = p_eval.add_run(obs["evaluacion_final"])
        r_ev.bold = True

    # --- FIRMA ---
    doc.add_paragraph()
    pfi = doc.add_paragraph()
    run_firma = pfi.add_run("Firma del cliente:")
    run_firma.bold = True
    run_firma.font.size = Pt(10)
    run_firma.font.color.rgb = GRIS_ETIQUETA

    if firma.get("data_base64"):
        try:
            img_bytes = base64.b64decode(firma["data_base64"])
            p_img = doc.add_paragraph()
            p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(6))
        except:
            pass
            
    if firma.get("texto"):
        pft = doc.add_paragraph(firma["texto"])
        pft.runs[0].font.size = Pt(8.5)
        pft.runs[0].font.color.rgb = GRIS_ETIQUETA

    if meta.get("fecha_hora_salida"):
        ps2 = doc.add_paragraph(f"Fecha y Hora de salida: {meta['fecha_hora_salida']}")
        ps2.runs[0].font.size = Pt(8.5)
        ps2.runs[0].font.color.rgb = GRIS_ETIQUETA

    # --- FOTOGRAFÍAS (Saltos de página por foto como en el PDF) ---
    if fotos:
        for foto in fotos:
            doc.add_page_break()
            p_foto_title = doc.add_paragraph()
            r_ft_title = p_foto_title.add_run(foto.get("titulo", "Fotografía:"))
            r_ft_title.bold = True
            r_ft_title.font.size = Pt(11)
            r_ft_title.font.color.rgb = AZUL_TITULO
            
            if foto.get("data_base64"):
                try:
                    img_bytes = base64.b64decode(foto["data_base64"])
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(16)) # Más grande, simulando la página completa
                except Exception as e:
                    pass

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
