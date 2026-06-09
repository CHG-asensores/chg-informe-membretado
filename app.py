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

import io
import base64
import re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """Manipula el XML para aplicar bordes finos estilo web."""
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

def set_cell_margins(cell, top=50, bottom=50, start=100, end=100):
    """Agrega padding interno a las celdas para que no se vean apretadas."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, value in [('top', top), ('bottom', bottom), ('left', start), ('right', end)]:
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_all_borders(table):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={"val": "nil"}, bottom={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})
            set_cell_margins(cell, top=80, bottom=80, start=0, end=0)

def limpiar_texto(texto):
    """Limpia caracteres basura (☑, V) que arrastra el OCR de los PDFs."""
    if not texto: return "-"
    limpio = re.sub(r'[☑\u2611]\s*|^V\s+', '', str(texto)).strip()
    return limpio if limpio else "-"

import io
import base64
import re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """Manipula el XML para aplicar bordes finos estilo web."""
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

def set_cell_margins(cell, top=50, bottom=50, start=100, end=100):
    """Agrega padding interno a las celdas para que no se vean apretadas."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, value in [('top', top), ('bottom', bottom), ('left', start), ('right', end)]:
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_all_borders(table):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={"val": "nil"}, bottom={"val": "nil"}, left={"val": "nil"}, right={"val": "nil"})
            set_cell_margins(cell, top=80, bottom=80, start=0, end=0)

def limpiar_texto(texto):
    """Limpia caracteres basura (☑, V) que arrastra el OCR de los PDFs."""
    if not texto: return "-"
    limpio = re.sub(r'[☑\u2611]\s*|^V\s+', '', str(texto)).strip()
    return limpio if limpio else "-"

import io
import base64
import re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE

def set_cell_border(cell, **kwargs):
    """Manipula el XML para aplicar bordes a celdas específicas."""
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

def limpiar_texto(texto):
    """Limpia caracteres basura (☑, V) que arrastra el OCR de los PDFs."""
    if not texto: return "-"
    limpio = re.sub(r'[☑\u2611]\s*|^V\s+', '', str(texto)).strip()
    return limpio if limpio else "-"

import io
import base64
import re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE

def set_cell_border(cell, **kwargs):
    """Aplica bordes personalizados a las celdas (ej. para la caja azul de Categorías)."""
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

def set_cell_bg_color(cell, color):
    """Establece el color de fondo de una celda (ej. caja negra)."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def limpiar_texto(texto):
    """Limpia caracteres basura (☑, V) que arrastra el OCR de los PDFs."""
    if not texto: return "-"
    limpio = re.sub(r'[☑\u2611]\s*|^V\s+', '', str(texto)).strip()
    return limpio if limpio else "-"

import io
import base64
import re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

def agregar_marca_agua(seccion, ruta_imagen):
    """Agrega imagen como fondo a página completa en el encabezado."""
    hdr = seccion.header
    # Limpiar encabezado
    for p in hdr.paragraphs:
        p.clear()
    p = hdr.paragraphs[0]
    run = p.add_run()
    pic = run.add_picture(ruta_imagen, width=Cm(21), height=Cm(29.7))

def set_cell_border(cell, **kwargs):
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

def set_cell_bg_color(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def limpiar_texto(texto):
    if not texto: return "-"
    limpio = re.sub(r'[☑\u2611]\s*|^V\s+', '', str(texto)).strip()
    return limpio if limpio else "-"

def generar_docx_con_observaciones(data, observaciones_extra):
    doc = Document()
    
    # 1. CONFIGURACIÓN DE PÁGINA A4
    seccion = doc.sections[0]
    seccion.page_height = Cm(29.7)
    seccion.page_width = Cm(21.0)
    
    # Márgenes para el contenido principal
    seccion.top_margin = Cm(3.0)
    seccion.bottom_margin = Cm(3.0)
    seccion.left_margin = Cm(1.5)
    seccion.right_margin = Cm(1.5)
    
    AZUL_LINK = RGBColor(0x25, 0x63, 0xEB)
    GRIS_ETIQUETA = RGBColor(0x6B, 0x72, 0x80)
    NEGRO_VALOR = RGBColor(0x00, 0x00, 0x00)
    VERDE_OK = RGBColor(0x10, 0xB9, 0x81)

    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(9.5)

    # 2. ENCABEZADO A PÁGINA COMPLETA (Patrón solicitado)
    LOGO = BASE_DIR / "template" / "assets" / "membrete.png"
    if LOGO.exists():
        agregar_marca_agua(seccion, str(LOGO))

    # 3. EXTRACCIÓN DE DATOS
    meta = data.get("meta", {})
    secciones = data.get("secciones", [])
    fotos = data.get("fotos", [])
    obs = data.get("observaciones", {})
    firma = data.get("firma", {})

    # 4. RECUADRO NEGRO SUPERIOR ("CHG Ascensores")
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.autofit = False
    tbl_top.columns[0].width = Cm(1.5)
    tbl_top.columns[1].width = Cm(16.5)
    
    cell_box = tbl_top.cell(0, 0)
    set_cell_bg_color(cell_box, "000000")
    p_box = cell_box.paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_box = p_box.add_run(" N ")
    r_box.font.color.rgb = RGBColor(255, 255, 255)
    r_box.font.bold = True
    
    cell_text = tbl_top.cell(0, 1)
    p_text = cell_text.paragraphs[0]
    r_text = p_text.add_run("   CHG Ascensores")
    r_text.font.bold = True
    r_text.font.size = Pt(11)

    doc.add_paragraph()

    # 5. METADATOS EN COLUMNAS (Estilo MaintainX)
    tbl_meta = doc.add_table(rows=4, cols=2)
    tbl_meta.autofit = False
    tbl_meta.columns[0].width = Cm(9.0)
    tbl_meta.columns[1].width = Cm(9.0)

    def fill_cell(cell, icon, label, value, is_status=False, is_category=False):
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(8)
        
        r_label = p.add_run(f"{icon} {label.upper()}\n")
        r_label.font.size = Pt(7.5)
        r_label.font.color.rgb = GRIS_ETIQUETA
        
        val_clean = limpiar_texto(value)
        
        if is_status:
            r_val = p.add_run(f"✓ {val_clean}")
            r_val.font.color.rgb = VERDE_OK
            r_val.font.bold = True
            r_val.font.size = Pt(10)
        elif is_category:
            p.text = "" 
            p.add_run(f"{icon} {label.upper()}").font.color.rgb = GRIS_ETIQUETA
            p.runs[0].font.size = Pt(7.5)
            
            t_cat = cell.add_table(rows=1, cols=1)
            c_cat = t_cat.cell(0, 0)
            set_cell_border(c_cat, 
                            top={"val": "single", "sz": "4", "color": "2563EB"},
                            bottom={"val": "single", "sz": "4", "color": "2563EB"},
                            left={"val": "single", "sz": "4", "color": "2563EB"},
                            right={"val": "single", "sz": "4", "color": "2563EB"})
            p_cat = c_cat.paragraphs[0]
            r_cat = p_cat.add_run(val_clean)
            r_cat.font.color.rgb = AZUL_LINK
            r_cat.font.size = Pt(9.5)
        else:
            r_val = p.add_run(val_clean)
            r_val.font.bold = True
            r_val.font.color.rgb = NEGRO_VALOR
            r_val.font.size = Pt(10)

    fill_cell(tbl_meta.cell(0, 0), "🔒", "ESTADO", meta.get("estado", "-"), is_status=True)
    fill_cell(tbl_meta.cell(0, 1), "🕒", "FECHA DE VENCIMIENTO", meta.get("fecha_vencimiento", "-"))
    fill_cell(tbl_meta.cell(1, 0), "⏱️", "TIEMPO ESTIMADO", meta.get("tiempo_estimado", "-"))
    fill_cell(tbl_meta.cell(1, 1), "🔨", "TIPO DE TRABAJO", meta.get("tipo_trabajo", "-"))
    fill_cell(tbl_meta.cell(2, 0), "👥", "ASIGNADOS", ", ".join(meta.get("asignados", [])) or "-")
    fill_cell(tbl_meta.cell(2, 1), "🏷️", "CATEGORÍAS", ", ".join(meta.get("categorias", [])) or "-", is_category=True)
    fill_cell(tbl_meta.cell(3, 0), "📍", "UBICACIÓN", meta.get("ubicacion", "-"))
    fill_cell(tbl_meta.cell(3, 1), "🏢", "ACTIVO", meta.get("activo", "-"))

    doc.add_paragraph()
    p_proc = doc.add_paragraph()
    p_proc.paragraph_format.space_before = Pt(12)
    p_proc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_p_etiq = p_proc.add_run("≡ PROCEDIMIENTO\n")
    r_p_etiq.font.size = Pt(7)
    r_p_etiq.font.color.rgb = GRIS_ETIQUETA
    r_p_val = p_proc.add_run(limpiar_texto(meta.get("procedimiento", "INFORME DE MANTENIMIENTO PREVENTIVO")).upper())
    r_p_val.font.size = Pt(11)
    r_p_val.bold = True
    r_p_val.font.color.rgb = AZUL_LINK
    r_p_val.underline = True

    # 6. CHECKLISTS APILADOS (Estilo visual MaintainX)
    for sec in secciones:
        ps = doc.add_paragraph()
        ps.paragraph_format.space_before = Pt(18)
        ps.paragraph_format.space_after = Pt(4)
        run_sec = ps.add_run(limpiar_texto(sec["nombre"]))
        run_sec.bold = True
        run_sec.font.size = Pt(10)
        run_sec.font.color.rgb = NEGRO_VALOR
        
        for campo in sec.get("campos", []):
            e = limpiar_texto(campo.get("etiqueta", ""))
            v = limpiar_texto(campo.get("valor", "-"))
            
            p_item = doc.add_paragraph()
            p_item.paragraph_format.space_after = Pt(0)
            p_item.paragraph_format.space_before = Pt(6)
            etiqueta_texto = e if e.endswith(":") else f"{e}:"
            r_e = p_item.add_run(etiqueta_texto)
            r_e.font.size = Pt(8.5)
            r_e.font.color.rgb = NEGRO_VALOR
            
            p_val = doc.add_paragraph()
            p_val.paragraph_format.space_after = Pt(0)
            
            r_icon = p_val.add_run("◉ ")
            r_icon.font.size = Pt(9.5)
            r_icon.font.color.rgb = AZUL_LINK
            
            r_v = p_val.add_run(v)
            r_v.font.size = Pt(9.5)
            r_v.bold = True
            r_v.font.color.rgb = NEGRO_VALOR

    # 7. OBSERVACIONES Y RECOMENDACIONES
    p_obs_title = doc.add_paragraph()
    p_obs_title.paragraph_format.space_before = Pt(24)
    p_obs_title.paragraph_format.space_after = Pt(6)
    r_obs_title = p_obs_title.add_run("Observaciones y Recomendaciones")
    r_obs_title.bold = True
    r_obs_title.font.size = Pt(10)
    
    todas_observaciones = []
    if obs.get("para_cliente"): todas_observaciones.append(limpiar_texto(obs["para_cliente"]))
    if obs.get("para_chg") and obs["para_chg"] != "-": todas_observaciones.append(f"Para CHG: {limpiar_texto(obs['para_chg'])}")
    if isinstance(observaciones_extra, list):
        todas_observaciones.extend([limpiar_texto(o) for o in observaciones_extra])
    elif isinstance(observaciones_extra, str) and observaciones_extra:
        todas_observaciones.append(limpiar_texto(observaciones_extra))

    for observacion in todas_observaciones:
        p_item = doc.add_paragraph(observacion)
        p_item.paragraph_format.space_after = Pt(4)
        p_item.runs[0].font.size = Pt(9)
        
    if obs.get("evaluacion_final"):
        p_eval = doc.add_paragraph()
        p_eval.paragraph_format.space_before = Pt(8)
        r_et_ev = p_eval.add_run("Evaluación final: ")
        r_et_ev.bold = True
        r_et_ev.font.color.rgb = GRIS_ETIQUETA
        r_ev = p_eval.add_run(limpiar_texto(obs["evaluacion_final"]))
        r_ev.bold = True

    # 8. FIRMA
    doc.add_paragraph()
    pfi = doc.add_paragraph()
    run_firma = pfi.add_run("Firma del cliente:")
    run_firma.bold = True
    run_firma.font.size = Pt(9)

    if firma.get("data_base64"):
        try:
            img_bytes = base64.b64decode(firma["data_base64"])
            p_img = doc.add_paragraph()
            p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(6))
        except Exception: pass
            
    if firma.get("texto"):
        pft = doc.add_paragraph(limpiar_texto(firma["texto"]))
        pft.runs[0].font.size = Pt(8)
        pft.runs[0].font.color.rgb = GRIS_ETIQUETA

    if meta.get("fecha_hora_salida"):
        ps2 = doc.add_paragraph(f"Fecha y Hora de salida: {limpiar_texto(meta['fecha_hora_salida'])}")
        ps2.runs[0].font.size = Pt(8)
        ps2.runs[0].font.color.rgb = GRIS_ETIQUETA

    # 9. FOTOGRAFÍAS
    if fotos:
        for foto in fotos:
            doc.add_page_break()
            p_foto_title = doc.add_paragraph()
            r_ft_title = p_foto_title.add_run(limpiar_texto(foto.get("titulo", "Fotografía:")))
            r_ft_title.bold = True
            r_ft_title.font.size = Pt(11)
            
            if foto.get("data_base64"):
                try:
                    img_bytes = base64.b64decode(foto["data_base64"])
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(16))
                except Exception: pass

    # 10. PIE DE PÁGINA
    ftr = seccion.footer
    tbl_ftr = ftr.add_table(rows=1, cols=2, width=Cm(17))
    tbl_ftr.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c in tbl_ftr.rows[0].cells:
        set_cell_border(c, top={"val": "single", "sz": "6", "color": "D1D5DB"})
        
    c_left = tbl_ftr.cell(0, 0)
    c_right = tbl_ftr.cell(0, 1)
    c_left.width = Cm(8)
    c_right.width = Cm(9)
    
    p_l = c_left.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_l = p_l.add_run("\nGenerado para CHG Ascensores")
    r_l.font.color.rgb = GRIS_ETIQUETA
    r_l.font.size = Pt(8)
    
    p_r = c_right.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_r = p_r.add_run("\n947234073\n(01) 627-9422\ncomercial@chgascensor.com\nAv. La Encalada N°110 - Surco")
    r_r.font.color.rgb = AZUL_LINK
    r_r.font.size = Pt(8.5)
    r_r.bold = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


UI = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CHG — Informe Membretado</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      background: #f0f2f5;
      color: #1a1a2e;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }

    header {
      background: #1a1a2e;
      color: #fff;
      padding: 18px 32px;
      display: flex;
      align-items: center;
      gap: 14px;
    }
    header .logo {
      font-size: 20px;
      font-weight: 700;
      letter-spacing: -0.5px;
    }
    header .logo span { color: #60a5fa; }
    header .sub {
      font-size: 13px;
      color: #94a3b8;
      margin-left: auto;
    }

    main {
      flex: 1;
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 24px;
      padding: 32px;
      max-width: 1100px;
      width: 100%;
      margin: 0 auto;
    }

    .panel {
      background: #fff;
      border-radius: 12px;
      padding: 28px;
      box-shadow: 0 1px 3px rgba(0,0,0,.08), 0 4px 16px rgba(0,0,0,.04);
      display: flex;
      flex-direction: column;
      gap: 18px;
    }

    .panel-title {
      font-size: 13px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: #64748b;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .panel-title .step {
      background: #1a1a2e;
      color: #fff;
      border-radius: 50%;
      width: 22px; height: 22px;
      font-size: 11px;
      display: flex; align-items: center; justify-content: center;
      flex-shrink: 0;
    }

    /* Drop zone */
    .drop-zone {
      border: 2px dashed #cbd5e1;
      border-radius: 10px;
      padding: 48px 24px;
      text-align: center;
      cursor: pointer;
      transition: border-color .2s, background .2s;
      position: relative;
    }
    .drop-zone:hover, .drop-zone.over {
      border-color: #3b82f6;
      background: #eff6ff;
    }
    .drop-zone input[type=file] {
      position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
    }
    .drop-icon { font-size: 40px; margin-bottom: 12px; }
    .drop-text { font-size: 15px; font-weight: 500; color: #334155; margin-bottom: 6px; }
    .drop-hint { font-size: 13px; color: #94a3b8; }

    .file-info {
      display: none;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 12px 16px;
      font-size: 14px;
      color: #334155;
      align-items: center;
      gap: 10px;
    }
    .file-info.show { display: flex; }
    .file-info .fname { font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .file-info .fsize { color: #94a3b8; font-size: 12px; flex-shrink: 0; }
    .file-info .remove-btn {
      background: none; border: none; cursor: pointer; color: #94a3b8; font-size: 18px; line-height: 1; padding: 0 2px;
    }
    .file-info .remove-btn:hover { color: #ef4444; }

    .btn-generate {
      background: #1a1a2e;
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 13px 20px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      transition: background .2s, opacity .2s;
    }
    .btn-generate:hover { background: #2d2d4e; }
    .btn-generate:disabled { opacity: .5; cursor: not-allowed; }

    /* Resultado */
    .result-empty {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 12px;
      color: #94a3b8;
      padding: 48px 24px;
      text-align: center;
    }
    .result-empty .icon { font-size: 48px; }
    .result-empty p { font-size: 14px; line-height: 1.6; }

    .result-card {
      display: none;
      flex-direction: column;
      gap: 16px;
    }
    .result-card.show { display: flex; }

    .result-badge {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 12px 16px;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 500;
    }
    .result-badge.success { background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }
    .result-badge.error   { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
    .result-badge.warning { background: #fffbeb; color: #92400e; border: 1px solid #fde68a; }

    .result-meta {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 14px 16px;
      font-size: 13px;
      color: #475569;
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px 16px;
    }
    .result-meta .item-label { color: #94a3b8; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; }
    .result-meta .item-value { font-weight: 500; color: #1e293b; margin-top: 2px; }

    .btn-download {
      background: #16a34a;
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 13px 20px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      text-align: center;
      text-decoration: none;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      transition: background .2s;
    }
    .btn-download:hover { background: #15803d; }

    .btn-reset {
      background: none;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 10px 20px;
      font-size: 14px;
      color: #64748b;
      cursor: pointer;
      transition: border-color .2s, color .2s;
    }
    .btn-reset:hover { border-color: #94a3b8; color: #334155; }

    /* Loading */
    .loading-overlay {
      display: none;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 16px;
      padding: 48px 24px;
      flex: 1;
    }
    .loading-overlay.show { display: flex; }
    .spinner {
      width: 44px; height: 44px;
      border: 4px solid #e2e8f0;
      border-top-color: #1a1a2e;
      border-radius: 50%;
      animation: spin .8s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .loading-text { font-size: 15px; font-weight: 500; color: #334155; }
    .loading-sub  { font-size: 13px; color: #94a3b8; }

    /* Validation warning */
    .val-details {
      background: #fffbeb;
      border: 1px solid #fde68a;
      border-radius: 8px;
      padding: 12px 16px;
      font-size: 12px;
      color: #78350f;
      display: none;
    }
    .val-details.show { display: block; }
    .val-details ul { margin-top: 6px; padding-left: 16px; }

    footer {
      text-align: center;
      padding: 16px;
      font-size: 12px;
      color: #94a3b8;
    }

    @media (max-width: 700px) {
      main { grid-template-columns: 1fr; padding: 16px; }
    }
  </style>
</head>
<body>

<header>
  <div class="logo">CHG <span>Ascensores</span></div>
  <div class="sub">Generador de Informes Membretados</div>
</header>

<main>
  <!-- ── Columna izquierda: subir PDF ── -->
  <div class="panel">
    <div class="panel-title">
      <div class="step">1</div>
      Subir PDF de MaintainX
    </div>

    <div class="drop-zone" id="dropZone">
      <input type="file" id="fileInput" accept=".pdf">
      <div class="drop-icon">📄</div>
      <div class="drop-text">Arrastrá el PDF aquí</div>
      <div class="drop-hint">o hacé clic para seleccionar</div>
    </div>

    <div class="file-info" id="fileInfo">
      <span>📎</span>
      <span class="fname" id="fileName"></span>
      <span class="fsize" id="fileSize"></span>
      <button class="remove-btn" id="removeBtn" title="Quitar archivo">✕</button>
    </div>

    <div style="display:flex;flex-direction:column;gap:6px;margin-top:4px;"><label for="observaciones" style="font-size:13px;font-weight:600;color:#334155;text-transform:uppercase;">Observaciones</label><textarea id="observaciones" rows="4" placeholder="Ej: Se recomienda cambiar baterias..." style="width:100%;padding:10px 14px;border:1.5px solid #cbd5e1;border-radius:8px;font-size:14px;font-family:inherit;resize:vertical;background:#f8fafc;box-sizing:border-box;"></textarea></div>
    <button class="btn-generate" id="btnGenerate" disabled>
      <span>⚙️</span>
      Generar informe membretado
    </button>
    <button class="btn-generate" id="btnGenerateWord" disabled style="background:#1d6f42;margin-top:4px;">
      <span>📄</span>
      Generar Word con observaciones
    </button>
  </div>

  <!-- ── Columna derecha: resultado ── -->
  <div class="panel">
    <div class="panel-title">
      <div class="step">2</div>
      Descargar informe
    </div>

    <div class="result-empty" id="resultEmpty">
      <div class="icon">📋</div>
      <p>El informe membretado aparecerá aquí<br>una vez que subas y proceses el PDF.</p>
    </div>

    <div class="loading-overlay" id="loadingOverlay">
      <div class="spinner"></div>
      <div class="loading-text">Procesando PDF…</div>
      <div class="loading-sub">Parseando datos y generando informe</div>
    </div>

    <div class="result-card" id="resultCard">
      <div class="result-badge" id="resultBadge"></div>

      <div class="result-meta" id="resultMeta"></div>

      <div class="val-details" id="valDetails">
        <strong>Advertencias de validación:</strong>
        <ul id="valList"></ul>
      </div>

      <a class="btn-download" id="btnDownload" href="#" download>
        ⬇️ Descargar PDF membretado
      </a>

      <button class="btn-reset" id="btnReset">Procesar otro PDF</button>
    </div>
  </div>
</main>

<footer>CHG Ascensores · Informes Membretados</footer>

<script>
  const dropZone    = document.getElementById('dropZone');
  const fileInput   = document.getElementById('fileInput');
  const fileInfo    = document.getElementById('fileInfo');
  const fileName    = document.getElementById('fileName');
  const fileSize    = document.getElementById('fileSize');
  const removeBtn   = document.getElementById('removeBtn');
  const btnGenerate = document.getElementById('btnGenerate');

  const resultEmpty   = document.getElementById('resultEmpty');
  const loadingOverlay = document.getElementById('loadingOverlay');
  const resultCard    = document.getElementById('resultCard');
  const resultBadge   = document.getElementById('resultBadge');
  const resultMeta    = document.getElementById('resultMeta');
  const valDetails    = document.getElementById('valDetails');
  const valList       = document.getElementById('valList');
  const btnDownload   = document.getElementById('btnDownload');
  const btnReset      = document.getElementById('btnReset');

  let selectedFile = null;

  function fmtSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024*1024) return (bytes/1024).toFixed(1) + ' KB';
    return (bytes/1024/1024).toFixed(1) + ' MB';
  }

  function setFile(file) {
    if (!file || file.type !== 'application/pdf') {
      alert('Por favor seleccioná un archivo PDF.');
      return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = fmtSize(file.size);
    fileInfo.classList.add('show');
    dropZone.style.display = 'none';
    btnGenerate.disabled = false;
    document.getElementById("btnGenerateWord").disabled = false;
    resetResult();
  }

  function resetFile() {
    selectedFile = null;
    fileInput.value = '';
    fileInfo.classList.remove('show');
    dropZone.style.display = '';
    btnGenerate.disabled = true;
    document.getElementById("btnGenerateWord").disabled = true;
    resetResult();
  }

  function resetResult() {
    resultCard.classList.remove('show');
    loadingOverlay.classList.remove('show');
    resultEmpty.style.display = '';
  }

  function showLoading() {
    resultEmpty.style.display = 'none';
    resultCard.classList.remove('show');
    loadingOverlay.classList.add('show');
  }

  function showResult(data) {
    loadingOverlay.classList.remove('show');
    resultCard.classList.add('show');

    if (data.ok) {
      resultBadge.className = 'result-badge success';
      resultBadge.innerHTML = '✅ Informe generado correctamente';
    } else if (data.warning) {
      resultBadge.className = 'result-badge warning';
      resultBadge.innerHTML = '⚠️ Generado con advertencias de validación';
    } else {
      resultBadge.className = 'result-badge error';
      resultBadge.innerHTML = '❌ ' + (data.error || 'Error al procesar');
      btnDownload.style.display = 'none';
      return;
    }

    btnDownload.style.display = '';
    btnDownload.href = data.download_url;
    btnDownload.download = data.filename;

    const m = data.meta || {};
    resultMeta.innerHTML = [
      ['N° de orden', '#' + (m.numero_orden || '—')],
      ['Estado',      m.estado || '—'],
      ['Ubicación',   m.ubicacion || '—'],
      ['Activo',      m.activo || '—'],
      ['Asignados',   (m.asignados || []).join(', ') || '—'],
      ['Campos',      m.campos_completados || '—'],
    ].map(([l, v]) => `
      <div>
        <div class="item-label">${l}</div>
        <div class="item-value">${v}</div>
      </div>`).join('');

    if (data.warnings && data.warnings.length) {
      valList.innerHTML = data.warnings.map(w =>
        `<li><strong>${w.etiqueta}</strong>: "${w.valor}"</li>`
      ).join('');
      valDetails.classList.add('show');
    } else {
      valDetails.classList.remove('show');
    }
  }

  // Drag & drop
  ['dragenter','dragover'].forEach(evt =>
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add('over'); })
  );
  ['dragleave','drop'].forEach(evt =>
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove('over'); })
  );
  dropZone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
  });

  removeBtn.addEventListener('click', resetFile);

  btnReset.addEventListener('click', () => {
    resetFile();
  });

  document.getElementById('btnGenerateWord').addEventListener('click', async () => {
    if (!selectedFile) return;
    const form = new FormData();
    form.append('pdf', selectedFile);
    form.append('observaciones', document.getElementById('observaciones').value.trim());
    try {
      const resp = await fetch('/generate-word', { method: 'POST', body: form });
      if (!resp.ok) { const d = await resp.json(); alert(d.error); return; }
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = selectedFile.name.replace('.pdf', '') + '_observaciones.docx';
      a.click();
      URL.revokeObjectURL(url);
    } catch(err) { alert('Error: ' + err.message); }
  });

  btnGenerate.addEventListener('click', async () => {
    if (!selectedFile) return;
    showLoading();
    btnGenerate.disabled = true;

    const form = new FormData();
    form.append('pdf', selectedFile);
    form.append('observaciones', document.getElementById('observaciones').value.trim());

    try {
      const resp = await fetch('/generate', { method: 'POST', body: form });
      const data = await resp.json();
      showResult(data);
    } catch (err) {
      showResult({ error: 'Error de red: ' + err.message });
    } finally {
      btnGenerate.disabled = false;
    }
  });
</script>
</body>
</html>"""

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
