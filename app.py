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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

import io
import base64
import re
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def agregar_marca_agua(seccion, ruta_imagen):
    """Agrega imagen como fondo a página completa, detrás del texto."""
    hdr = seccion.header
    # Limpiar encabezado
    for p in hdr.paragraphs:
        p.clear()
    
    p = hdr.paragraphs[0]
    run = p.add_run()
    # Inserción base permitida
    pic = run.add_picture(str(ruta_imagen), width=Cm(21), height=Cm(29.7))
    
    # Transformación a 'Detrás del texto' con reporte de errores
    try:
        drawing = run._r.find(qn('w:drawing'))
        if drawing is not None:
            inline = drawing.find(qn('wp:inline'))
            if inline is not None:
                anchor = OxmlElement('wp:anchor')
                anchor.set('distT', '0')
                anchor.set('distB', '0')
                anchor.set('distL', '0')
                anchor.set('distR', '0')
                anchor.set('simplePos', '0')
                anchor.set('relativeHeight', '0')
                anchor.set('behindDoc', '1') # 1 = Detrás del texto
                anchor.set('locked', '0')
                anchor.set('layoutInCell', '1')
                anchor.set('allowOverlap', '1')
                
                simplePos = OxmlElement('wp:simplePos')
                simplePos.set('x', '0')
                simplePos.set('y', '0')
                anchor.append(simplePos)
                
                posH = OxmlElement('wp:positionH')
                posH.set('relativeFrom', 'page')
                offsetH = OxmlElement('wp:posOffset')
                offsetH.text = '0'
                posH.append(offsetH)
                anchor.append(posH)
                
                posV = OxmlElement('wp:positionV')
                posV.set('relativeFrom', 'page')
                offsetV = OxmlElement('wp:posOffset')
                offsetV.text = '0'
                posV.append(offsetV)
                anchor.append(posV)
                
                for child in list(inline):
                    anchor.append(child)
                    
                drawing.replace(inline, anchor)
    except Exception as e:
        raise RuntimeError(f"Error al insertar marca de agua: {e}")

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

def limpiar_texto(texto):
    if not texto: return "-"
    limpio = re.sub(r'[☑\u2611]\s*|^V\s+', '', str(texto)).strip()
    return limpio if limpio else "-"

import io
import base64
import re
import zipfile
import shutil
import os
import tempfile
from pathlib import Path
from lxml import etree
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def aplicar_fondo_pagina(docx_bytes, ruta_imagen_png):
    """Inserta imagen como fondo de página completa en el docx manipulando su ZIP interno."""
    if not os.path.exists(ruta_imagen_png):
        return docx_bytes

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_docx_path = os.path.join(tmpdir, 'temp.docx')
        with open(tmp_docx_path, 'wb') as f:
            f.write(docx_bytes)
        
        extract_dir = os.path.join(tmpdir, 'extracted')
        with zipfile.ZipFile(tmp_docx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
            
        # 1. Copiar la imagen a word/media/
        media_dir = os.path.join(extract_dir, 'word', 'media')
        os.makedirs(media_dir, exist_ok=True)
        shutil.copy2(ruta_imagen_png, os.path.join(media_dir, 'bg_watermark.png'))
        
        # 2. Actualizar word/_rels/document.xml.rels
        rels_path = os.path.join(extract_dir, 'word', '_rels', 'document.xml.rels')
        ns_rels = 'http://schemas.openxmlformats.org/package/2006/relationships'
        etree.register_namespace('', ns_rels)
        tree_rels = etree.parse(rels_path)
        root_rels = tree_rels.getroot()
        
        bg_rId = "rIdBgWatermark"
        new_rel = etree.Element(f"{{{ns_rels}}}Relationship", 
                                Id=bg_rId, 
                                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", 
                                Target="media/bg_watermark.png")
        root_rels.append(new_rel)
        tree_rels.write(rels_path, xml_declaration=True, encoding='UTF-8')
        
        # 3. Actualizar word/document.xml
        doc_path = os.path.join(extract_dir, 'word', 'document.xml')
        ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        ns_v = 'urn:schemas-microsoft-com:vml'
        ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        ns_o = 'urn:schemas-microsoft-com:office:office'
        
        tree_doc = etree.parse(doc_path)
        root_doc = tree_doc.getroot()
        
        bg_elem = etree.Element(f"{{{ns_w}}}background", {f"{{{ns_w}}}color": "FFFFFF"})
        v_bg = etree.SubElement(bg_elem, f"{{{ns_v}}}background", id="_x0000_s1025")
        v_bg.set(f"{{{ns_o}}}bwmode", "white")
        v_bg.set(f"{{{ns_o}}}targetscreensize", "1024,768")
        
        v_fill = etree.SubElement(v_bg, f"{{{ns_v}}}fill")
        v_fill.set(f"{{{ns_r}}}id", bg_rId)
        v_fill.set(f"{{{ns_o}}}title", "Fondo")
        v_fill.set("type", "frame")
        
        # Insertar el background justo antes de <w:body>
        body_elem = root_doc.find(f"{{{ns_w}}}body")
        if body_elem is not None:
            body_index = root_doc.index(body_elem)
            root_doc.insert(body_index, bg_elem)
        else:
            root_doc.insert(0, bg_elem)
            
        tree_doc.write(doc_path, xml_declaration=True, encoding='UTF-8')

        # 4. Forzar configuración para que el fondo sea visible e imprimible en Word
        settings_path = os.path.join(extract_dir, 'word', 'settings.xml')
        if os.path.exists(settings_path):
            tree_set = etree.parse(settings_path)
            root_set = tree_set.getroot()
            if root_set.find(f"{{{ns_w}}}displayBackgroundShape") is None:
                disp = etree.Element(f"{{{ns_w}}}displayBackgroundShape")
                root_set.append(disp)
            tree_set.write(settings_path, xml_declaration=True, encoding='UTF-8')
        
        # 5. Volver a empaquetar el ZIP (.docx)
        new_docx_path = os.path.join(tmpdir, 'final.docx')
        with zipfile.ZipFile(new_docx_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zipf.write(file_path, arcname)
                    
        with open(new_docx_path, 'rb') as f:
            return f.read()

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

def aplicar_fondo_pagina(docx_bytes, ruta_imagen_png):
    if not os.path.exists(ruta_imagen_png):
        return docx_bytes

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = os.path.join(tmpdir, 'temp.docx')
        with open(tmp_path, 'wb') as f:
            f.write(docx_bytes)

        extract_dir = os.path.join(tmpdir, 'extracted')
        with zipfile.ZipFile(tmp_path, 'r') as z:
            z.extractall(extract_dir)

        # Copiar imagen
        media_dir = os.path.join(extract_dir, 'word', 'media')
        os.makedirs(media_dir, exist_ok=True)
        shutil.copy2(ruta_imagen_png, os.path.join(media_dir, 'bg.png'))

        # Agregar relacion en document.xml.rels
        rels_path = os.path.join(extract_dir, 'word', '_rels', 'document.xml.rels')
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels = f.read()
        if 'rIdBg' not in rels:
            rels = rels.replace('</Relationships>',
                '<Relationship Id="rIdBg" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                'Target="media/bg.png"/></Relationships>')
            with open(rels_path, 'w', encoding='utf-8') as f:
                f.write(rels)

        # Agregar background en document.xml
        doc_path = os.path.join(extract_dir, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            doc_xml = f.read()
        bg_xml = (
            '<w:background w:color="FFFFFF" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<v:background id="_x0000_s1025" o:bwmode="white">'
            '<v:fill r:id="rIdBg" o:title="bg" type="frame"/>'
            '</v:background></w:background>'
        )
        if '<w:background' not in doc_xml:
            doc_xml = doc_xml.replace('<w:body>', bg_xml + '<w:body>', 1)
            with open(doc_path, 'w', encoding='utf-8') as f:
                f.write(doc_xml)

        # Activar displayBackgroundShape en settings.xml
        settings_path = os.path.join(extract_dir, 'word', 'settings.xml')
        if os.path.exists(settings_path):
            with open(settings_path, 'r', encoding='utf-8') as f:
                settings = f.read()
            if 'displayBackgroundShape' not in settings:
                settings = settings.replace('</w:settings>',
                    '<w:displayBackgroundShape/></w:settings>')
                with open(settings_path, 'w', encoding='utf-8') as f:
                    f.write(settings)

        # Reempaquetar preservando tipos MIME originales
        new_path = os.path.join(tmpdir, 'final.docx')
        with zipfile.ZipFile(tmp_path, 'r') as orig_zip:
            orig_names = {i.filename: i for i in orig_zip.infolist()}
        with zipfile.ZipFile(new_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    fp = os.path.join(root, file)
                    arcname = os.path.relpath(fp, extract_dir)
                    zout.write(fp, arcname)

        with open(new_path, 'rb') as f:
            return f.read()


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
