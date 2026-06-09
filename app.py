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
