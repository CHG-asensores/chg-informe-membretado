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
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF para overlay del membrete
from flask import Flask, jsonify, render_template_string, request, send_file
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent / "src"))
from maintainx_parser import parse_pdf
from render_html import render_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150 MB (varios PDFs)

BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "template" / "informe.html"
MEMBRETE_PATH = BASE_DIR / "template" / "assets" / "membrete.png"


# ─── Nombre de archivo de salida ──────────────────────────────────────────
# Formato requerido: "INF - MANT. {MES} EDIF. {NOMBRE DEL EDIFICIO} {N° INFORME}"
# Ejemplo:           "INF - MANT. JUNIO EDIF. DEL PARQUE 340"

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}


def extraer_nombre_edificio(activo: str) -> str:
    """A partir del campo ACTIVO (p.ej. "E. Alzadia - Ascensor 1" o
    "E. Santa Maria - Montacarga 1") extrae el nombre completo en
    mayúsculas, conservando el sufijo (ascensor, montacarga, etc.).
    Solo quita el prefijo "E." / "E ".
    """
    nombre = (activo or "").strip()
    nombre = re.sub(r"^\s*E\.?\s*", "", nombre, flags=re.IGNORECASE)
    return nombre.strip().upper()


def extraer_mes(meta: dict) -> str:
    """Determina el mes (en español, mayúsculas) a partir de la fecha de
    ingreso del informe. Si no se puede determinar, usa el mes actual."""
    fecha_ingreso = meta.get("fecha_hora_ingreso", "") or ""
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})", fecha_ingreso)
    if m:
        mes_num = int(m.group(2))
        if mes_num in MESES_ES:
            return MESES_ES[mes_num]

    import datetime
    return MESES_ES.get(datetime.date.today().month, "")


def build_output_filename(meta: dict) -> str:
    """Construye el nombre de salida del PDF según el formato:
    "INF - MANT. {MES} EDIF. {NOMBRE DEL EDIFICIO} [{NÚMERO}].pdf"
    El número del edificio se incluye solo si está presente en ubicacion.
    El número de orden ya no se incluye.
    """
    nombre_edif = extraer_nombre_edificio(meta.get("activo", ""))
    mes         = extraer_mes(meta)

    partes = ["INF - MANT.", mes, "EDIF."]
    if nombre_edif:
        partes.append(nombre_edif)

    nombre_archivo = " ".join(p for p in partes if p)
    nombre_archivo = re.sub(r"\s+", " ", nombre_archivo).strip()
    return f"{nombre_archivo}.pdf"


def html_to_pdf(html_bytes: bytes) -> bytes:
    """Renderiza HTML a PDF usando Chromium headless."""
    with sync_playwright() as p:
        browser = p.chromium.launch(args=[
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--no-sandbox",
            "--single-process",
            "--disable-extensions",
        ])
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
    out = src.tobytes(garbage=4, deflate=True, deflate_images=True, clean=True)
    src.close()
    return out


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
      grid-template-columns: 40% 60%;
      align-items: start;
      gap: 20px;
      padding: 24px;
      max-width: 1080px;
      width: 100%;
      margin: 0 auto;
    }

    .panel {
      background: #fff;
      border-radius: 12px;
      padding: 24px 16px;
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

    .file-list { display: flex; flex-direction: column; gap: 8px; }
    .file-item {
      display: flex;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 10px 14px;
      font-size: 14px;
      color: #334155;
      align-items: center;
      gap: 10px;
    }
    .file-item .ficon { font-size: 18px; flex-shrink: 0; line-height: 1; }
    .file-item .fname { font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .file-item .fsize { color: #94a3b8; font-size: 12px; flex-shrink: 0; }
    .file-item .remove-btn {
      background: none; border: none; cursor: pointer; color: #94a3b8; font-size: 18px; line-height: 1; padding: 0 2px;
    }
    .file-item .remove-btn:hover { color: #ef4444; }
    .file-list-summary { font-size: 12px; color: #94a3b8; padding: 0 2px; }

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
    .result-meta .card-actions {
      grid-column: 1 / -1;
      display: flex;
      gap: 10px;
      margin-top: 4px;
    }
    .result-meta .card-actions > * {
      flex: 1;
      margin-top: 0 !important;
      height: 38px;
      padding: 0 10px;
      font-size: 13px;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      border-radius: 10px;
      box-sizing: border-box;
    }

    .result-meta { padding: 0; border: none; background: none; display: block; }

    .carousel { display: flex; flex-direction: column; gap: 12px; }
    .carousel-track {
      position: relative;
      width: 100%;
    }
    .carousel-arrow {
      position: absolute;
      top: 50%;
      transform: translateY(-50%);
      background: #fff;
      border: 1px solid #e2e8f0;
      border-radius: 50%;
      width: 32px;
      height: 32px;
      flex-shrink: 0;
      cursor: pointer;
      font-size: 16px;
      color: #475569;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 2px 8px rgba(15,23,42,.10);
      z-index: 5;
      transition: background .15s, color .15s, opacity .15s;
    }
    .carousel-arrow:hover:not(:disabled) { background: #1a1a2e; color: #fff; }
    .carousel-arrow:disabled { opacity: .3; cursor: not-allowed; }
    .carousel-arrow.prev { left: 0; }
    .carousel-arrow.next { right: 0; }
    .carousel-slide {
      width: 100%;
      box-sizing: border-box;
      padding: 0 40px;
    }
    .carousel-dots {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
    }
    .carousel-dots .dot {
      width: 7px; height: 7px;
      border-radius: 50%;
      background: #cbd5e1;
      transition: background .15s, transform .15s;
    }
    .carousel-dots .dot.active {
      background: #1a1a2e;
      transform: scale(1.25);
    }
    .carousel-pagination {
      text-align: center; font-size: 12px; color: #94a3b8;
      font-variant-numeric: tabular-nums;
    }

    .result-meta-card {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 16px;
      padding: 18px 16px;
      display: flex;
      flex-direction: column;
      gap: 12px;
      box-shadow: 0 1px 2px rgba(15,23,42,.04);
      width: 100%;
      height: 350px;
      box-sizing: border-box;
      overflow: hidden;
    }
    .result-meta-card .card-header {
      display: flex;
      align-items: center;
      gap: 12px;
      flex-shrink: 0;
      height: 40px;
    }
    .result-meta-card .card-icon {
      width: 36px; height: 36px;
      flex-shrink: 0;
      border-radius: 50%;
      background: #e8efff;
      display: flex; align-items: center; justify-content: center;
      font-size: 17px;
    }
    .result-meta-card .card-title {
      font-weight: 800;
      font-size: 15px;
      line-height: 1.2;
      color: #0f172a;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      min-width: 0;
    }
    .result-meta-card .card-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      grid-template-rows: repeat(3, 1fr);
      gap: 8px;
      flex: 1;
      overflow: hidden;
    }
    .result-meta-card .card-grid > div {
      background: #fff;
      border: 1px solid #eef2f7;
      border-radius: 10px;
      padding: 7px 10px;
      box-sizing: border-box;
      display: flex;
      flex-direction: column;
      justify-content: center;
      overflow: hidden;
    }
    .result-meta-card .item-label {
      color: #94a3b8; font-size: 9.5px; font-weight: 600;
      text-transform: uppercase; letter-spacing: .06em;
      flex-shrink: 0;
    }
    .result-meta-card .item-value {
      font-weight: 700; font-size: 13px; color: #1e293b;
      margin-top: 2px; line-height: 1.2;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .result-meta-card.error { border-color: #fecaca; background: #fef2f2; }
    .result-meta-card.error .item-value { color: #991b1b; }
    .result-meta-card .card-approve {
      background: #fff; border: 1px solid #eef2f7;
      border-radius: 10px; padding: 8px 12px;
      flex-shrink: 0;
    }
    .result-meta-card .card-actions {
      display: flex; gap: 10px; flex-shrink: 0;
    }
    .result-meta-card .card-actions > * {
      flex: 1; margin-top: 0 !important;
      height: 38px; padding: 0 10px; font-size: 13px;
      display: flex; align-items: center; justify-content: center;
      gap: 6px; border-radius: 10px; box-sizing: border-box;
      white-space: nowrap;
    }

    .btn-preview-item {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      margin-top: 10px;
      background: #fff;
      color: #334155;
      border: 1.5px solid #cbd5e1;
      border-radius: 8px;
      padding: 10px 16px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: border-color .2s, color .2s;
    }
    .btn-preview-item:hover { border-color: #94a3b8; color: #1e293b; }

    .preview-modal-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(15, 23, 42, .55);
      z-index: 1000;
      align-items: center;
      justify-content: center;
      padding: 12px;
    }
    .preview-modal-overlay.show { display: flex; }
    .preview-modal {
      background: #fff;
      border-radius: 12px;
      width: 100%;
      max-width: 1100px;
      height: 94vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      box-shadow: 0 10px 40px rgba(0,0,0,.25);
    }
    .preview-modal-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px 16px;
      border-bottom: 1px solid #e2e8f0;
      font-size: 14px;
      font-weight: 600;
      color: #1e293b;
    }
    .preview-modal-header .fname-title {
      overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding-right: 12px;
    }
    .preview-modal-close {
      background: none;
      border: none;
      font-size: 22px;
      line-height: 1;
      cursor: pointer;
      color: #64748b;
      flex-shrink: 0;
      padding: 0 4px;
    }
    .preview-modal-close:hover { color: #ef4444; }
    .preview-modal iframe {
      flex: 1;
      width: 100%;
      border: none;
      background: #f1f5f9;
    }

    .btn-download-item {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      margin-top: 10px;
      background: #16a34a;
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 10px 16px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      text-decoration: none;
      transition: background .2s;
    }
    .btn-download-item:hover { background: #15803d; }

    .btn-download {
      background: #16a34a;
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 12px 20px;
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

    .btn-drive {
      background: #fff;
      color: #1a73e8;
      border: 1.5px solid #1a73e8;
      border-radius: 8px;
      padding: 12px 20px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      text-align: center;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      transition: background .2s, color .2s;
    }
    .btn-drive:hover:not(:disabled) { background: #1a73e8; color: #fff; }
    .btn-drive:disabled { opacity: .5; cursor: not-allowed; }

    .approve-row {
      display: flex; align-items: center; gap: 8px;
      font-size: 13px; color: #334155; margin-top: 16px; line-height: 1.5;
      user-select: none; cursor: pointer;
    }
    .approve-row input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; }

    .drive-status { font-size: 12px; margin-top: 0; }
    .drive-status.ok    { color: #166534; }
    .drive-status.error { color: #991b1b; }

    .btn-reset {
      background: none;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 10px 20px;
      font-size: 14px;
      color: #64748b;
      cursor: pointer;
      transition: border-color .2s, color .2s;
      margin-top: 0;
    }
    .btn-reset:hover { border-color: #94a3b8; color: #334155; }

    .result-card.show {
      display: flex;
      flex-direction: column;
      gap: 12px;
    }

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

<main id="mainGrid">
  <!-- ── Columna izquierda: subir PDF ── -->
  <div class="panel">
    <div class="panel-title">
      <div class="step">1</div>
      Subir PDF(s) de MaintainX
    </div>

    <div class="drop-zone" id="dropZone">
      <input type="file" id="fileInput" accept=".pdf" multiple>
      <div class="drop-icon">📄</div>
      <div class="drop-text">Arrastra aquí tus archivos PDF</div>
      <div class="drop-hint">o haz clic para seleccionarlos desde tu equipo</div>
    </div>

    <div class="file-list" id="fileList"></div>

    <button class="btn-generate" id="btnGenerate" disabled>
      <span>⚙️</span>
      Generar informe(s) membretado(s)
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
      <p>El/los informe(s) membretado(s) aparecerán aquí<br>una vez que subas y proceses el/los PDF.</p>
    </div>

    <div class="loading-overlay" id="loadingOverlay">
      <div class="spinner"></div>
      <div class="loading-text" id="loadingText">Procesando PDF…</div>
      <div class="loading-sub">Parseando datos y generando informe</div>
    </div>

    <div class="result-card" id="resultCard">
      <div class="result-badge" id="resultBadge"></div>

      <div class="result-meta" id="resultMeta"></div>

      <a class="btn-download" id="btnDownload" href="#" download>
        ⬇️ Descargar informe
      </a>

      <button class="btn-drive" id="btnDrive">
        📤 Subir aprobados a Drive
      </button>
      <div class="drive-status" id="driveStatus"></div>

      <button class="btn-reset" id="btnReset">Procesar otro(s) PDF</button>
    </div>
  </div>
</main>

<footer>CHG Ascensores · Informes Membretados</footer>

<div class="preview-modal-overlay" id="previewOverlay">
  <div class="preview-modal">
    <div class="preview-modal-header">
      <span class="fname-title" id="previewTitle">Vista previa</span>
      <button class="preview-modal-close" id="previewClose" title="Cerrar">✕</button>
    </div>
    <iframe id="previewFrame" src="" title="Vista previa del PDF"></iframe>
  </div>
</div>

<script>
  const dropZone    = document.getElementById('dropZone');
  const mainGrid    = document.getElementById('mainGrid');
  const fileInput   = document.getElementById('fileInput');
  const fileList    = document.getElementById('fileList');
  const btnGenerate = document.getElementById('btnGenerate');

  const resultEmpty   = document.getElementById('resultEmpty');
  const loadingOverlay = document.getElementById('loadingOverlay');
  const loadingText   = document.getElementById('loadingText');
  const resultCard    = document.getElementById('resultCard');
  const resultBadge   = document.getElementById('resultBadge');
  const resultMeta    = document.getElementById('resultMeta');
  const btnDownload   = document.getElementById('btnDownload');
  const btnDrive      = document.getElementById('btnDrive');
  const driveStatus   = document.getElementById('driveStatus');
  const btnReset      = document.getElementById('btnReset');

  let selectedFiles = [];

  function fmtSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024*1024) return (bytes/1024).toFixed(1) + ' KB';
    return (bytes/1024/1024).toFixed(1) + ' MB';
  }

  function renderFileList() {
    fileList.innerHTML = selectedFiles.map((file, idx) => `
      <div class="file-item">
        <span class="ficon">📄</span>
        <span class="fname">${file.name}</span>
        <span class="fsize">${fmtSize(file.size)}</span>
        <button class="remove-btn" data-idx="${idx}" title="Quitar archivo">✕</button>
      </div>
    `).join('');

    if (selectedFiles.length > 1) {
      fileList.innerHTML += `<div class="file-list-summary">${selectedFiles.length} archivos seleccionados</div>`;
    }

    fileList.querySelectorAll('.remove-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.idx, 10);
        selectedFiles.splice(idx, 1);
        renderFileList();
        updateGenerateState();
      });
    });
  }

  function updateGenerateState() {
    btnGenerate.disabled = selectedFiles.length === 0;
    resetResult();
  }

  function addFiles(fileListInput) {
    let rejected = false;
    for (const file of fileListInput) {
      if (file.type !== 'application/pdf') {
        rejected = true;
        continue;
      }
      selectedFiles.push(file);
    }
    if (rejected) {
      alert('Solo se admiten archivos PDF. Se ignoraron los archivos que no lo son.');
    }
    renderFileList();
    updateGenerateState();
  }

  function resetFiles() {
    selectedFiles = [];
    fileInput.value = '';
    renderFileList();
    updateGenerateState();
  }

  function resetResult() {
    resultCard.classList.remove('show');
    loadingOverlay.classList.remove('show');
    resultEmpty.style.display = '';
    if (driveStatus) { driveStatus.textContent = ''; driveStatus.className = 'drive-status'; }
  }

  function showLoading() {
    resultEmpty.style.display = 'none';
    resultCard.classList.remove('show');
    loadingOverlay.classList.add('show');
    if (selectedFiles.length > 1) {
      loadingText.textContent = `Procesando ${selectedFiles.length} PDFs…`;
    } else {
      loadingText.textContent = 'Procesando PDF…';
    }
  }

  function metaGridHtml(m) {
    return [
      ['N° de orden', '#' + (m.numero_orden || '—')],
      ['Estado',      m.estado || '—'],
      ['Ubicación',   m.ubicacion || '—'],
      ['Activo',      m.activo || '—'],
      ['Asignados',   (m.asignados || []).join(', ') || '—'],
      ['Campos',      m.campos_completados || '—'],
    ].map(([l, v]) => `
        <div>
          <div class="item-label">${l}</div>
          <div class="item-value" title="${v}">${v}</div>
        </div>`).join('');
  }

  // ── Carrusel de resúmenes (cuando se procesan varios PDFs) ──────────────
  let carouselSlides = [];   // array de strings HTML, una por archivo
  let carouselIndex  = 0;
  let approvedIds    = new Set();   // file_id de tarjetas marcadas "Aprobado"

  function renderCarousel() {
    const total = carouselSlides.length;
    if (total === 0) {
      resultMeta.innerHTML = '';
      return;
    }
    if (carouselIndex < 0) carouselIndex = 0;
    if (carouselIndex > total - 1) carouselIndex = total - 1;

    const dots = Array.from({ length: total }, (_, i) =>
      `<span class="dot${i === carouselIndex ? ' active' : ''}"></span>`
    ).join('');

    resultMeta.innerHTML = `
      <div class="carousel">
        <div class="carousel-track">
          <button class="carousel-arrow prev" id="carouselPrev" ${carouselIndex === 0 ? 'disabled' : ''} title="Anterior">‹</button>
          <div class="carousel-slide">${carouselSlides[carouselIndex]}</div>
          <button class="carousel-arrow next" id="carouselNext" ${carouselIndex === total - 1 ? 'disabled' : ''} title="Siguiente">›</button>
        </div>
        <div class="carousel-dots">${dots}</div>
        <div class="carousel-pagination">${carouselIndex + 1} / ${total}</div>
      </div>
    `;

    const prevBtn = document.getElementById('carouselPrev');
    const nextBtn = document.getElementById('carouselNext');
    if (prevBtn) prevBtn.addEventListener('click', () => { carouselIndex--; renderCarousel(); });
    if (nextBtn) nextBtn.addEventListener('click', () => { carouselIndex++; renderCarousel(); });

    const approveCb = document.getElementById('approveCheckbox');
    if (approveCb) {
      approveCb.addEventListener('change', () => {
        const fid = approveCb.dataset.fileId;
        if (approveCb.checked) approvedIds.add(fid);
        else approvedIds.delete(fid);
        updateDriveButtonState();
      });
    }
  }

  function approveCheckboxHtml(fileId) {
    if (!fileId) return '';
    const checked = approvedIds.has(fileId) ? 'checked' : '';
    return `
      <label class="approve-row card-approve" style="margin-top:0;">
        <input type="checkbox" id="approveCheckbox" data-file-id="${fileId}" ${checked}>
        Aprobado para subir a Drive
      </label>`;
  }

  function itemPreviewHtml(item, showDownload) {
    if (!item.file_id) return '';
    const dl = `/download/${item.file_id}`;
    const safeName = (item.filename || '').replace(/'/g, "\\'");
    let html = `<button class="btn-preview-item" onclick="openPreview('${item.file_id}', '${safeName}')">👁️ Ver vista previa</button>`;
    if (showDownload) {
      html += `<a class="btn-download-item" href="${dl}" download="${item.filename}">⬇️ Descargar este PDF</a>`;
    }
    return html;
  }

  const previewOverlay = document.getElementById('previewOverlay');
  const previewFrame   = document.getElementById('previewFrame');
  const previewTitle   = document.getElementById('previewTitle');
  const previewClose   = document.getElementById('previewClose');

  window.openPreview = function(fileId, filename) {
    previewTitle.textContent = filename || 'Vista previa';
    previewFrame.src = `/download/${fileId}?inline=1`;
    previewOverlay.classList.add('show');
  };

  function closePreview() {
    previewOverlay.classList.remove('show');
    previewFrame.src = '';
  }

  previewClose.addEventListener('click', closePreview);
  previewOverlay.addEventListener('click', (e) => {
    if (e.target === previewOverlay) closePreview();
  });

  function updateDriveButtonState() {
    btnDrive.disabled = approvedIds.size === 0;
    btnDrive.textContent = approvedIds.size > 0
      ? `📤 Subir ${approvedIds.size} aprobado(s) a Drive`
      : '📤 Subir aprobados a Drive';
    driveStatus.textContent = '';
    driveStatus.className = 'drive-status';
  }

  function showResult(data) {
    loadingOverlay.classList.remove('show');
    resultCard.classList.add('show');
    carouselSlides = [];
    carouselIndex = 0;
    approvedIds = new Set();
    btnDrive.style.display = '';
    updateDriveButtonState();

    if (data.ok) {
      if (data.multi) {
        const errCount = (data.errors || []).length;
        if (errCount > 0) {
          resultBadge.className = 'result-badge warning';
          resultBadge.innerHTML = `✅ ${data.count} informe(s) generado(s), ${errCount} con error`;
        } else {
          resultBadge.className = 'result-badge success';
          resultBadge.innerHTML = `✅ ${data.count} informes generados correctamente`;
        }
      } else {
        resultBadge.className = 'result-badge success';
        resultBadge.innerHTML = '✅ Informe generado correctamente';
      }
    } else {
      resultBadge.className = 'result-badge error';
      resultBadge.innerHTML = '❌ ' + (data.error || 'Error al procesar');
      btnDownload.style.display = 'none';
      btnDrive.style.display = 'none';
      if (data.errors && data.errors.length) {
        carouselSlides = data.errors.map(e => `
          <div class="result-meta-card error">
            <div class="card-title">${e.archivo}</div>
            <div class="item-value">${e.error}</div>
          </div>`);
      }
      renderCarousel();
      return;
    }

    btnDownload.style.display = '';
    btnDownload.href = data.download_url;
    btnDownload.download = data.filename;
    btnDownload.innerHTML = data.multi
      ? '⬇️ Descargar .zip con informes membretados'
      : '⬇️ Descargar PDF membretado';

    if (data.multi) {
      carouselSlides = (data.items || []).map(item => {
        const m = item.meta || {};
        const titulo = m.numero_orden ? `#${m.numero_orden} — ${m.ubicacion || item.filename}` : item.filename;
        return `
          <div class="result-meta-card">
            <div class="card-header">
              <div class="card-icon">🏢</div>
              <div class="card-title" title="${titulo}">${titulo}</div>
            </div>
            <div class="card-grid">${metaGridHtml(m)}</div>
            ${approveCheckboxHtml(item.file_id)}
            <div class="card-actions">${itemPreviewHtml(item, true)}</div>
          </div>`;
      });

      if (data.errors && data.errors.length) {
        carouselSlides = carouselSlides.concat(data.errors.map(e => `
          <div class="result-meta-card error">
            <div class="card-title">⚠️ ${e.archivo}</div>
            <div class="item-value">${e.error}</div>
          </div>`));
      }

      renderCarousel();
    } else {
      const m = data.meta || {};
      const titulo = m.numero_orden ? `#${m.numero_orden} — ${m.ubicacion || data.filename}` : data.filename;
      resultMeta.innerHTML = `
        <div class="result-meta-card">
          <div class="card-header">
            <div class="card-icon">🏢</div>
            <div class="card-title" title="${titulo}">${titulo}</div>
          </div>
          <div class="card-grid">${metaGridHtml(m)}</div>
          ${approveCheckboxHtml(data.file_id)}
          <div class="card-actions">${itemPreviewHtml(data, false)}</div>
        </div>`;
      const approveCb = document.getElementById('approveCheckbox');
      if (approveCb) {
        approveCb.addEventListener('change', () => {
          const fid = approveCb.dataset.fileId;
          if (approveCb.checked) approvedIds.add(fid);
          else approvedIds.delete(fid);
          updateDriveButtonState();
        });
      }
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
    if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) addFiles(fileInput.files);
    fileInput.value = '';
  });

  btnReset.addEventListener('click', () => {
    resetFiles();
  });

  btnGenerate.addEventListener('click', async () => {
    if (!selectedFiles.length) return;
    showLoading();
    btnGenerate.disabled = true;

    const form = new FormData();
    selectedFiles.forEach(file => form.append('pdf', file));

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

  btnDrive.addEventListener('click', async () => {
    if (approvedIds.size === 0) return;
    btnDrive.disabled = true;
    driveStatus.className = 'drive-status';
    driveStatus.textContent = '⏳ Subiendo a Drive…';

    try {
      const resp = await fetch('/upload-to-drive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_ids: Array.from(approvedIds) }),
      });
      const data = await resp.json();

      if (resp.status === 501) {
        driveStatus.className = 'drive-status error';
        driveStatus.textContent = '⚠️ ' + data.error;
      } else if (data.ok) {
        driveStatus.className = 'drive-status ok';
        driveStatus.textContent = `✅ ${data.results.length} archivo(s) subido(s) a Drive correctamente.`;
      } else {
        const failed = (data.results || []).filter(r => !r.ok);
        driveStatus.className = 'drive-status error';
        driveStatus.textContent = `⚠️ ${failed.length} archivo(s) fallaron: `
          + failed.map(f => `${f.filename || f.file_id} (${f.error})`).join(', ');
      }
    } catch (err) {
      driveStatus.className = 'drive-status error';
      driveStatus.textContent = '❌ Error de red: ' + err.message;
    } finally {
      btnDrive.disabled = approvedIds.size === 0;
    }
  });
</script>
</body>
</html>"""

@app.route("/")


def index():
    return render_template_string(UI)


_pdf_store: dict = {}


def process_one_pdf(pdf_bytes: bytes):
    """Procesa un PDF de MaintainX y retorna (filename, pdf_out_bytes, meta).
    Lanza excepción con mensaje descriptivo si algo falla."""
    try:
        data = parse_pdf(pdf_bytes)
    except Exception as exc:
        raise RuntimeError(f"Error al parsear el PDF: {exc}")

    try:
        html_b64 = render_html(data, str(TEMPLATE_PATH), str(MEMBRETE_PATH))
        html_bytes = base64.b64decode(html_b64)
    except Exception as exc:
        raise RuntimeError(f"Error al renderizar el HTML: {exc}")

    try:
        pdf_raw = html_to_pdf(html_bytes)
    except Exception as exc:
        raise RuntimeError(f"Error al generar PDF con Playwright: {exc}")

    try:
        pdf_out = apply_letterhead(pdf_raw, str(MEMBRETE_PATH))
    except Exception as exc:
        raise RuntimeError(f"Error al aplicar membrete: {exc}")

    meta     = data.get("meta", {})
    filename = build_output_filename(meta)
    return filename, pdf_out, meta


@app.route("/generate", methods=["POST"])
def generate():
    files = request.files.getlist("pdf")
    if not files:
        return jsonify({"error": "No se recibió ningún archivo PDF."}), 400

    import hashlib, time, zipfile

    # ─── Un solo archivo: comportamiento original (PDF directo) ───────────
    if len(files) == 1:
        original_name = files[0].filename or "archivo.pdf"
        pdf_bytes = files[0].read()
        if not pdf_bytes:
            return jsonify({"error": "El archivo está vacío."}), 400

        try:
            filename, pdf_out, meta = process_one_pdf(pdf_bytes)
        except RuntimeError as exc:
            return jsonify({"error": str(exc)}), 500

        file_id = hashlib.md5(f"{filename}{time.time()}".encode()).hexdigest()[:12]
        _pdf_store[file_id] = (filename, pdf_out, filename)

        return jsonify({
            "ok":           True,
            "multi":        False,
            "meta":         meta,
            "filename":     filename,
            "file_id":      file_id,
            "download_url": f"/download/{file_id}",
        })

    # ─── Múltiples archivos: procesar en cola y empaquetar en .zip ─────────
    results = []   # [(filename, pdf_bytes)]
    items   = []   # [{"filename": ..., "meta": {...}}]
    errors  = []   # [{"archivo": nombre_original, "error": mensaje}]

    for f in files:
        original_name = f.filename or "archivo.pdf"
        pdf_bytes = f.read()
        if not pdf_bytes:
            errors.append({"archivo": original_name, "error": "Archivo vacío"})
            continue
        try:
            filename, pdf_out, meta = process_one_pdf(pdf_bytes)
            # Evitar nombres duplicados dentro del zip
            base, ext = os.path.splitext(filename)
            candidate = filename
            n = 1
            existentes = {r[0] for r in results}
            while candidate in existentes:
                candidate = f"{base}_{n}{ext}"
                n += 1
            results.append((candidate, pdf_out))
            individual_id = hashlib.md5(f"{candidate}{time.time()}{len(items)}".encode()).hexdigest()[:12]
            _pdf_store[individual_id] = (candidate, pdf_out, candidate)
            items.append({"filename": candidate, "meta": meta, "file_id": individual_id})
        except RuntimeError as exc:
            errors.append({"archivo": original_name, "error": str(exc)})

    if not results:
        return jsonify({
            "ok":     False,
            "error":  "No se pudo procesar ningún archivo.",
            "errors": errors,
        }), 500

    # Crear .zip en memoria
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, pdf_bytes_ in results:
            zf.writestr(fname, pdf_bytes_)
    zip_buf.seek(0)

    zip_filename = f"informes-membretados-{int(time.time())}.zip"
    file_id = hashlib.md5(f"{zip_filename}{time.time()}".encode()).hexdigest()[:12]
    _pdf_store[file_id] = (zip_filename, zip_buf.getvalue(), zip_filename)

    return jsonify({
        "ok":           True,
        "multi":        True,
        "count":        len(results),
        "items":        items,
        "errors":       errors,
        "filename":     zip_filename,
        "download_url": f"/download/{file_id}",
    })


@app.route("/download/<file_id>")
def download(file_id):
    if file_id not in _pdf_store:
        return "Archivo no encontrado o expirado.", 404
    filename, file_bytes, _ = _pdf_store[file_id]
    mimetype = "application/zip" if filename.lower().endswith(".zip") else "application/pdf"

    # Si se pide como vista previa (inline), no forzar descarga
    inline = request.args.get("inline", "").strip() in ("1", "true", "yes")

    resp = send_file(
        io.BytesIO(file_bytes),
        mimetype=mimetype,
        as_attachment=not inline,
        download_name=filename,
    )
    resp.headers["Content-Length"] = str(len(file_bytes))
    return resp


@app.route("/upload-to-drive", methods=["POST"])
def upload_to_drive():
    """Reenvía los PDFs aprobados (por file_id) a un webhook de n8n,
    que se encarga de subirlos a Google Drive.

    Configurar la variable de entorno N8N_WEBHOOK_URL con la URL del
    webhook de n8n (nodo Webhook -> Google Drive Upload).
    """
    webhook_url = os.environ.get("N8N_WEBHOOK_URL", "").strip()
    if not webhook_url:
        return jsonify({
            "ok": False,
            "error": "La integración con n8n aún no está configurada "
                     "(falta la variable de entorno N8N_WEBHOOK_URL).",
        }), 501

    payload = request.get_json(silent=True) or {}
    file_ids = payload.get("file_ids", [])
    if not file_ids:
        return jsonify({"ok": False, "error": "No se indicó ningún archivo aprobado."}), 400

    import requests

    results = []
    for fid in file_ids:
        if fid not in _pdf_store:
            results.append({"file_id": fid, "ok": False, "error": "Archivo no encontrado o expirado."})
            continue

        filename, file_bytes, drive_name = _pdf_store[fid]
        try:
            resp = requests.post(
                webhook_url,
                files={"data": (drive_name, file_bytes, "application/pdf")},
                data={"filename": drive_name},
                timeout=60,
            )
            if resp.ok:
                results.append({"file_id": fid, "filename": drive_name, "ok": True})
            else:
                results.append({
                    "file_id": fid, "filename": drive_name, "ok": False,
                    "error": f"n8n respondió con estado {resp.status_code}",
                })
        except Exception as exc:
            results.append({"file_id": fid, "filename": drive_name, "ok": False, "error": str(exc)})

    all_ok = all(r["ok"] for r in results)
    return jsonify({"ok": all_ok, "results": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("\n  CHG Informe Membretado")
    print("  -------------------------------------")
    print(f"  Servidor:  http://localhost:{port}")
    print("  Motor PDF: Playwright (Chromium)")
    print("  -------------------------------------\n")
    app.run(host="0.0.0.0", port=port, debug=True)
