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
from flask import Flask, jsonify, render_template_string, request, send_file, session, redirect, url_for
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent / "src"))
from maintainx_parser import parse_pdf
from render_html import render_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150 MB (varios PDFs)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "chg-ascensores-secret-2024")
APP_PASSWORD = "BVm8i5nM9YEbtB11M"
LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAC0AhwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+f+iiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiijB9KACiijB9KACijB9KMH0oAKKMH0owfSgAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACgdaKBjPNADyCvRcfWhcDjfjj0rS8NeG9U8Wa/beHNDtHuLq7mEcMa9zX054f8A2Lvhta6Gtn4ivr27vnj/AH1zDMIljb/YU/8As9fRZLwzmufuTwsfdj3PIzLPMuyhR+sP4j5TBBG/POelG7KlWOMnOa7r43/BfWPg34gWwkY3OnXQZ7C8MZXcP7rf7S9xWX8IfDOj+MfiJpPhvXr1re1vbsRTPH94j+6P977v4151XK8ZRzD6lVjy1OblO2GOw1bB/WYS5oW5jmMx/wCSaN0f+Sa+xf8Ahkf4F/8AQs3H/gdL/wDF07/hkb4Ff9CxN/4Hy/8Axdfarwy4ga+KH/gT/wDkT5f/AF4ya/wz+5f/ACR8cbo/8k0Zj7fzNfY//DI3wK/6Fib/AMD5f/i6P+GRvgV/0LE3/gfL/wDF0f8AEMeIP5of+BP/AORD/XjJv5Z/cv8A5I+OASxAIJNLgDIYED2r7GP7JHwLPXwzN/4Hy/8AxdfOn7QPgPwz8OfiPd+GPCdw7WyRROI3k3tEzICVLd//ALKvHzzhDNMgw3t8RKPLfl909TK+I8uzis6VHm5l3R5+evFFOSMyybEqdLOP/lpJXyJ7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"

BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "template" / "informe.html"
MEMBRETE_PATH = BASE_DIR / "template" / "assets" / "membrete.png"

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

_SOLO_EQUIPO_RE = re.compile(
    r"^(Plataforma|Montacarga|Ascensor|Monta\s*carga|Monta\s*plato)\s*\d*\s*$",
    re.IGNORECASE,
)


def extraer_nombre_edificio(activo: str, ubicacion: str = "") -> str:
    nombre = (activo or "").strip()
    nombre = re.sub(r"^\s*E\.?\s*", "", nombre, flags=re.IGNORECASE).strip()
    if _SOLO_EQUIPO_RE.match(nombre):
        nombre = re.sub(r"^\s*Edificio\s+", "", (ubicacion or ""), flags=re.IGNORECASE).strip()
        if not nombre:
            nombre = (activo or "").strip()
    return nombre.upper()


def extraer_mes(meta: dict) -> str:
    fecha_ingreso = meta.get("fecha_hora_ingreso", "") or ""
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})", fecha_ingreso)
    if m:
        mes_num = int(m.group(2))
        if mes_num in MESES_ES:
            return MESES_ES[mes_num]
    import datetime
    return MESES_ES.get(datetime.date.today().month, "")


def build_output_filename(meta: dict) -> str:
    nombre_edif = extraer_nombre_edificio(
        meta.get("activo", ""),
        meta.get("ubicacion", ""),
    )
    mes = extraer_mes(meta)
    partes = ["INF - MANT.", mes, "EDIF."]
    if nombre_edif:
        partes.append(nombre_edif)
    nombre_archivo = " ".join(p for p in partes if p)
    nombre_archivo = re.sub(r"\s+", " ", nombre_archivo).strip()
    return f"{nombre_archivo}.pdf"


def html_to_pdf(html_bytes: bytes) -> bytes:
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
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    with open(membrete_path, "rb") as f:
        membrete_bytes = f.read()
    for page in src:
        page.insert_image(page.rect, stream=membrete_bytes, overlay=False)
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
    header .logo { font-size: 20px; font-weight: 700; letter-spacing: -0.5px; }
    header .logo span { color: #60a5fa; }
    header .sub { font-size: 13px; color: #94a3b8; margin-left: auto; }
    header .btn-logout {
      background: none; border: 1px solid #334155; color: #94a3b8;
      border-radius: 6px; padding: 6px 14px; font-size: 12px; cursor: pointer;
      margin-left: 12px; transition: border-color .2s, color .2s; text-decoration: none;
    }
    header .btn-logout:hover { border-color: #94a3b8; color: #fff; }

    main {
      flex: 1; display: grid; grid-template-columns: 40% 60%;
      align-items: start; gap: 20px; padding: 24px;
      max-width: 1080px; width: 100%; margin: 0 auto;
    }

    .panel {
      background: #fff; border-radius: 12px; padding: 24px 16px;
      box-shadow: 0 1px 3px rgba(0,0,0,.08), 0 4px 16px rgba(0,0,0,.04);
      display: flex; flex-direction: column; gap: 18px;
    }

    .panel-title {
      font-size: 13px; font-weight: 600; text-transform: uppercase;
      letter-spacing: .06em; color: #64748b; display: flex; align-items: center; gap: 8px;
    }
    .panel-title .step {
      background: #1a1a2e; color: #fff; border-radius: 50%;
      width: 22px; height: 22px; font-size: 11px;
      display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }

    .drop-zone {
      border: 2px dashed #cbd5e1; border-radius: 10px; padding: 48px 24px;
      text-align: center; cursor: pointer; transition: border-color .2s, background .2s; position: relative;
    }
    .drop-zone:hover, .drop-zone.over { border-color: #3b82f6; background: #eff6ff; }
    .drop-zone input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
    .drop-icon { font-size: 40px; margin-bottom: 12px; }
    .drop-text { font-size: 15px; font-weight: 500; color: #334155; margin-bottom: 6px; }
    .drop-hint { font-size: 13px; color: #94a3b8; }

    .file-list { display: flex; flex-direction: column; gap: 8px; }
    .file-item {
      display: flex; background: #f8fafc; border: 1px solid #e2e8f0;
      border-radius: 8px; padding: 10px 14px; font-size: 14px; color: #334155; align-items: center; gap: 10px;
    }
    .file-item .ficon { font-size: 18px; flex-shrink: 0; line-height: 1; }
    .file-item .fname { font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .file-item .fsize { color: #94a3b8; font-size: 12px; flex-shrink: 0; }
    .file-item .remove-btn { background: none; border: none; cursor: pointer; color: #94a3b8; font-size: 18px; line-height: 1; padding: 0 2px; }
    .file-item .remove-btn:hover { color: #ef4444; }
    .file-list-summary { font-size: 12px; color: #94a3b8; padding: 0 2px; }

    .btn-generate {
      background: #1a1a2e; color: #fff; border: none; border-radius: 8px;
      padding: 13px 20px; font-size: 15px; font-weight: 600; cursor: pointer;
      display: flex; align-items: center; justify-content: center; gap: 8px;
      transition: background .2s, opacity .2s;
    }
    .btn-generate:hover { background: #2d2d4e; }
    .btn-generate:disabled { opacity: .5; cursor: not-allowed; }

    .result-empty {
      flex: 1; display: flex; flex-direction: column; align-items: center;
      justify-content: center; gap: 12px; color: #94a3b8; padding: 48px 24px; text-align: center;
    }
    .result-empty .icon { font-size: 48px; }
    .result-empty p { font-size: 14px; line-height: 1.6; }

    .result-card { display: none; flex-direction: column; gap: 12px; }
    .result-card.show { display: flex; }

    .result-badge {
      display: flex; align-items: center; gap: 10px; padding: 12px 16px;
      border-radius: 8px; font-size: 14px; font-weight: 500;
    }
    .result-badge.success { background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }
    .result-badge.error   { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
    .result-badge.warning { background: #fffbeb; color: #92400e; border: 1px solid #fde68a; }

    .result-meta { background: none; border: none; display: block; }

    .carousel { display: flex; flex-direction: column; gap: 12px; }
    .carousel-track { position: relative; width: 100%; }
    .carousel-arrow {
      position: absolute; top: 50%; transform: translateY(-50%);
      background: #fff; border: 1px solid #e2e8f0; border-radius: 50%;
      width: 32px; height: 32px; flex-shrink: 0; cursor: pointer;
      font-size: 16px; color: #475569; display: flex; align-items: center; justify-content: center;
      box-shadow: 0 2px 8px rgba(15,23,42,.10); z-index: 5; transition: background .15s, color .15s, opacity .15s;
    }
    .carousel-arrow:hover:not(:disabled) { background: #1a1a2e; color: #fff; }
    .carousel-arrow:disabled { opacity: .3; cursor: not-allowed; }
    .carousel-arrow.prev { left: 0; }
    .carousel-arrow.next { right: 0; }
    .carousel-slide { width: 100%; box-sizing: border-box; padding: 0 40px; }
    .carousel-dots { display: flex; align-items: center; justify-content: center; gap: 6px; }
    .carousel-dots .dot { width: 7px; height: 7px; border-radius: 50%; background: #cbd5e1; transition: background .15s, transform .15s; }
    .carousel-dots .dot.active { background: #1a1a2e; transform: scale(1.25); }
    .carousel-pagination { text-align: center; font-size: 12px; color: #94a3b8; font-variant-numeric: tabular-nums; }

    .result-meta-card {
      background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 16px;
      padding: 18px 16px; display: flex; flex-direction: column; gap: 12px;
      box-shadow: 0 1px 2px rgba(15,23,42,.04); width: 100%; height: 350px;
      box-sizing: border-box; overflow: hidden;
    }
    .result-meta-card .card-header { display: flex; align-items: center; gap: 12px; flex-shrink: 0; height: 40px; }
    .result-meta-card .card-icon {
      width: 36px; height: 36px; flex-shrink: 0; border-radius: 50%;
      background: #e8efff; display: flex; align-items: center; justify-content: center; font-size: 17px;
    }
    .result-meta-card .card-title {
      font-weight: 800; font-size: 15px; line-height: 1.2; color: #0f172a;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0;
    }
    .result-meta-card .card-grid {
      display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: repeat(3, 1fr);
      gap: 8px; flex: 1; overflow: hidden;
    }
    .result-meta-card .card-grid > div {
      background: #fff; border: 1px solid #eef2f7; border-radius: 10px;
      padding: 7px 10px; box-sizing: border-box; display: flex; flex-direction: column; justify-content: center; overflow: hidden;
    }
    .result-meta-card .item-label { color: #94a3b8; font-size: 9.5px; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; flex-shrink: 0; }
    .result-meta-card .item-value { font-weight: 700; font-size: 13px; color: #1e293b; margin-top: 2px; line-height: 1.2; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .result-meta-card.error { border-color: #fecaca; background: #fef2f2; }
    .result-meta-card.error .item-value { color: #991b1b; }
    .result-meta-card .card-approve { background: #fff; border: 1px solid #eef2f7; border-radius: 10px; padding: 8px 12px; flex-shrink: 0; }
    .result-meta-card .card-actions { display: flex; gap: 10px; flex-shrink: 0; }
    .result-meta-card .card-actions > * {
      flex: 1; margin-top: 0 !important; height: 38px; padding: 0 10px; font-size: 13px;
      display: flex; align-items: center; justify-content: center;
      gap: 6px; border-radius: 10px; box-sizing: border-box; white-space: nowrap;
    }

    .btn-preview-item {
      display: flex; align-items: center; justify-content: center; gap: 6px; margin-top: 10px;
      background: #fff; color: #334155; border: 1.5px solid #cbd5e1; border-radius: 8px;
      padding: 10px 16px; font-size: 13px; font-weight: 600; cursor: pointer;
      transition: border-color .2s, color .2s;
    }
    .btn-preview-item:hover { border-color: #94a3b8; color: #1e293b; }

    .preview-modal-overlay {
      display: none; position: fixed; inset: 0; background: rgba(15, 23, 42, .55);
      z-index: 1000; align-items: center; justify-content: center; padding: 12px;
    }
    .preview-modal-overlay.show { display: flex; }
    .preview-modal {
      background: #fff; border-radius: 12px; width: 100%; max-width: 1100px; height: 94vh;
      display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 10px 40px rgba(0,0,0,.25);
    }
    .preview-modal-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 12px 16px; border-bottom: 1px solid #e2e8f0; font-size: 14px; font-weight: 600; color: #1e293b;
    }
    .preview-modal-header .fname-title { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding-right: 12px; }
    .preview-modal-close { background: none; border: none; font-size: 22px; line-height: 1; cursor: pointer; color: #64748b; flex-shrink: 0; padding: 0 4px; }
    .preview-modal-close:hover { color: #ef4444; }
    .preview-modal iframe { flex: 1; width: 100%; border: none; background: #f1f5f9; }

    .btn-download-item {
      display: flex; align-items: center; justify-content: center; gap: 6px; margin-top: 10px;
      background: #16a34a; color: #fff; border: none; border-radius: 8px;
      padding: 10px 16px; font-size: 13px; font-weight: 600; cursor: pointer;
      text-decoration: none; transition: background .2s;
    }
    .btn-download-item:hover { background: #15803d; }

    .btn-download {
      background: #16a34a; color: #fff; border: none; border-radius: 8px;
      padding: 12px 20px; font-size: 15px; font-weight: 600; cursor: pointer;
      text-align: center; text-decoration: none; display: flex;
      align-items: center; justify-content: center; gap: 8px; transition: background .2s;
    }
    .btn-download:hover { background: #15803d; }

    .btn-drive {
      background: #fff; color: #1a73e8; border: 1.5px solid #1a73e8; border-radius: 8px;
      padding: 12px 20px; font-size: 15px; font-weight: 600; cursor: pointer;
      text-align: center; display: flex; align-items: center; justify-content: center; gap: 8px;
      transition: background .2s, color .2s;
    }
    .btn-drive:hover:not(:disabled) { background: #1a73e8; color: #fff; }
    .btn-drive:disabled { opacity: .5; cursor: not-allowed; }

    .approve-row {
      display: flex; align-items: center; gap: 8px; font-size: 13px; color: #334155;
      margin-top: 16px; line-height: 1.5; user-select: none; cursor: pointer;
    }
    .approve-row input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; }

    .drive-status { font-size: 12px; margin-top: 0; }
    .drive-status.ok    { color: #166534; }
    .drive-status.error { color: #991b1b; }

    .btn-reset {
      background: none; border: 1px solid #e2e8f0; border-radius: 8px;
      padding: 10px 20px; font-size: 14px; color: #64748b; cursor: pointer;
      transition: border-color .2s, color .2s; margin-top: 0;
    }
    .btn-reset:hover { border-color: #94a3b8; color: #334155; }

    .loading-overlay {
      display: none; flex-direction: column; align-items: center;
      justify-content: center; gap: 16px; padding: 48px 24px; flex: 1;
    }
    .loading-overlay.show { display: flex; }
    .spinner {
      width: 44px; height: 44px; border: 4px solid #e2e8f0;
      border-top-color: #1a1a2e; border-radius: 50%; animation: spin .8s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .loading-text { font-size: 15px; font-weight: 500; color: #334155; }
    .loading-sub  { font-size: 13px; color: #94a3b8; }

    footer { text-align: center; padding: 16px; font-size: 12px; color: #94a3b8; }

    @media (max-width: 700px) { main { grid-template-columns: 1fr; padding: 16px; } }
  </style>
</head>
<body>

<header>
  <div class="logo">CHG <span>Ascensores</span></div>
  <div class="sub">Generador de Informes Membretados</div>
  <a href="/logout" class="btn-logout">Cerrar sesión</a>
</header>

<main id="mainGrid">
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
      <a class="btn-download" id="btnDownload" href="#" download>⬇️ Descargar informe</a>
      <button class="btn-drive" id="btnDrive">📤 Subir aprobados a Drive</button>
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
        selectedFiles.splice(parseInt(btn.dataset.idx, 10), 1);
        renderFileList(); updateGenerateState();
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
      if (file.type !== 'application/pdf') { rejected = true; continue; }
      selectedFiles.push(file);
    }
    if (rejected) alert('Solo se admiten archivos PDF.');
    renderFileList(); updateGenerateState();
  }

  function resetFiles() {
    selectedFiles = []; fileInput.value = '';
    renderFileList(); updateGenerateState();
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
    loadingText.textContent = selectedFiles.length > 1 ? `Procesando ${selectedFiles.length} PDFs…` : 'Procesando PDF…';
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

  let carouselSlides = [];
  let carouselIndex  = 0;
  let approvedIds    = new Set();

  function renderCarousel() {
    const total = carouselSlides.length;
    if (total === 0) { resultMeta.innerHTML = ''; return; }
    if (carouselIndex < 0) carouselIndex = 0;
    if (carouselIndex > total - 1) carouselIndex = total - 1;
    const dots = Array.from({ length: total }, (_, i) =>
      `<span class="dot${i === carouselIndex ? ' active' : ''}"></span>`).join('');
    resultMeta.innerHTML = `
      <div class="carousel">
        <div class="carousel-track">
          <button class="carousel-arrow prev" id="carouselPrev" ${carouselIndex === 0 ? 'disabled' : ''}>‹</button>
          <div class="carousel-slide">${carouselSlides[carouselIndex]}</div>
          <button class="carousel-arrow next" id="carouselNext" ${carouselIndex === total - 1 ? 'disabled' : ''}>›</button>
        </div>
        <div class="carousel-dots">${dots}</div>
        <div class="carousel-pagination">${carouselIndex + 1} / ${total}</div>
      </div>`;
    document.getElementById('carouselPrev')?.addEventListener('click', () => { carouselIndex--; renderCarousel(); });
    document.getElementById('carouselNext')?.addEventListener('click', () => { carouselIndex++; renderCarousel(); });
    const approveCb = document.getElementById('approveCheckbox');
    if (approveCb) {
      approveCb.addEventListener('change', () => {
        const fid = approveCb.dataset.fileId;
        if (approveCb.checked) approvedIds.add(fid); else approvedIds.delete(fid);
        updateDriveButtonState();
      });
    }
  }

  function approveCheckboxHtml(fileId) {
    if (!fileId) return '';
    const checked = approvedIds.has(fileId) ? 'checked' : '';
    return `<label class="approve-row card-approve" style="margin-top:0;">
      <input type="checkbox" id="approveCheckbox" data-file-id="${fileId}" ${checked}>
      Aprobado para subir a Drive
    </label>`;
  }

  function itemPreviewHtml(item, showDownload) {
    if (!item.file_id) return '';
    const dl = `/download/${item.file_id}`;
    const safeName = (item.filename || '').replace(/'/g, "\\'");
    let html = `<button class="btn-preview-item" onclick="openPreview('${item.file_id}', '${safeName}')">👁️ Ver vista previa</button>`;
    if (showDownload) html += `<a class="btn-download-item" href="${dl}" download="${item.filename}">⬇️ Descargar este PDF</a>`;
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
  previewClose.addEventListener('click', () => { previewOverlay.classList.remove('show'); previewFrame.src = ''; });
  previewOverlay.addEventListener('click', (e) => { if (e.target === previewOverlay) { previewOverlay.classList.remove('show'); previewFrame.src = ''; } });

  function updateDriveButtonState() {
    btnDrive.disabled = approvedIds.size === 0;
    btnDrive.textContent = approvedIds.size > 0 ? `📤 Subir ${approvedIds.size} aprobado(s) a Drive` : '📤 Subir aprobados a Drive';
    driveStatus.textContent = ''; driveStatus.className = 'drive-status';
  }

  function showResult(data) {
    loadingOverlay.classList.remove('show');
    resultCard.classList.add('show');
    carouselSlides = []; carouselIndex = 0; approvedIds = new Set();
    btnDrive.style.display = ''; updateDriveButtonState();

    if (data.ok) {
      if (data.multi) {
        const errCount = (data.errors || []).length;
        resultBadge.className = errCount > 0 ? 'result-badge warning' : 'result-badge success';
        resultBadge.innerHTML = errCount > 0 ? `✅ ${data.count} informe(s) generado(s), ${errCount} con error` : `✅ ${data.count} informes generados correctamente`;
      } else {
        resultBadge.className = 'result-badge success';
        resultBadge.innerHTML = '✅ Informe generado correctamente';
      }
    } else {
      resultBadge.className = 'result-badge error';
      resultBadge.innerHTML = '❌ ' + (data.error || 'Error al procesar');
      btnDownload.style.display = 'none'; btnDrive.style.display = 'none';
      if (data.errors && data.errors.length) {
        carouselSlides = data.errors.map(e => `<div class="result-meta-card error"><div class="card-title">${e.archivo}</div><div class="item-value">${e.error}</div></div>`);
      }
      renderCarousel(); return;
    }

    btnDownload.style.display = '';
    btnDownload.href = data.download_url;
    btnDownload.download = data.filename;
    btnDownload.innerHTML = data.multi ? '⬇️ Descargar .zip con informes membretados' : '⬇️ Descargar PDF membretado';

    if (data.multi) {
      carouselSlides = (data.items || []).map(item => {
        const m = item.meta || {};
        const titulo = m.numero_orden ? `#${m.numero_orden} — ${m.ubicacion || item.filename}` : item.filename;
        return `<div class="result-meta-card">
          <div class="card-header"><div class="card-icon">🏢</div><div class="card-title" title="${titulo}">${titulo}</div></div>
          <div class="card-grid">${metaGridHtml(m)}</div>
          ${approveCheckboxHtml(item.file_id)}
          <div class="card-actions">${itemPreviewHtml(item, true)}</div>
        </div>`;
      });
      if (data.errors && data.errors.length) {
        carouselSlides = carouselSlides.concat(data.errors.map(e => `<div class="result-meta-card error"><div class="card-title">⚠️ ${e.archivo}</div><div class="item-value">${e.error}</div></div>`));
      }
      renderCarousel();
    } else {
      const m = data.meta || {};
      const titulo = m.numero_orden ? `#${m.numero_orden} — ${m.ubicacion || data.filename}` : data.filename;
      resultMeta.innerHTML = `<div class="result-meta-card">
        <div class="card-header"><div class="card-icon">🏢</div><div class="card-title" title="${titulo}">${titulo}</div></div>
        <div class="card-grid">${metaGridHtml(m)}</div>
        ${approveCheckboxHtml(data.file_id)}
        <div class="card-actions">${itemPreviewHtml(data, false)}</div>
      </div>`;
      const approveCb = document.getElementById('approveCheckbox');
      if (approveCb) {
        approveCb.addEventListener('change', () => {
          const fid = approveCb.dataset.fileId;
          if (approveCb.checked) approvedIds.add(fid); else approvedIds.delete(fid);
          updateDriveButtonState();
        });
      }
    }
  }

  ['dragenter','dragover'].forEach(evt => dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add('over'); }));
  ['dragleave','drop'].forEach(evt => dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove('over'); }));
  dropZone.addEventListener('drop', e => { if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files); });
  fileInput.addEventListener('change', () => { if (fileInput.files.length) addFiles(fileInput.files); fileInput.value = ''; });
  btnReset.addEventListener('click', resetFiles);

  btnGenerate.addEventListener('click', async () => {
    if (!selectedFiles.length) return;
    showLoading(); btnGenerate.disabled = true;
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
        driveStatus.textContent = `⚠️ ${failed.length} archivo(s) fallaron: ` + failed.map(f => `${f.filename || f.file_id} (${f.error})`).join(', ');
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

LOGIN_UI = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CHG — Acceso</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif; background: #f0f2f5; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
    .login-card { background: #fff; border-radius: 16px; padding: 40px 36px; width: 100%; max-width: 380px; box-shadow: 0 4px 24px rgba(0,0,0,.10); display: flex; flex-direction: column; align-items: center; gap: 24px; }
    .login-logo { display: flex; flex-direction: column; align-items: center; gap: 8px; }
    .login-logo-img { width: 260px; height: auto; }
    .login-title { font-size: 15px; color: #64748b; text-align: center; margin-top: -10px; }
    .login-form { width: 100%; display: flex; flex-direction: column; gap: 14px; }
    .login-form label { font-size: 13px; font-weight: 600; color: #475569; margin-bottom: 4px; display: block; }
    .login-form input[type=password] { width: 100%; padding: 11px 14px; border: 1.5px solid #e2e8f0; border-radius: 8px; font-size: 15px; color: #1e293b; outline: none; transition: border-color .2s; }
    .login-form input[type=password]:focus { border-color: #1a1a2e; }
    .login-btn { width: 100%; padding: 12px; background: #1a1a2e; color: #fff; border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; transition: background .2s; margin-top: 4px; }
    .login-btn:hover { background: #2d2d4e; }
    .login-error { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; border-radius: 8px; padding: 10px 14px; font-size: 13px; width: 100%; text-align: center; }
    .login-footer { font-size: 12px; color: #94a3b8; }
  </style>
</head>
<body>
  <div class="login-card">
    <div class="login-logo">
      <img class="login-logo-img" src="/static-logo" alt="CHG Logo">
    </div>
    <div class="login-title">Ingresa la contraseña para continuar</div>
    {% if error %}<div class="login-error">⚠️ Contraseña incorrecta. Intenta de nuevo.</div>{% endif %}
    <form class="login-form" method="POST" action="/login">
      <div>
        <label for="pwd">Contraseña</label>
        <input type="password" id="pwd" name="password" autofocus autocomplete="current-password">
      </div>
      <button type="submit" class="login-btn">Entrar</button>
    </form>
    <div class="login-footer">CHG Ascensores · Informes Membretados</div>
  </div>
</body>
</html>"""

import functools

def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("authenticated"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

@app.route("/login", methods=["GET", "POST"])
def login():
    error = False
    if request.method == "POST":
        pwd = request.form.get("password", "")
        if pwd == APP_PASSWORD:
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = True
    return render_template_string(LOGIN_UI, error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/static-logo")
def static_logo():
    img_bytes = base64.b64decode(LOGO_B64)
    return send_file(io.BytesIO(img_bytes), mimetype="image/jpeg")

@app.route("/")
@login_required
def index():
    return render_template_string(UI)

# _pdf_store guarda: (filename, pdf_bytes, drive_name, meta)
_pdf_store: dict = {}

def process_one_pdf(pdf_bytes: bytes):
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
    # Agregar observaciones para CHG al meta para enviarlas al webhook
    meta["para_chg"] = data.get("observaciones", {}).get("para_chg", "")
    filename = build_output_filename(meta)
    return filename, pdf_out, meta

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    files = request.files.getlist("pdf")
    if not files:
        return jsonify({"error": "No se recibió ningún archivo PDF."}), 400

    import hashlib, time, zipfile

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
        _pdf_store[file_id] = (filename, pdf_out, filename, meta)
        return jsonify({
            "ok":           True,
            "multi":        False,
            "meta":         meta,
            "filename":     filename,
            "file_id":      file_id,
            "download_url": f"/download/{file_id}",
        })

    results = []
    items   = []
    errors  = []

    for f in files:
        original_name = f.filename or "archivo.pdf"
        pdf_bytes = f.read()
        if not pdf_bytes:
            errors.append({"archivo": original_name, "error": "Archivo vacío"})
            continue
        try:
            filename, pdf_out, meta = process_one_pdf(pdf_bytes)
            base, ext = os.path.splitext(filename)
            candidate = filename
            n = 1
            existentes = {r[0] for r in results}
            while candidate in existentes:
                candidate = f"{base}_{n}{ext}"
                n += 1
            results.append((candidate, pdf_out))
            individual_id = hashlib.md5(f"{candidate}{time.time()}{len(items)}".encode()).hexdigest()[:12]
            _pdf_store[individual_id] = (candidate, pdf_out, candidate, meta)
            items.append({"filename": candidate, "meta": meta, "file_id": individual_id})
        except RuntimeError as exc:
            errors.append({"archivo": original_name, "error": str(exc)})

    if not results:
        return jsonify({"ok": False, "error": "No se pudo procesar ningún archivo.", "errors": errors}), 500

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, pdf_bytes_ in results:
            zf.writestr(fname, pdf_bytes_)
    zip_buf.seek(0)

    zip_filename = f"informes-membretados-{int(time.time())}.zip"
    file_id = hashlib.md5(f"{zip_filename}{time.time()}".encode()).hexdigest()[:12]
    _pdf_store[file_id] = (zip_filename, zip_buf.getvalue(), zip_filename, {})

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
@login_required
def download(file_id):
    if file_id not in _pdf_store:
        return "Archivo no encontrado o expirado.", 404
    filename, file_bytes, _, _meta = _pdf_store[file_id]
    mimetype = "application/zip" if filename.lower().endswith(".zip") else "application/pdf"
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
@login_required
def upload_to_drive():
    webhook_url = os.environ.get("N8N_WEBHOOK_URL", "").strip()
    if not webhook_url:
        return jsonify({
            "ok": False,
            "error": "La integración con n8n aún no está configurada (falta la variable de entorno N8N_WEBHOOK_URL).",
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

        filename, file_bytes, drive_name, meta = _pdf_store[fid]

        # Extraer datos del meta para enviar al webhook
        fecha_vencimiento = meta.get("fecha_vencimiento", "")
        activo            = meta.get("activo", "")
        numero_orden      = meta.get("numero_orden", "")
        obs_chg           = meta.get("para_chg", "") or meta.get("observaciones_chg", "") or ""

        try:
            resp = requests.post(
                webhook_url,
                files={"data": (drive_name, file_bytes, "application/pdf")},
                data={
                    "filename":          drive_name,
                    "fecha_vencimiento": fecha_vencimiento,
                    "activo":            activo,
                    "numero_orden":      numero_orden,
                    "observaciones_chg": obs_chg,
                },
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
