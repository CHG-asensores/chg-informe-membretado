"""
Servidor web para generar informes membretados CHG Ascensores.
Reemplaza el workflow n8n — corre localmente sin Docker.

Uso:
    pip install flask pymupdf
    python app.py
"""
import io
import os
import sys
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request, send_file

sys.path.insert(0, str(Path(__file__).parent / "src"))
from maintainx_parser import parse_pdf

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150 MB (varios PDFs)

BASE_DIR      = Path(__file__).parent
MEMBRETE_PATH = BASE_DIR / "template" / "assets" / "membrete.png"

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

    .carousel { display: flex; flex-direction: column; gap: 10px; }
    .carousel-track {
      display: flex; align-items: stretch; gap: 10px;
    }
    .carousel-arrow {
      background: #fff; border: 1px solid #e2e8f0; border-radius: 8px;
      width: 36px; flex-shrink: 0; cursor: pointer; font-size: 16px;
      color: #475569; display: flex; align-items: center; justify-content: center;
      transition: background .2s, color .2s, opacity .2s;
    }
    .carousel-arrow:hover:not(:disabled) { background: #f1f5f9; color: #1e293b; }
    .carousel-arrow:disabled { opacity: .35; cursor: not-allowed; }
    .carousel-slide { flex: 1; min-width: 0; }
    .carousel-pagination {
      text-align: center; font-size: 12px; color: #94a3b8;
      font-variant-numeric: tabular-nums;
    }

    .result-meta-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px 16px; }
    .result-meta-card .card-title { font-weight: 600; font-size: 13px; color: #1e293b; margin-bottom: 8px; }
    .result-meta-card .card-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 16px; font-size: 13px; color: #475569; }
    .result-meta-card .item-label { color: #94a3b8; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; }
    .result-meta-card .item-value { font-weight: 500; color: #1e293b; margin-top: 2px; }
    .result-meta-card.error { border-color: #fecaca; background: #fef2f2; }
    .result-meta-card.error .item-value { color: #991b1b; }

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
      font-size: 13px; color: #334155; margin-top: 8px;
      user-select: none; cursor: pointer;
    }
    .approve-row input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; }

    .drive-status { font-size: 12px; margin-top: 2px; }
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
  <div class="panel">
    <div class="panel-title">
      <div class="step">1</div>
      Subir PDF(s) de MaintainX
    </div>

    <div class="drop-zone" id="dropZone">
      <input type="file" id="fileInput" accept=".pdf" multiple>
      <div class="drop-icon">📄</div>
      <div class="drop-text">Arrastrá uno o más PDF aquí</div>
      <div class="drop-hint">o hacé clic para seleccionar</div>
    </div>

    <div class="file-list" id="fileList"></div>

    <div style="display:flex;flex-direction:column;gap:6px;margin-top:4px;"><label for="observaciones" style="font-size:13px;font-weight:600;color:#334155;text-transform:uppercase;">Observaciones</label><textarea id="observaciones" rows="4" placeholder="Ej: Se recomienda cambiar baterias..." style="width:100%;padding:10px 14px;border:1.5px solid #cbd5e1;border-radius:8px;font-size:14px;font-family:inherit;resize:vertical;background:#f8fafc;box-sizing:border-box;"></textarea></div>
    <button class="btn-generate" id="btnGenerate" disabled>
      <span>⚙️</span>
      Generar informe(s) membretado(s)
    </button>
  </div>
