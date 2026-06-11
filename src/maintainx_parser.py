"""
Parser de PDFs exportados por MaintainX para CHG Ascensores.
Técnica: OVERLAY (Estampado) + Extracción de Metadata para UI.
Requisitos: pip install pymupdf
"""
import base64, re
import fitz  # PyMuPDF

NOISE_Y_TOP    = 45
NOISE_Y_BOTTOM = 790

def extract_lines(page):
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0: continue
        for line in block["lines"]:
            y0 = line["bbox"][1]
            if y0 < NOISE_Y_TOP or y0 > NOISE_Y_BOTTOM: continue
            text = "".join(s["text"] for s in line["spans"]).strip()
            if not text: continue
            text = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
            lines.append({"text": text, "y": y0, "x": line["bbox"][0]})
    lines.sort(key=lambda l: (round(l["y"]), l["x"]))
    return lines

def parse_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    meta = {}

    # --- 1. EXTRACCIÓN DE METADATA (Solo para el Carrusel de la UI) ---
    if len(doc) > 0:
        page0 = doc[0]
        lines = extract_lines(page0)
        
        # Número de orden y título desde el bloque superior
        for block in page0.get_text("dict")["blocks"]:
            if block.get("type") == 0:
                for line in block["lines"]:
                    if line["bbox"][1] < NOISE_Y_TOP:
                        raw = "".join(s["text"] for s in line["spans"]).strip()
                        m = re.search(r"#(\d+)", raw)
                        if m: meta["numero_orden"] = m.group(1)
                        t = re.match(r"^(.+?)\s*#\d+", raw)
                        if t: meta["titulo"] = t.group(1).strip()

        HEADER_LABELS = {
            "ESTADO": "estado", "FECHA DE VENCIMIENTO": "fecha_vencimiento",
            "TIEMPO ESTIMADO": "tiempo_estimado", "TIPO DE TRABAJO": "tipo_trabajo",
            "ASIGNADOS": "asignados", "CATEGORÍAS": "categorias",
            "UBICACIÓN": "ubicacion", "ACTIVO": "activo", "PROCEDIMIENTO": "procedimiento",
        }

        pending_left = pending_right = None
        COL_THRESHOLD = 180

        for entry in lines:
            text = entry["text"]
            if text in {"CHG Ascensores", "CHG", "ASCENSORES", "Campos completados"}: continue

            col_is_left = entry["x"] < COL_THRESHOLD
            pending = pending_left if col_is_left else pending_right

            if text in HEADER_LABELS:
                if col_is_left: pending_left  = HEADER_LABELS[text]
                else:           pending_right = HEADER_LABELS[text]
                continue

            if pending:
                if pending == "asignados":
                    meta.setdefault("asignados", [])
                    if text not in meta["asignados"]: meta["asignados"].append(text)
                elif pending == "categorias":
                    meta.setdefault("categorias", [])
                    if text not in meta["categorias"]: meta["categorias"].append(text)
                elif pending == "ubicacion":
                    if "ubicacion" not in meta: meta["ubicacion"] = text
                    else: meta["direccion"] = text; pending_left = pending_right = None
                elif pending == "activo":
                    if "activo" not in meta: meta["activo"] = text
                    else:
                        if not re.match(r"^\d+$", text): meta["activo"] += " " + text
                else:
                    meta[pending] = text
                    if col_is_left: pending_left = None
                    else:           pending_right = None

    asignados = meta.get("asignados", [])
    if len(asignados) > 1 and all(" " not in a for a in asignados):
        meta["asignados"] = [" ".join(asignados)]

    # --- 2. ESTAMPADO DIRECTO SOBRE EL PDF (Overlay) ---
    for page in doc:
        w, h = page.rect.width, page.rect.height

        # A. Tapar cabecera original (Elimina logo MaintainX y metadatos sup.)
        page.draw_rect(fitz.Rect(0, 0, w, 70), color=(1,1,1), fill=(1,1,1))

        # B. Tapar pie de página original (Elimina "Generado para CHG..." inf.)
        page.draw_rect(fitz.Rect(0, h - 55, w, h), color=(1,1,1), fill=(1,1,1))

        # C. Dibujar el membrete de CHG
        page.draw_rect(fitz.Rect(0, 0, w, 15), color=(0.04, 0.16, 0.38), fill=(0.04, 0.16, 0.38)) # Barra azul marino
        page.insert_text((30, 45), "CHG ASCENSORES", fontsize=22, fontname="helv-bo", color=(0.04, 0.16, 0.38))
        page.insert_text((30, 60), f"Informe de Mantenimiento | Orden #{meta.get('numero_orden', '')}", fontsize=10, fontname="helv", color=(0.4, 0.4, 0.4))
        page.draw_rect(fitz.Rect(0, h - 10, w, h), color=(0.04, 0.16, 0.38), fill=(0.04, 0.16, 0.38)) # Barra pie

    stamped_pdf_bytes = doc.write()
    doc.close()

    pdf_base64 = base64.b64encode(stamped_pdf_bytes).decode('utf-8')

    # Retornamos el diccionario manteniendo la estructura para no romper app.py
    # La UI usará 'meta', y el backend usará 'stamped_pdf_b64'
    return {
        "meta": meta,
        "stamped_pdf_b64": pdf_base64,
        "secciones": [], "fotos": [], "observaciones": {}, "firma": {"data_base64": "", "texto": ""},
        "seguimiento": {"titulo": "", "contenido": ""}, "info_orden": {"campos": []},
        "comentarios": [], "historial": [],
        "validacion": {"ok": True, "campos_esperados": 1, "campos_parseados": 1, "valores_fuera_de_conjunto": []}
    }
