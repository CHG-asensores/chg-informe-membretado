"""
Parser de PDFs exportados por MaintainX para CHG Ascensores.
Se importa desde el nodo Code (Python) del workflow n8n.

Requisitos: pip install pymupdf
"""
import base64, re
import fitz  # PyMuPDF

NOISE_Y_TOP    = 45     # pt — encabezado azul MaintainX
NOISE_Y_BOTTOM = 790    # pt — pie de página MaintainX (A4 = 842 pt)
KNOWN_VALUES   = {"Bueno", "Malo", "No Aplica", "condición estable", "Operativo", "-"}
BOLD_FLAG      = 1 << 4


def extract_lines(page):
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            y0 = line["bbox"][1]
            if y0 < NOISE_Y_TOP or y0 > NOISE_Y_BOTTOM:
                continue
            text = "".join(s["text"] for s in line["spans"]).strip()
            if not text:
                continue
            text = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
            is_bold = any(s["flags"] & BOLD_FLAG for s in line["spans"] if s["text"].strip())
            lines.append({"text": text, "y": y0, "x": line["bbox"][0], "bold": is_bold})
    lines.sort(key=lambda l: (round(l["y"]), l["x"]))
    return lines


def get_image_positions(page):
    """Devuelve lista de (xref, y_top) para cada imagen en la página, ordenadas por y."""
    positions = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") == 1:  # tipo imagen
            y0 = block["bbox"][1]
            # Buscar xref correspondiente comparando bbox
            positions.append({"bbox": block["bbox"], "y": y0, "xref": None})

    # Mapear xrefs a posiciones usando get_images
    page_imgs = page.get_images(full=True)
    img_list = list(page_imgs)

    # Asociar xrefs a bloques por orden de aparición en la página
    for i, pos in enumerate(positions):
        if i < len(img_list):
            pos["xref"] = img_list[i][0]

    return sorted(positions, key=lambda p: p["y"])


def parse_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    meta      = {}
    secciones = []
    fotos     = []   # fotos de inspección: [{etiqueta, data_base64, ext}]
    obs       = {
        "evaluacion_final": "",
        "para_cliente": "",
        "para_chg": "",
        "detalles_repuesto": {
            "info_tecnica": "",
            "fotos": []
        }
    }
    firma = {"data_base64": "", "texto": ""}

    # Extraer número de orden del encabezado azul
    for block in doc[0].get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            if line["bbox"][1] < NOISE_Y_TOP:
                raw = "".join(s["text"] for s in line["spans"]).strip()
                m = re.search(r"#(\d+)", raw)
                if m:
                    meta["numero_orden"] = m.group(1)
                t = re.match(r"^(.+?)\s*#\d+", raw)
                if t:
                    meta["titulo"] = t.group(1).strip()

    HEADER_LABELS = {
        "ESTADO":               "estado",
        "FECHA DE VENCIMIENTO": "fecha_vencimiento",
        "TIEMPO ESTIMADO":      "tiempo_estimado",
        "TIPO DE TRABAJO":      "tipo_trabajo",
        "ASIGNADOS":            "asignados",
        "CATEGORÍAS":           "categorias",
        "UBICACIÓN":            "ubicacion",
        "ACTIVO":               "activo",
        "PROCEDIMIENTO":        "procedimiento",
    }
    SKIP_LINES = {
        "CHG Ascensores",
        "Campos completados",
        "* Indica que la pregunta es obligatoria",
    }

    state            = "HEADER"
    current_section  = None
    pending_label    = None
    pending_left     = None
    pending_right    = None
    COL_THRESHOLD    = 180
    campos_esperados = 0
    campos_parseados = 0
    valores_invalidos = []

    # Para saber en qué estado estaba el parser cuando encontró cada imagen
    # Rastreamos el y_position del inicio de "Detalles del repuesto" por página
    repuesto_y_por_pagina = {}  # page_num -> y donde empieza "Detalles del repuesto"

    for page_num, page in enumerate(doc):
        lines = extract_lines(page)

        # Primera pasada: detectar y_position de "Detalles del repuesto" en esta página
        repuesto_start_y = None
        firma_start_y = None
        for entry in lines:
            if entry["text"] == "Detalles del repuesto a reparar o cambiar" and entry["bold"]:
                repuesto_start_y = entry["y"]
            if "Firma del cliente" in entry["text"] or re.match(r"^Firmado por .+", entry["text"]):
                firma_start_y = entry["y"]

        repuesto_y_por_pagina[page_num] = repuesto_start_y

        i = 0
        while i < len(lines):
            entry   = lines[i]
            text    = entry["text"]
            is_bold = entry["bold"]

            if "INFORMACIÓN DE ORDEN DE TRABAJO" in text:
                state = "DONE"
                break

            if text in SKIP_LINES:
                i += 1
                continue

            # ─── HEADER ──────────────────────────────────────────────────────
            if state == "HEADER":
                m = re.match(r"^(\d+)\s*/\s*(\d+)$", text)
                if m:
                    campos_esperados = int(m.group(2))
                    meta["campos_completados"] = f"{m.group(1)} / {m.group(2)}"
                    i += 1
                    continue

                if text == "Fecha y Hora de ingreso:":
                    if i + 1 < len(lines):
                        meta["fecha_hora_ingreso"] = lines[i + 1]["text"]
                        i += 2
                    else:
                        i += 1
                    state = "SECTIONS"
                    pending_left = pending_right = None
                    continue

                col_is_left = entry["x"] < COL_THRESHOLD
                pending     = pending_left if col_is_left else pending_right

                if text in HEADER_LABELS:
                    if col_is_left:
                        pending_left  = HEADER_LABELS[text]
                    else:
                        pending_right = HEADER_LABELS[text]
                    i += 1
                    continue

                if pending:
                    if pending == "asignados":
                        meta.setdefault("asignados", [])
                        if text not in meta["asignados"]:
                            meta["asignados"].append(text)
                    elif pending == "categorias":
                        meta.setdefault("categorias", [])
                        if text not in meta["categorias"]:
                            meta["categorias"].append(text)
                    elif pending == "ubicacion":
                        if "ubicacion" not in meta:
                            meta["ubicacion"] = text
                        else:
                            meta.setdefault("direccion", text)
                            if col_is_left:  pending_left  = None
                            else:            pending_right = None
                    elif pending == "activo":
                        if "activo" not in meta:
                            meta["activo"] = text
                        else:
                            if not re.match(r"^\d+$", text):
                                meta["activo"] += " " + text
                    elif pending == "procedimiento":
                        meta["procedimiento"] = text
                        if col_is_left:  pending_left  = None
                        else:            pending_right = None
                    else:
                        meta[pending] = text
                        if col_is_left:  pending_left  = None
                        else:            pending_right = None
                    i += 1
                    continue

            # ─── SECTIONS ────────────────────────────────────────────────────
            elif state == "SECTIONS":
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    state = "PHOTOS"
                    fotos.append({"etiqueta": text.rstrip(":"), "data_base64": "", "ext": "jpeg"})
                    i += 1
                    continue

                if text == "Observaciones y Recomendaciones" and is_bold:
                    state         = "OBSERVATIONS"
                    pending_label = None
                    i += 1
                    continue

                if is_bold and not text.endswith(":") and len(text) > 3:
                    current_section = {"nombre": text, "campos": []}
                    secciones.append(current_section)
                    pending_label   = None
                    i += 1
                    continue

                if text.endswith(":"):
                    pending_label = text[:-1].strip()
                    i += 1
                    continue

                if pending_label and current_section is not None:
                    current_section["campos"].append({"etiqueta": pending_label, "valor": text})
                    campos_parseados += 1
                    if text not in KNOWN_VALUES:
                        valores_invalidos.append({"etiqueta": pending_label, "valor": text})
                    pending_label = None

            # ─── PHOTOS ──────────────────────────────────────────────────────
            elif state == "PHOTOS":
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    fotos.append({"etiqueta": text.rstrip(":"), "data_base64": "", "ext": "jpeg"})
                    i += 1
                    continue
                if text == "Observaciones y Recomendaciones" and is_bold:
                    state         = "OBSERVATIONS"
                    pending_label = None

            # ─── OBSERVATIONS ────────────────────────────────────────────────
            elif state == "OBSERVATIONS":
                if text == "Detalles del repuesto a reparar o cambiar" and is_bold:
                    state         = "REPUESTO"
                    pending_label = None
                    i += 1
                    continue

                if text == "Evaluación final:":
                    pending_label = "evaluacion_final"
                elif text == "Observaciones y Recomendaciones para cliente:":
                    pending_label = "para_cliente"
                elif text == "Observaciones para CHG Ascensores:":
                    pending_label = "para_chg"
                elif text == "Fecha y Hora de salida:":
                    pending_label = "fecha_salida"
                elif re.match(r"^Firmado por .+", text):
                    firma["texto"] = text
                    pending_label  = None
                elif pending_label:
                    if pending_label == "evaluacion_final":
                        obs["evaluacion_final"] = text
                        pending_label = None
                    elif pending_label == "para_cliente":
                        if obs["para_cliente"]:
                            obs["para_cliente"] += "\n" + text
                        else:
                            obs["para_cliente"] = text
                    elif pending_label == "para_chg":
                        if obs["para_chg"]:
                            obs["para_chg"] += "\n" + text
                        else:
                            obs["para_chg"] = text
                    elif pending_label == "fecha_salida":
                        meta["fecha_hora_salida"] = text
                        pending_label = None

            # ─── REPUESTO ────────────────────────────────────────────────────
            elif state == "REPUESTO":
                if text == "Observaciones y Recomendaciones" and is_bold:
                    state         = "OBSERVATIONS"
                    pending_label = None
                    i += 1
                    continue

                if text == "Información de repuesto (medidas, datos técnicos):":
                    pending_label = "info_tecnica"
                    i += 1
                    continue

                if re.match(r"^Fotos del repuesto a reparar", text, re.I):
                    pending_label = "fotos_repuesto"
                    i += 1
                    continue

                if pending_label == "info_tecnica":
                    if obs["detalles_repuesto"]["info_tecnica"]:
                        obs["detalles_repuesto"]["info_tecnica"] += "\n" + text
                    else:
                        obs["detalles_repuesto"]["info_tecnica"] = text

            i += 1

        # ── Asignación de imágenes por posición Y relativa al texto ──────────
        img_positions = get_image_positions(page)
        repuesto_y = repuesto_y_por_pagina.get(page_num)

        page_text = page.get_text()
        has_firma = "Firma del cliente" in page_text or "Firmado por" in page_text

        for img_pos in img_positions:
            xref = img_pos.get("xref")
            if xref is None:
                continue
            try:
                base_img = doc.extract_image(xref)
                img_b64  = base64.b64encode(base_img["image"]).decode("utf-8")
                img_ext  = base_img.get("ext", "jpeg")

                img_y = img_pos["y"]

                if has_firma and repuesto_y is None:
                    # Página de firma
                    if not firma["data_base64"]:
                        firma["data_base64"] = img_b64
                        firma["ext"]         = img_ext
                elif repuesto_y is not None and img_y > repuesto_y:
                    # Imagen está DESPUÉS del título "Detalles del repuesto" → va al repuesto
                    obs["detalles_repuesto"]["fotos"].append({
                        "data_base64": img_b64,
                        "ext": img_ext
                    })
                else:
                    # Imagen está ANTES del título "Detalles del repuesto" → foto de inspección
                    for foto in fotos:
                        if not foto["data_base64"]:
                            foto["data_base64"] = img_b64
                            foto["ext"]         = img_ext
                            break
            except Exception:
                pass

        if state == "DONE":
            break

    doc.close()

    asignados = meta.get("asignados", [])
    if len(asignados) > 1 and all(" " not in a for a in asignados):
        meta["asignados"] = [" ".join(asignados)]

    return {
        "meta":      meta,
        "secciones": secciones,
        "fotos":     fotos,
        "observaciones": obs,
        "firma":     firma,
        "validacion": {
            "ok":                        campos_parseados == campos_esperados and not valores_invalidos,
            "campos_esperados":          campos_esperados,
            "campos_parseados":          campos_parseados,
            "valores_fuera_de_conjunto": valores_invalidos,
        },
    }
