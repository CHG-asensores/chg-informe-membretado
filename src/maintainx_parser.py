"""
Parser de PDFs exportados por MaintainX para CHG Ascensores.
Requisitos: pip install pymupdf
"""
import base64, re
import fitz  # PyMuPDF

NOISE_Y_TOP    = 45
NOISE_Y_BOTTOM = 790
KNOWN_VALUES   = {"Bueno", "Malo", "No Aplica", "condición estable", "condición inestable", "Operativo", "Si Requiere", "Crítico", "-"}
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
    positions = []
    page_imgs = list(page.get_images(full=True))
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") == 1:
            positions.append({"bbox": block["bbox"], "y": block["bbox"][1], "xref": None})
    for i, pos in enumerate(positions):
        if i < len(page_imgs):
            pos["xref"] = page_imgs[i][0]
    return sorted(positions, key=lambda p: p["y"])


def parse_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    meta      = {}
    secciones = []
    fotos     = []
    obs       = {
        "evaluacion_final": "",
        "para_cliente": "",
        "para_chg": "",
        "detalles_repuesto": {"info_tecnica": "", "fotos": []}
    }
    firma = {"data_base64": "", "texto": ""}
    seguimiento = {"titulo": "", "contenido": ""}
    info_orden  = {"campos": []}    # INFORMACIÓN DE ORDEN DE TRABAJO
    comentarios = []                # lista de {texto, autor_fecha}
    historial   = []                # lista de {accion, fecha}

    # Número de orden del encabezado azul
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
        "ESTADO": "estado", "FECHA DE VENCIMIENTO": "fecha_vencimiento",
        "TIEMPO ESTIMADO": "tiempo_estimado", "TIPO DE TRABAJO": "tipo_trabajo",
        "ASIGNADOS": "asignados", "CATEGORÍAS": "categorias",
        "UBICACIÓN": "ubicacion", "ACTIVO": "activo", "PROCEDIMIENTO": "procedimiento",
        "PRIORIDAD": "prioridad",
    }
    SKIP_LINES = {"CHG Ascensores", "Campos completados",
                  "* Indica que la pregunta es obligatoria",
                  "Generado para CHG Ascensores"}

    state            = "HEADER"
    current_section  = None
    pending_label    = None
    pending_left     = None
    pending_right    = None
    COL_THRESHOLD    = 180
    campos_esperados = 0
    campos_parseados = 0
    valores_invalidos = []
    current_comentario = None

    for page_num, page in enumerate(doc):
        lines = extract_lines(page)

        # Detectar y_position de "Detalles del repuesto" en esta página
        repuesto_start_y = None
        for entry in lines:
            if entry["text"] == "Detalles del repuesto a reparar o cambiar" and entry["bold"]:
                repuesto_start_y = entry["y"]
                break

        i = 0
        while i < len(lines):
            entry   = lines[i]
            text    = entry["text"]
            is_bold = entry["bold"]

            if text in SKIP_LINES:
                i += 1
                continue

            # Líneas de paginación tipo "Página X de Y" (algunos PDFs las incluyen)
            if re.match(r"^Página \d+ de \d+$", text, re.I):
                i += 1
                continue

            # Etiquetas de prioridad sueltas (Alto/Medio/Bajo) antes del header,
            # cuando aparecen como línea independiente fuera de la columna de PRIORIDAD
            if state == "HEADER" and text in ("Alto", "Medio", "Bajo") and "prioridad" not in meta \
                    and "estado" not in meta:
                meta["prioridad"] = text
                i += 1; continue

            # ─── HEADER ──────────────────────────────────────────────────────
            if state == "HEADER":
                m = re.match(r"^(\d+)\s*/\s*(\d+)$", text)
                if m:
                    campos_esperados = int(m.group(2))
                    meta["campos_completados"] = f"{m.group(1)} / {m.group(2)}"
                    i += 1; continue

                if text == "Fecha y Hora de ingreso:":
                    if i + 1 < len(lines):
                        meta["fecha_hora_ingreso"] = lines[i+1]["text"]
                        i += 2
                    else:
                        i += 1
                    state = "SECTIONS"
                    pending_left = pending_right = None
                    continue

                col_is_left = entry["x"] < COL_THRESHOLD
                pending = pending_left if col_is_left else pending_right

                if text in HEADER_LABELS:
                    if col_is_left: pending_left  = HEADER_LABELS[text]
                    else:           pending_right = HEADER_LABELS[text]
                    i += 1; continue

                if pending:
                    if pending == "asignados":
                        meta.setdefault("asignados", [])
                        if text not in meta["asignados"]: meta["asignados"].append(text)
                    elif pending == "categorias":
                        meta.setdefault("categorias", [])
                        if text not in meta["categorias"]: meta["categorias"].append(text)
                    elif pending == "ubicacion":
                        if "ubicacion" not in meta: meta["ubicacion"] = text
                        else:
                            meta.setdefault("direccion", text)
                            if col_is_left: pending_left  = None
                            else:           pending_right = None
                    elif pending == "activo":
                        if "activo" not in meta: meta["activo"] = text
                        else:
                            if not re.match(r"^\d+$", text): meta["activo"] += " " + text
                    elif pending == "procedimiento":
                        meta["procedimiento"] = text
                        if col_is_left: pending_left  = None
                        else:           pending_right = None
                    else:
                        meta[pending] = text
                        if col_is_left: pending_left  = None
                        else:           pending_right = None
                    i += 1; continue

                # ── Fallback: ninguna condición de HEADER coincidió.
                # Esto ocurre cuando, tras el título del procedimiento, viene
                # directamente la primera sección del checklist (p.ej.
                # "Limpieza general") sin pasar por "Fecha y Hora de ingreso:".
                # Se considera terminado el HEADER y se reprocesa esta misma
                # línea ya en estado SECTIONS.
                state = "SECTIONS"
                pending_left = pending_right = None
                continue

            # ─── SECTIONS ────────────────────────────────────────────────────
            elif state == "SECTIONS":
                # Seguimiento de tiempos y costos
                if text == "Seguimiento de tiempos y costos":
                    seguimiento["titulo"] = text
                    state = "SEGUIMIENTO"
                    i += 1; continue

                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    state = "PHOTOS"
                    fotos.append({"etiqueta": text.rstrip(":"), "data_base64": "", "ext": "jpeg"})
                    i += 1; continue

                if text == "Observaciones y Recomendaciones" and is_bold:
                    state = "OBSERVATIONS"
                    pending_label = None
                    i += 1; continue

                if is_bold and not text.endswith(":") and len(text) > 3:
                    current_section = {"nombre": text, "campos": []}
                    secciones.append(current_section)
                    pending_label = None
                    i += 1; continue

                if text.endswith(":"):
                    pending_label = text[:-1].strip()
                    i += 1; continue

                if pending_label and current_section is not None:
                    current_section["campos"].append({"etiqueta": pending_label, "valor": text})
                    campos_parseados += 1
                    if text not in KNOWN_VALUES:
                        valores_invalidos.append({"etiqueta": pending_label, "valor": text})
                    pending_label = None

            # ─── SEGUIMIENTO ─────────────────────────────────────────────────
            elif state == "SEGUIMIENTO":
                # Cualquier línea hasta encontrar sección bold siguiente
                if is_bold and len(text) > 3 and not text.endswith(":"):
                    # Fin del seguimiento, volver a SECTIONS
                    state = "SECTIONS"
                    current_section = {"nombre": text, "campos": []}
                    secciones.append(current_section)
                    pending_label = None
                    i += 1; continue
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    state = "PHOTOS"
                    fotos.append({"etiqueta": text.rstrip(":"), "data_base64": "", "ext": "jpeg"})
                    i += 1; continue
                if text == "Observaciones y Recomendaciones" and is_bold:
                    state = "OBSERVATIONS"
                    pending_label = None
                    i += 1; continue
                # Acumular contenido del seguimiento
                if seguimiento["contenido"]:
                    seguimiento["contenido"] += "\n" + text
                else:
                    seguimiento["contenido"] = text

            # ─── PHOTOS ──────────────────────────────────────────────────────
            elif state == "PHOTOS":
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    fotos.append({"etiqueta": text.rstrip(":"), "data_base64": "", "ext": "jpeg"})
                    i += 1; continue
                if text == "Observaciones y Recomendaciones" and is_bold:
                    state = "OBSERVATIONS"
                    pending_label = None

            # ─── OBSERVATIONS ────────────────────────────────────────────────
            elif state == "OBSERVATIONS":
                if text == "Detalles del repuesto a reparar o cambiar" and is_bold:
                    state = "REPUESTO"
                    pending_label = None
                    i += 1; continue
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
                    pending_label = None
                elif text == "Firma del cliente:":
                    pending_label = None
                elif text == "INFORMACIÓN DE ORDEN DE TRABAJO":
                    state = "INFO_ORDEN"
                    pending_label = None
                    i += 1; continue
                elif pending_label:
                    if pending_label == "evaluacion_final":
                        obs["evaluacion_final"] = text
                        pending_label = None
                    elif pending_label == "para_cliente":
                        obs["para_cliente"] = (obs["para_cliente"] + "\n" + text).strip("\n")
                    elif pending_label == "para_chg":
                        obs["para_chg"] = (obs["para_chg"] + "\n" + text).strip("\n")
                    elif pending_label == "fecha_salida":
                        meta["fecha_hora_salida"] = text
                        pending_label = None

            # ─── REPUESTO ────────────────────────────────────────────────────
            elif state == "REPUESTO":
                if text == "Observaciones y Recomendaciones" and is_bold:
                    state = "OBSERVATIONS"
                    pending_label = None
                    i += 1; continue
                if text == "Información de repuesto (medidas, datos técnicos):":
                    pending_label = "info_tecnica"
                    i += 1; continue
                if re.match(r"^Fotos del repuesto a reparar", text, re.I):
                    pending_label = "fotos_repuesto"
                    i += 1; continue
                # Si en una nueva página aparecen directamente las etiquetas
                # de Observaciones (sin el encabezado bold "Observaciones y
                # Recomendaciones" repetido), volver a OBSERVATIONS y procesar
                if text in (
                    "Observaciones y Recomendaciones para cliente:",
                    "Observaciones para CHG Ascensores:",
                    "Fecha y Hora de salida:",
                    "Evaluación final:",
                ) or re.match(r"^Firmado por .+", text) or text == "INFORMACIÓN DE ORDEN DE TRABAJO":
                    state = "OBSERVATIONS"
                    pending_label = None
                    continue  # reprocesar esta línea en estado OBSERVATIONS
                if pending_label == "info_tecnica":
                    dr = obs["detalles_repuesto"]
                    dr["info_tecnica"] = (dr["info_tecnica"] + "\n" + text).strip("\n")

            # ─── INFO_ORDEN ───────────────────────────────────────────────────
            elif state == "INFO_ORDEN":
                if text == "COMENTARIOS":
                    state = "COMENTARIOS"
                    i += 1; continue
                if text == "HISTORIAL DE ORDEN DE TRABAJO":
                    state = "HISTORIAL"
                    i += 1; continue
                # Acumular campos de info de orden
                info_orden["campos"].append(text)

            # ─── COMENTARIOS ─────────────────────────────────────────────────
            elif state == "COMENTARIOS":
                if text == "HISTORIAL DE ORDEN DE TRABAJO":
                    state = "HISTORIAL"
                    i += 1; continue
                # "Comentado por X el fecha" → metadato del comentario anterior
                if re.match(r"^Comentado por .+", text):
                    if current_comentario is not None:
                        current_comentario["autor_fecha"] = text
                        comentarios.append(current_comentario)
                        current_comentario = None
                else:
                    if current_comentario is None:
                        current_comentario = {"texto": text, "autor_fecha": ""}
                    else:
                        current_comentario["texto"] += " " + text

            # ─── HISTORIAL ───────────────────────────────────────────────────
            elif state == "HISTORIAL":
                if text in ("Firmado por", "Fecha"):
                    state = "DONE_HIST"
                    i += 1; continue
                # Líneas alternas: acción / fecha
                if re.match(r"^\d{2}/\d{2}/\d{4}", text):
                    if historial:
                        historial[-1]["fecha"] = text
                else:
                    historial.append({"accion": text, "fecha": ""})

            elif state == "DONE_HIST":
                pass

            i += 1

        # Guardar comentario pendiente
        if current_comentario is not None:
            comentarios.append(current_comentario)
            current_comentario = None

        # ── Asignación de imágenes por posición Y ────────────────────────────
        img_positions = get_image_positions(page)
        page_text_raw = page.get_text()
        has_firma = "Firma del cliente" in page_text_raw or "Firmado por" in page_text_raw

        for img_pos in img_positions:
            xref = img_pos.get("xref")
            if xref is None:
                continue
            try:
                base_img = doc.extract_image(xref)
                img_b64  = base64.b64encode(base_img["image"]).decode("utf-8")
                img_ext  = base_img.get("ext", "jpeg")
                img_y    = img_pos["y"]

                if has_firma and repuesto_start_y is None:
                    if not firma["data_base64"]:
                        firma["data_base64"] = img_b64
                        firma["ext"]         = img_ext
                elif repuesto_start_y is not None and img_y > repuesto_start_y:
                    obs["detalles_repuesto"]["fotos"].append({"data_base64": img_b64, "ext": img_ext})
                else:
                    for foto in fotos:
                        if not foto["data_base64"]:
                            foto["data_base64"] = img_b64
                            foto["ext"]         = img_ext
                            break
            except Exception:
                pass

    doc.close()

    asignados = meta.get("asignados", [])
    if len(asignados) > 1 and all(" " not in a for a in asignados):
        meta["asignados"] = [" ".join(asignados)]

    return {
        "meta":        meta,
        "secciones":   secciones,
        "fotos":       fotos,
        "observaciones": obs,
        "firma":       firma,
        "seguimiento": seguimiento,
        "info_orden":  info_orden,
        "comentarios": comentarios,
        "historial":   historial,
        "validacion": {
            "ok":                        campos_parseados == campos_esperados and not valores_invalidos,
            "campos_esperados":          campos_esperados,
            "campos_parseados":          campos_parseados,
            "valores_fuera_de_conjunto": valores_invalidos,
        },
    }
