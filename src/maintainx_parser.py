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

# Color azul del círculo seleccionado en MaintainX
BLUE_FILL = (0.0, 0.45879998803138733, 1.0)
# Tolerancia en px para emparejar centro del círculo con centro del texto
SEL_TOL   = 8


def get_selected_ys(page) -> list:
    """Devuelve lista de coordenadas Y (centro) de los círculos seleccionados
    en la página, identificados por su relleno azul en los drawings."""
    ys = []
    for d in page.get_drawings():
        if d.get("fill") == BLUE_FILL:
            rect = d.get("rect")
            if rect:
                ys.append((rect.y0 + rect.y1) / 2)
    return ys

# Umbrales para las 3 columnas del header (x < LEFT → col izq, x < MID → col central, x >= MID → col der)
COL_LEFT_MAX   = 180   # columna izquierda: x < 180
COL_MID_MAX    = 340   # columna central:   180 ≤ x < 340
                       # columna derecha:   x ≥ 340


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
            cy = (y0 + line["bbox"][3]) / 2   # centro vertical de la línea
            lines.append({"text": text, "y": y0, "cy": cy, "x": line["bbox"][0], "bold": is_bold})
    lines.sort(key=lambda l: (round(l["y"]), l["x"]))
    return lines


def get_image_positions(page):
    """Devuelve lista de imágenes en la página con su posición Y y xref correcto.
    Busca el xref real usando clip de cada bloque de imagen para evitar
    desincronización entre el orden de bloques y el orden de get_images()."""
    positions = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 1:
            continue
        bbox = block["bbox"]
        # Buscar el xref cuya imagen se renderiza en este bbox
        xref = None
        clip = fitz.Rect(bbox)
        for img in page.get_images(full=True):
            # Verificar que la imagen aparece dentro del clip usando get_image_rects
            try:
                rects = page.get_image_rects(img[0])
                for r in rects:
                    if abs(r.x0 - clip.x0) < 5 and abs(r.y0 - clip.y0) < 5:
                        xref = img[0]
                        break
            except Exception:
                pass
            if xref:
                break
        # Fallback: si no encontró por rect, usar orden posicional
        if xref is None:
            idx = len(positions)
            all_imgs = page.get_images(full=True)
            if idx < len(all_imgs):
                xref = all_imgs[idx][0]
        positions.append({"bbox": bbox, "y": bbox[1], "xref": xref})
    return sorted(positions, key=lambda p: p["y"])


def get_firma_y(page):
    """Devuelve la posición Y del texto 'Firma del cliente:' en la página,
    o None si no aparece."""
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            txt = "".join(s["text"] for s in line["spans"]).strip()
            if txt == "Firma del cliente:":
                return line["bbox"][1]
    return None


def get_foto_labels_y(page):
    """Devuelve posiciones Y de etiquetas 'Fotografía ...' en la página."""
    labels = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            txt = "".join(s["text"] for s in line["spans"]).strip()
            if re.match(r"^Fotograf[íi]a\b", txt, re.I):
                labels.append(line["bbox"][1])
    return sorted(labels)


def _col(x):
    """Devuelve 'left', 'mid' o 'right' según la posición X en el header."""
    if x < COL_LEFT_MAX:
        return "left"
    if x < COL_MID_MAX:
        return "mid"
    return "right"


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
    info_orden  = {"campos": []}
    comentarios = []
    historial   = []

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

    # Mapping de labels del header a claves de meta
    # Soporta 3 columnas: izquierda, central y derecha
    HEADER_LABELS = {
        "ESTADO":                "estado",
        "PRIORIDAD":             "prioridad",
        "FECHA DE VENCIMIENTO":  "fecha_vencimiento",
        "TIEMPO ESTIMADO":       "tiempo_estimado",
        "TIPO DE TRABAJO":       "tipo_trabajo",
        "ASIGNADOS":             "asignados",
        "CATEGORÍAS":            "categorias",
        "UBICACIÓN":             "ubicacion",
        "ACTIVO":                "activo",
        "PROCEDIMIENTO":         "procedimiento",
    }
    SKIP_LINES = {"CHG Ascensores", "Campos completados",
                  "* Indica que la pregunta es obligatoria",
                  "Generado para CHG Ascensores"}

    state            = "HEADER"
    current_section  = None
    pending_label    = None

    # pending por columna: left / mid / right
    pending_left  = None
    pending_mid   = None
    pending_right = None

    campos_esperados = 0
    campos_parseados = 0
    valores_invalidos = []
    current_comentario = None

    seen_repuesto_marker = False
    repuesto_photos_max  = 3

    # ── Estado de fotos tipo "Foto" (Montacarga/Plataforma) ──────────────────
    # En estos tipos, la sección bold es "Foto" (texto corto), y la etiqueta
    # del slot viene en la línea siguiente (ej: "Foto de cuarto de motor:").
    # pending_foto_label guarda la etiqueta hasta que llegue la imagen.
    pending_foto_label = None

    for page_num, page in enumerate(doc):
        lines = extract_lines(page)

        repuesto_start_y = None
        for entry in lines:
            if entry["text"] == "Detalles del repuesto a reparar o cambiar" and entry["bold"]:
                repuesto_start_y = entry["y"]
                seen_repuesto_marker = True
                break

        firma_text_y = get_firma_y(page)

        # Círculos seleccionados (azul relleno) en esta página
        selected_ys = get_selected_ys(page)

        i = 0
        while i < len(lines):
            entry   = lines[i]
            text    = entry["text"]
            is_bold = entry["bold"]
            x       = entry["x"]

            if text in SKIP_LINES:
                i += 1
                continue

            if re.match(r"^Página \d+ de \d+$", text, re.I):
                i += 1
                continue

            # Ignorar líneas "Rellenado por..." y "Exportado por..." en cualquier estado
            if re.match(r"^(Rellenado|Exportado) por ", text):
                i += 1
                continue

            # Ignorar sección "Descripción" y su texto explicativo del formulario
            if text == "Descripción":
                i += 1
                continue
            if re.match(r"^Este formulario debe ser llenado", text, re.I):
                i += 1
                continue

            # ─── HEADER ──────────────────────────────────────────────────────
            if state == "HEADER":
                m_campos = re.match(r"^(\d+)\s*/\s*(\d+)$", text)
                if m_campos:
                    campos_esperados = int(m_campos.group(2))
                    meta["campos_completados"] = f"{m_campos.group(1)} / {m_campos.group(2)}"
                    i += 1; continue

                if text == "Fecha y Hora de ingreso:":
                    if i + 1 < len(lines):
                        meta["fecha_hora_ingreso"] = lines[i+1]["text"]
                        i += 2
                    else:
                        i += 1
                    state = "SECTIONS"
                    pending_left = pending_mid = pending_right = None
                    continue

                if text == "Seguimiento de tiempos y costos":
                    seguimiento["titulo"] = text
                    pending_left = pending_mid = pending_right = None
                    i += 1
                    while i < len(lines):
                        nxt = lines[i]
                        if nxt["bold"]:
                            break
                        if nxt["text"] in HEADER_LABELS:
                            break
                        if re.match(r"^(\d+)\s*/\s*(\d+)$", nxt["text"]):
                            break
                        if nxt["text"] == "Fecha y Hora de ingreso:":
                            break
                        if seguimiento["contenido"]:
                            seguimiento["contenido"] += "\n" + nxt["text"]
                        else:
                            seguimiento["contenido"] = nxt["text"]
                        i += 1
                    continue

                col = _col(x)

                if text in HEADER_LABELS:
                    key = HEADER_LABELS[text]
                    if col == "left":
                        pending_left  = key
                    elif col == "mid":
                        pending_mid   = key
                    else:
                        pending_right = key
                    i += 1; continue

                # Determinar qué pending corresponde a esta columna
                pending = {"left": pending_left, "mid": pending_mid, "right": pending_right}.get(col)

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
                            if col == "left":   pending_left  = None
                            elif col == "mid":  pending_mid   = None
                            else:               pending_right = None
                    elif pending == "activo":
                        if "activo" not in meta:
                            meta["activo"] = text
                        else:
                            if not re.match(r"^\d+$", text):
                                meta["activo"] += " " + text
                    elif pending == "procedimiento":
                        meta["procedimiento"] = text
                        if col == "left":   pending_left  = None
                        elif col == "mid":  pending_mid   = None
                        else:               pending_right = None
                    else:
                        meta[pending] = text
                        if col == "left":   pending_left  = None
                        elif col == "mid":  pending_mid   = None
                        else:               pending_right = None
                    i += 1; continue

                if is_bold or text.endswith(":"):
                    state = "SECTIONS"
                    pending_left = pending_mid = pending_right = None
                    continue

                if "direccion" in meta:
                    meta["direccion"] += " " + text
                else:
                    meta["direccion"] = text
                i += 1; continue

            # ─── SECTIONS ────────────────────────────────────────────────────
            elif state == "SECTIONS":
                if text == "Seguimiento de tiempos y costos":
                    seguimiento["titulo"] = text
                    state = "SEGUIMIENTO"
                    i += 1; continue

                # Tipo Ascensor: "Fotografía 1:", "Fotografía 2:", etc.
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    state = "PHOTOS"
                    fotos.append({"etiqueta": text.rstrip(":"), "imagenes": [], "_y": entry["y"], "_pn": page_num})
                    pending_foto_label = None
                    i += 1; continue

                # Tipo Montacarga/Plataforma: sección bold "Foto" sola
                if text == "Foto" and is_bold:
                    state = "PHOTOS"
                    pending_foto_label = None  # la etiqueta vendrá en la siguiente línea
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
                    # Verificar si esta línea es la opción seleccionada.
                    # En Montacarga/Plataforma aparecen todas las opciones;
                    # en Ascensor solo aparece la opción seleccionada.
                    # Usamos el círculo azul (selected_ys) cuando está disponible.
                    is_selected = any(abs(sy - entry["cy"]) < SEL_TOL for sy in selected_ys)
                    if selected_ys:
                        # Hay círculos en la página → solo guardar si este está seleccionado
                        if is_selected:
                            current_section["campos"].append({"etiqueta": pending_label, "valor": text})
                            campos_parseados += 1
                            if text not in KNOWN_VALUES:
                                valores_invalidos.append({"etiqueta": pending_label, "valor": text})
                            pending_label = None
                    else:
                        # Sin círculos detectados (fallback): tomar el primer valor como antes
                        current_section["campos"].append({"etiqueta": pending_label, "valor": text})
                        campos_parseados += 1
                        if text not in KNOWN_VALUES:
                            valores_invalidos.append({"etiqueta": pending_label, "valor": text})
                        pending_label = None

            # ─── SEGUIMIENTO ─────────────────────────────────────────────────
            elif state == "SEGUIMIENTO":
                if is_bold and len(text) > 3 and not text.endswith(":"):
                    if text == "Foto":
                        state = "PHOTOS"
                        pending_foto_label = None
                        i += 1; continue
                    state = "SECTIONS"
                    current_section = {"nombre": text, "campos": []}
                    secciones.append(current_section)
                    pending_label = None
                    i += 1; continue
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    state = "PHOTOS"
                    fotos.append({"etiqueta": text.rstrip(":"), "imagenes": [], "_y": entry["y"], "_pn": page_num})
                    pending_foto_label = None
                    i += 1; continue
                if text == "Observaciones y Recomendaciones" and is_bold:
                    state = "OBSERVATIONS"
                    pending_label = None
                    i += 1; continue
                if seguimiento["contenido"]:
                    seguimiento["contenido"] += "\n" + text
                else:
                    seguimiento["contenido"] = text

            # ─── PHOTOS ──────────────────────────────────────────────────────
            elif state == "PHOTOS":
                # Tipo Ascensor: siguiente sección "Fotografía X"
                if re.match(r"^Fotograf[íi]a\b.+", text, re.I):
                    fotos.append({"etiqueta": text.rstrip(":"), "imagenes": [], "_y": entry["y"], "_pn": page_num})
                    pending_foto_label = None
                    i += 1; continue

                # Tipo Montacarga/Plataforma: nueva sección "Foto" bold
                if text == "Foto" and is_bold:
                    pending_foto_label = None
                    i += 1; continue

                if text == "Observaciones y Recomendaciones" and is_bold:
                    state = "OBSERVATIONS"
                    pending_label = None
                    i += 1; continue

                # Etiqueta del slot: línea que termina en ":"
                if text.endswith(":") and not is_bold:
                    etiqueta = text.rstrip(":").strip()
                    already = any(f["etiqueta"] == etiqueta for f in fotos)
                    if not already:
                        fotos.append({"etiqueta": etiqueta, "imagenes": [], "_y": entry["y"], "_pn": page_num})
                    pending_foto_label = etiqueta
                    i += 1; continue

            # ─── OBSERVATIONS ────────────────────────────────────────────────
            elif state == "OBSERVATIONS":
                if text == "Detalles del repuesto a reparar o cambiar" and is_bold:
                    state = "REPUESTO"
                    seen_repuesto_marker = True
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
                    m2 = re.search(r"máximo\s*(\d+)", text, re.I)
                    if m2:
                        repuesto_photos_max = int(m2.group(1))
                    i += 1; continue
                if text in (
                    "Observaciones y Recomendaciones para cliente:",
                    "Observaciones para CHG Ascensores:",
                    "Fecha y Hora de salida:",
                    "Evaluación final:",
                ) or re.match(r"^Firmado por .+", text) or text == "INFORMACIÓN DE ORDEN DE TRABAJO":
                    state = "OBSERVATIONS"
                    pending_label = None
                    continue
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
                info_orden["campos"].append(text)

            # ─── COMENTARIOS ─────────────────────────────────────────────────
            elif state == "COMENTARIOS":
                if text == "HISTORIAL DE ORDEN DE TRABAJO":
                    state = "HISTORIAL"
                    i += 1; continue
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
                if re.match(r"^\d{2}/\d{2}/\d{4}", text):
                    if historial:
                        historial[-1]["fecha"] = text
                else:
                    historial.append({"accion": text, "fecha": ""})

            elif state == "DONE_HIST":
                pass

            i += 1

        if current_comentario is not None:
            comentarios.append(current_comentario)
            current_comentario = None

        # ── Asignación de imágenes por posición Y ────────────────────────────
        img_positions = get_image_positions(page)

        for img_pos in img_positions:
            xref = img_pos.get("xref")
            if xref is None:
                continue
            try:
                base_img = doc.extract_image(xref)
                img_b64  = base64.b64encode(base_img["image"]).decode("utf-8")
                img_ext  = base_img.get("ext", "jpeg")
                img_y    = img_pos["y"]

                repuesto_fotos = obs["detalles_repuesto"]["fotos"]

                # ── Caso 1: imagen de repuesto ──
                if repuesto_start_y is not None and img_y > repuesto_start_y:
                    repuesto_fotos.append({"data_base64": img_b64, "ext": img_ext})
                    continue

                if (
                    repuesto_start_y is None
                    and seen_repuesto_marker
                    and len(repuesto_fotos) < repuesto_photos_max
                    and (firma_text_y is None or img_y < firma_text_y)
                ):
                    repuesto_fotos.append({"data_base64": img_b64, "ext": img_ext})
                    continue

                # ── Caso 2: imagen de firma ──
                if firma_text_y is not None and img_y >= firma_text_y and not firma["data_base64"]:
                    firma["data_base64"] = img_b64
                    firma["ext"]         = img_ext
                    continue

                # ── Caso 3: fotos de inspección ──
                # Saltear imágenes pequeñas (logos, avatares)
                bbox   = img_pos["bbox"]
                width  = bbox[2] - bbox[0]
                height = bbox[3] - bbox[1]
                if width < 50 or height < 50:
                    continue  # logo / avatar del técnico

                # Asignar al slot cuya etiqueta esté más cerca hacia arriba
                # en la misma página, o en una página anterior.
                best_slot = None
                best_score = (-1, float("inf"))   # (page_num, dy) — mayor pn y menor dy
                for foto in fotos:
                    slot_pn = foto.get("_pn", -1)
                    slot_y  = foto.get("_y", -1)
                    if slot_pn < 0:
                        continue
                    if slot_pn == page_num:
                        # Misma página: la etiqueta debe estar encima de la imagen
                        dy = img_y - slot_y
                        if dy >= 0:
                            score = (slot_pn, dy)
                            if score[0] > best_score[0] or (score[0] == best_score[0] and score[1] < best_score[1]):
                                best_score = score
                                best_slot  = foto
                    elif slot_pn < page_num:
                        # Página anterior: candidato solo si no hay ninguno en la misma página
                        score = (slot_pn, 0)
                        if best_score[0] < slot_pn:
                            best_score = score
                            best_slot  = foto

                if best_slot is not None:
                    best_slot["imagenes"].append({"data_base64": img_b64, "ext": img_ext})

            except Exception:
                pass

    doc.close()

    asignados = meta.get("asignados", [])
    if len(asignados) > 1 and all(" " not in a for a in asignados):
        meta["asignados"] = [" ".join(asignados)]

    # Quitar campo interno _y de los slots de fotos
    for foto in fotos:
        foto.pop("_y", None)

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
