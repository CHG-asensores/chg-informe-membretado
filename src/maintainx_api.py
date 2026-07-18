"""
Lee una OT de MaintainX por API y arma el mismo dict `data` que producia parse_pdf().
Reemplaza el parseo por coordenadas: los campos vienen etiquetados, no posicionados.

Requiere la variable de entorno MAINTAINX_API_KEY.
"""
import base64
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone

import requests

API_BASE = "https://api.getmaintainx.com/v1"
TZ_LIMA = timezone(timedelta(hours=-5))

# Campos que NO deben aparecer en el informe del cliente (comparacion en minusculas)
LABELS_EXCLUIDOS = {
    "informacion para chg",
    "informacion para chg ascensores",
    "observaciones para chg ascensores",
}

# ─────────────────────────────────────────────────────────────────────────────
# LISTA BLANCA. Un campo de texto o foto SOLO sale en el informe del cliente si
# su label esta aqui. Todo lo demas se ignora.
#
# El default es EXCLUIR, no incluir: Anyela edita los campos OT por OT, y un
# campo nuevo no puede llegar al cliente porque nadie lo listo.
# Lo que no se publique queda en data['no_publicado'] para revisarlo.
# ─────────────────────────────────────────────────────────────────────────────

# Labels cuyo contenido va al bloque de repuestos
LABELS_REPUESTO_INFO = {
    # Template viejo (Ascensor / Plataforma): confirmado, el informe.html
    # aprobado por CHG ya lo pinta como 'Detalles del repuesto a reparar'.
    "informacion de repuesto (medidas, datos tecnicos)",

    # MODULO 1. Decision de Zota (17/07/2026): se publica al cliente.
    # NO esta confirmado por CHG; se incluye por analogia de nombre con el
    # campo de arriba. Si CHG dice que es interno, borrar esta linea y el
    # texto pasa a data['no_publicado'] automaticamente.
    "informacion de reepuestos",
}
LABELS_REPUESTO_FOTOS = {
    "adjuntar imagen",
    "fotos del repuesto a reparar (maximo 3)",
}
LABELS_OBS_CLIENTE = {
    "observaciones y recomendaciones para cliente",
}


def _norm(s):
    """Minusculas, sin tildes, sin espacios dobles. Para comparar labels."""
    s = (s or "").strip().lower()
    for a, b in zip("áéíóúüñ", "aeiouun"):
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s)


# ─────────────────────────── Extraccion del ID desde el PDF ───────────────────────────

def extraer_workorder_id(pdf_bytes):
    """
    Saca el workOrderId interno del enlace 'Abrir en la aplicacion MaintainX'
    que MaintainX embebe en el pie de cada pagina del export.

    Es una anotacion de enlace, no texto posicionado: no le afecta que cambie
    el layout ni los campos del procedure.
    """
    import fitz

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page in doc:
            for link in page.get_links():
                uri = link.get("uri") or ""
                m = re.search(r"workOrderId=(\d+)", uri)
                if m:
                    return int(m.group(1))
                m = re.search(r"/workorders/(\d+)", uri)
                if m:
                    return int(m.group(1))
    finally:
        doc.close()
    raise ValueError(
        "No se encontro el enlace de MaintainX en el PDF. "
        "Verifica que sea el export original de MaintainX y no un PDF reimpreso."
    )


# ─────────────────────────── Llamada a la API ───────────────────────────

def _fetch_via_n8n(workorder_id):
    """
    Ruta sin API key propia: n8n ya tiene la credencial de MaintainX guardada
    (cifrada, no se puede leer desde la UI). El workflow "Proxy OT Membretado"
    devuelve {workOrder, usuarios} en una sola llamada.
    """
    url = os.environ.get("N8N_OT_WEBHOOK_URL", "").strip()
    if not url:
        raise RuntimeError(
            "Falta configuracion: define MAINTAINX_API_KEY (ruta directa) "
            "o N8N_OT_WEBHOOK_URL (ruta via n8n)."
        )
    headers = {}
    nombre_header = os.environ.get("N8N_OT_HEADER_NAME", "").strip()
    token = os.environ.get("N8N_OT_TOKEN", "").strip()
    if nombre_header and token:
        headers[nombre_header] = token

    r = requests.post(url, json={"workOrderId": workorder_id}, headers=headers, timeout=90)
    r.raise_for_status()
    payload = r.json()
    if isinstance(payload, list):  # respondWith=allIncomingItems devuelve array
        payload = payload[0] if payload else {}
    w = payload.get("workOrder")
    if not w:
        raise RuntimeError(
            "El webhook de n8n no devolvio 'workOrder'. Revisa el workflow "
            "'Proxy OT Membretado' y que el webhook este publicado."
        )
    return w, (payload.get("usuarios") or {})


def obtener_ot(workorder_id, api_key=None):
    """
    Devuelve (workOrder, usuarios).

    usuarios = {} o {"1230555": "Jordan Davila", ...} si vino por n8n.
    usuarios = None si se fue directo a MaintainX: ahi los nombres se
    resuelven uno a uno con /users/{id}.

    El dia que CHG entregue una API key propia, basta con setear
    MAINTAINX_API_KEY en EasyPanel: el rodeo por n8n se apaga solo.
    """
    key = (api_key or os.environ.get("MAINTAINX_API_KEY", "")).strip()
    if key:
        return fetch_workorder(workorder_id, api_key=key), None
    return _fetch_via_n8n(workorder_id)


def fetch_workorder(workorder_id, api_key=None):
    key = (api_key or os.environ.get("MAINTAINX_API_KEY", "")).strip()
    if not key:
        raise RuntimeError("Falta la variable de entorno MAINTAINX_API_KEY.")

    url = f"{API_BASE}/workorders/{workorder_id}"
    r = requests.get(
        url,
        headers={"Authorization": f"Bearer {key}"},
        # expand=assignees NO trae nombres (comprobado 17/07/2026: devuelve
        # [{id, type}]). Los nombres se resuelven aparte con /users/{id}.
        params=[("expand", "asset"), ("expand", "location")],
        timeout=30,
    )
    r.raise_for_status()
    payload = r.json()
    return payload.get("workOrder") or payload


# ─────────────────────────── Helpers de formato ───────────────────────────

def _dt(iso):
    if not iso:
        return None
    try:
        return datetime.fromisoformat(str(iso).replace("Z", "+00:00")).astimezone(TZ_LIMA)
    except Exception:
        return None


def _fmt_fecha_hora(iso):
    d = _dt(iso)
    return d.strftime("%d/%m/%Y, %H:%M") if d else ""


def _fmt_fecha(iso):
    d = _dt(iso)
    return d.strftime("%d/%m/%Y") if d else ""


def _fmt_duracion(segundos):
    if not segundos:
        return ""
    total_min = int(segundos) // 60
    h, m = divmod(total_min, 60)
    if h and m:
        return f"{h}h {m}m"
    if h:
        return f"{h}h"
    return f"{m}m"


def _adjuntos(value, descargar=True, _fetch=None):
    """
    Devuelve [{data_base64, ext}] desde value.attachments.

    OJO: value.attachment (singular) es SIEMPRE una copia del primer elemento de
    value.attachments. Si se usan los dos, cada foto sale duplicada.
    Se usa unicamente el array.
    """
    value = value or {}
    atts = value.get("attachments") or []
    out = []
    for a in atts:
        url = a.get("url")
        if not url:
            continue
        ext = (a.get("fileName", "") or "").rsplit(".", 1)[-1].lower() or "jpeg"
        if ext == "jpg":
            ext = "jpeg"
        if not descargar:
            out.append({"data_base64": "", "ext": ext, "fileName": a.get("fileName", "")})
            continue
        fn = _fetch or _descargar_bytes
        try:
            out.append({"data_base64": fn(url), "ext": ext, "fileName": a.get("fileName", "")})
        except Exception:
            # Una foto caida no debe tumbar el informe entero
            continue
    return out


# ─────────────────────────── Nombres de usuario ───────────────────────────

_USERS_CACHE = {}


def _nombre_usuario(uid, api_key):
    """
    GET /v1/users/{id} -> {"user": {"firstName": "Jordan", "lastName": "Davila", ...}}
    No hay campo fullName: el nombre viene partido.
    Se cachea: son 6 usuarios distintos en 151 OTs (medido 17/07/2026),
    asi que en la practica son 6 llamadas por arranque del proceso.
    """
    if uid in _USERS_CACHE:
        return _USERS_CACHE[uid]
    nombre = ""
    try:
        r = requests.get(
            f"{API_BASE}/users/{uid}",
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=20,
        )
        r.raise_for_status()
        u = (r.json() or {}).get("user") or {}
        nombre = " ".join(x for x in [u.get("firstName") or "", u.get("lastName") or ""] if x).strip()
    except Exception:
        nombre = ""
    _USERS_CACHE[uid] = nombre
    return nombre


def _descargar_bytes(url):
    """Las URLs de MaintainX son S3 pre-firmadas (X-Amz-Expires=3600). No llevan auth."""
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return base64.b64encode(r.content).decode("utf-8")


def _prefetch_adjuntos(w, max_workers=8):
    """
    Baja todas las fotos de la OT en paralelo y devuelve {url: base64}.

    Antes las fotos venian dentro del PDF (cero red). Ahora son ~24 descargas
    por informe: en serie se comen el timeout de gunicorn (120s, 1 worker),
    sobre todo si el operador arrastra varios PDFs de golpe.
    """
    urls = []
    for f in (w.get("procedure") or {}).get("fields") or []:
        v = f.get("value") or {}
        for a in (v.get("attachments") or []):
            u = a.get("url")
            if u and u not in urls:
                urls.append(u)
    cache = {}
    if not urls:
        return cache
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futuros = {ex.submit(_descargar_bytes, u): u for u in urls}
        for fut in as_completed(futuros):
            u = futuros[fut]
            try:
                cache[u] = fut.result()
            except Exception:
                cache[u] = None  # una foto caida no tumba el informe
    return cache


# ─────────────────────────── Construccion del contrato `data` ───────────────────────────

def construir_data(w, descargar_fotos=True, _fetch=None, resolver_usuario=None):
    """
    Arma el mismo dict que consumia render_html(). Recorre procedure.fields
    usando parentId para agrupar, sin ninguna lista de labels hardcodeada:
    Anyela edita los campos OT por OT y cualquier lista fija se rompe.
    """
    proc = w.get("procedure") or {}
    fields = proc.get("fields") or []
    byid = {f["id"]: f for f in fields}

    asset = w.get("asset") or {}
    location = w.get("location") or {}

    # assigneeIds esta siempre; assignees solo aparece con expand y no trae nombre.
    ids = list(w.get("assigneeIds") or [])
    if not ids:
        ids = [a.get("id") for a in (w.get("assignees") or []) if a.get("id")]
    asignados = []
    for uid in ids:
        nombre = resolver_usuario(uid) if resolver_usuario else ""
        if nombre:
            asignados.append(nombre)

    meta = {
        "numero_orden": str(w.get("sequentialId", "")),
        "titulo": w.get("title", "") or "",
        "estado": "Hecho" if w.get("status") == "DONE" else (w.get("status") or ""),
        "prioridad": w.get("priority", "") or "",
        "fecha_vencimiento": _fmt_fecha(w.get("dueDate")),
        "tiempo_estimado": _fmt_duracion(w.get("estimatedTime")),
        "tipo_trabajo": "Preventivo" if w.get("type") == "PREVENTIVE" else (w.get("type") or ""),
        "asignados": asignados,
        "categorias": list(w.get("categories") or []),
        "ubicacion": location.get("name", "") or "",
        "direccion": location.get("address", "") or "",
        "activo": asset.get("name", "") or "",
        "procedimiento": proc.get("title", "") or "",
        # El MODULO 1 no tiene 'Fecha y Hora de ingreso'. completedAt es la fecha real
        # del mantenimiento y es la que debe nombrar el archivo, no la fecha de hoy.
        "completado_en": w.get("completedAt") or "",
        "fecha_hora_ingreso": "",
        "fecha_hora_salida": "",
    }

    secciones = []
    fotos = []
    obs = {
        "descripcion": "",
        "evaluacion_final": "",
        "para_cliente": "",
        "para_chg": "",
        "detalles_repuesto": {"info_tecnica": "", "fotos": []},
    }
    firma = {"data_base64": "", "ext": "png", "texto": ""}

    no_publicado = []

    seccion_actual = None
    en_observaciones = False

    def nombre_seccion(f):
        p = byid.get(f.get("parentId"))
        while p is not None and p.get("type") != "HEADING":
            p = byid.get(p.get("parentId"))
        return (p.get("label") or "").strip() if p else ""

    for f in fields:
        tipo = f.get("type")
        label = (f.get("label") or "").strip()
        nlabel = _norm(label)
        value = f.get("value") or {}
        texto = (value.get("text") or "").strip()
        es_raiz = f.get("parentId") is None

        # Nunca sacar al cliente lo que es interno de CHG
        if nlabel in LABELS_EXCLUIDOS:
            obs["para_chg"] = texto
            continue

        # ── Encabezados de raiz = secciones
        if es_raiz and tipo == "HEADING":
            en_observaciones = "observacion" in nlabel
            if en_observaciones:
                # Sus hijos van al bloque de observaciones, pero la prosa del
                # HEADING es parte del informe y antes se perdia entera.
                obs["descripcion"] = (f.get("description") or "").strip()
                seccion_actual = None
                continue
            seccion_actual = {
                "fotos": [],
                "nombre": label,
                # La prosa tecnica vive en el description del HEADING, en el template
                # de MaintainX. Se lee en runtime: si CHG la edita, el informe se
                # actualiza solo.
                "descripcion": (f.get("description") or "").strip(),
                "campos": [],
            }
            secciones.append(seccion_actual)
            continue

        # ── Campos de raiz sueltos (Firma, Horario de Salida)
        if es_raiz:
            if tipo == "SIGNATURE":
                imgs = _adjuntos(value, descargar_fotos, _fetch)
                if imgs:
                    firma["data_base64"] = imgs[0]["data_base64"]
                    firma["ext"] = imgs[0]["ext"]
                firma["texto"] = texto
            elif tipo == "DATE":
                if "salida" in nlabel:
                    meta["fecha_hora_salida"] = _fmt_fecha_hora(texto)
                elif "ingreso" in nlabel:
                    meta["fecha_hora_ingreso"] = _fmt_fecha_hora(texto)
            continue

        # ── Hijos
        if tipo == "SIGNATURE":
            imgs = _adjuntos(value, descargar_fotos, _fetch)
            if imgs:
                firma["data_base64"] = imgs[0]["data_base64"]
                firma["ext"] = imgs[0]["ext"]
            firma["texto"] = texto
            continue

        if tipo == "DATE":
            if "salida" in nlabel:
                meta["fecha_hora_salida"] = _fmt_fecha_hora(texto)
            elif "ingreso" in nlabel:
                meta["fecha_hora_ingreso"] = _fmt_fecha_hora(texto)
            continue

        if tipo == "MULTIPLE_CHOICE":
            if "evaluacion final" in nlabel:
                obs["evaluacion_final"] = texto
            elif seccion_actual is not None and texto:
                seccion_actual["campos"].append({"etiqueta": label, "valor": texto})
            continue

        if tipo == "TEXT":
            if nlabel in LABELS_REPUESTO_INFO:
                obs["detalles_repuesto"]["info_tecnica"] = texto
            elif nlabel in LABELS_OBS_CLIENTE:
                obs["para_cliente"] = texto
            elif texto:
                # Campo de texto no listado: NO se publica.
                no_publicado.append({"label": label, "tipo": tipo, "valor": texto})
            continue

        if tipo == "FILE":
            imgs = _adjuntos(value, descargar_fotos, _fetch)
            if not imgs:
                continue
            if nlabel in LABELS_REPUESTO_FOTOS:
                obs["detalles_repuesto"]["fotos"].extend(imgs)
            elif seccion_actual is None:
                # Fotos fuera de toda seccion: no se publican.
                no_publicado.append({"label": label, "tipo": tipo, "valor": f"{len(imgs)} imagen(es)"})
            else:
                # Las fotos van al final de SU seccion, como en el export de
                # MaintainX. Ya no hace falta prefijar con el nombre de la
                # seccion: el titulo esta justo arriba.
                #
                # Los labels si se repiten dentro de una seccion ('Evidencia
                # fotografica del antes' x2 en el MODULO 1): esos se fusionan
                # en un solo bloque con todas sus imagenes.
                destino = seccion_actual["fotos"]
                existente = next((x for x in destino if x["etiqueta"] == label), None)
                if existente:
                    existente["imagenes"].extend(imgs)
                else:
                    destino.append({"etiqueta": label, "imagenes": list(imgs)})
            continue

        # UNSUPPORTED y cualquier tipo nuevo: se ignora en silencio.

    # Secciones sin nada dentro (p.ej. el HEADING de portada) no se pintan
    secciones = [s for s in secciones if s["campos"] or s["descripcion"] or s["fotos"]]

    hechos = sum(len(s["campos"]) for s in secciones)
    meta["campos_completados"] = f"{hechos} / {hechos}"

    return {
        "meta": meta,
        "secciones": secciones,
        "fotos": fotos,
        "observaciones": obs,
        "firma": firma,
        # Campos que la API trajo y NO se publicaron. Si aparece algo inesperado
        # aqui, es que CHG agrego un campo y hay que decidir que hacer con el.
        "no_publicado": no_publicado,
        "seguimiento": {"titulo": "", "contenido": ""},
        "info_orden": {"campos": []},
        "comentarios": [],
        "historial": [],
    }


def parse_ot_desde_pdf(pdf_bytes, api_key=None):
    """Reemplazo directo de parse_pdf(pdf_bytes)."""
    key = (api_key or os.environ.get("MAINTAINX_API_KEY", "")).strip()
    wo_id = extraer_workorder_id(pdf_bytes)
    w, usuarios = obtener_ot(wo_id, api_key=key)

    # Las fotos NO necesitan key: las URLs de S3 vienen pre-firmadas.
    # Se bajan siempre directo, sin pasar por n8n.
    cache = _prefetch_adjuntos(w)

    def _desde_cache(url):
        b64 = cache.get(url)
        if not b64:
            raise ValueError("descarga fallida")
        return b64

    if usuarios is None:
        resolver = lambda uid: _nombre_usuario(uid, key)
    else:
        resolver = lambda uid: usuarios.get(str(uid), "")

    return construir_data(w, _fetch=_desde_cache, resolver_usuario=resolver)
