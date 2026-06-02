"""
Extrae el PNG de membrete del DOCX de CHG Ascensores.
Corre UNA sola vez localmente para obtener template/assets/membrete.png.

Uso:
    pip install python-docx Pillow
    python extract_letterhead.py

El DOCX tiene cuerpo vacío; el membrete es una imagen PNG de página completa
anclada en el encabezado (header1.xml). Este script la extrae y la guarda.
"""

import zipfile
import shutil
from pathlib import Path

DOCX_PATH = Path(r"C:\Users\1abel\Downloads\Orden de trabajo #488 - M. Preventivo EDIF. SAN BORJA NORTE (1).docx")
OUT_PATH = Path(__file__).parent.parent / "template" / "assets" / "membrete.png"


def extract_header_image(docx_path: Path, out_path: Path) -> None:
    # Un .docx es un ZIP. Buscamos la imagen referenciada desde word/header1.xml.
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()

        # Buscar relaciones del encabezado para encontrar la imagen
        header_rels = [n for n in names if "header" in n and "_rels" in n]
        image_path_in_zip = None

        for rel_file in header_rels:
            content = z.read(rel_file).decode("utf-8")
            # Buscar Target con extensión de imagen
            for line in content.split(">"):
                if 'Target="' in line and any(ext in line.lower() for ext in [".png", ".jpg", ".jpeg", ".emf", ".wmf"]):
                    start = line.index('Target="') + 8
                    end = line.index('"', start)
                    target = line[start:end]
                    # Resolver path relativo desde word/
                    if target.startswith(".."):
                        candidate = "word/" + target.lstrip("../")
                    elif not target.startswith("word/"):
                        candidate = "word/" + target
                    else:
                        candidate = target
                    # Normalizar separadores
                    candidate = candidate.replace("\\", "/")
                    if candidate in names:
                        image_path_in_zip = candidate
                        break
                if image_path_in_zip:
                    break

        # Fallback: si no se encontró por rels, tomar la imagen más grande de word/media/
        if not image_path_in_zip:
            media_files = [n for n in names if n.startswith("word/media/")]
            if not media_files:
                raise FileNotFoundError("No se encontró ninguna imagen en el DOCX.")
            # Tomar la más grande (el membrete de página completa suele ser la mayor)
            sizes = {n: z.getinfo(n).file_size for n in media_files}
            image_path_in_zip = max(sizes, key=sizes.get)
            print(f"Fallback: usando imagen más grande: {image_path_in_zip}")

        print(f"Imagen encontrada: {image_path_in_zip}")
        out_path.parent.mkdir(parents=True, exist_ok=True)

        with z.open(image_path_in_zip) as src, open(out_path, "wb") as dst:
            shutil.copyfileobj(src, dst)

    print(f"Membrete guardado en: {out_path}")

    # Verificar dimensiones con Pillow
    try:
        from PIL import Image
        img = Image.open(out_path)
        print(f"Dimensiones: {img.size[0]}x{img.size[1]} px, modo: {img.mode}")
    except ImportError:
        print("Pillow no instalado — dimensiones no verificadas.")


if __name__ == "__main__":
    extract_header_image(DOCX_PATH, OUT_PATH)
