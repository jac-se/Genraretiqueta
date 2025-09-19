import os
import textwrap
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

# ================= CONFIG =================
INPUT_CSV = "equipos.csv"
SALIDA_DIR = "salida_etiquetas"
NOMBRE_PDF = "Etiquetas_QR.pdf"

# --- Tamaño JANEL J-5262 (14 etiquetas: 2x7, 4.0" x 1.33") ---
PAGE_SIZE = letter
ETIQUETA_ANCHO_IN = 3.90
ETIQUETA_ALTO_IN  = 1.33
COLUMNAS = 2
FILAS    = 7

# Márgenes/espaciado (ajusta si tu impresora desplaza)
MARGEN_IZQ_IN = 0.25
MARGEN_SUP_IN = 0.50
ESPACIO_H_INTER_IN = 0.12
ESPACIO_V_INTER_IN = 0.12

# Offsets de calibración fina (mueven TODO el pliego)
OFFSET_X_IN = 0.10   # negativo = izquierda ; positivo = derecha MUEVE -2.5MM
OFFSET_Y_IN = -0.14   # negativo = arriba    ; positivo = abajo SUBE -2.5 M A LA DERECHA

DELTA_POR_FILA_IN = 0.00   # ajuste acumulado por fila (usa ±0.01 si ves deriva vertical)
DELTA_POR_COL_IN  = 0.00   # ajuste acumulado por columna (usa ±0.01 si la 2ª col. se “va”)


DPI = 300

# Fuentes (Arial; cambia si quieres otra). Usa rutas válidas en tu Windows.
FUENTE_BOLD_PATH = "C:/Windows/Fonts/arialbd.ttf"  # negritas
FUENTE_REG_PATH  = "C:/Windows/Fonts/arial.ttf"    # regular

# Tamaños relativos
SCALE_FONTS    = 1.00
TITLE_SIZE_REL = 0.22   # Usuario (negritas)
BODY_SIZE_REL  = 0.17   # Monitor/CPU/UPS

# Datos visibles (debajo del título)
CAMPOS_TEXTO_VISIBLES = ["Monitor", "CPU", "UPS", "Area", "Ubicacion", "Ext"]

# Datos que SÍ van dentro del QR (mantiene InventarioID)
CAMPOS_PARA_QR = [
    "Usuario", "InventarioID", "UPS", "CPU", "Monitor",
    "Area", "Ubicacion", "Ext", "FechaAlta"
]

ALIAS_ETIQUETAS = {
    "Usuario": "Usuario",
    "UPS": "UPS",
    "CPU": "CPU",
    "Monitor": "Monitor",
    "Area": "Área",
    "Ubicacion": "Ubicación",
    "Ext": "Ext."
}

# Si no tienes columna UPS y quieres mostrar algo fijo, ponlo aquí (o deja vacío)
UPS_DEFAULT = ""

# QR
QR_BOX_SIZE = 8
QR_BORDER   = 2
QR_LADO_REL = 0.36   # % del lado menor de la etiqueta

# ======== Utilidades de fuente/medición ========
def _truetype_or_default(path, size):
    try:
        if path and os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    except Exception:
        pass
    return ImageFont.load_default()

def load_font(size, bold=False):
    return _truetype_or_default(FUENTE_BOLD_PATH if bold else FUENTE_REG_PATH, size)

def text_fits(draw, text, font, max_w):
    return draw.textlength(text, font=font) <= max_w

def wrap_to_width(draw, text, font, max_w, max_lines=2):
    """Envuelve por palabras hasta máx. 2 líneas; recorta con '…' si excede."""
    if text_fits(draw, text, font, max_w):
        return [text]

    words = text.split()
    if len(words) == 1:  # palabra larguísima
        t = text
        while not text_fits(draw, t, font, max_w) and len(t) > 1:
            t = t[:-1]
        return [t + "…"] if t != text else [t]

    lines = []
    current = ""
    for w in words:
        test = (current + " " + w).strip()
        if text_fits(draw, test, font, max_w):
            current = test
        else:
            if current:
                lines.append(current)
            current = w
            if len(lines) == max_lines - 1:  # última línea; recortar
                # llena lo que quepa en la última línea
                last = current
                while not text_fits(draw, last + "…", font, max_w) and len(last) > 1:
                    last = last[:-1]
                lines.append(last + "…")
                return lines
    if current:
        lines.append(current)
    return lines[:max_lines]

# ================= QR =================
def make_qr(texto):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=QR_BOX_SIZE,
        border=QR_BORDER,
    )
    qr.add_data(texto)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")

def construir_texto_qr(row: dict) -> str:
    url = row.get("URL", "")
    if isinstance(url, str) and url.strip():
        return url.strip()
    partes = []
    for campo in CAMPOS_PARA_QR:
        val = str(row.get(campo, "")).strip()
        if not val and campo == "UPS" and UPS_DEFAULT:
            val = UPS_DEFAULT
        if val:
            partes.append(f"{campo}={val}")
    return " | ".join(partes) if partes else "SIN_DATOS"

# ================= Dibujo etiqueta =================
def crear_imagen_etiqueta(row: dict):
    w = int(ETIQUETA_ANCHO_IN * DPI)
    h = int(ETIQUETA_ALTO_IN * DPI)
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)

    pad = int(min(w, h) * 0.06)

    # QR
    qr_text = construir_texto_qr(row)
    qr_side = int(min(w, h) * QR_LADO_REL)
    qr_img = make_qr(qr_text).resize((qr_side, qr_side), resample=Image.LANCZOS)
    qr_x = pad
    qr_y = (h - qr_side) // 2
    img.paste(qr_img, (qr_x, qr_y))

    # Área de texto
    text_x = qr_x + qr_side + pad
    text_w = w - text_x - pad

    # Fuentes base
    title_px_max = max(11, int(h * TITLE_SIZE_REL * SCALE_FONTS))
    body_font = load_font(size=max(10, int(h * BODY_SIZE_REL * SCALE_FONTS)), bold=False)

    # ===== TÍTULO: Usuario en negritas con auto-fit y hasta 2 líneas =====
    usuario_val = str(row.get("Usuario", "")).strip() or "Equipo"
    size = title_px_max
    title_font_dyn = load_font(size=size, bold=True)

    # Primero, intenta caber en una línea ajustando tamaño hacia abajo
    MIN_TITLE_PX = 10
    while size > MIN_TITLE_PX and not text_fits(draw, usuario_val, title_font_dyn, text_w):
        size -= 1
        title_font_dyn = load_font(size=size, bold=True)

    # Si aún no cabe en una línea, permite dos líneas y, si hace falta, baja un poco más
    title_lines = wrap_to_width(draw, usuario_val, title_font_dyn, text_w, max_lines=2)
    # Si alguna línea aún no cabe (por cambiar tamaño), baja un poco más
    while any(not text_fits(draw, ln, title_font_dyn, text_w) for ln in title_lines) and size > MIN_TITLE_PX:
        size -= 1
        title_font_dyn = load_font(size=size, bold=True)
        title_lines = wrap_to_width(draw, usuario_val, title_font_dyn, text_w, max_lines=2)

    # ===== Cuerpos: Monitor, CPU, UPS (máx 3 líneas) =====
    body_lines = []
    orden = ["Monitor", "CPU", "UPS", "Area", "Ubicacion", "Ext"]
    for c in orden:
        val = str(row.get(c, "")).strip()
        if not val and c == "UPS" and UPS_DEFAULT:
            val = UPS_DEFAULT
        if val:
            alias = ALIAS_ETIQUETAS.get(c, c)
            # recorte suave por ancho
            v = val
            while not text_fits(draw, f"{alias}: {v}", body_font, text_w) and len(v) > 3:
                v = v[:-4] + "…"
            body_lines.append(f"{alias}: {v}")
        if len(body_lines) >= 3:
            break

    # Medición y render centrado vertical
    ESP = max(1, int(body_font.size * 0.25))
    alturas = []

    for ln in title_lines:
        bbox = draw.textbbox((0, 0), ln, font=title_font_dyn)
        alturas.append(bbox[3] - bbox[1])

    for ln in body_lines:
        bbox = draw.textbbox((0, 0), ln, font=body_font)
        alturas.append(bbox[3] - bbox[1])

    total_h = sum(alturas) + ESP * (len(alturas) - 1)
    y = max(pad, (h - total_h) // 2)

    # Render: títulos (negritas) y cuerpos
    for ln in title_lines:
        draw.text((text_x, y), ln, font=title_font_dyn, fill="black")
        y += draw.textbbox((0, 0), ln, font=title_font_dyn)[3] + ESP

    for i, ln in enumerate(body_lines):
        draw.text((text_x, y), ln, font=body_font, fill="black")
        y += draw.textbbox((0, 0), ln, font=body_font)[3] + (ESP if i < len(body_lines) - 1 else 0)

    # Borde sutil
    draw.rectangle([(0, 0), (w - 1, h - 1)], outline="black", width=1)
    return img

# ================= Exportación =================
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import os

def exportar_pngs_y_pdf(df):
    """
    Genera PNGs individuales y compone un PDF en rejilla 2x7 (JANEL J-5262).
    Incluye:
      - Offsets globales OFFSET_X_IN/OFFSET_Y_IN para mover todo el pliego.
      - Corrección acumulada por fila/col: DELTA_POR_FILA_IN / DELTA_POR_COL_IN.
    """
    os.makedirs(SALIDA_DIR, exist_ok=True)
    png_dir = os.path.join(SALIDA_DIR, "png")
    os.makedirs(png_dir, exist_ok=True)

    # ===== 1) Render de etiquetas a PNG =====
    rutas = []
    for idx, row in df.iterrows():
        rd = {k: ("" if pd.isna(v) else v) for k, v in row.items()}
        img = crear_imagen_etiqueta(rd)
        fname = f"etiqueta_{idx+1:03d}.png"
        fpath = os.path.join(png_dir, fname)
        img.save(fpath, dpi=(DPI, DPI))
        rutas.append(fpath)

    # ===== 2) PDF en rejilla con offsets y correcciones =====
    c = canvas.Canvas(os.path.join(SALIDA_DIR, NOMBRE_PDF), pagesize=PAGE_SIZE)
    page_w, page_h = PAGE_SIZE

    etiqueta_w_pt = ETIQUETA_ANCHO_IN * inch
    etiqueta_h_pt = ETIQUETA_ALTO_IN * inch
    margen_izq_pt = MARGEN_IZQ_IN * inch
    margen_sup_pt = MARGEN_SUP_IN * inch
    esp_h_pt = ESPACIO_H_INTER_IN * inch
    esp_v_pt = ESPACIO_V_INTER_IN * inch

    # Offsets globales (mueven TODO el pliego)
    off_x_pt = OFFSET_X_IN * inch     # + derecha / - izquierda
    off_y_pt = OFFSET_Y_IN * inch     # + abajo   / - arriba

    # Corrección acumulativa por columna/fila
    delta_col_pt  = DELTA_POR_COL_IN  * inch  # se suma col*delta
    delta_fila_pt = DELTA_POR_FILA_IN * inch  # se suma fila*delta

    # Posiciones X por columna
    x_positions = []
    for col in range(COLUMNAS):
        base_x = margen_izq_pt + col * (etiqueta_w_pt + esp_h_pt)
        x_positions.append(base_x + off_x_pt + col * delta_col_pt)

    # Posiciones Y por fila (recordar que ReportLab dibuja desde abajo)
    y_positions = []
    for fila in range(FILAS):
        base_y = page_h - margen_sup_pt - fila * (etiqueta_h_pt + esp_v_pt) - etiqueta_h_pt
        y_positions.append(base_y + off_y_pt + fila * delta_fila_pt)

    # Coloca las imágenes en la rejilla
    col = 0
    fila = 0
    for ruta in rutas:
        if fila >= FILAS:
            fila = 0
            col += 1
        if col >= COLUMNAS:
            c.showPage()
            col = 0
            fila = 0

        x = x_positions[col]
        y = y_positions[fila]
        c.drawImage(
            ruta, x, y,
            width=etiqueta_w_pt,
            height=etiqueta_h_pt,
            preserveAspectRatio=True,
            anchor='sw'
        )
        fila += 1

    c.save()
    return rutas, os.path.join(SALIDA_DIR, NOMBRE_PDF)

# ================= IO robusto + Main =================
def leer_csv_robusto(path):
    try:
        return pd.read_csv(path, encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin1")

def main():
    try:
        if not os.path.exists(INPUT_CSV):
            print("❌ ERROR: No se encontró el archivo CSV de entrada.")
            print(f"   Debes crear el archivo: {INPUT_CSV}")
            print("   Columnas típicas: Usuario, UPS, CPU, Monitor, Area, Ubicacion, Ext, InventarioID, FechaAlta")
            input("\nPresiona ENTER para cerrar...")
            return

        df = leer_csv_robusto(INPUT_CSV)
        exportar_pngs_y_pdf(df)
        print("✅ Etiquetas generadas.")
        print(f"📄 PDF: {os.path.join(SALIDA_DIR, NOMBRE_PDF)}")
        print(f"🖨️ Rejilla: {COLUMNAS}x{FILAS} = {COLUMNAS*FILAS} por hoja.")
        input("\nPresiona ENTER para salir...")

    except Exception as e:
        print("⚠️ Ocurrió un error inesperado:")
        print(str(e))
        input("\nPresiona ENTER para cerrar...")

if __name__ == "__main__":
    main()
