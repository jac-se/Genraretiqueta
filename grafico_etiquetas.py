# -*- coding: utf-8 -*-
"""
Generador de etiquetas QR con vista gráfica para Windows (UI compacta)
- Interfaz compacta y organizada
- Vista previa: Hoja completa o Etiqueta individual
- Impresión directa GDI
- Perfiles configurables
- Parámetros editables para etiquetas JANEL J-5262
"""

import os
import json
import pandas as pd
import qrcode
from dataclasses import dataclass, asdict, fields
from typing import List, Tuple, Optional
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageWin

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# === Impresión directa GDI ===
try:
    import win32print
    import win32ui

    HAS_PYWIN32 = True
except Exception:
    HAS_PYWIN32 = False

# === Datos de contacto ===
CONTACTO = {
    "autor": "Ing. Jesús Arturo Cisneros Cantero",
    "correo": "tu.correo@ejemplo.com",
    "telefono": "+52 55 0000 0000",
    "sitio": "https://tusitio.ejemplo"
}

# === Constantes de Perfiles ===
PERFILES_DIR = os.path.join(os.path.dirname(__file__), "perfiles")
PERFIL_DEFECTO = "Default"
AUTO_GUARDAR_ARCH = os.path.join(PERFILES_DIR, ".auto_guardar.flag")

# === Utilidades de fuentes/impresoras ===
FONTS_DIRS = [r"C:\\Windows\\Fonts"]
FONT_MAPPING = {
    "Arial": "arial.ttf",
    "Arial Bold": "arialbd.ttf",
    "Arial Black": "ariblk.ttf",
    "Calibri": "calibri.ttf",
    "Calibri Bold": "calibrib.ttf",
    "Times New Roman": "times.ttf",
    "Times New Roman Bold": "timesbd.ttf",
    "Courier New": "cour.ttf",
    "Courier New Bold": "courbd.ttf",
    "Verdana": "verdana.ttf",
    "Verdana Bold": "verdanab.ttf",
    "Tahoma": "tahoma.ttf",
    "Tahoma Bold": "tahomabd.ttf",
    "Comic Sans MS": "comic.ttf",
    "Comic Sans MS Bold": "comicbd.ttf",
    "Georgia": "georgia.ttf",
    "Georgia Bold": "georgiab.ttf",
}


def get_font_path(font_name: str) -> str:
    if font_name in FONT_MAPPING:
        for d in FONTS_DIRS:
            path = os.path.join(d, FONT_MAPPING[font_name])
            if os.path.exists(path):
                return path
    for d in FONTS_DIRS:
        path = os.path.join(d, "arial.ttf")
        if os.path.exists(path):
            return path
    return ""


def list_printers():
    if not HAS_PYWIN32:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        printers = win32print.EnumPrinters(flags)
        return sorted({p[2] for p in printers if p[2]})
    except Exception:
        return []


# === Configuración de hoja ===
@dataclass
class SheetConfig:
    # VALORES PREDETERMINADOS ORIGINALES
    columnas: int = 2
    filas: int = 7
    etiqueta_ancho_in: float = 4.00
    etiqueta_alto_in: float = 1.333
    margen_izq_in: float = 0.19
    margen_sup_in: float = 0.48
    espacio_h_inter_in: float = 0.18
    espacio_v_inter_in: float = 0.10
    offset_x_in: float = 0.00
    offset_y_in: float = 0.00
    dpi: int = 300
    fuente_bold_path: str = r"C:\\Windows\\Fonts\\arialbd.ttf"
    fuente_reg_path: str = r"C:\\Windows\\Fonts\\arial.ttf"
    title_px_manual: int = 0
    body_px_manual: int = 0
    qr_scale_manual: int = 100
    copias_por_etiqueta: int = 3
    block_offset_x_px: int = 0
    block_offset_y_px: int = 0
    mostrar_borde: bool = True
    grosor_borde: int = 1
    fuente_bold_name: str = "Arial Bold"
    fuente_reg_name: str = "Arial"

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "SheetConfig":
        valid_keys = {f.name for f in fields(cls)}
        clean = {k: v for k, v in (data or {}).items() if k in valid_keys}
        return cls(**clean)


# === Funciones de renderizado ===
def load_font(path: str, size: int) -> ImageFont.FreeTypeFont:
    try:
        return ImageFont.truetype(path, size=size)
    except Exception:
        return ImageFont.load_default()


def crear_imagen_etiqueta(cfg: SheetConfig, row: dict) -> Image.Image:
    w = int(cfg.etiqueta_ancho_in * cfg.dpi)
    h = int(cfg.etiqueta_alto_in * cfg.dpi)
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)

    # QR
    qr_side = int(min(w, h) * 0.35 * (cfg.qr_scale_manual / 100))
    qr_payload = json.dumps(row, ensure_ascii=False)
    qr_img = qrcode.make(qr_payload).resize((qr_side, qr_side))
    qr_x = 10 + cfg.block_offset_x_px
    qr_y = (h - qr_side) // 2 + cfg.block_offset_y_px
    img.paste(qr_img, (qr_x, qr_y))

    # Texto
    text_x = qr_x + qr_side + 10
    title_size = cfg.title_px_manual or 24
    body_size = cfg.body_px_manual or 16

    title_font = load_font(get_font_path(cfg.fuente_bold_name), title_size)
    body_font = load_font(get_font_path(cfg.fuente_reg_name), body_size)

    # Limpiar valores NaN y manejar datos vacíos
    usuario_val = str(row.get("Usuario", "")).strip()
    if not usuario_val or usuario_val.lower() == "nan":
        usuario_val = "Equipo"

    draw.text((text_x, 8), usuario_val, font=title_font, fill="black")
    y = 10 + title_size

    for campo in ["Monitor", "CPU", "UPS", "Area", "Ext"]:
        val = str(row.get(campo, "")).strip()
        if not val or val.lower() == "nan":
            continue
        draw.text((text_x, y), f"{campo}: {val}", font=body_font, fill="black")
        y += body_size + 2

    if cfg.mostrar_borde:
        draw.rectangle([(0, 0), (w - 1, h - 1)],
                       outline="black", width=cfg.grosor_borde)

    return img


def crear_hoja_completa(cfg: SheetConfig, etiquetas: List[Image.Image]) -> Image.Image:
    sheet_w = int(8.5 * cfg.dpi)
    sheet_h = int(11 * cfg.dpi)
    sheet = Image.new("RGB", (sheet_w, sheet_h), "white")

    # Conversión a píxeles
    margen_izq_px = int(cfg.margen_izq_in * cfg.dpi)
    margen_sup_px = int(cfg.margen_sup_in * cfg.dpi)
    espacio_h_px = int(cfg.espacio_h_inter_in * cfg.dpi)
    espacio_v_px = int(cfg.espacio_v_inter_in * cfg.dpi)
    offset_x_px = int(cfg.offset_x_in * cfg.dpi)
    offset_y_px = int(cfg.offset_y_in * cfg.dpi)

    etiq_w = int(cfg.etiqueta_ancho_in * cfg.dpi)
    etiq_h = int(cfg.etiqueta_alto_in * cfg.dpi)

    # Colocar etiquetas
    idx = 0
    for fila in range(cfg.filas):
        for col in range(cfg.columnas):
            if idx >= len(etiquetas):
                break
            x = margen_izq_px + offset_x_px + col * (etiq_w + espacio_h_px)
            y = margen_sup_px + offset_y_px + fila * (etiq_h + espacio_v_px)
            sheet.paste(etiquetas[idx], (x, y))
            idx += 1
        if idx >= len(etiquetas):
            break

    # Borde de hoja
    draw = ImageDraw.Draw(sheet)
    draw.rectangle([(0, 0), (sheet_w - 1, sheet_h - 1)],
                   outline="lightgray", width=2)
    return sheet


# === Armar HOJAS desde rows ===
def crear_hojas_desde_rows(cfg: SheetConfig, rows: List[dict]) -> List[Image.Image]:
    capacidad = cfg.columnas * cfg.filas
    hojas = []
    if not rows:
        return hojas
    etiquetas = []
    for r in rows:
        for _ in range(cfg.copias_por_etiqueta):
            etiquetas.append(crear_imagen_etiqueta(cfg, r))
    for i in range(0, len(etiquetas), capacidad):
        chunk = etiquetas[i:i + capacidad]
        hoja = crear_hoja_completa(cfg, chunk)
        hojas.append(hoja)
    return hojas


# === Utilidades de archivo ===
def leer_csv(path: str) -> pd.DataFrame:
    if path.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path, encoding="utf-8")

    # Limpiar valores NaN en el DataFrame
    return df.fillna('')


# === Gestión de perfiles ===
def asegurar_dir_perfiles():
    os.makedirs(PERFILES_DIR, exist_ok=True)


def ruta_perfil(nombre: str) -> str:
    asegurar_dir_perfiles()
    safe = "".join(c for c in nombre if c.isalnum() or c in ("-", "_", " ")).strip()
    if not safe:
        safe = PERFIL_DEFECTO
    return os.path.join(PERFILES_DIR, f"{safe}.json")


def listar_perfiles() -> List[str]:
    asegurar_dir_perfiles()
    out = []
    for fn in os.listdir(PERFILES_DIR):
        if fn.lower().endswith(".json") and not fn.startswith("."):
            out.append(os.path.splitext(fn)[0])
    out.sort()
    if PERFIL_DEFECTO not in out:
        out.insert(0, PERFIL_DEFECTO)
    return out


def guardar_perfil(nombre: str, cfg: SheetConfig):
    asegurar_dir_perfiles()
    with open(ruta_perfil(nombre), "w", encoding="utf-8") as f:
        json.dump(cfg.to_dict(), f, ensure_ascii=False, indent=2)


def cargar_perfil(nombre: str) -> SheetConfig:
    try:
        with open(ruta_perfil(nombre), "r", encoding="utf-8") as f:
            data = json.load(f)
            return SheetConfig.from_dict(data)
    except Exception:
        cfg = SheetConfig()
        guardar_perfil(nombre, cfg)
        return cfg


def leer_auto_guardar_flag() -> bool:
    try:
        return os.path.exists(AUTO_GUARDAR_ARCH)
    except Exception:
        return False


def escribir_auto_guardar_flag(enabled: bool):
    asegurar_dir_perfiles()
    try:
        if enabled:
            with open(AUTO_GUARDAR_ARCH, "w", encoding="utf-8") as f:
                f.write("1")
        else:
            if os.path.exists(AUTO_GUARDAR_ARCH):
                os.remove(AUTO_GUARDAR_ARCH)
    except Exception:
        pass


# === GUI compacta ===
class EtiquetasGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Etiquetas QR – GUI con Perfiles")
        self.geometry("1100x700")
        self.minsize(900, 500)

        asegurar_dir_perfiles()
        self.cfg = cargar_perfil(PERFIL_DEFECTO)
        self.auto_guardar_var = tk.BooleanVar(value=leer_auto_guardar_flag())

        # Variables de control
        self.perfil_var = tk.StringVar(value=PERFIL_DEFECTO)
        self.csv_path = tk.StringVar()
        self.vista_var = tk.StringVar(value="Hoja completa")
        self.printer_var = tk.StringVar(value="(Predeterminada)")
        self.font_reg_var = tk.StringVar(value=self.cfg.fuente_reg_name)
        self.font_bold_var = tk.StringVar(value=self.cfg.fuente_bold_name)
        self.borde_var = tk.BooleanVar(value=self.cfg.mostrar_borde)

        self._vars = {}
        self._preview_scale = 1.0
        self._drag_start = None
        self._start_offset = (0, 0)

        self._build_compact_ui()
        self._apply_cfg_to_controls()
        self._preview()

    def _build_compact_ui(self):
        # Frame principal
        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Configuración de grid
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # Panel izquierdo (controles)
        left = ttk.Frame(main, width=350)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        left.columnconfigure(0, weight=1)

        # Panel derecho (vista previa)
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        # ===== CONTROLES IZQUIERDOS =====

        # Perfiles
        f_prof = ttk.LabelFrame(left, text="Perfiles de configuración")
        f_prof.pack(fill=tk.X, pady=3)

        row = ttk.Frame(f_prof)
        row.pack(fill=tk.X, padx=2, pady=2)
        ttk.Label(row, text="Perfil:").pack(side=tk.LEFT)
        self.cmb_perfil = ttk.Combobox(row, textvariable=self.perfil_var,
                                       state="readonly", width=12)
        self.cmb_perfil.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.cmb_perfil.bind("<<ComboboxSelected>>", lambda e: self._on_select_perfil())

        ttk.Button(row, text="Guardar", width=8,
                   command=self._guardar_perfil_actual).pack(side=tk.LEFT, padx=1)
        ttk.Button(row, text="+", width=3,
                   command=self._guardar_como).pack(side=tk.LEFT, padx=1)

        # Auto-guardar
        chk_auto = ttk.Checkbutton(f_prof, text="Auto-guardar cambios",
                                   variable=self.auto_guardar_var,
                                   command=self._toggle_auto_guardar)
        chk_auto.pack(anchor=tk.W, padx=2, pady=2)

        # Archivo
        f_arch = ttk.LabelFrame(left, text="Archivo de entrada")
        f_arch.pack(fill=tk.X, pady=3)

        row = ttk.Frame(f_arch)
        row.pack(fill=tk.X, padx=2, pady=2)
        ttk.Entry(row, textvariable=self.csv_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="Buscar…", command=self._pick_csv).pack(side=tk.LEFT, padx=2)

        # Vista e Impresión
        f_vista = ttk.LabelFrame(left, text="Vista e Impresión")
        f_vista.pack(fill=tk.X, pady=3)

        ttk.Radiobutton(f_vista, text="Hoja completa", variable=self.vista_var,
                        value="Hoja completa", command=self._preview).pack(anchor=tk.W, padx=2, pady=1)
        ttk.Radiobutton(f_vista, text="Etiqueta individual", variable=self.vista_var,
                        value="Etiqueta", command=self._preview).pack(anchor=tk.W, padx=2, pady=1)

        row = ttk.Frame(f_vista)
        row.pack(fill=tk.X, padx=2, pady=2)
        ttk.Label(row, text="Impresora:").pack(side=tk.LEFT)
        self.cmb_printer = ttk.Combobox(row, textvariable=self.printer_var,
                                        state="readonly", width=15)
        self.cmb_printer.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(row, text="🔄", width=3,
                   command=self._reload_printers).pack(side=tk.LEFT)

        # Configuración básica
        f_conf = ttk.LabelFrame(left, text="Configuración Básica")
        f_conf.pack(fill=tk.X, pady=3)

        # Disposición
        frame_disp = ttk.Frame(f_conf)
        frame_disp.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_disp, text="Columnas:").pack(side=tk.LEFT)
        self._spin_int(frame_disp, '', 'columnas', 1, 5, 1, width=4)
        ttk.Label(frame_disp, text="Filas:").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_int(frame_disp, '', 'filas', 1, 20, 1, width=4)
        ttk.Label(frame_disp, text="Copias:").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_int(frame_disp, '', 'copias_por_etiqueta', 1, 10, 1, width=4)

        # Fuentes
        frame_font = ttk.Frame(f_conf)
        frame_font.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_font, text="Fuente Regular:").pack(side=tk.LEFT)
        self.cmb_font_reg = ttk.Combobox(frame_font, textvariable=self.font_reg_var,
                                         state="readonly", width=12)
        self.cmb_font_reg.pack(side=tk.LEFT, padx=2)
        self.cmb_font_reg.bind("<<ComboboxSelected>>", lambda e: self._update_font_reg())

        ttk.Label(frame_font, text="Negrita:").pack(side=tk.LEFT, padx=(10, 0))
        self.cmb_font_bold = ttk.Combobox(frame_font, textvariable=self.font_bold_var,
                                          state="readonly", width=12)
        self.cmb_font_bold.pack(side=tk.LEFT, padx=2)
        self.cmb_font_bold.bind("<<ComboboxSelected>>", lambda e: self._update_font_bold())

        # Tamaños
        frame_size = ttk.Frame(f_conf)
        frame_size.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_size, text="Título (px):").pack(side=tk.LEFT)
        self._spin_int(frame_size, '', 'title_px_manual', 0, 60, 1, width=4)
        ttk.Label(frame_size, text="Cuerpo (px):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_int(frame_size, '', 'body_px_manual', 0, 40, 1, width=4)
        ttk.Label(frame_size, text="QR (%):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_int(frame_size, '', 'qr_scale_manual', 20, 200, 5, width=4)

        # Borde
        frame_borde = ttk.Frame(f_conf)
        frame_borde.pack(fill=tk.X, padx=2, pady=1)
        ttk.Checkbutton(frame_borde, text="Mostrar borde", variable=self.borde_var,
                        command=self._toggle_borde).pack(side=tk.LEFT)
        ttk.Label(frame_borde, text="Grosor:").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_int(frame_borde, '', 'grosor_borde', 1, 10, 1, width=3)

        # Parámetros de hoja (JANEL J-5262)
        f_sheet = ttk.LabelFrame(left, text="Parámetros de Hoja")
        f_sheet.pack(fill=tk.X, pady=3)

        # Dimensiones etiqueta
        frame_dim = ttk.Frame(f_sheet)
        frame_dim.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_dim, text="Ancho (in):").pack(side=tk.LEFT)
        self._spin_float(frame_dim, '', 'etiqueta_ancho_in', 0.50, 8.50, 0.001, width=6)
        ttk.Label(frame_dim, text="Alto (in):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_float(frame_dim, '', 'etiqueta_alto_in', 0.30, 11.00, 0.001, width=6)

        # Márgenes
        frame_marg = ttk.Frame(f_sheet)
        frame_marg.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_marg, text="Margen Izq (in):").pack(side=tk.LEFT)
        self._spin_float(frame_marg, '', 'margen_izq_in', 0.00, 2.00, 0.001, width=6)
        ttk.Label(frame_marg, text="Margen Sup (in):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_float(frame_marg, '', 'margen_sup_in', 0.00, 2.00, 0.001, width=6)

        # Espaciado
        frame_esp = ttk.Frame(f_sheet)
        frame_esp.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_esp, text="Espacio H (in):").pack(side=tk.LEFT)
        self._spin_float(frame_esp, '', 'espacio_h_inter_in', 0.00, 2.00, 0.001, width=6)
        ttk.Label(frame_esp, text="Espacio V (in):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_float(frame_esp, '', 'espacio_v_inter_in', 0.00, 2.00, 0.001, width=6)

        # Offsets
        frame_off = ttk.Frame(f_sheet)
        frame_off.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_off, text="Offset X (in):").pack(side=tk.LEFT)
        self._spin_float(frame_off, '', 'offset_x_in', -1.00, 1.00, 0.001, width=6)
        ttk.Label(frame_off, text="Offset Y (in):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_float(frame_off, '', 'offset_y_in', -1.00, 1.00, 0.001, width=6)

        # DPI
        frame_dpi = ttk.Frame(f_sheet)
        frame_dpi.pack(fill=tk.X, padx=2, pady=1)
        ttk.Label(frame_dpi, text="DPI:").pack(side=tk.LEFT)
        self._spin_int(frame_dpi, '', 'dpi', 72, 1200, 1, width=5)

        # Offset del bloque (arrastre)
        f_block = ttk.LabelFrame(left, text="Offset de Bloque (Arrastre)")
        f_block.pack(fill=tk.X, pady=3)

        frame_block = ttk.Frame(f_block)
        frame_block.pack(fill=tk.X, padx=2, pady=2)
        ttk.Label(frame_block, text="X (px):").pack(side=tk.LEFT)
        self._spin_int(frame_block, '', 'block_offset_x_px', -500, 500, 1, width=5)
        ttk.Label(frame_block, text="Y (px):").pack(side=tk.LEFT, padx=(10, 0))
        self._spin_int(frame_block, '', 'block_offset_y_px', -500, 500, 1, width=5)

        # Botones de acción
        f_actions = ttk.Frame(left)
        f_actions.pack(fill=tk.X, pady=5)

        ttk.Button(f_actions, text="Vista Previa",
                   command=self._preview).pack(fill=tk.X, pady=2)
        ttk.Button(f_actions, text="Imprimir",
                   command=self._print_gdi).pack(fill=tk.X, pady=2)
        ttk.Button(f_actions, text="Exportar (PNG/PDF)",
                   command=self._guardar_preview).pack(fill=tk.X, pady=2)
        ttk.Button(f_actions, text="Ayuda",
                   command=self._mostrar_ayuda).pack(fill=tk.X, pady=2)

        # ===== VISTA PREVIA DERECHA =====
        self.canvas = tk.Canvas(right, bg="#f5f5f5")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Button-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_move)
        self.canvas.bind("<ButtonRelease-1>", self._on_drag_end)

        # Cargar datos iniciales
        self._reload_perfiles()
        self._reload_printers()
        self._reload_fonts()

    # === Métodos auxiliares para controles ===
    def _spin_int(self, parent, label, attr, f, t, inc, width=6):
        var = tk.StringVar(value=str(getattr(self.cfg, attr)))
        spin = ttk.Spinbox(parent, from_=f, to=t, increment=inc,
                           textvariable=var, width=width)
        spin.pack(side=tk.LEFT, padx=2)
        self._vars[attr] = var

        def update(*_):
            try:
                value = var.get()
                if value.strip():  # Solo convertir si no está vacío
                    setattr(self.cfg, attr, int(value))
                    self._tal_vez_auto_guardar()
                    self._preview()
            except (ValueError, tk.TclError):
                pass  # Ignorar errores de conversión

        var.trace_add("write", update)

    def _spin_float(self, parent, label, attr, f, t, inc, width=6):
        var = tk.StringVar(value=str(getattr(self.cfg, attr)))
        spin = ttk.Spinbox(parent, from_=f, to=t, increment=inc,
                           textvariable=var, width=width, format="%.3f")
        spin.pack(side=tk.LEFT, padx=2)
        self._vars[attr] = var

        def update(*_):
            try:
                value = var.get()
                if value.strip():  # Solo convertir si no está vacío
                    setattr(self.cfg, attr, float(value))
                    self._tal_vez_auto_guardar()
                    self._preview()
            except (ValueError, tk.TclError):
                pass  # Ignorar errores de conversión

        var.trace_add("write", update)

    def _apply_cfg_to_controls(self):
        self.font_reg_var.set(self.cfg.fuente_reg_name)
        self.font_bold_var.set(self.cfg.fuente_bold_name)
        self.borde_var.set(self.cfg.mostrar_borde)
        for attr, var in self._vars.items():
            try:
                var.set(str(getattr(self.cfg, attr)))
            except Exception:
                pass

    # === Métodos de funcionalidad ===
    def _reload_perfiles(self):
        self.cmb_perfil["values"] = listar_perfiles()

    def _on_select_perfil(self):
        nombre = self.perfil_var.get().strip() or PERFIL_DEFECTO
        self.cfg = cargar_perfil(nombre)
        self._apply_cfg_to_controls()
        self._preview()

    def _guardar_perfil_actual(self):
        guardar_perfil(self.perfil_var.get() or PERFIL_DEFECTO, self.cfg)
        self._reload_perfiles()
        messagebox.showinfo("Perfil", "Perfil guardado.")

    def _guardar_como(self):
        nombre = simpledialog.askstring("Guardar como", "Nombre del nuevo perfil:")
        if nombre:
            guardar_perfil(nombre, self.cfg)
            self.perfil_var.set(nombre)
            self._reload_perfiles()

    def _toggle_auto_guardar(self):
        escribir_auto_guardar_flag(self.auto_guardar_var.get())

    def _tal_vez_auto_guardar(self):
        if self.auto_guardar_var.get():
            try:
                guardar_perfil(self.perfil_var.get() or PERFIL_DEFECTO, self.cfg)
            except Exception:
                pass

    def _pick_csv(self):
        p = filedialog.askopenfilename(
            filetypes=[("CSV o Excel", "*.csv *.xlsx *.xls")]
        )
        if p:
            self.csv_path.set(p)
            self._preview()

    def _reload_printers(self):
        items = ["(Predeterminada)"] + list_printers()
        self.cmb_printer["values"] = items

    def _reload_fonts(self):
        font_names = sorted(FONT_MAPPING.keys())
        self.cmb_font_reg["values"] = font_names
        self.cmb_font_bold["values"] = font_names

    def _update_font_reg(self):
        self.cfg.fuente_reg_name = self.font_reg_var.get()
        self._tal_vez_auto_guardar()
        self._preview()

    def _update_font_bold(self):
        self.cfg.fuente_bold_name = self.font_bold_var.get()
        self._tal_vez_auto_guardar()
        self._preview()

    def _toggle_borde(self):
        self.cfg.mostrar_borde = self.borde_var.get()
        self._tal_vez_auto_guardar()
        self._preview()

    def _build_rows_for_preview(self) -> List[dict]:
        if self.csv_path.get():
            try:
                df = leer_csv(self.csv_path.get())
                return [r.to_dict() for _, r in df.iterrows()][: self.cfg.columnas * self.cfg.filas]
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")
                return []
        # Fallback de ejemplo
        rows = []
        for i in range(min(6, self.cfg.columnas * self.cfg.filas)):
            rows.append({
                "Usuario": f"Usuario {i + 1}",
                "Monitor": f"MON{i + 1}",
                "CPU": f"CPU{i + 1}",
                "UPS": f"UPS{i + 1}",
                "Area": "Oficina",
                "Ext": f"{1000 + i}",
            })
        return rows

    def _get_current_image(self) -> Optional[Image.Image]:
        try:
            self.cfg.fuente_reg_name = self.font_reg_var.get()
            self.cfg.fuente_bold_name = self.font_bold_var.get()
            rows = self._build_rows_for_preview()
            if self.vista_var.get() == "Hoja completa":
                etiquetas = [crear_imagen_etiqueta(self.cfg, r) for r in rows]
                return crear_hoja_completa(self.cfg, etiquetas)
            else:
                row = rows[0] if rows else {
                    "Usuario": "Ejemplo", "Monitor": "CN41", "CPU": "MXL12",
                    "UPS": "2162K", "Area": "Oficina", "Ext": "1234",
                }
                return crear_imagen_etiqueta(self.cfg, row)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar la imagen actual:\n{e}")
            return None

    def _preview(self):
        img = self._get_current_image()
        if img is None:
            return
        try:
            w, h = img.size
            cw, ch = max(1, self.canvas.winfo_width()), max(1, self.canvas.winfo_height())
            scale = min(cw / w, ch / h) * 0.95 if cw and ch else 1.0
            self._preview_scale = max(0.01, scale)
            disp = img.resize((max(1, int(w * self._preview_scale)), max(1, int(h * self._preview_scale))),
                              Image.LANCZOS)
            self._tk_img = ImageTk.PhotoImage(disp)
            self.canvas.delete("all")
            self.canvas.create_image(cw // 2, ch // 2, image=self._tk_img, anchor="center")
        except Exception as e:
            messagebox.showerror("Error vista previa", str(e))

    def _on_drag_start(self, e):
        if self.vista_var.get() != "Etiqueta":
            return
        self._drag_start = (e.x, e.y)
        self._start_offset = (self.cfg.block_offset_x_px, self.cfg.block_offset_y_px)

    def _on_drag_move(self, e):
        if not self._drag_start or self.vista_var.get() != "Etiqueta":
            return
        dx = (e.x - self._drag_start[0]) / (self._preview_scale or 1)
        dy = (e.y - self._drag_start[1]) / (self._preview_scale or 1)
        self.cfg.block_offset_x_px = int(self._start_offset[0] + dx)
        self.cfg.block_offset_y_px = int(self._start_offset[1] + dy)
        self._tal_vez_auto_guardar()
        self._preview()

    def _on_drag_end(self, _):
        self._drag_start = None

    def _guardar_preview(self):
        img = self._get_current_image()
        if img is None:
            return
        inicial = "hoja" if self.vista_var.get() == "Hoja completa" else "etiqueta"
        path = filedialog.asksaveasfilename(
            title="Exportar vista actual",
            defaultextension=".png",
            initialfile=f"{inicial}.png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf")]
        )
        if not path:
            return
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".pdf":
                # Multipágina con todas las hojas del CSV
                MULTIPAGINA = True
                if MULTIPAGINA and self.csv_path.get():
                    df = leer_csv(self.csv_path.get())
                    rows = [r.to_dict() for _, r in df.iterrows()]
                    hojas = crear_hojas_desde_rows(self.cfg, rows) or [img]
                    hojas_rgb = [h.convert("RGB") for h in hojas]
                    hojas_rgb[0].save(path, "PDF", resolution=self.cfg.dpi,
                                      save_all=True, append_images=hojas_rgb[1:])
                else:
                    img.convert("RGB").save(path, "PDF", resolution=self.cfg.dpi)
            else:
                img.save(path, "PNG")
            messagebox.showinfo("Exportar", f"Archivo guardado:\n{path}")
        except Exception as e:
            messagebox.showerror("Exportar", f"No se pudo guardar el archivo:\n{e}")

    def _print_gdi(self):
        if not HAS_PYWIN32:
            messagebox.showerror(
                "pywin32 requerido",
                "Para imprimir sin PDF necesitas instalar pywin32:\n\npy -m pip install pywin32",
            )
            return
        try:
            if not self.csv_path.get():
                messagebox.showwarning("Imprimir", "Primero selecciona un archivo de datos (CSV/Excel).")
                return
            df = leer_csv(self.csv_path.get())
            rows = [r.to_dict() for _, r in df.iterrows()]
            if not rows:
                messagebox.showwarning("Imprimir", "El archivo de datos no tiene filas.")
                return

            hojas = crear_hojas_desde_rows(self.cfg, rows)
            if not hojas:
                messagebox.showwarning("Imprimir", "No hay nada que imprimir.")
                return

            printer = self.printer_var.get()
            if printer == "(Predeterminada)":
                try:
                    printer = win32print.GetDefaultPrinter()
                except Exception:
                    messagebox.showerror("Impresora", "No hay impresora predeterminada configurada.")
                    return

            hDC = win32ui.CreateDC()
            try:
                hDC.CreatePrinterDC(printer)
            except Exception as e:
                messagebox.showerror("Impresora", f"No se pudo abrir la impresora '{printer}'.\n\n{e}")
                return

            HORZRES, VERTRES = 8, 10
            try:
                printable_w = int(hDC.GetDeviceCaps(HORZRES))
                printable_h = int(hDC.GetDeviceCaps(VERTRES))
            except Exception as e:
                hDC.DeleteDC()
                messagebox.showerror("Impresión", f"No se pudieron leer capacidades de la impresora.\n\n{e}")
                return

            if printable_w <= 0 or printable_h <= 0:
                hDC.DeleteDC()
                messagebox.showerror(
                    "Impresión",
                    "La impresora devolvió un área imprimible inválida (0x0). "
                    "Revisa el driver o selecciona otra impresora."
                )
                return

            try:
                hDC.StartDoc("Etiquetas QR")
            except Exception as e:
                hDC.DeleteDC()
                messagebox.showerror("Impresión", f"No se pudo iniciar el documento de impresión.\n\n{e}")
                return

            try:
                for hoja in hojas:
                    hDC.StartPage()
                    ratio = min(printable_w / hoja.width, printable_h / hoja.height)
                    ratio = max(0.01, min(ratio, 10.0))
                    new_w, new_h = int(hoja.width * ratio), int(hoja.height * ratio)
                    dib = ImageWin.Dib(hoja.resize((new_w, new_h), Image.LANCZOS))
                    x0 = (printable_w - new_w) // 2
                    y0 = (printable_h - new_h) // 2
                    dib.draw(hDC.GetHandleOutput(), (x0, y0, x0 + new_w, y0 + new_h))
                    hDC.EndPage()
                hDC.EndDoc()
            except Exception as e:
                try:
                    hDC.EndDoc()
                except Exception:
                    pass
                messagebox.showerror("Impresión", f"Error durante la impresión:\n\n{e}")
            finally:
                hDC.DeleteDC()

            messagebox.showinfo("Impresión completada", f"Se imprimió en: {printer}")

        except Exception as e:
            messagebox.showerror("Error al imprimir", str(e))

    def _mostrar_ayuda(self):
        info = (
            "Generador de Etiquetas QR (Windows)\n"
            "• Vista previa: Hoja completa o Etiqueta\n"
            "• Arrastre del bloque (Etiqueta individual)\n"
            "• Impresión directa GDI (por hojas)\n"
            "• Perfiles (JSON) y parámetros de hoja\n"
            "• Exportación a PNG/PDF\n\n"
            f"Contacto:\n"
            f"Autor: {CONTACTO.get('autor', '')}\n"
            f"Correo: {CONTACTO.get('correo', '')}\n"
            f"Teléfono: {CONTACTO.get('telefono', '')}\n"
            f"Sitio: {CONTACTO.get('sitio', '')}\n"
        )
        messagebox.showinfo("Acerca de…", info)


# === Main ===
if __name__ == "__main__":
    app = EtiquetasGUI()
    app.mainloop()