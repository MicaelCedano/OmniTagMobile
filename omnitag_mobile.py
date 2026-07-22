# -*- coding: utf-8 -*-
"""
OmniTag Mobile - Generador de Etiquetas y Registro Automático Multimarca
Versión: 4.0.7 (Ajuste Proporcional Completo de Etiqueta 2x1 y Código de Barras)
Autor: Micael Cedano
"""
from PIL import Image, ImageDraw, ImageFont, ImageTk
import barcode
from barcode.writer import ImageWriter
import io
import customtkinter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import platform
import subprocess
import tempfile
import atexit
import json
import re
import threading
import time
import pandas as pd
import urllib.request
import urllib.parse
import sys

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module='pymobiledevice3')

CURRENT_VERSION = "v4.0.7"
REPO_OWNER = "MicaelCedano"
REPO_NAME = "OmniTagMobile"

# --- Determinación de Rutas (Modo Script vs Modo PyInstaller .EXE) ---
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(os.path.abspath(sys.executable))
    bundle_dir = getattr(sys, '_MEIPASS', script_dir)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    bundle_dir = script_dir

CONFIG_FILE_NAME = os.path.join(script_dir, "etiqueta_config.json")
EXCEL_FILE_NAME = os.path.join(script_dir, "plantilla_compra_iphone.xlsx")

def _get_asset_path(filename):
    local_p = os.path.join(script_dir, filename)
    if os.path.exists(local_p): return local_p
    return os.path.join(bundle_dir, filename)

FONT_BOLD_PATH_TTF = _get_asset_path("arialbd.ttf")
FONT_REGULAR_PATH_TTF = _get_asset_path("arial.ttf")

RL_FONT_BOLD_NAME = "ArialBoldRegistered"
RL_FONT_REGULAR_NAME = "ArialRegularRegistered"

LABEL_WIDTH_INCHES = 2.0
LABEL_HEIGHT_INCHES = 1.0
PREVIEW_MAX_WIDTH = 380

def parse_version(v_str):
    v_clean = re.sub(r'[^0-9.]', '', str(v_str))
    try:
        return tuple(map(int, v_clean.split('.')))
    except Exception:
        return (0, 0, 0)

# --- Dependencias iOS (pymobiledevice3) ---
PYMOBILEDEVICE_AVAILABLE = False
try:
    from pymobiledevice3.usbmux import list_devices
    from pymobiledevice3.lockdown import create_using_usbmux
    from pymobiledevice3.services.diagnostics import DiagnosticsService
    from pymobiledevice3.exceptions import NoDeviceConnectedError
    PYMOBILEDEVICE_AVAILABLE = True
except Exception as e:
    print(f"ADVERTENCIA: pymobiledevice3 no disponible: {e}")

# --- Dependencias Android (adbutils) ---
ADBUTILS_AVAILABLE = False
try:
    import adbutils
    ADBUTILS_AVAILABLE = True
except Exception as e:
    print(f"ADVERTENCIA: adbutils no disponible: {e}")

# --- Dependencias ReportLab (PDF) ---
PDF_SAVE_ENABLED = False
try:
    from reportlab.pdfgen import canvas as reportlab_canvas
    from reportlab.lib.pagesizes import inch
    from reportlab.lib.utils import ImageReader as ReportLabImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_SAVE_ENABLED = True
except ImportError:
    print("ADVERTENCIA: ReportLab no instalado.")

# --- Mapeo de Modelos de iPhone ---
IPHONE_MODEL_MAPPING = {
    "iPhone1,1": "iPhone", "iPhone1,2": "iPhone 3G", "iPhone2,1": "iPhone 3GS",
    "iPhone3,1": "iPhone 4", "iPhone3,2": "iPhone 4", "iPhone3,3": "iPhone 4",
    "iPhone4,1": "iPhone 4S",
    "iPhone5,1": "iPhone 5", "iPhone5,2": "iPhone 5", "iPhone5,3": "iPhone 5c", "iPhone5,4": "iPhone 5c",
    "iPhone6,1": "iPhone 5s", "iPhone6,2": "iPhone 5s",
    "iPhone7,2": "iPhone 6", "iPhone7,1": "iPhone 6 Plus",
    "iPhone8,1": "iPhone 6s", "iPhone8,2": "iPhone 6s Plus",
    "iPhone8,4": "iPhone SE (1st gen)",
    "iPhone9,1": "iPhone 7", "iPhone9,3": "iPhone 7", "iPhone9,2": "iPhone 7 Plus", "iPhone9,4": "iPhone 7 Plus",
    "iPhone10,1": "iPhone 8", "iPhone10,4": "iPhone 8", "iPhone10,2": "iPhone 8 Plus", "iPhone10,5": "iPhone 8 Plus",
    "iPhone10,3": "iPhone X", "iPhone10,6": "iPhone X",
    "iPhone11,8": "iPhone XR", "iPhone11,2": "iPhone XS", "iPhone11,4": "iPhone XS Max", "iPhone11,6": "iPhone XS Max",
    "iPhone12,1": "iPhone 11", "iPhone12,3": "iPhone 11 Pro", "iPhone12,5": "iPhone 11 Pro Max",
    "iPhone12,8": "iPhone SE (2nd gen)",
    "iPhone13,2": "iPhone 12", "iPhone13,1": "iPhone 12 mini", "iPhone13,3": "iPhone 12 Pro", "iPhone13,4": "iPhone 12 Pro Max",
    "iPhone14,5": "iPhone 13", "iPhone14,4": "iPhone 13 mini", "iPhone14,2": "iPhone 13 Pro", "iPhone14,3": "iPhone 13 Pro Max",
    "iPhone14,6": "iPhone SE (3rd gen)",
    "iPhone14,7": "iPhone 14", "iPhone14,8": "iPhone 14 Plus", "iPhone15,2": "iPhone 14 Pro", "iPhone15,3": "iPhone 14 Pro Max",
    "iPhone15,4": "iPhone 15", "iPhone15,5": "iPhone 15 Plus", "iPhone16,1": "iPhone 15 Pro", "iPhone16,2": "iPhone 15 Pro Max",
    "iPhone17,1": "iPhone 16 Pro", "iPhone17,2": "iPhone 16 Pro Max", "iPhone17,3": "iPhone 16", "iPhone17,4": "iPhone 16 Plus",
    "iPhone18,1": "iPhone 17 Pro", "iPhone18,2": "iPhone 17 Pro Max", "iPhone18,3": "iPhone 17", "iPhone18,4": "iPhone 17 Plus",
}

# --- Mapeo Inteligente por Prefijos para Samsung ---
SAMSUNG_BASE_MAPPING = {
    # Z Flip Series
    "SM-F741": "Samsung Galaxy Z Flip6",
    "SM-F731": "Samsung Galaxy Z Flip5",
    "SM-F721": "Samsung Galaxy Z Flip4",
    "SM-F711": "Samsung Galaxy Z Flip3",
    "SM-F707": "Samsung Galaxy Z Flip 5G",
    "SM-F700": "Samsung Galaxy Z Flip",

    # Z Fold Series
    "SM-F956": "Samsung Galaxy Z Fold6",
    "SM-F946": "Samsung Galaxy Z Fold5",
    "SM-F936": "Samsung Galaxy Z Fold4",
    "SM-F926": "Samsung Galaxy Z Fold3",
    "SM-F916": "Samsung Galaxy Z Fold2",
    "SM-F900": "Samsung Galaxy Fold",

    # S24 Series
    "SM-S928": "Samsung Galaxy S24 Ultra",
    "SM-S926": "Samsung Galaxy S24+",
    "SM-S921": "Samsung Galaxy S24",
    "SM-S924": "Samsung Galaxy S24 FE",

    # S23 Series
    "SM-S918": "Samsung Galaxy S23 Ultra",
    "SM-S916": "Samsung Galaxy S23+",
    "SM-S911": "Samsung Galaxy S23",
    "SM-S914": "Samsung Galaxy S23 FE",

    # S22 Series
    "SM-S908": "Samsung Galaxy S22 Ultra",
    "SM-S906": "Samsung Galaxy S22+",
    "SM-S901": "Samsung Galaxy S22",

    # S21 Series
    "SM-G998": "Samsung Galaxy S21 Ultra",
    "SM-G996": "Samsung Galaxy S21+",
    "SM-G991": "Samsung Galaxy S21",
    "SM-G990": "Samsung Galaxy S21 FE",

    # S20 Series
    "SM-G988": "Samsung Galaxy S20 Ultra",
    "SM-G986": "Samsung Galaxy S20+",
    "SM-G981": "Samsung Galaxy S20",
    "SM-G781": "Samsung Galaxy S20 FE",
    "SM-G780": "Samsung Galaxy S20 FE",

    # Note Series
    "SM-N986": "Samsung Galaxy Note 20 Ultra",
    "SM-N981": "Samsung Galaxy Note 20",
    "SM-N975": "Samsung Galaxy Note 10+",
    "SM-N970": "Samsung Galaxy Note 10",
    "SM-N960": "Samsung Galaxy Note 9",
    "SM-N950": "Samsung Galaxy Note 8",

    # Galaxy A Series
    "SM-A556": "Samsung Galaxy A55 5G",
    "SM-A546": "Samsung Galaxy A54 5G",
    "SM-A536": "Samsung Galaxy A53 5G",
    "SM-A526": "Samsung Galaxy A52 5G",
    "SM-A525": "Samsung Galaxy A52",
    "SM-A515": "Samsung Galaxy A51",
    "SM-A356": "Samsung Galaxy A35 5G",
    "SM-A346": "Samsung Galaxy A34 5G",
    "SM-A336": "Samsung Galaxy A33 5G",
    "SM-A256": "Samsung Galaxy A25 5G",
    "SM-A245": "Samsung Galaxy A24",
    "SM-A156": "Samsung Galaxy A15 5G",
    "SM-A155": "Samsung Galaxy A15",
    "SM-A146": "Samsung Galaxy A14 5G",
    "SM-A145": "Samsung Galaxy A14",
    "SM-A136": "Samsung Galaxy A13 5G",
    "SM-A135": "Samsung Galaxy A13",
    "SM-A125": "Samsung Galaxy A12",
    "SM-A105": "Samsung Galaxy A10",
    "SM-A057": "Samsung Galaxy A05s",
    "SM-A055": "Samsung Galaxy A05",
    "SM-A047": "Samsung Galaxy A04s",
    "SM-A045": "Samsung Galaxy A04",
    "SM-A037": "Samsung Galaxy A03s",
    "SM-A035": "Samsung Galaxy A03",
}

def resolver_nombre_android(brand, model_code, dev=None):
    model_upper = model_code.upper().strip()
    if "SAMSUNG" in brand.upper() or model_upper.startswith("SM-"):
        for prefix, nombre_comercial in SAMSUNG_BASE_MAPPING.items():
            if model_upper.startswith(prefix):
                return nombre_comercial
                
    if dev:
        try:
            mname = dev.shell("getprop ro.product.marketname").strip()
            if mname and len(mname) > 3: return mname
        except Exception: pass
        try:
            mname2 = dev.shell("getprop bluetooth.device.default_name").strip()
            if mname2 and ("Galaxy" in mname2 or "Pixel" in mname2): return mname2
        except Exception: pass

    if "PIXEL" in model_upper:
        return f"Google {model_code}"
        
    return f"{brand} {model_upper}"

INT_COLOR_MAP = {
    "1": "Black", "2": "White", "3": "Gold", "4": "Rose Gold", 
    "5": "Jet Black", "6": "Red", "7": "Silver"
}

COLORES_PARA_REMOVER = [
    "BLACK", "WHITE", "SILVER", "GOLD", "ROSE GOLD", "SPACE GRAY", "SPACE BLACK",
    "GRAPHITE", "MIDNIGHT", "STARLIGHT", "RED", "PRODUCT RED", "BLUE", "SIERRA BLUE",
    "PACIFIC BLUE", "PINK", "GREEN", "MIDNIGHT GREEN", "YELLOW", "PURPLE",
    "DEEP PURPLE", "TITANIUM", "NATURAL TITANIUM", "BLUE TITANIUM", "WHITE TITANIUM",
    "BLACK TITANIUM", "GRAY", "CORAL", "TEAL", "ULTRAMARINE", "DESERT TITANIUM",
    "NEGRO", "BLANCO", "PLATA", "PLATEADO", "DORADO", "ORO", "ROJO", "AZUL",
    "VERDE", "AMARILLO", "ROSA", "PURPURA", "MORADO", "GRIS", "GRIS ESPACIAL"
]

HASH_COLOR_MAP = {
    "#3b3b3b": "Space Gray", "#000000": "Black", "#ffffff": "White",
    "#ff3b30": "Product Red", "#e1e4e3": "Silver", "#f9e5c9": "Gold",
    "#d8a1bc": "Rose Gold", "#121212": "Graphite", "#28344e": "Midnight",
    "#faf7f2": "Starlight", "#a0b4d6": "Sierra Blue", "#574f6f": "Deep Purple",
    "#464c48": "Midnight Green", "#1d4c7b": "Pacific Blue", "#f2f0eb": "Silver",
    "#2e3c4e": "Pacific Blue", "#1f4663": "Pacific Blue", "#003f5d": "Pacific Blue", "#36495d": "Pacific Blue",
    "#e2e4e1": "Silver", "#fad7bd": "Gold", "#afe3b2": "Green", "#ffe681": "Yellow",
    "#fec2dc": "Pink", "#b8afe6": "Purple", "#1e1e1e": "Space Black",
    "#4a4a4c": "Space Gray",
    "#c2bcb2": "Natural Titanium", "#bfa48f": "Desert Titanium", "#3c3c3d": "Black Titanium", "#f2f1ed": "White Titanium",
    "#f2adda": "Pink", "#9aadf6": "Ultramarine", "#b0d4d2": "Teal", "#3c4042": "Black", "#fafafa": "White",
    "#2f4452": "Blue Titanium", "#837f7d": "Natural Titanium", "#1b1b1b": "Black Titanium", "#dddddd": "White Titanium",
    "#e3c8ca": "Pink", "#ced5d9": "Blue", "#cad4c5": "Green", "#e5e0c1": "Yellow", "#35393b": "Black",
    "#594f63": "Deep Purple", "#403e3d": "Space Black", "#f0f2f2": "Silver", "#f4e8ce": "Gold",
    "#e6ddeb": "Purple", "#a0b4c7": "Blue", "#f9e479": "Yellow", "#222930": "Midnight", 
    "#faf6f2": "Starlight", "#fc0324": "Product Red", "#5e5566": "Deep Purple",
    "#a7c1d9": "Sierra Blue", "#576856": "Alpine Green", "#54524f": "Graphite", "#f1f2ed": "Silver", "#fae7cf": "Gold",
    "#276787": "Blue", "#1c5c78": "Blue", "#376e8a": "Blue", "#394c38": "Green", "#faddd7": "Pink", "#232a31": "Midnight", "#bf0013": "Product Red",
    "#95aec5": "Sierra Blue", "#9bb5ce": "Sierra Blue",
    "#201d24": "Black", "#043458": "Blue", "#4c4a46": "Graphite", "#feedd8": "Gold", "#2e4755": "Pacific Blue",
    "#e1f8dc": "Green", "#e23636": "Red", "#fbf7f4": "White",
    "#4e5851": "Midnight Green", "#4e5850": "Midnight Green", "#003e30": "Midnight Green", "#1f352e": "Midnight Green", "#535150": "Space Gray", "#ebebe3": "Silver",
    "#272729": "Space Gray", "#e2e3e4": "Silver", "#f7e8dd": "Gold", "#960111": "Product Red",
}

SUMATRA_PDF_PATH = None
temporary_files_to_delete = []

def _read_config():
    if os.path.exists(CONFIG_FILE_NAME):
        with open(CONFIG_FILE_NAME, 'r', encoding='utf-8') as f:
            try: return json.load(f)
            except json.JSONDecodeError: return {}
    return {}

def _write_config(config_data):
    try:
        with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
        print(f"--- Configuración guardada en: {CONFIG_FILE_NAME} ---")
    except Exception as e:
        print(f"Error guardando config: {e}")

def cargar_config_inicial():
    global SUMATRA_PDF_PATH
    config = _read_config()
    path_guardado = config.get("sumatra_pdf_path")
    if path_guardado and os.path.exists(path_guardado) and os.path.isfile(path_guardado):
        SUMATRA_PDF_PATH = path_guardado
    else:
        detectar_sumatra_si_no_configurado()

def cargar_logo_config():
    config = _read_config()
    logo_path = config.get("logo_path")
    if logo_path and os.path.exists(logo_path) and os.path.isfile(logo_path):
        return logo_path
    return _get_asset_path("logo.png")

def guardar_logo_config(logo_path):
    if logo_path:
        config = _read_config()
        config["logo_path"] = logo_path
        _write_config(config)

def cargar_excel_config():
    config = _read_config()
    path = config.get("last_excel_path")
    if path and os.path.exists(path) and os.path.isfile(path):
        return path
    return EXCEL_FILE_NAME

def guardar_excel_config(excel_path):
    if excel_path:
        config = _read_config()
        config["last_excel_path"] = excel_path
        _write_config(config)

def guardar_config_sumatra():
    if SUMATRA_PDF_PATH and platform.system() == "Windows":
        config = _read_config()
        config["sumatra_pdf_path"] = SUMATRA_PDF_PATH
        _write_config(config)

def detectar_sumatra_si_no_configurado():
    global SUMATRA_PDF_PATH
    if SUMATRA_PDF_PATH or platform.system() != "Windows": return
    paths = [
        "SumatraPDF.exe",
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
    ]
    for p in paths:
        try:
            res = subprocess.run(["where", os.path.basename(p)], capture_output=True, text=True, check=False, shell=True)
            if res.returncode == 0 and res.stdout.strip():
                SUMATRA_PDF_PATH = res.stdout.strip().splitlines()[0]
                return
            elif os.path.exists(p) and os.path.isfile(p):
                SUMATRA_PDF_PATH = p
                return
        except Exception: pass

def cleanup_temp_files():
    for temp_file_path in list(temporary_files_to_delete):
        try:
            if os.path.exists(temp_file_path): os.remove(temp_file_path)
            if temp_file_path in temporary_files_to_delete:
                temporary_files_to_delete.remove(temp_file_path)
        except Exception: pass

atexit.register(cleanup_temp_files)
atexit.register(guardar_config_sumatra)

def cargar_fuentes_pdf():
    global RL_FONT_BOLD_NAME, RL_FONT_REGULAR_NAME
    if not PDF_SAVE_ENABLED: return
    try:
        if os.path.exists(FONT_BOLD_PATH_TTF):
            pdfmetrics.registerFont(TTFont(RL_FONT_BOLD_NAME, FONT_BOLD_PATH_TTF))
        else: raise IOError()
    except Exception:
        RL_FONT_BOLD_NAME = 'Helvetica-Bold'
    try:
        if os.path.exists(FONT_REGULAR_PATH_TTF):
            pdfmetrics.registerFont(TTFont(RL_FONT_REGULAR_NAME, FONT_REGULAR_PATH_TTF))
        else: raise IOError()
    except Exception:
        RL_FONT_REGULAR_NAME = 'Helvetica'

# --- Decodificador Parcel UTF-16 LE para Android ---
def decode_parcel_utf16(parcel_output):
    hex_words = re.findall(r"([0-9a-fA-F]{8})", parcel_output)
    if not hex_words or len(hex_words) < 3:
        return None
        
    chars = []
    for word in hex_words[2:]:
        try:
            u1 = int(word[4:8], 16)
            u2 = int(word[0:4], 16)
            if 32 <= u1 <= 126: chars.append(chr(u1))
            if 32 <= u2 <= 126: chars.append(chr(u2))
        except Exception: pass
        
    result = "".join(chars)
    digits = "".join([c for c in result if c.isdigit()])
    if len(digits) >= 14:
        return digits[:15]
    return None

# --- Extraer IMEI en Android ---
def obtener_imei_android(dev):
    for code in [1, 2, 3, 4, 5, 7, 8, 9, 10]:
        try:
            out = dev.shell(f'service call iphonesubinfo {code} s16 "com.android.shell"')
            if "Result: Parcel" in out:
                imei = decode_parcel_utf16(out)
                if imei: return imei
        except Exception: pass

    for code in [1, 2, 3, 4]:
        try:
            out = dev.shell(f"service call iphonesubinfo {code}")
            if "Result: Parcel" in out:
                imei = decode_parcel_utf16(out)
                if imei: return imei
        except Exception: pass

    cmd_list = [
        'cmd phone get-imei 0 s16 "com.android.shell"',
        'cmd phone get-imei 1 s16 "com.android.shell"',
        'cmd phone get-imei',
        'cmd phone get-device-id'
    ]
    for cmd in cmd_list:
        try:
            res = dev.shell(cmd).strip()
            digits = "".join([c for c in res if c.isdigit()])
            if len(digits) >= 14:
                return digits[:15]
        except Exception: pass

    for d_cmd in ["dumpsys iphonesubinfo", "dumpsys telephony.registry"]:
        try:
            out = dev.shell(d_cmd)
            matches = re.findall(r"(?:mImei|mDeviceId|IMEI|imei)\s*[=:]\s*(\d{14,15})", out)
            for m in matches:
                if len(m) >= 14:
                    return m[:15]
        except Exception: pass

    props = [
        "gsm.baseband.imei", "ril.imei", "ril.IMEI", 
        "ro.ril.oem_imei", "persist.radio.imei", "gsm.imei1", "gsm.imei", "ril.serialnumber"
    ]
    for prop in props:
        try:
            val = dev.shell(f"getprop {prop}").strip()
            digits = "".join([c for c in val if c.isdigit()])
            if len(digits) >= 14:
                return digits[:15]
        except Exception: pass

    return None

# --- Previsualización y Generación de PDF Adaptadas para Etiquetas 2" x 1" ---
def _generar_etiqueta_pil_image(modelo, numero_serie, especificacion, path_logo_pil):
    DPI = 300
    LABEL_WIDTH_PX, LABEL_HEIGHT_PX = int(LABEL_WIDTH_INCHES * DPI), int(LABEL_HEIGHT_INCHES * DPI)  # 600x300 px
    
    TOP_MARGIN_PX = int(0.06 * DPI)    # ~18px
    SIDE_MARGIN_PX = int(0.08 * DPI)   # ~24px
    
    image = Image.new("RGB", (LABEL_WIDTH_PX, LABEL_HEIGHT_PX), "white")
    draw = ImageDraw.Draw(image)
    
    # Fuentes con tamaño optimizado para encajar perfectamente sin cortar el código de barras
    try:
        font_bold = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=24)
        font_regular = ImageFont.truetype(FONT_REGULAR_PATH_TTF, size=20)
        font_sn = ImageFont.truetype(FONT_REGULAR_PATH_TTF, size=18)
    except IOError:
        font_bold, font_regular, font_sn = ImageFont.load_default(), ImageFont.load_default(), ImageFont.load_default()
    
    current_y = TOP_MARGIN_PX
    
    # 1. Logo superior
    if path_logo_pil and os.path.exists(path_logo_pil):
        try:
            with Image.open(path_logo_pil) as logo_img:
                logo_img = logo_img.convert("RGBA")
                logo_max_width = LABEL_WIDTH_PX - 2 * SIDE_MARGIN_PX
                logo_max_height = 42  # Altura máxima adecuada
                logo_img.thumbnail((logo_max_width, logo_max_height), Image.Resampling.LANCZOS)
                logo_x = (LABEL_WIDTH_PX - logo_img.width) // 2
                image.paste(logo_img, (logo_x, current_y), logo_img)
                current_y += logo_img.height + 8
        except Exception as e:
            print(f"Error procesando logo: {e}")

    # 2. Línea Modelo
    if modelo:
        txt_mod = f"Modelo: {modelo}"
        x_mod = (LABEL_WIDTH_PX - draw.textlength(txt_mod, font=font_bold)) // 2
        draw.text((x_mod, current_y), txt_mod, fill="black", font=font_bold)
        current_y += 28

    # 3. Especificación opcional
    if especificacion:
        x_spec = (LABEL_WIDTH_PX - draw.textlength(especificacion, font=font_regular)) // 2
        draw.text((x_spec, current_y), especificacion, fill="black", font=font_regular)
        current_y += 24

    # 4. Línea IMEI
    if numero_serie:
        txt_imei = f"IMEI: {numero_serie}"
        x_imei = (LABEL_WIDTH_PX - draw.textlength(txt_imei, font=font_bold)) // 2
        draw.text((x_imei, current_y), txt_imei, fill="black", font=font_bold)
        current_y += 30

    # 5. Código de Barras e IMEI inferior (Encajado completo)
    if numero_serie:
        try:
            barcode_options = {
                'module_height': 8.0, 'module_width': 0.25,
                'quiet_zone': 3.0, 'write_text': False,
                'font_size': 8
            }
            code128 = barcode.get_barcode_class('code128')
            writer = ImageWriter()
            barcode_obj = code128(numero_serie, writer=writer)
            barcode_pil = barcode_obj.render(barcode_options).convert('RGB')
            
            max_bc_w = LABEL_WIDTH_PX - 2 * SIDE_MARGIN_PX
            if barcode_pil.width > max_bc_w:
                ratio = max_bc_w / barcode_pil.width
                new_w, new_h = int(barcode_pil.width * ratio), int(barcode_pil.height * ratio)
                barcode_pil = barcode_pil.resize((new_w, new_h), Image.Resampling.LANCZOS)

            bc_x = (LABEL_WIDTH_PX - barcode_pil.width) // 2
            image.paste(barcode_pil, (bc_x, current_y))
            current_y += barcode_pil.height + 4

            sn_text_w = draw.textlength(numero_serie, font=font_sn)
            sn_x = (LABEL_WIDTH_PX - sn_text_w) // 2
            draw.text((sn_x, current_y), numero_serie, fill="black", font=font_sn)
        except Exception as e:
            print(f"Error código de barras PIL: {e}")
            
    return image

def _generar_etiqueta_pdf_temporal(modelo, numero_serie, especificacion, path_logo_pil):
    if not PDF_SAVE_ENABLED: return None
    fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="etiqueta_omni_")
    os.close(fd)
    temporary_files_to_delete.append(temp_pdf_path)
    
    c = reportlab_canvas.Canvas(temp_pdf_path, pagesize=(LABEL_WIDTH_INCHES * inch, LABEL_HEIGHT_INCHES * inch))
    width, height = LABEL_WIDTH_INCHES * inch, LABEL_HEIGHT_INCHES * inch
    margin_top, margin_sides = 0.05 * inch, 0.08 * inch
    current_y = height - margin_top

    try:
        if path_logo_pil and os.path.exists(path_logo_pil):
            logo_pil = Image.open(path_logo_pil)
            logo_max_width_pt, logo_max_height_pt = width - 2 * margin_sides, 0.20 * height
            w_px, h_px = logo_pil.size
            aspect = h_px / float(w_px) if w_px > 0 else 0
            logo_w_pt = logo_max_width_pt
            logo_h_pt = logo_w_pt * aspect
            if logo_h_pt > logo_max_height_pt:
                logo_h_pt = logo_max_height_pt
                logo_w_pt = logo_h_pt / aspect if aspect > 0 else 0

            img_reader = ReportLabImageReader(logo_pil)
            current_y -= logo_h_pt
            c.drawImage(img_reader, (width - logo_w_pt) / 2, current_y, width=logo_w_pt, height=logo_h_pt, mask='auto')
            current_y -= 4
    except Exception as e:
        print(f"Error logo PDF: {e}")

    info_items_pdf = []
    if modelo:
        info_items_pdf.append((f"Modelo: {modelo}", RL_FONT_BOLD_NAME, 8))
    if especificacion:
        info_items_pdf.append((especificacion, RL_FONT_REGULAR_NAME, 7))
    if numero_serie:
        info_items_pdf.append((f"IMEI: {numero_serie}", RL_FONT_BOLD_NAME, 8))
    
    for texto, font, size in info_items_pdf:
        current_y -= size
        c.setFont(font, size)
        c.drawCentredString(width / 2, current_y, texto)
        current_y -= 2

    if numero_serie:
        try:
            current_y -= 2
            barcode_options = {
                'module_height': 7.0, 'module_width': 0.25,
                'quiet_zone': 2.0, 'write_text': False,
                'font_size': 7
            }
            code128 = barcode.get_barcode_class('code128')
            writer = ImageWriter()
            barcode_obj = code128(numero_serie, writer=writer)
            barcode_pil_pdf = barcode_obj.render(barcode_options).convert('RGB')
            
            barcode_io = io.BytesIO()
            barcode_pil_pdf.save(barcode_io, format='PNG')
            barcode_io.seek(0)
            
            img_reader = ReportLabImageReader(barcode_io)
            bc_w, bc_h = img_reader.getSize()
            max_bc_w = width - (2 * margin_sides)
            if bc_w > max_bc_w:
                ratio = max_bc_w / bc_w
                bc_w, bc_h = bc_w * ratio, bc_h * ratio
            
            current_y -= bc_h
            c.drawImage(img_reader, (width - bc_w) / 2, current_y, width=bc_w, height=bc_h, mask='auto')
            
            sn_font_size = 7
            current_y -= 1 + sn_font_size
            c.setFont(RL_FONT_REGULAR_NAME, sn_font_size)
            c.drawCentredString(width / 2, current_y, numero_serie)
        except Exception as e:
            print(f"Error código de barras PDF: {e}")
            
    c.save()
    return temp_pdf_path

# --- Manejador Excel ---
class ExcelManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.ensure_file_exists()

    def set_filepath(self, new_path):
        self.filepath = new_path
        self.ensure_file_exists()

    def ensure_file_exists(self):
        if not os.path.exists(self.filepath):
            try:
                df = pd.DataFrame(columns=["IMEI", "Modelo"])
                df.to_excel(self.filepath, index=False)
            except Exception as e:
                print(f"Error creando Excel: {e}")

    def registrar_dispositivo(self, imei, modelo_completo):
        try:
            try: df = pd.read_excel(self.filepath)
            except FileNotFoundError: df = pd.DataFrame(columns=["IMEI", "Modelo"])
            
            df['IMEI'] = df['IMEI'].astype(str)
            if str(imei) in df['IMEI'].values:
                return False, f"El IMEI {imei} ya está registrado.", len(df)

            new_row = pd.DataFrame([{"IMEI": str(imei), "Modelo": modelo_completo}])
            df = pd.concat([df, new_row], ignore_index=True)
            df.to_excel(self.filepath, index=False)
            return True, "Registrado en Excel.", len(df)
        except PermissionError:
            return False, "Error: El archivo Excel está abierto. Ciérralo para guardar.", -1
        except Exception as e:
            return False, f"Error Excel: {e}", -1

    def actualizar_registro(self, imei, nuevo_modelo):
        try:
            df = pd.read_excel(self.filepath)
            df['IMEI'] = df['IMEI'].astype(str)
            imei_str = str(imei)
            if imei_str in df['IMEI'].values:
                idx = df.index[df['IMEI'] == imei_str].tolist()[0]
                df.at[idx, 'Modelo'] = nuevo_modelo
                df.to_excel(self.filepath, index=False)
                return True
            return False
        except Exception as e:
            print(f"Error actualizando Excel: {e}")
            return False

    def eliminar_registro(self, imei):
        try:
            df = pd.read_excel(self.filepath)
            df['IMEI'] = df['IMEI'].astype(str)
            imei_str = str(imei)
            if imei_str in df['IMEI'].values:
                df = df[df['IMEI'] != imei_str]
                df.to_excel(self.filepath, index=False)
                return True
            return False
        except Exception as e:
            print(f"Error eliminando de Excel: {e}")
            return False

    def obtener_conteo(self):
        try:
            if os.path.exists(self.filepath):
                return len(pd.read_excel(self.filepath))
        except: pass
        return 0

    def obtener_registros(self):
        try:
            if os.path.exists(self.filepath):
                return pd.read_excel(self.filepath).iloc[::-1].to_dict('records')
        except: pass
        return []

# --- Monitor de Dispositivos Unificado (iOS + Android) ---
class UnifiedDeviceMonitor(threading.Thread):
    def __init__(self, success_callback, trust_callback, disconnected_callback):
        super().__init__()
        self.success_callback = success_callback
        self.trust_callback = trust_callback
        self.disconnected_callback = disconnected_callback
        self.running = True
        self.last_udid = None
        self.daemon = True

    def run(self):
        print("--- Monitor de Dispositivos Unificado Iniciado (iOS & Android) ---")
        while self.running:
            device_found = False
            
            # 1. Intentar detectar iPhone por USBmux
            if PYMOBILEDEVICE_AVAILABLE:
                try:
                    devices = list_devices()
                    if devices:
                        device = devices[0]
                        udid = device.serial
                        if udid != self.last_udid:
                            lockdown = create_using_usbmux(serial=udid)
                            try: lockdown.validate_pairing()
                            except Exception as ep:
                                if "SessionActive" not in str(ep): raise ep
                            
                            product_type = lockdown.get_value(key='ProductType')
                            serial_number = lockdown.get_value(key='SerialNumber')
                            imei = lockdown.get_value(key='InternationalMobileEquipmentIdentity')
                            model_name = IPHONE_MODEL_MAPPING.get(product_type, "iPhone Desconocido")
                            
                            if not imei: raise Exception("IMEI vacío")
                            
                            capacidad = ""
                            try:
                                bytes_cap = lockdown.get_value(domain='com.apple.disk_usage', key='TotalDiskCapacity') or lockdown.get_value(key='TotalDiskCapacity')
                                if bytes_cap:
                                    gb = int(bytes_cap) / 1000000000
                                    closest = min([16, 32, 64, 128, 256, 512, 1024], key=lambda x: abs(x - gb))
                                    capacidad = "1TB" if closest == 1024 else f"{closest}GB"
                            except: pass

                            color = ""
                            try:
                                color_hex = lockdown.get_value(key='DeviceColor')
                                color = HASH_COLOR_MAP.get(str(color_hex).lower().strip(), "")
                            except: pass

                            info = {
                                'udid': udid,
                                'brand': 'Apple',
                                'product_name': model_name,
                                'serial_number': serial_number,
                                'imei': imei,
                                'capacity': capacidad,
                                'color': color,
                                'is_android': False
                            }
                            self.last_udid = udid
                            self.success_callback(info)
                        device_found = True
                except Exception as e:
                    if "PairingDialogResponsePending" in str(e) or "PasswordProtected" in str(e):
                        self.trust_callback("iOS")
                    else: pass

            # 2. Si no hay iPhone, intentar detectar Android por ADB
            if not device_found and ADBUTILS_AVAILABLE:
                try:
                    adb_devices = adbutils.adb.device_list()
                    if adb_devices:
                        dev = adb_devices[0]
                        udid = dev.serial
                        if udid != self.last_udid:
                            brand = dev.shell("getprop ro.product.brand").strip().capitalize()
                            model_code = dev.shell("getprop ro.product.model").strip()
                            
                            model_name = resolver_nombre_android(brand, model_code, dev)
                            serial = dev.shell("getprop ro.serialno").strip() or udid
                            
                            imei = obtener_imei_android(dev)
                            if not imei:
                                imei = serial

                            capacidad = ""
                            try:
                                df_out = dev.shell("df -h /data")
                                size_str = df_out.strip().split("\n")[-1].split()[1]
                                gb_val = float(re.sub(r'[^\d.]', '', size_str))
                                closest = min([32, 64, 128, 256, 512, 1024], key=lambda x: abs(x - gb_val))
                                capacidad = "1TB" if closest == 1024 else f"{closest}GB"
                            except: pass

                            info = {
                                'udid': udid,
                                'brand': brand,
                                'product_name': model_name,
                                'serial_number': serial,
                                'imei': imei,
                                'capacity': capacidad,
                                'color': "",
                                'is_android': True
                            }
                            self.last_udid = udid
                            self.success_callback(info)
                        device_found = True
                except Exception as e:
                    print(f"Error ADB: {e}")

            if not device_found and self.last_udid is not None:
                self.last_udid = None
                self.disconnected_callback()

            time.sleep(2)

    def stop(self):
        self.running = False

# --- Colores de UI ---
COLOR_BG_DARK = "#0F172A"
COLOR_CARD_BG = "#1E293B"
COLOR_CARD_BORDER = "#334155"
COLOR_TEXT_PRIMARY = "#F8FAFC"
COLOR_TEXT_SECONDARY = "#94A3B8"
COLOR_ACCENT_PRIMARY = "#6366F1"
COLOR_ACCENT_PRIMARY_HOVER = "#4F46E5"
COLOR_ACCENT_SUCCESS = "#10B981"
COLOR_ACCENT_SUCCESS_HOVER = "#059669"
COLOR_ACCENT_SECONDARY = "#334155"
COLOR_ACCENT_SECONDARY_HOVER = "#475569"
COLOR_ACCENT_DANGER = "#EF4444"
COLOR_ACCENT_DANGER_HOVER = "#DC2626"

# --- Aplicación Principal ---
class OmniTagMobileApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.preview_ctk_image = None
        self._preview_update_job = None
        self._excel_update_job = None
        self.trust_window = None
        self.current_udid = None
        self.current_device_info = None
        
        # Variables de Actualización Estilo MCTools
        self.update_ready = False
        self.downloaded_new_exe = None

        start_excel = cargar_excel_config()
        self.excel_manager = ExcelManager(start_excel)
        
        self.title(f"OmniTag Mobile {CURRENT_VERSION} - Detección Multimarca & Etiquetas")
        self.geometry("1300x740")
        self.minsize(1200, 640)
        self.configure(fg_color=COLOR_BG_DARK)
        
        cargar_config_inicial()
        cargar_fuentes_pdf()
        
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=1)

        # Header Banner
        self.header_frame = customtkinter.CTkFrame(self, fg_color=COLOR_CARD_BG, corner_radius=12, border_width=1, border_color=COLOR_CARD_BORDER, height=80)
        self.header_frame.grid(row=0, column=0, columnspan=3, padx=20, pady=(15, 10), sticky="ew")
        self.header_frame.grid_propagate(False)
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)
        
        title_box = customtkinter.CTkFrame(self.header_frame, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        lbl_main_title = customtkinter.CTkLabel(title_box, text="OMNITAG MOBILE", font=customtkinter.CTkFont(family="Segoe UI", size=20, weight="bold"), text_color=COLOR_TEXT_PRIMARY)
        lbl_main_title.pack(anchor="w")
        
        lbl_sub_title = customtkinter.CTkLabel(title_box, text=f"{CURRENT_VERSION} • Lectura Multimarca (iPhone / Samsung / Pixel) & Control de Inventario Excel", font=customtkinter.CTkFont(family="Segoe UI", size=11), text_color=COLOR_TEXT_SECONDARY)
        lbl_sub_title.pack(anchor="w")
        
        # Header Derecho: Badge Estado + Botón de Actualización
        right_header_box = customtkinter.CTkFrame(self.header_frame, fg_color="transparent")
        right_header_box.grid(row=0, column=1, padx=20, pady=12, sticky="e")

        self.badge_frame = customtkinter.CTkFrame(right_header_box, fg_color="#334155", corner_radius=18, height=36, border_width=1, border_color="#475569")
        self.badge_frame.grid(row=0, column=0, padx=(0, 10))
        self.badge_frame.grid_propagate(False)
        
        self.badge_label = customtkinter.CTkLabel(self.badge_frame, text="Estado: Desconectado ❌", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#94A3B8", padx=15, pady=0)
        self.badge_label.pack(expand=True, fill="both")

        self.btn_update = customtkinter.CTkButton(
            right_header_box, 
            text="Buscar act.", 
            command=self.accion_boton_actualizacion, 
            width=120, 
            height=36, 
            corner_radius=18, 
            fg_color="#334155", 
            hover_color="#475569", 
            font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold")
        )
        self.btn_update.grid(row=0, column=1)

        # Barra de progreso inline para actualización (Estilo MCTools)
        self.update_progress = customtkinter.CTkProgressBar(self.header_frame, height=4, corner_radius=2, fg_color="#0F172A", progress_color="#06B6D4")
        self.update_progress.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 2))
        self.update_progress.set(0)
        self.update_progress.grid_remove()

        # Contenedores Principales
        self.controls_frame = customtkinter.CTkFrame(self, width=320, fg_color=COLOR_CARD_BG, corner_radius=16, border_width=1, border_color=COLOR_CARD_BORDER)
        self.controls_frame.grid(row=1, column=0, padx=(20, 10), pady=(10, 20), sticky="nsew")
        self.controls_frame.grid_rowconfigure(2, weight=1)
        self.controls_frame.grid_columnconfigure(0, weight=1)
        self.controls_frame.grid_propagate(False)

        self.preview_frame = customtkinter.CTkFrame(self, width=420, fg_color=COLOR_CARD_BG, corner_radius=16, border_width=1, border_color=COLOR_CARD_BORDER)
        self.preview_frame.grid(row=1, column=1, padx=10, pady=(10, 20), sticky="nsew")
        self.preview_frame.grid_propagate(False)

        self.excel_frame = customtkinter.CTkFrame(self, fg_color=COLOR_CARD_BG, corner_radius=16, border_width=1, border_color=COLOR_CARD_BORDER)
        self.excel_frame.grid(row=1, column=2, padx=(10, 20), pady=(10, 20), sticky="nsew")

        # Monitor Unificado
        self.auto_print_var = tk.BooleanVar(value=False)
        self.device_monitor = UnifiedDeviceMonitor(
            success_callback=self.on_device_info_received,
            trust_callback=self.on_device_trust_needed,
            disconnected_callback=self.on_device_disconnected
        )
        self.device_monitor.start()

        self._setup_ui()
        self._bind_events()
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.after(100, self.force_preview_update)
        
        # Chequeo automático de actualizaciones al iniciar (estilo MCTools)
        self.after(1500, lambda: self.chequear_actualizaciones_async(manual=False))

    def on_closing(self):
        """Intercepta el cierre de la app. Si hay una actualización lista, la aplica al salir."""
        if getattr(self, 'update_ready', False) and getattr(self, 'downloaded_new_exe', None) and os.path.exists(self.downloaded_new_exe):
            self.ejecutar_instalacion_inmediata()
        else:
            if self.device_monitor: self.device_monitor.stop()
            try: guardar_logo_config(self.logo_path_var.get().strip())
            except Exception: pass
            finally:
                cleanup_temp_files()
                try: self.destroy()
                except Exception: pass
                os._exit(0)

    # --- Lógica de Auto-Update Estilo MCTools ---
    def accion_boton_actualizacion(self):
        if getattr(self, 'update_ready', False) and getattr(self, 'downloaded_new_exe', None):
            self.ejecutar_instalacion_inmediata()
        else:
            self.chequear_actualizaciones_async(manual=True)

    def chequear_actualizaciones_async(self, manual=False):
        if getattr(self, 'update_ready', False): return
        if manual:
            self.btn_update.configure(text="Buscando...", fg_color="#334155", state="disabled")
        
        thread = threading.Thread(target=self._buscar_actualizaciones_hilo, args=(manual,))
        thread.daemon = True
        thread.start()

    def _buscar_actualizaciones_hilo(self, manual):
        try:
            url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode('utf-8'))
                
            latest_version_tag = data.get("tag_name", "")
            latest_version = latest_version_tag.lstrip('v')
            if not any(c.isdigit() for c in latest_version):
                release_title = data.get("name", "")
                if release_title:
                    latest_version_tag = release_title
                    latest_version = release_title.lstrip('v')
            
            current_ver = CURRENT_VERSION.lstrip('v')
            
            if parse_version(latest_version) > parse_version(current_ver):
                assets = data.get("assets", [])
                exe_url = None
                for asset in assets:
                    name = asset.get("name", "")
                    if name.endswith(".exe"):
                        exe_url = asset.get("browser_download_url")
                        break
                
                if exe_url and getattr(sys, 'frozen', False):
                    self.after(100, lambda: self.iniciar_descarga_inline(exe_url, latest_version_tag))
                else:
                    self.after(100, lambda: self.btn_update.configure(text="¡Nueva v" + latest_version + "!", fg_color="#EF4444", state="normal"))
            else:
                self.after(100, lambda: self.btn_update.configure(text="Al día", fg_color="#10B981", state="normal"))
                if manual:
                    self.after(200, lambda: messagebox.showinfo("Actualizado", f"Ya tienes la versión más reciente ({CURRENT_VERSION})."))
        except Exception as e:
            print(f"Error buscando actualización: {e}")
            self.after(100, lambda: self.btn_update.configure(text="Buscar act.", fg_color="#334155", state="normal"))
            if manual:
                self.after(200, lambda: messagebox.showerror("Error", f"No se pudo buscar actualizaciones:\n{e}"))
        finally:
            self.after(900000, lambda: self.chequear_actualizaciones_async(manual=False))

    def iniciar_descarga_inline(self, exe_url, nueva_version_tag):
        self.btn_update.configure(text="Descargando 0%", fg_color="#0284C7", state="disabled")
        self.update_progress.grid()
        self.update_progress.set(0)
        
        thread = threading.Thread(target=self._hilo_descarga_inline, args=(exe_url, nueva_version_tag))
        thread.daemon = True
        thread.start()

    def _hilo_descarga_inline(self, exe_url, nueva_version_tag):
        try:
            temp_dir = tempfile.gettempdir()
            new_exe = os.path.join(temp_dir, f"omnitag_update_{int(time.time())}.exe")
            
            req = urllib.request.Request(exe_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            with urllib.request.urlopen(req, timeout=45) as response:
                total_size = int(response.info().get('Content-Length', 0))
                bytes_downloaded = 0
                
                with open(new_exe, 'wb') as f:
                    while True:
                        chunk = response.read(8192)
                        if not chunk: break
                        f.write(chunk)
                        bytes_downloaded += len(chunk)
                        
                        if total_size > 0:
                            progreso = bytes_downloaded / total_size
                            porcentaje = int(progreso * 100)
                            self.after(0, lambda p=progreso, pct=porcentaje: self._actualizar_progreso_inline(p, pct))
            
            self.after(0, lambda: self._finalizar_descarga_inline(nueva_version_tag, new_exe))
        except Exception as e:
            print(f"Error en descarga inline: {e}")
            self.after(0, self._error_descarga_inline)

    def _actualizar_progreso_inline(self, progreso, porcentaje):
        self.update_progress.set(progreso)
        self.btn_update.configure(text=f"Descargando {porcentaje}%")

    def _finalizar_descarga_inline(self, nueva_version_tag, new_exe):
        self.update_progress.grid_remove()
        self.update_ready = True
        self.downloaded_new_exe = new_exe
        self.btn_update.configure(
            text="✨ Instalar ahora", 
            fg_color="#10B981", 
            hover_color="#059669", 
            text_color="#FFFFFF",
            state="normal"
        )

    def _error_descarga_inline(self):
        self.update_progress.grid_remove()
        self.btn_update.configure(text="Buscar act.", fg_color="#334155", state="normal")

    def ejecutar_instalacion_inmediata(self):
        if not hasattr(self, 'downloaded_new_exe') or not self.downloaded_new_exe or not os.path.exists(self.downloaded_new_exe):
            messagebox.showerror("Error", "No se encontró el archivo de actualización listo para instalar.")
            return
            
        try:
            current_exe = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
            temp_dir = tempfile.gettempdir()
            new_exe = self.downloaded_new_exe
            bat_path = os.path.join(temp_dir, "omnitag_updater.bat")
            vbs_path = os.path.join(temp_dir, "omnitag_launcher.vbs")
            
            exe_basename = os.path.basename(current_exe)
            bat_lines = [
                "@echo off",
                ":wait_exit",
                "ping 127.0.0.1 -n 2 >nul",
                f'tasklist /FI "IMAGENAME eq {exe_basename}" 2>nul | find /I "{exe_basename}" >nul',
                "if not errorlevel 1 goto wait_exit",
                "ping 127.0.0.1 -n 3 >nul",
                ":retry_copy",
                f'copy /Y "{new_exe}" "{current_exe}" >nul 2>&1 || goto retry_copy',
                "ping 127.0.0.1 -n 2 >nul",
                f'start "" "{current_exe}"',
                "ping 127.0.0.1 -n 2 >nul",
                f'if exist "{vbs_path}" del /F /Q "{vbs_path}" >nul 2>&1',
                f'if exist "{new_exe}" del /F /Q "{new_exe}" >nul 2>&1',
                '(goto) 2>nul & del "%~f0" >nul 2>&1'
            ]
            
            with open(bat_path, 'w', encoding='cp1252') as f:
                f.write("\r\n".join(bat_lines))
                
            vbs_code = f'CreateObject("WScript.Shell").Run Chr(34) & "{bat_path}" & Chr(34), 0, False'
            with open(vbs_path, 'w', encoding='cp1252') as f:
                f.write(vbs_code)
                
            if self.device_monitor: self.device_monitor.stop()
            subprocess.Popen(['wscript.exe', vbs_path])
            time.sleep(0.3)
            os._exit(0)
        except Exception as e:
            messagebox.showerror("Error de Instalación", f"No se pudo iniciar la actualización:\n{e}")

    # --- UI Base ---
    def _setup_ui(self):
        self.controls_frame.grid_columnconfigure(0, weight=1)
        
        self.modelo_var = tk.StringVar()
        self.imei_var = tk.StringVar()
        self.logo_path_var = tk.StringVar(value=cargar_logo_config())
        self.status_label = self.badge_label
        
        main_controls_frame = customtkinter.CTkFrame(self.controls_frame, fg_color="transparent")
        main_controls_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        main_controls_frame.grid_columnconfigure(0, weight=1)

        # Modelo
        customtkinter.CTkLabel(main_controls_frame, text="Modelo Dispositivo:", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY).grid(row=0, column=0, sticky="w")
        modelo_entry_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        modelo_entry_frame.grid(row=1, column=0, pady=(2,10), sticky="ew")
        modelo_entry_frame.grid_columnconfigure(0, weight=1)
        
        self.modelo_entry = customtkinter.CTkEntry(modelo_entry_frame, textvariable=self.modelo_var, placeholder_text="Ej: Samsung S23 Ultra Black 256GB", fg_color="#0F172A", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, corner_radius=8, height=32)
        self.modelo_entry.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        btn_paste_model = customtkinter.CTkButton(modelo_entry_frame, text="Pegar", width=55, command=self.pegar_modelo, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, corner_radius=8, height=32)
        btn_paste_model.grid(row=0, column=1)

        # IMEI / S/N
        customtkinter.CTkLabel(main_controls_frame, text="IMEI:", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY).grid(row=2, column=0, sticky="w")
        imei_entry_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        imei_entry_frame.grid(row=3, column=0, pady=(2,15), sticky="ew")
        imei_entry_frame.grid_columnconfigure(0, weight=1)
        
        self.imei_entry = customtkinter.CTkEntry(imei_entry_frame, textvariable=self.imei_var, placeholder_text="Ingrese IMEI...", fg_color="#0F172A", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, corner_radius=8, height=32)
        self.imei_entry.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        btn_paste_imei = customtkinter.CTkButton(imei_entry_frame, text="Pegar", width=55, command=self.pegar_imei, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, corner_radius=8, height=32)
        btn_paste_imei.grid(row=0, column=1)

        # Logo
        logo_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="#0F172A", corner_radius=10, border_width=1, border_color=COLOR_CARD_BORDER)
        logo_frame.grid(row=4, column=0, sticky='ew', pady=(0,15))
        logo_frame.grid_columnconfigure(0, weight=1)
        
        customtkinter.CTkLabel(logo_frame, text="Ruta del Logo:", font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=COLOR_TEXT_SECONDARY).grid(row=0, column=0, columnspan=2, padx=12, pady=(8,2), sticky="w")
        self.logo_entry = customtkinter.CTkEntry(logo_frame, textvariable=self.logo_path_var, fg_color="#1E293B", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, corner_radius=6, height=28)
        self.logo_entry.grid(row=1, column=0, padx=(12,5), pady=(0,10), sticky='ew')
        
        btn_search_logo = customtkinter.CTkButton(logo_frame, text="Buscar...", width=70, command=self.buscar_logo, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, corner_radius=6, height=28)
        btn_search_logo.grid(row=1, column=1, padx=(0,12), pady=(0,10))

        # Botones PDF & Impresión
        pdf_buttons_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        pdf_buttons_frame.grid(row=5, column=0, sticky='ew', pady=(0,15))
        pdf_buttons_frame.grid_columnconfigure((0,1), weight=1)
        
        btn_save_pdf = customtkinter.CTkButton(pdf_buttons_frame, text="Guardar PDF 💾", command=self.generar_y_guardar_pdf, fg_color=COLOR_ACCENT_PRIMARY, hover_color=COLOR_ACCENT_PRIMARY_HOVER, corner_radius=10, height=38)
        btn_save_pdf.grid(row=0, column=0, padx=(0,5), sticky='ew')
        
        btn_print = customtkinter.CTkButton(pdf_buttons_frame, text="Imprimir 🖨️", command=self.imprimir, fg_color=COLOR_ACCENT_SUCCESS, hover_color=COLOR_ACCENT_SUCCESS_HOVER, corner_radius=10, height=38)
        btn_print.grid(row=0, column=1, padx=(5,0), sticky='ew')

        # Switches
        switch_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        switch_frame.grid(row=6, column=0, pady=(5,0), sticky="ew")
        
        self.sw_auto_print = customtkinter.CTkSwitch(switch_frame, text="Auto-Imprimir al Conectar", variable=self.auto_print_var, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), progress_color=COLOR_ACCENT_SUCCESS, text_color=COLOR_TEXT_SECONDARY)
        self.sw_auto_print.pack(anchor="w", pady=(0,8))
        
        self.auto_shutdown_var = tk.BooleanVar(value=False)
        self.sw_auto_shutdown = customtkinter.CTkSwitch(switch_frame, text="Apagar al Terminar", variable=self.auto_shutdown_var, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), progress_color=COLOR_ACCENT_DANGER, text_color="#EF4444")
        self.sw_auto_shutdown.pack(anchor="w")

        # Pie
        bottom_frame = customtkinter.CTkFrame(self.controls_frame, fg_color="transparent")
        bottom_frame.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="s")
        bottom_frame.grid_columnconfigure(0, weight=1)
        
        btn_sumatra = customtkinter.CTkButton(bottom_frame, text="⚙️ Configurar SumatraPDF", command=self.configurar_ruta_sumatra_manualmente, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, height=30)
        btn_sumatra.grid(row=0, column=0, sticky="ew", pady=(0,10))
        
        lbl_author = customtkinter.CTkLabel(bottom_frame, text="OmniTag Mobile • Micael Cedano", font=customtkinter.CTkFont(family="Segoe UI", size=10, slant="italic"), text_color="gray50")
        lbl_author.grid(row=1, column=0, sticky="w")

        # Vista Previa
        self.preview_frame.grid_rowconfigure(1, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        
        lbl_preview_title = customtkinter.CTkLabel(self.preview_frame, text="VISTA PREVIA DE ETIQUETA", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY)
        lbl_preview_title.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="w")
        
        preview_container = customtkinter.CTkFrame(self.preview_frame, fg_color="#111827", corner_radius=12, border_width=1, border_color="#1F2937")
        preview_container.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")
        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)

        self.preview_image_label = customtkinter.CTkLabel(preview_container, text="La previsualización aparecerá aquí.", font=customtkinter.CTkFont(family="Segoe UI", size=11, slant="italic"), text_color="gray50")
        self.preview_image_label.grid(row=0, column=0)

        # Excel Section
        self._setup_excel_view()

    def _setup_excel_view(self):
        self.excel_frame.grid_rowconfigure(4, weight=1)
        self.excel_frame.grid_columnconfigure(0, weight=1)
        
        db_header_frame = customtkinter.CTkFrame(self.excel_frame, fg_color="transparent")
        db_header_frame.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        db_header_frame.grid_columnconfigure(0, weight=1)
        db_header_frame.grid_columnconfigure(1, weight=0)
        
        lbl_db_title = customtkinter.CTkLabel(db_header_frame, text="HISTORIAL DE REGISTROS EXCEL", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY)
        lbl_db_title.grid(row=0, column=0, sticky="w")
        
        self.count_label = customtkinter.CTkLabel(db_header_frame, text=f"Total: {self.excel_manager.obtener_conteo()}", font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color="#06B6D4")
        self.count_label.grid(row=0, column=1, sticky="e")
        
        project_frame = customtkinter.CTkFrame(self.excel_frame, fg_color="#0F172A", corner_radius=10, border_width=1, border_color=COLOR_CARD_BORDER, height=36)
        project_frame.grid(row=1, column=0, padx=15, pady=(5, 10), sticky="ew")
        project_frame.grid_propagate(False)
        project_frame.grid_columnconfigure(1, weight=1)
        
        lbl_proj = customtkinter.CTkLabel(project_frame, text=" Excel:", font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=COLOR_TEXT_SECONDARY)
        lbl_proj.grid(row=0, column=0, padx=(12, 2), pady=4, sticky="w")
        
        self.lbl_filename = customtkinter.CTkLabel(project_frame, text=os.path.basename(self.excel_manager.filepath), font=customtkinter.CTkFont(family="Consolas", size=11, weight="bold"), text_color="#10B981", anchor="w")
        self.lbl_filename.grid(row=0, column=1, padx=(2, 10), pady=4, sticky="ew")

        btn_box = customtkinter.CTkFrame(self.excel_frame, fg_color="transparent")
        btn_box.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")
        btn_box.grid_columnconfigure((0,1,2), weight=1)
        
        btn_new = customtkinter.CTkButton(btn_box, text="Nuevo Excel", command=self.nuevo_proyecto, fg_color=COLOR_ACCENT_SUCCESS, hover_color=COLOR_ACCENT_SUCCESS_HOVER, corner_radius=8, height=32)
        btn_new.grid(row=0, column=0, padx=(0,4), sticky="ew")
        
        btn_load = customtkinter.CTkButton(btn_box, text="Cargar Excel", command=self.cargar_proyecto, fg_color="#3B82F6", hover_color="#2563EB", corner_radius=8, height=32)
        btn_load.grid(row=0, column=1, padx=2, sticky="ew")
        
        btn_export = customtkinter.CTkButton(btn_box, text="Exportar Copia", command=self.exportar_proyecto, fg_color=COLOR_ACCENT_PRIMARY, hover_color=COLOR_ACCENT_PRIMARY_HOVER, corner_radius=8, height=32)
        btn_export.grid(row=0, column=2, padx=(4,0), sticky="ew")
        
        search_frame = customtkinter.CTkFrame(self.excel_frame, fg_color="transparent")
        search_frame.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_change)
        
        search_entry = customtkinter.CTkEntry(search_frame, textvariable=self.search_var, placeholder_text="🔍  Buscar IMEI o Modelo...", fg_color="#0F172A", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, corner_radius=8, height=32)
        search_entry.grid(row=0, column=0, sticky="ew")

        table_container = customtkinter.CTkFrame(self.excel_frame, fg_color="#0F172A", corner_radius=12, border_width=1, border_color=COLOR_CARD_BORDER)
        table_container.grid(row=4, column=0, padx=15, pady=(0, 15), sticky="nsew")
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#1E293B", foreground="#F8FAFC", fieldbackground="#1E293B", bordercolor="#334155", rowheight=28, font=('Segoe UI', 10))
        style.map('Treeview', background=[('selected', '#4F46E5')], foreground=[('selected', '#FFFFFF')])
        style.configure("Treeview.Heading", background="#334155", foreground="#F8FAFC", borderwidth=0, font=('Segoe UI', 10, 'bold'))
        
        self.tree = ttk.Treeview(table_container, columns=("IMEI", "Modelo"), show="headings", style="Treeview")
        self.tree.heading("IMEI", text="IMEI")
        self.tree.heading("Modelo", text="Modelo del Dispositivo")
        self.tree.column("IMEI", width=140, anchor="center")
        self.tree.column("Modelo", width=250, anchor="w")
        
        self.tree.tag_configure("evenrow", background="#1E293B")
        self.tree.tag_configure("oddrow", background="#161E2E")
        
        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(1, 0), pady=1)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 1), pady=1)

        self.tree_menu = tk.Menu(self.tree, tearoff=0, bg="#1E293B", fg="#F8FAFC", activebackground="#4F46E5", activeforeground="#FFFFFF")
        self.tree_menu.add_command(label="✏️ Editar Modelo", command=self.edit_selected_item)
        self.tree_menu.add_command(label="❌ Eliminar Registro", command=self.delete_selected_item)
        self.tree.bind("<Button-3>", self.show_context_menu)

        self.refresh_excel_table()

    def show_context_menu(self, event):
        try:
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self.tree_menu.post(event.x_root, event.y_root)
        except: pass

    def edit_selected_item(self):
        selected_item = self.tree.selection()
        if not selected_item: return
        item_values = self.tree.item(selected_item[0], 'values')
        if not item_values: return
        
        imei_val, current_model = item_values[0], item_values[1]
        dialog = customtkinter.CTkInputDialog(text=f"Editar Modelo para IMEI {imei_val}:", title="Editar Modelo")
        new_model = dialog.get_input()
        
        if new_model is not None:
            new_model = new_model.strip()
            if new_model:
                self.excel_manager.actualizar_registro(imei_val, new_model)
                self.refresh_excel_table()
                if self.imei_var.get() == imei_val:
                    self.modelo_var.set(new_model)

    def delete_selected_item(self):
        selected_item = self.tree.selection()
        if not selected_item: return
        item_values = self.tree.item(selected_item[0], 'values')
        if not item_values: return
        
        imei_val = item_values[0]
        if messagebox.askyesno("Confirmar Eliminación", f"¿Eliminar registro?\nIMEI: {imei_val}"):
            if self.excel_manager.eliminar_registro(imei_val):
                self.refresh_excel_table()

    def nuevo_proyecto(self):
        filepath = filedialog.asksaveasfilename(title="Crear Nuevo Proyecto Excel", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialdir=script_dir)
        if filepath:
            try:
                df = pd.DataFrame(columns=["IMEI", "Modelo"])
                df.to_excel(filepath, index=False)
                self._cambiar_archivo_excel(filepath)
                messagebox.showinfo("Proyecto Nuevo", f"Creado: {os.path.basename(filepath)}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear:\n{e}")

    def cargar_proyecto(self):
        filepath = filedialog.askopenfilename(title="Cargar Proyecto Excel", filetypes=[("Excel Files", "*.xlsx")], initialdir=script_dir)
        if filepath: self._cambiar_archivo_excel(filepath)

    def exportar_proyecto(self):
        if not os.path.exists(self.excel_manager.filepath): return
        import shutil
        dest_path = filedialog.asksaveasfilename(title="Exportar Copia Como", defaultextension=".xlsx", initialfile=f"Copia_{os.path.basename(self.excel_manager.filepath)}", filetypes=[("Excel Files", "*.xlsx")])
        if dest_path:
            try:
                shutil.copy2(self.excel_manager.filepath, dest_path)
                messagebox.showinfo("Exportado", f"Copia guardada en:\n{dest_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Falló la exportación:\n{e}")

    def _cambiar_archivo_excel(self, filepath):
        self.excel_manager.set_filepath(filepath)
        guardar_excel_config(filepath)
        self.lbl_filename.configure(text=os.path.basename(filepath))
        self.refresh_excel_table()
    
    def refresh_excel_table(self, filter_query=""):
        for item in self.tree.get_children(): self.tree.delete(item)
        registros = self.excel_manager.obtener_registros()
        query = filter_query.lower().strip()
        
        visible_count = 0
        for reg in registros:
            val_imei = str(reg.get("IMEI", "")).replace("nan", "")
            val_mod = str(reg.get("Modelo", "")).replace("nan", "")
            if query and (query not in val_imei.lower() and query not in val_mod.lower()): continue
            row_tag = "evenrow" if visible_count % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=(val_imei, val_mod), tags=(row_tag,))
            visible_count += 1
            
        if self.count_label:
            if query: self.count_label.configure(text=f"Filtrados: {visible_count} de {len(registros)}", text_color="#06B6D4")
            else: self.count_label.configure(text=f"Registros: {len(registros)}", text_color="#10B981")

    def on_search_change(self, *args):
        self.refresh_excel_table(self.search_var.get())

    def _bind_events(self):
        for var in [self.modelo_var, self.imei_var, self.logo_path_var]:
            var.trace_add("write", self.schedule_preview_update)
        self.modelo_var.trace_add("write", self.schedule_excel_update)

    def schedule_preview_update(self, *args):
        if self._preview_update_job: self.after_cancel(self._preview_update_job)
        self._preview_update_job = self.after(500, self.force_preview_update)

    def schedule_excel_update(self, *args):
        if self._excel_update_job: self.after_cancel(self._excel_update_job)
        self._excel_update_job = self.after(1500, self.perform_excel_update)

    def perform_excel_update(self):
        imei = self.imei_var.get().strip()
        nuevo_modelo = self.modelo_var.get().strip()
        if imei and nuevo_modelo:
            self.excel_manager.actualizar_registro(imei, nuevo_modelo)
            self.refresh_excel_table()

    def force_preview_update(self):
        try:
            texto_modelo_completo = self.modelo_var.get().strip().upper()
            modelo_limpio = self._limpiar_modelo_para_impresion(texto_modelo_completo)
            
            pil_image = _generar_etiqueta_pil_image(
                modelo_limpio,
                self.imei_var.get().strip().upper(),
                "", 
                self.logo_path_var.get().strip()
            )
            
            w_new = min(pil_image.width, PREVIEW_MAX_WIDTH)
            h_new = int(pil_image.height * (w_new / pil_image.width))
            img_resized = pil_image.resize((w_new, h_new))
            
            self.preview_ctk_image = customtkinter.CTkImage(light_image=img_resized, dark_image=img_resized, size=(w_new, h_new))
            self.preview_image_label.configure(image=self.preview_ctk_image, text="")
        except Exception as e:
            print(f"Error preview: {e}")

    def _limpiar_modelo_para_impresion(self, texto):
        if not texto: return ""
        texto_limpio = texto
        for color in sorted(COLORES_PARA_REMOVER, key=len, reverse=True):
            pattern = re.compile(re.escape(color), re.IGNORECASE)
            texto_limpio = pattern.sub("", texto_limpio)
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip()
        return texto_limpio

    def generar_y_guardar_pdf(self): self._procesar_generacion(guardar_permanente=True)
    def imprimir(self): self._procesar_generacion(imprimir_despues=True)

    def _procesar_generacion(self, guardar_permanente=False, imprimir_despues=False):
        if not PDF_SAVE_ENABLED:
            messagebox.showerror("Función Deshabilitada", "ReportLab no está instalado.")
            return
            
        texto_modelo_completo = self.modelo_var.get().strip().upper()
        modelo_limpio = self._limpiar_modelo_para_impresion(texto_modelo_completo)
        imei = self.imei_var.get().strip().upper()
        
        temp_pdf_path = _generar_etiqueta_pdf_temporal(modelo_limpio, imei, "", self.logo_path_var.get().strip())
        if not temp_pdf_path or not os.path.exists(temp_pdf_path):
            messagebox.showerror("Error", "No se pudo crear el PDF.")
            return

        if guardar_permanente:
            dest_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile=f"Etiqueta_{imei}.pdf" if imei else "Etiqueta.pdf")
            if dest_path:
                try:
                    import shutil
                    shutil.copy2(temp_pdf_path, dest_path)
                    messagebox.showinfo("Guardado", "Etiqueta guardada exitosamente.")
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

        if imprimir_despues:
            self._enviar_a_impresora(temp_pdf_path)
            if self.auto_shutdown_var.get():
                self.after(2000, self.intentar_apagar_dispositivo)

    def _enviar_a_impresora(self, pdf_path):
        if not os.path.exists(pdf_path): return
        current_os = platform.system()
        try:
            if current_os == "Windows":
                if SUMATRA_PDF_PATH and os.path.exists(SUMATRA_PDF_PATH):
                    subprocess.Popen([SUMATRA_PDF_PATH, "-print-to-default", "-silent", pdf_path])
                else:
                    os.startfile(pdf_path, "print")
            elif current_os in ["Darwin", "Linux"]:
                cmd = "lpr" if current_os == "Darwin" else "lp"
                subprocess.run([cmd, pdf_path], check=True)
        except Exception as e:
            messagebox.showerror("Error Impresión", f"No se pudo imprimir:\n{e}")

    def intentar_apagar_dispositivo(self):
        if not self.current_udid: return
        try:
            if self.current_device_info and self.current_device_info.get('is_android'):
                if ADBUTILS_AVAILABLE:
                    dev = adbutils.adb.device(self.current_udid)
                    dev.shell("reboot -p")
            else:
                lockdown = create_using_usbmux(serial=self.current_udid)
                diag = DiagnosticsService(lockdown=lockdown)
                diag.shutdown()
        except Exception as e:
            print(f"Error al apagar: {e}")

    def pegar_modelo(self):
        try:
            contenido = self.clipboard_get()
            if contenido:
                texto_limpio = re.sub(r'\s+', ' ', contenido.strip())
                self.modelo_entry.delete(0, tk.END)
                self.modelo_entry.insert(0, texto_limpio)
        except tk.TclError: pass

    def pegar_imei(self):
        try:
            contenido = self.clipboard_get()
            if contenido:
                self.imei_entry.delete(0, tk.END)
                self.imei_entry.insert(0, contenido.strip())
        except tk.TclError: pass

    def buscar_logo(self):
        filepath = filedialog.askopenfilename(title="Seleccionar Logo", filetypes=[("Imágenes", "*.png *.jpg *.jpeg"), ("Todos", "*.*")])
        if filepath:
            self.logo_path_var.set(filepath)
            guardar_logo_config(filepath)

    def configurar_ruta_sumatra_manualmente(self):
        if platform.system() != "Windows": return
        filepath = filedialog.askopenfilename(title="Localizar SumatraPDF.exe", filetypes=[("Ejecutable", "SumatraPDF.exe")])
        if filepath and os.path.basename(filepath).lower() == 'sumatrapdf.exe':
            global SUMATRA_PDF_PATH
            SUMATRA_PDF_PATH = filepath
            guardar_config_sumatra()
            messagebox.showinfo("Éxito", f"SumatraPDF establecido:\n{filepath}")

    def actualizar_estado(self, estado, brand=""):
        if estado == "conectado":
            self.badge_frame.configure(fg_color="#065F46", border_color="#047857")
            self.badge_label.configure(text=f"Conectado: {brand} ✅", text_color="#A7F3D0")
        elif estado == "confiar":
            self.badge_frame.configure(fg_color="#78350F", border_color="#B45309")
            self.badge_label.configure(text="¡Falta Confiar / Depuración! ⚠️", text_color="#FDE68A")
        else:
            self.badge_frame.configure(fg_color="#334155", border_color="#475569")
            self.badge_label.configure(text="Estado: Desconectado ❌", text_color="#CBD5E1")

    def on_device_info_received(self, info):
        self.current_device_info = info
        self.after(0, self.hide_trust_alert)
        brand = info.get('brand', 'Dispositivo')
        self.after(0, lambda: self.actualizar_estado("conectado", brand))
        self.after(0, lambda: self._update_ui_from_device(info))

    def on_device_disconnected(self):
        self.current_device_info = None
        self.after(0, self.hide_trust_alert)
        self.after(0, lambda: self.actualizar_estado("desconectado"))

    def on_device_trust_needed(self, os_type="iOS"):
        self.after(0, lambda: self.actualizar_estado("confiar"))
        self.after(0, lambda: self.show_trust_alert(os_type))

    def show_trust_alert(self, os_type="iOS"):
        if self.trust_window is None or not self.trust_window.winfo_exists():
            self.trust_window = customtkinter.CTkToplevel(self)
            self.trust_window.title("PERMISO REQUERIDO")
            self.trust_window.geometry("500x250")
            self.trust_window.attributes("-topmost", True)
            
            customtkinter.CTkLabel(self.trust_window, text="⚠️ ACCIÓN REQUERIDA EN TELÉFONO ⚠️", font=("Arial", 16, "bold"), text_color="yellow").pack(pady=(20, 10))
            if os_type == "iOS":
                msg = "El iPhone detectado no ha confiado en este equipo.\n\n1. Desbloquea el iPhone.\n2. Presiona 'Confiar' en la pantalla del iPhone."
            else:
                msg = "El teléfono Android requiere autorización ADB.\n\n1. Desbloquea la pantalla.\n2. Acepta 'Permitir depuración por USB'."
            customtkinter.CTkLabel(self.trust_window, text=msg, font=("Arial", 13), justify="center").pack(pady=10)

    def hide_trust_alert(self):
        if self.trust_window and self.trust_window.winfo_exists():
            self.trust_window.destroy()
        self.trust_window = None

    def _update_ui_from_device(self, info):
        model_name = info.get('product_name', 'Dispositivo')
        imei = info.get('imei') or info.get('serial_number') or ""
        capacidad = info.get('capacity', '')
        color = info.get('color', '')
        
        self.current_udid = info.get('udid')
        full_model_text = f"{model_name} {color} {capacidad}".strip()
        while "  " in full_model_text: full_model_text = full_model_text.replace("  ", " ")

        self.modelo_var.set(full_model_text)
        self.imei_var.set(imei)
        
        if imei:
            exito, msg, count = self.excel_manager.registrar_dispositivo(imei, full_model_text)
            if exito:
                if count >= 0: self.count_label.configure(text=f"Registrados: {count}", text_color="#10B981")
                self.after(0, self.refresh_excel_table)
            else:
                if "abierto" in msg.lower(): messagebox.showwarning("Error Excel", msg)
        
        if self.auto_print_var.get():
            self.after(500, self.imprimir)

if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")
    
    app = OmniTagMobileApp()
    app.mainloop()
