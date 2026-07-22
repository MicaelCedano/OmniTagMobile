# -*- coding: utf-8 -*-
"""
Generador de Etiquetas para iPhone
Versión: 3.2.5 (Personalizado)
Autor: Micael (Modificado por Asistente de IA)
Fecha: 2025-08-15

Cambios en v3.2.5:
- Adaptado para iPhones: S/N cambiado a IMEI
- Eliminadas las plantillas
- Eliminadas las especificaciones
- Eliminada la condición final que impedía el dibujado del código de barras en el PDF.
- Reajustados los márgenes y espaciados para garantizar que todos los elementos quepan.
- Se fuerza el dibujado secuencial y completo de todos los elementos.
"""
from PIL import Image, ImageDraw, ImageFont, ImageTk
import barcode
from barcode.writer import ImageWriter
import io
import customtkinter
import tkinter as tk
from tkinter import ttk  # Importar ttk para Treeview
from tkinter import filedialog, messagebox
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
from openpyxl import load_workbook

import warnings
# Suprimir warnings de criptografía que ensucian la consola
warnings.filterwarnings("ignore", category=UserWarning, module='pymobiledevice3')

# --- Dependencias para Detección de Dispositivos (iPhone) ---
PYMOBILEDEVICE_AVAILABLE = False
try:
    from pymobiledevice3.usbmux import list_devices
    # CORRECCIÓN: Usar create_using_usbmux para instanciar el cliente correctamente
    from pymobiledevice3.lockdown import create_using_usbmux
    from pymobiledevice3.services.diagnostics import DiagnosticsService
    from pymobiledevice3.exceptions import NoDeviceConnectedError
    PYMOBILEDEVICE_AVAILABLE = True
except ImportError as e:
    print(f"ADVERTENCIA: Error al importar 'pymobiledevice3': {e}")
    # Fallback o reintento si la estructura es diferente
    try:
        from pymobiledevice3.services.lockdown import create_using_usbmux
        PYMOBILEDEVICE_AVAILABLE = True
    except ImportError:
        print("La detección automática no funcionará.")
except Exception as e:
    print(f"ADVERTENCIA: Error inesperado al importar dependencias: {e}")


# --- Dependencias para Guardado/Impresión en PDF ---
PDF_SAVE_ENABLED = False
try:
    from reportlab.pdfgen import canvas as reportlab_canvas
    from reportlab.lib.pagesizes import inch
    from reportlab.lib.utils import ImageReader as ReportLabImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_SAVE_ENABLED = True
except ImportError:
    print("ADVERTENCIA: La librería 'ReportLab' no está instalada. El guardado en PDF y la impresión estarán deshabilitados.")
    PDF_SAVE_ENABLED = False

# --- Constantes ---
# Obtener la ruta absoluta del directorio del script
script_dir = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE_NAME = os.path.join(script_dir, "etiqueta_config.json")
# Nombre del archivo Excel (basado en la plantilla del usuario)
EXCEL_FILE_NAME = os.path.join(script_dir, "plantilla_compra_iphone (7).xlsx")

LABEL_WIDTH_INCHES = 4
LABEL_HEIGHT_INCHES = 3
PREVIEW_MAX_WIDTH = 380
PREVIEW_MAX_HEIGHT = int(PREVIEW_MAX_WIDTH * (LABEL_HEIGHT_INCHES / LABEL_WIDTH_INCHES))

# --- Rutas de Fuentes (Asegúrate de que estos archivos .ttf estén en la misma carpeta) ---
FONT_BOLD_PATH_TTF = "arialbd.ttf"
FONT_REGULAR_PATH_TTF = "arial.ttf"

# --- Nombres de Fuentes para ReportLab (PDF) ---
RL_FONT_BOLD_NAME = "ArialBoldRegistered"
RL_FONT_REGULAR_NAME = "ArialRegularRegistered"


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

# Fallback para colores numéricos (Adivinanza basada en comunes)
INT_COLOR_MAP = {
    "1": "Black", "2": "White", "3": "Gold", "4": "Rose Gold", 
    "5": "Jet Black", "6": "Red", "7": "Silver"
}

# Colores conocidos para eliminar de la etiqueta impresa (se mantienen en Excel)
COLORES_PARA_REMOVER = [
    "BLACK", "WHITE", "SILVER", "GOLD", "ROSE GOLD", "SPACE GRAY", "SPACE BLACK",
    "GRAPHITE", "MIDNIGHT", "STARLIGHT", "RED", "PRODUCT RED", "BLUE", "SIERRA BLUE",
    "PACIFIC BLUE", "PINK", "GREEN", "MIDNIGHT GREEN", "YELLOW", "PURPLE",
    "DEEP PURPLE", "TITANIUM", "NATURAL TITANIUM", "BLUE TITANIUM", "WHITE TITANIUM",
    "BLACK TITANIUM", "GRAY", "CORAL", "TEAL", "ULTRAMARINE", "DESERT TITANIUM",
    # EspaÃ±ol
    "NEGRO", "BLANCO", "PLATA", "PLATEADO", "DORADO", "ORO", "ROJO", "AZUL",
    "VERDE", "AMARILLO", "ROSA", "PURPURA", "MORADO", "GRIS", "GRIS ESPACIAL"
]

# Mapeo aproximado de colores (Hex -> Nombre)
# Estos son valores comunes devueltos por DeviceColor
HASH_COLOR_MAP = {
    # Generic / Common
    "#3b3b3b": "Space Gray", "#000000": "Black", "#ffffff": "White",
    "#ff3b30": "Product Red", "#e1e4e3": "Silver", "#f9e5c9": "Gold",
    "#d8a1bc": "Rose Gold", "#121212": "Graphite", "#28344e": "Midnight",
    "#faf7f2": "Starlight", "#a0b4d6": "Sierra Blue", "#574f6f": "Deep Purple",
    "#464c48": "Midnight Green", "#1d4c7b": "Pacific Blue", "#f2f0eb": "Silver",
    "#2e3c4e": "Pacific Blue", "#1f4663": "Pacific Blue", "#003f5d": "Pacific Blue", "#36495d": "Pacific Blue",
    "#e2e4e1": "Silver", "#fad7bd": "Gold", "#afe3b2": "Green", "#ffe681": "Yellow",
    "#fec2dc": "Pink", "#b8afe6": "Purple", "#1e1e1e": "Space Black",
    "#4a4a4c": "Space Gray",

    # iPhone 16 Series
    "#c2bcb2": "Natural Titanium", "#bfa48f": "Desert Titanium", "#3c3c3d": "Black Titanium", "#f2f1ed": "White Titanium",
    "#f2adda": "Pink", "#9aadf6": "Ultramarine", "#b0d4d2": "Teal", "#3c4042": "Black", "#fafafa": "White",

    # iPhone 15 Series
    "#2f4452": "Blue Titanium", "#837f7d": "Natural Titanium", "#1b1b1b": "Black Titanium", "#dddddd": "White Titanium",
    "#e3c8ca": "Pink", "#ced5d9": "Blue", "#cad4c5": "Green", "#e5e0c1": "Yellow", "#35393b": "Black",

    # iPhone 14 Series
    "#594f63": "Deep Purple", "#403e3d": "Space Black", "#f0f2f2": "Silver", "#f4e8ce": "Gold",
    "#e6ddeb": "Purple", "#a0b4c7": "Blue", "#f9e479": "Yellow", "#222930": "Midnight", 
    "#faf6f2": "Starlight", "#fc0324": "Product Red", "#5e5566": "Deep Purple",

    # iPhone 13 Series
    "#a7c1d9": "Sierra Blue", "#576856": "Alpine Green", "#54524f": "Graphite", "#f1f2ed": "Silver", "#fae7cf": "Gold",
    "#276787": "Blue", "#1c5c78": "Blue", "#376e8a": "Blue", "#394c38": "Green", "#faddd7": "Pink", "#232a31": "Midnight", "#bf0013": "Product Red",
    "#95aec5": "Sierra Blue", "#9bb5ce": "Sierra Blue",

    # iPhone 12 Series
    "#201d24": "Black", "#043458": "Blue", "#4c4a46": "Graphite", "#feedd8": "Gold", "#2e4755": "Pacific Blue",
    "#e1f8dc": "Green", "#e23636": "Red", "#fbf7f4": "White",

    # iPhone 11 Series
    "#4e5851": "Midnight Green", "#4e5850": "Midnight Green", "#003e30": "Midnight Green", "#1f352e": "Midnight Green", "#535150": "Space Gray", "#ebebe3": "Silver",

    # iPhone 8 / SE
    "#272729": "Space Gray", "#e2e3e4": "Silver", "#f7e8dd": "Gold", "#960111": "Product Red",
}


# --- Variables Globales ---
SUMATRA_PDF_PATH = None
temporary_files_to_delete = []

# --- Funciones de Configuración y Limpieza ---

def _read_config():
    """Lee el archivo de configuración JSON de forma segura."""
    if os.path.exists(CONFIG_FILE_NAME):
        with open(CONFIG_FILE_NAME, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return {}
    return {}

def _write_config(config_data):
    """Escribe en el archivo de configuración JSON."""
    try:
        with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
    except Exception as e:
        print(f"Error al guardar configuración: {e}")

def cargar_config_inicial():
    """Carga la configuración de SumatraPDF al inicio."""
    global SUMATRA_PDF_PATH
    config = _read_config()
    path_guardado = config.get("sumatra_pdf_path")
    if path_guardado and os.path.exists(path_guardado) and os.path.isfile(path_guardado):
        SUMATRA_PDF_PATH = path_guardado
        print(f"SumatraPDF cargado desde config: {SUMATRA_PDF_PATH}")
    else:
        detectar_sumatra_si_no_configurado()

def cargar_logo_config():
    """Carga la ruta del logo guardada en la configuración."""
    config = _read_config()
    logo_path = config.get("logo_path")
    if logo_path and os.path.exists(logo_path) and os.path.isfile(logo_path):
        return logo_path
    return "logo.png"  # Valor por defecto

def guardar_logo_config(logo_path):
    """Guarda la ruta del logo en el archivo de configuración."""
    if logo_path:
        config = _read_config()
        config["logo_path"] = logo_path
        _write_config(config)

def cargar_excel_config():
    """Carga la ruta del último Excel abierto."""
    config = _read_config()
    path = config.get("last_excel_path")
    if path and os.path.exists(path) and os.path.isfile(path):
        return path
    return EXCEL_FILE_NAME # Default

def guardar_excel_config(excel_path):
    """Guarda la ruta del Excel en uso."""
    if excel_path:
        config = _read_config()
        config["last_excel_path"] = excel_path
        _write_config(config)

def guardar_config_sumatra():
    """Guarda solo la ruta de SumatraPDF en el archivo de configuración."""
    if SUMATRA_PDF_PATH and platform.system() == "Windows":
        config = _read_config()
        config["sumatra_pdf_path"] = SUMATRA_PDF_PATH
        _write_config(config)

def detectar_sumatra_si_no_configurado():
    """Intenta encontrar SumatraPDF en rutas comunes si no está configurado."""
    global SUMATRA_PDF_PATH
    if SUMATRA_PDF_PATH or platform.system() != "Windows": return
    SUMATRA_PDF_DEFAULT_PATHS = [
        "SumatraPDF.exe",
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
    ]
    for path_candidate in SUMATRA_PDF_DEFAULT_PATHS:
        try:
            result = subprocess.run(["where", os.path.basename(path_candidate)], capture_output=True, text=True, check=False, shell=True)
            if result.returncode == 0 and result.stdout.strip():
                SUMATRA_PDF_PATH = result.stdout.strip().splitlines()[0]
                print(f"SumatraPDF detectado en el PATH: {SUMATRA_PDF_PATH}")
                return
            elif os.path.exists(path_candidate) and os.path.isfile(path_candidate):
                SUMATRA_PDF_PATH = path_candidate
                print(f"SumatraPDF detectado en: {SUMATRA_PDF_PATH}")
                return
        except Exception: pass
    print("ADVERTENCIA: SumatraPDF no se detectó automáticamente.")

def cleanup_temp_files():
    """Elimina los archivos temporales creados durante la sesión."""
    for temp_file_path in list(temporary_files_to_delete):
        try:
            if os.path.exists(temp_file_path): os.remove(temp_file_path)
            if temp_file_path in temporary_files_to_delete:
                temporary_files_to_delete.remove(temp_file_path)
        except Exception as e:
            print(f"Error al limpiar archivo temporal {temp_file_path}: {e}")

atexit.register(cleanup_temp_files)
atexit.register(guardar_config_sumatra)

def cargar_fuentes_pdf():
    """Registra las fuentes TTF en ReportLab, usando Helvetica como fallback seguro."""
    global RL_FONT_BOLD_NAME, RL_FONT_REGULAR_NAME
    if not PDF_SAVE_ENABLED: return
    try:
        if os.path.exists(FONT_BOLD_PATH_TTF):
            pdfmetrics.registerFont(TTFont(RL_FONT_BOLD_NAME, FONT_BOLD_PATH_TTF))
        else: raise IOError(f"No se encontró '{FONT_BOLD_PATH_TTF}'.")
    except Exception:
        RL_FONT_BOLD_NAME = 'Helvetica-Bold'
    try:
        if os.path.exists(FONT_REGULAR_PATH_TTF):
            pdfmetrics.registerFont(TTFont(RL_FONT_REGULAR_NAME, FONT_REGULAR_PATH_TTF))
        else: raise IOError(f"No se encontró '{FONT_REGULAR_PATH_TTF}'.")
    except Exception:
        RL_FONT_REGULAR_NAME = 'Helvetica'

# --- FUNCIÓN DE PREVISUALIZACIÓN ---
def _generar_etiqueta_pil_image(modelo, numero_serie, especificacion, path_logo_pil):
    """Genera la etiqueta como una imagen PIL, replicando la lógica del PDF."""
    DPI = 300
    LABEL_WIDTH_PX, LABEL_HEIGHT_PX = int(LABEL_WIDTH_INCHES * DPI), int(LABEL_HEIGHT_INCHES * DPI)
    
    TOP_MARGIN_PX = int(0.20 * DPI)
    SIDE_MARGIN_PX = int(0.15 * DPI)
    BOTTOM_MARGIN_PX = int(0.20 * DPI)
    
    image = Image.new("RGB", (LABEL_WIDTH_PX, LABEL_HEIGHT_PX), "white")
    draw = ImageDraw.Draw(image)
    try:
        font_bold = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=int(12 * DPI / 72))
        font_regular = ImageFont.truetype(FONT_REGULAR_PATH_TTF, size=int(10 * DPI / 72))
    except IOError:
        font_bold, font_regular = ImageFont.load_default(), ImageFont.load_default()
    
    current_y = TOP_MARGIN_PX
    
    # 1. Logo
    if path_logo_pil and os.path.exists(path_logo_pil):
        try:
            with Image.open(path_logo_pil) as logo_img:
                logo_img = logo_img.convert("RGBA")
                logo_max_width = LABEL_WIDTH_PX - 2 * SIDE_MARGIN_PX
                logo_max_height = int(0.28 * LABEL_HEIGHT_PX)
                logo_img.thumbnail((logo_max_width, logo_max_height), Image.Resampling.LANCZOS)
                
                logo_x = (LABEL_WIDTH_PX - logo_img.width) // 2
                image.paste(logo_img, (logo_x, current_y), logo_img)
                current_y += logo_img.height + int(0.1 * DPI)
        except Exception as e:
            print(f"Error procesando logo: {e}")

            current_y += logo_img.height + int(0.1 * DPI)
        except Exception as e:
            print(f"Error procesando logo: {e}")

    # 2. Texto
    # Construir línea de Modelo + Capacidad + Color
    # Ej: "iPhone 11 - 64GB Black" o en líneas separadas si es muy largo
    
    # Vamos a intentar ponerlo en una o dos líneas
    texto_modelo = f"Modelo: {modelo}"
    if especificacion: # Usaremos este campo para "Capacidad + Color"
        texto_detalles = especificacion
    else:
        texto_detalles = ""

    info_items = []
    info_items.append((texto_modelo, font_bold, modelo))
    if texto_detalles:
        info_items.append((texto_detalles, font_regular, texto_detalles))
    
    info_items.append((f"IMEI: {numero_serie}", font_bold, numero_serie))
    
    for texto, font, valor in info_items:
        if not valor.strip(): continue
        x_pos = (LABEL_WIDTH_PX - draw.textlength(texto, font=font)) // 2
        draw.text((x_pos, current_y), texto, fill="black", font=font)
        current_y += font.size + 8 # Un poco más de espaciado

    # 3. Código de Barras
    if numero_serie:
        try:
            padding_before_bc = int(0.1 * DPI)
            current_y += padding_before_bc
            
            # Configurar parámetros optimizados para lectura
            # quiet_zone mínimo recomendado: 6.5 módulos para Code128
            # module_height: altura suficiente para escaneo confiable
            barcode_options = {
                'module_height': 15.0,  # Altura adecuada para lectura
                'module_width': 0.3,    # Ancho de módulo optimizado  
                'quiet_zone': 6.5,       # Zona silenciosa amplia (mínimo recomendado: 6.5)
                'write_text': False,    # El texto se dibuja por separado
                'text_distance': 5.0,
                'font_size': 10
            }
            
            code128 = barcode.get_barcode_class('code128')
            writer = ImageWriter()
            barcode_obj = code128(numero_serie, writer=writer)
            
            # Usar render() que devuelve directamente una imagen PIL
            barcode_pil = barcode_obj.render(barcode_options)
            barcode_pil = barcode_pil.convert('RGB')
            
            # Ajustar tamaño manteniendo proporción, pero sin degradar demasiado
            max_bc_w = LABEL_WIDTH_PX - 2 * SIDE_MARGIN_PX
            if barcode_pil.width > max_bc_w:
                ratio = max_bc_w / barcode_pil.width
                new_width = int(barcode_pil.width * ratio)
                new_height = int(barcode_pil.height * ratio)
                # Usar LANCZOS para mejor calidad en redimensionamiento
                barcode_pil = barcode_pil.resize((new_width, new_height), Image.Resampling.LANCZOS)

            sn_font = ImageFont.truetype(FONT_REGULAR_PATH_TTF, size=int(9 * DPI / 72))
            sn_text_w = draw.textlength(numero_serie, font=sn_font)
            
            bc_x = (LABEL_WIDTH_PX - barcode_pil.width) // 2
            image.paste(barcode_pil, (bc_x, current_y))
            current_y += barcode_pil.height + int(0.03 * DPI)

            sn_x = (LABEL_WIDTH_PX - sn_text_w) // 2
            draw.text((sn_x, current_y), numero_serie, fill="black", font=sn_font)
        except Exception as e:
            print(f"Error generando código de barras en previsualización: {e}")
            
    return image

# --- FUNCIÓN DE GENERACIÓN DE PDF (REVISADA Y SECUENCIAL) ---
def _generar_etiqueta_pdf_temporal(modelo, numero_serie, especificacion, path_logo_pil):
    if not PDF_SAVE_ENABLED: return None
    fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="etiqueta_")
    os.close(fd)
    temporary_files_to_delete.append(temp_pdf_path)
    
    c = reportlab_canvas.Canvas(temp_pdf_path, pagesize=(LABEL_WIDTH_INCHES * inch, LABEL_HEIGHT_INCHES * inch))
    width, height = LABEL_WIDTH_INCHES * inch, LABEL_HEIGHT_INCHES * inch
    
    # Márgenes
    margin_top = 0.20 * inch
    margin_sides = 0.15 * inch
    
    # Coordenada Y, se mueve de arriba hacia abajo
    current_y = height - margin_top

    # 1. Dibujar Logo
    try:
        if path_logo_pil and os.path.exists(path_logo_pil):
            logo_pil = Image.open(path_logo_pil)
            logo_max_width_pt = width - 2 * margin_sides
            logo_max_height_pt = 0.28 * height
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
            current_y -= 0.1 * inch
    except Exception as e:
        print(f"Error al procesar logo para PDF: {e}")

    # 2. Dibujar Texto
    # Lógica similar a PIL para incluir detalles
    info_items_pdf = []
    
    # Modelo
    info_items_pdf.append((f"Modelo: {modelo}", RL_FONT_BOLD_NAME, 12, modelo))
    
    # Capacidad y Color (pasado en 'especificacion')
    if especificacion:
        info_items_pdf.append((especificacion, RL_FONT_REGULAR_NAME, 11, especificacion))
        
    # IMEI
    info_items_pdf.append((f"IMEI: {numero_serie}", RL_FONT_BOLD_NAME, 12, numero_serie))
    
    for texto, font, size, valor in info_items_pdf:
        if not valor.strip(): continue
        current_y -= size
        c.setFont(font, size)
        c.drawCentredString(width / 2, current_y, texto)
        current_y -= 6

    # 3. Dibujar Código de Barras (Sin 'if' de espacio)
    if numero_serie:
        try:
            current_y -= 0.1 * inch
            
            # Configurar parámetros optimizados para lectura
            # quiet_zone mínimo recomendado: 6.5 módulos para Code128
            barcode_options = {
                'module_height': 15.0,  # Altura adecuada para lectura
                'module_width': 0.3,    # Ancho de módulo optimizado
                'quiet_zone': 6.5,      # Zona silenciosa amplia (mínimo recomendado: 6.5)
                'write_text': False,    # El texto se dibuja por separado
                'text_distance': 5.0,
                'font_size': 10
            }
            
            code128 = barcode.get_barcode_class('code128')
            writer = ImageWriter()
            barcode_obj = code128(numero_serie, writer=writer)
            
            # Generar imagen PIL y guardar en buffer para PDF
            barcode_pil_pdf = barcode_obj.render(barcode_options)
            barcode_pil_pdf = barcode_pil_pdf.convert('RGB')
            
            barcode_io = io.BytesIO()
            barcode_pil_pdf.save(barcode_io, format='PNG')
            barcode_io.seek(0)
            
            img_reader = ReportLabImageReader(barcode_io)
            bc_w, bc_h = img_reader.getSize()
            
            max_bc_w = width - (2 * margin_sides)
            if bc_w > max_bc_w:
                ratio = max_bc_w / bc_w
                bc_w, bc_h = bc_w * ratio, bc_h * ratio
            
            # Dibujar barcode
            current_y -= bc_h
            c.drawImage(img_reader, (width - bc_w) / 2, current_y, width=bc_w, height=bc_h, mask='auto')
            
            # Dibujar texto IMEI
            sn_font_size = 9
            current_y -= 3
            current_y -= sn_font_size
            c.setFont(RL_FONT_REGULAR_NAME, sn_font_size)
            c.drawCentredString(width / 2, current_y, numero_serie)

        except Exception as e:
            print(f"Error generando código de barras para PDF: {e}")
            
    c.save()
    return temp_pdf_path


# --- MANEJADOR DE EXCEL ---
class ExcelManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.ensure_file_exists()

    def set_filepath(self, new_path):
        """Cambia el archivo de trabajo actual."""
        self.filepath = new_path
        self.ensure_file_exists()

    def ensure_file_exists(self):
        """Si el archivo no existe, lo crea con las cabeceras."""
        if not os.path.exists(self.filepath):
            try:
                df = pd.DataFrame(columns=["IMEI", "Modelo"])
                df.to_excel(self.filepath, index=False)
            except Exception as e:
                print(f"Error creando Excel: {e}")

    def registrar_dispositivo(self, imei, modelo_completo):
        """Registra el dispositivo. 'modelo_completo' incluye nombre, gb y color."""
        try:
            # Leer archivo existente
            try:
                df = pd.read_excel(self.filepath)
            except FileNotFoundError:
                df = pd.DataFrame(columns=["IMEI", "Modelo"])
            
            # Verificar duplicados
            df['IMEI'] = df['IMEI'].astype(str)
            if str(imei) in df['IMEI'].values:
                return False, f"El IMEI {imei} ya está registrado.", len(df)

            # Agregar nuevo registro (2 columnas)
            new_row = pd.DataFrame([{
                "IMEI": str(imei), 
                "Modelo": modelo_completo
            }])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # Guardar
            df.to_excel(self.filepath, index=False)
            return True, "Registrado en Excel.", len(df)
            
        except PermissionError:
            return False, "Error: El archivo Excel está abierto. Ciérralo para guardar.", -1
        except Exception as e:
            return False, f"Error Excel: {e}", -1

    def actualizar_registro(self, imei, nuevo_modelo):
        """Actualiza el modelo de un IMEI existente."""
        try:
            df = pd.read_excel(self.filepath)
            df['IMEI'] = df['IMEI'].astype(str)
            imei_str = str(imei)
            
            if imei_str in df['IMEI'].values:
                # Localizar índice
                idx = df.index[df['IMEI'] == imei_str].tolist()[0]
                # Actualizar
                df.at[idx, 'Modelo'] = nuevo_modelo
                df.to_excel(self.filepath, index=False)
                return True
            return False
        except Exception as e:
            print(f"Error actualizando Excel: {e}")
            return False
    def eliminar_registro(self, imei):
        """Elimina un registro dado su IMEI."""
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
                df = pd.read_excel(self.filepath)
                return len(df)
        except:
            pass
        return 0

    def obtener_registros(self):
        """Retorna una lista de diccionarios con los datos, más reciente primero."""
        try:
            if os.path.exists(self.filepath):
                df = pd.read_excel(self.filepath)
                # Invertir orden para ver los últimos arriba
                return df.iloc[::-1].to_dict('records')
        except:
            pass
        return []


# --- MONITOR DE DISPOSITIVOS ---
class DeviceMonitor(threading.Thread):
    def __init__(self, success_callback, trust_callback, disconnected_callback):
        super().__init__()
        self.success_callback = success_callback
        self.trust_callback = trust_callback
        self.disconnected_callback = disconnected_callback
        self.running = True
        self.last_successful_udid = None
        self.daemon = True

    def run(self):
        print("--- DEBUG: Iniciando Monitor de Dispositivos ---")
        if not PYMOBILEDEVICE_AVAILABLE:
            print("--- DEBUG: PYMOBILEDEVICE_AVAILABLE es False ---")
            return
        
        print("--- DEBUG: Monitor corriendo. Esperando dispositivos... ---")
        while self.running:
            try:
                # list_devices devuelve una lista de objetos MuxDevice
                devices = list_devices()
                
                if not devices:
                    if self.last_successful_udid is not None:
                        self.last_successful_udid = None
                        self.disconnected_callback()
                
                for device in devices:
                    udid = device.serial
                    
                    # Si ya procesamos este dispositivo con éxito y sigue conectado, lo ignoramos
                    if udid == self.last_successful_udid:
                        continue
                    
                    # Nuevo dispositivo o reintento de uno fallido
                    print(f"Intentando conectar a: {udid}")
                    try:
                        # Intentar conectar usando create_using_usbmux
                        lockdown = create_using_usbmux(serial=udid)
                        
                        # Validar paridad pero ignorar error de SessionActive (que es benigno)
                        try:
                            lockdown.validate_pairing()
                        except Exception as e_pair:
                            # Si es SessionActive, significa que ya hay sesión (es bueno/confiado)
                            if "SessionActive" in str(e_pair):
                                pass 
                            else:
                                # Cualquier otro error (PasswordProtected, InvalidHostID) es falta de confianza
                                raise e_pair
                        
                        print("--- DEBUG: Confianza (Trust) verificada correctamente ---")

                        # Obtener informaciÃ³n bÃ¡sica
                        product_type = lockdown.get_value(key='ProductType') 
                        serial_number = lockdown.get_value(key='SerialNumber')
                        imei = lockdown.get_value(key='InternationalMobileEquipmentIdentity')
                        
                        # --- Determinar Nombre del Modelo (ANTES del color para usarlo en la lÃ³gica) ---
                        model_name = IPHONE_MODEL_MAPPING.get(product_type, "iPhone Desconocido")
                        
                        # Si no leemos IMEI, es muy probable que no haya paridad/confianza
                        if not imei:
                            raise Exception("IMEI vacío. Posible falta de confianza (Trust).")
                        
                        # Capacidad (Intentar varios métodos)
                        capacidad = ""
                        try:
                            # Método 1: Domain com.apple.disk_usage
                            bytes_cap = lockdown.get_value(domain='com.apple.disk_usage', key='TotalDiskCapacity')
                            if not bytes_cap:
                                # Método 2: Global (a veces funciona en antiguos)
                                bytes_cap = lockdown.get_value(key='TotalDiskCapacity')
                            
                            if bytes_cap:
                                gb = int(bytes_cap) / 1000000000
                                known_sizes = [16, 32, 64, 128, 256, 512, 1024]
                                closest = min(known_sizes, key=lambda x: abs(x - gb))
                                if closest == 1024:
                                    capacidad = "1TB"
                                else:
                                    capacidad = f"{closest}GB"
                        except Exception as e:
                            print(f"Error Capacity: {e}")

                        # Color
                        try:
                            # DeviceColor es un string hexadecimal ej "#3b3b3b"
                            color_hex = lockdown.get_value(key='DeviceColor')
                            print(f"--- DEBUG RAW COLOR: {color_hex}") # Debug para ver qué llega
                            
                            color_hex_lower = str(color_hex).lower().strip()
                            color = HASH_COLOR_MAP.get(color_hex_lower, color_hex)
                            
                            # Si es un número pequeño (ej "1", "2"), ignorarlo porque no sabemos qué color es
                            if color and len(color) < 2 and color.isdigit():
                                color = ""
                            
                            
                            if not color or color == "None" or len(color) < 2:
                                enc_color = lockdown.get_value(key='DeviceEnclosureColor')
                                if enc_color:
                                     # Convertir a str primero
                                     enc_val = str(enc_color).lower().strip()
                                     color = HASH_COLOR_MAP.get(enc_val, enc_color)
                                     
                                     if "iPhone 12 Pro" in model_name:
                                         if enc_val == "1":
                                             color = "Pacific Blue"
                                         elif enc_val == "2" or enc_val == "white":
                                              color = "Silver"
                                         elif enc_val == "3":
                                              color = "Gold"
                                     
                                     elif "iPhone 12" in model_name: # 12 y 12 mini
                                         if enc_val == "1":
                                             color = "Blue"

                                     # Fallbacks generales si aún no tenemos color
                                     if not color:
                                          if enc_val == "1": color = "Black"
                                          elif enc_val == "2": color = "White"

                            # Si no se detectó color válido, intentar domain iTunes
                            if not color:
                                try:
                                    # A veces está aquí como string
                                    itunes_color = lockdown.get_value(domain='com.apple.mobile.iTunes', key='DeviceColor')
                                    if itunes_color:
                                        color = str(itunes_color)
                                except: pass

                            # Limpiar si sigue siendo un número (ej "1", "2") 
                            # INTENTO: Mapear 1 y 2 a colores comunes si no hay nada más
                            if color and str(color).strip().isdigit():
                                digit = str(color).strip()
                                color = INT_COLOR_MAP.get(digit, "") # Si es 1->Black, 2->White, sino borra
                            
                            # Si no se detectó color válido, dejarlo vacío para que el usuario lo ponga
                            if not color: color = ""
                            
                            if color is None: color = ""
                        except Exception as e:
                            print(f"Error Color: {e}")
                            color = ""
                        

                        # Corrección: iPhone 13/mini Pink a veces se detecta como Gold, Rose Gold o Yellow
                        if "iPhone 13" in model_name and "Pro" not in model_name:
                            # Estos colores NO existen en el iPhone 13 base, así que si se detectan, es un error y probablemente sea Pink
                            if color in ["Gold", "Rose Gold", "Yellow", "Orange"]: 
                                color = "Pink"

                        # Corrección: iPhone 11 Pro (y Max) Silver a veces se detecta como White
                        if "iPhone 11 Pro" in model_name and color == "White":
                            color = "Silver"

                        print(f"--- DEBUG: Info -> {product_type}, {imei}, {capacidad}, {color}")

                        device_info = {
                            'udid': udid, # Importante para diagnÃ³sticos (apagar)
                            'product_type': product_type,
                            'product_name': model_name, # AÃ‘ADIDO: Pasar el nombre real calculado
                            'serial_number': serial_number,
                            'imei': imei,
                            'capacity': capacidad,
                            'color': color
                        }
                        
                        self.last_successful_udid = udid
                        self.success_callback(device_info)
                        
                    except Exception as e:
                        # Si falla (probablemente Trust Dialog), notificamos para pedir confianza
                        # No guardamos last_successful_udid, así que en el siguiente loop reintenta
                        if "PairingDialogResponsePending" in str(e) or "PasswordProtected" in str(e):
                             print(f"--- DEBUG: Error de Confianza detectado: {e}")
                             self.trust_callback()
                        else:
                             print(f"--- DEBUG: Error conectando (no confianza): {e}")
                             self.trust_callback() # Asumimos confianza por defecto en fallo conex
                            
                time.sleep(2)
                
            except Exception as e:
                # Error general en el loop
                print(f"Monitor loop error: {e}")
                time.sleep(5)

    def stop(self):
        self.running = False

# --- Constantes de Diseño y Paleta de Colores Modernos ---
COLOR_BG_DARK = "#0F172A"       # Fondo general de la ventana principal (Slate-900)
COLOR_CARD_BG = "#1E293B"       # Fondo de las tarjetas/paneles (Slate-800)
COLOR_CARD_BORDER = "#334155"   # Color del borde de las tarjetas (Slate-700)
COLOR_TEXT_PRIMARY = "#F8FAFC"  # Texto principal en blanco suave (Slate-50)
COLOR_TEXT_SECONDARY = "#94A3B8"# Texto secundario en gris azulado (Slate-400)

COLOR_ACCENT_PRIMARY = "#6366F1"# Indigo para acciones principales (PDF, etc.)
COLOR_ACCENT_PRIMARY_HOVER = "#4F46E5"

COLOR_ACCENT_SUCCESS = "#10B981"# Esmeralda para imprimir o éxito
COLOR_ACCENT_SUCCESS_HOVER = "#059669"

COLOR_ACCENT_SECONDARY = "#334155"# Pizarra para botones secundarios
COLOR_ACCENT_SECONDARY_HOVER = "#475569"

COLOR_ACCENT_DANGER = "#EF4444" # Rojo coral para borrar o apagar
COLOR_ACCENT_DANGER_HOVER = "#DC2626"

# --- Clase Principal de la Aplicación ---
class AppGeneradorEtiquetas(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.preview_ctk_image = None
        self._preview_update_job = None
        self._excel_update_job = None # Job para guardar cambios en Excel al editar
        
        # Cargar último Excel usado
        start_excel = cargar_excel_config()
        self.excel_manager = ExcelManager(start_excel)
        
        self.trust_window = None 
        self.current_udid = None # Guardar UDID para acciones posteriores (apagar)
        # Ventana de alerta "Confiar"
        
        # Configuración Inicial
        self.title("Generador de Etiquetas iPhone v3.2.5")
        self.geometry("1300x700") # Aumentado el alto para acomodar el header superior
        self.minsize(1200, 600)
        self.configure(fg_color=COLOR_BG_DARK)
        
        cargar_config_inicial()
        cargar_fuentes_pdf()
        
        # Configurar Grid Layout Principal (Fila 0 = Header, Fila 1 = Contenido)
        self.grid_rowconfigure(0, weight=0) # Header
        self.grid_rowconfigure(1, weight=1) # Contenido principal
        self.grid_columnconfigure(0, weight=0) # Controles (fijo 320)
        self.grid_columnconfigure(1, weight=0) # Vista Previa (fijo 420)
        self.grid_columnconfigure(2, weight=1) # Excel (expansible)

        # --- ENCABEZADO SUPERIOR (Header Banner) ---
        self.header_frame = customtkinter.CTkFrame(self, fg_color=COLOR_CARD_BG, corner_radius=12, border_width=1, border_color=COLOR_CARD_BORDER, height=70)
        self.header_frame.grid(row=0, column=0, columnspan=3, padx=20, pady=(20, 10), sticky="ew")
        self.header_frame.grid_propagate(False)
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)
        
        # Título y Subtítulo en Header
        title_box = customtkinter.CTkFrame(self.header_frame, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        lbl_main_title = customtkinter.CTkLabel(title_box, text="GENERADOR DE ETIQUETAS IPHONE", font=customtkinter.CTkFont(family="Segoe UI", size=18, weight="bold"), text_color=COLOR_TEXT_PRIMARY)
        lbl_main_title.pack(anchor="w")
        
        lbl_sub_title = customtkinter.CTkLabel(title_box, text="v3.2.5 • Control de Inventario, RMA y Excel Automático", font=customtkinter.CTkFont(family="Segoe UI", size=11), text_color=COLOR_TEXT_SECONDARY)
        lbl_sub_title.pack(anchor="w")
        
        # Badge de Estado de Conexión en Header
        self.badge_frame = customtkinter.CTkFrame(self.header_frame, fg_color="#334155", corner_radius=18, height=36, border_width=1, border_color="#475569")
        self.badge_frame.grid(row=0, column=1, padx=20, pady=17, sticky="e")
        self.badge_frame.grid_propagate(False)
        
        self.badge_label = customtkinter.CTkLabel(self.badge_frame, text="Estado: Desconectado ❌", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#94A3B8", padx=15, pady=0)
        self.badge_label.pack(expand=True, fill="both")

        # Crear Tarjetas principales
        # 1. Tarjeta de Controles (Izquierda - Ancho fijo 320)
        self.controls_frame = customtkinter.CTkFrame(self, width=320, fg_color=COLOR_CARD_BG, corner_radius=16, border_width=1, border_color=COLOR_CARD_BORDER)
        self.controls_frame.grid(row=1, column=0, padx=(20, 10), pady=(10, 20), sticky="nsew")
        self.controls_frame.grid_rowconfigure(2, weight=1)
        self.controls_frame.grid_columnconfigure(0, weight=1)
        self.controls_frame.grid_propagate(False)

        # 2. Tarjeta de Vista Previa (Centro - Ancho fijo 420)
        self.preview_frame = customtkinter.CTkFrame(self, width=420, fg_color=COLOR_CARD_BG, corner_radius=16, border_width=1, border_color=COLOR_CARD_BORDER)
        self.preview_frame.grid(row=1, column=1, padx=10, pady=(10, 20), sticky="nsew")
        self.preview_frame.grid_propagate(False)

        # 3. Tarjeta de Base de Datos Excel (Derecha - Expansible)
        self.excel_frame = customtkinter.CTkFrame(self, fg_color=COLOR_CARD_BG, corner_radius=16, border_width=1, border_color=COLOR_CARD_BORDER)
        self.excel_frame.grid(row=1, column=2, padx=(10, 20), pady=(10, 20), sticky="nsew")

        # --- Configurar Monitor de Dispositivos ---
        self.auto_print_var = tk.BooleanVar(value=False)
        if PYMOBILEDEVICE_AVAILABLE:
            self.device_monitor = DeviceMonitor(
                success_callback=self.on_device_info_received,
                trust_callback=self.on_device_trust_needed,
                disconnected_callback=self.on_device_disconnected
            )
            self.device_monitor.start()
        else:
            self.device_monitor = None

        self._setup_ui()
        self._bind_events()
        
        # Guardar configuración al cerrar la ventana
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.after(100, self.force_preview_update)

    def on_closing(self):
        """Guarda la configuración (logo) antes de cerrar la aplicación."""
        if self.device_monitor:
            self.device_monitor.stop()
        try:
            # Guardar el logo actual
            guardar_logo_config(self.logo_path_var.get().strip())
        except Exception as e:
            print(f"Error al guardar configuración al cerrar: {e}")
        finally:
            # Limpiar archivos temporales
            for f in temporary_files_to_delete:
                if os.path.exists(f):
                    try:
                        os.remove(f)
                    except Exception as e:
                        print(f"Error al borrar archivo temporal {f}: {e}")
            self.destroy()

    def _setup_ui(self):
        # Frame de Controles (Izquierda)
        self.controls_frame.grid_columnconfigure(0, weight=1)
        
        # Variables
        self.modelo_var = tk.StringVar()
        self.imei_var = tk.StringVar()
        logo_path_inicial = cargar_logo_config()
        self.logo_path_var = tk.StringVar(value=logo_path_inicial)
        
        # Mapear status_label al badge del header para retrocompatibilidad total
        self.status_label = self.badge_label
        
        # --- UI Elements de Control ---
        main_controls_frame = customtkinter.CTkFrame(self.controls_frame, fg_color="transparent")
        main_controls_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        main_controls_frame.grid_columnconfigure(0, weight=1)

        # Campo Modelo
        customtkinter.CTkLabel(main_controls_frame, text="Modelo:", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY).grid(row=0, column=0, padx=0, pady=(0,2), sticky="w")
        
        modelo_entry_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        modelo_entry_frame.grid(row=1, column=0, padx=0, pady=(0,10), sticky="ew")
        modelo_entry_frame.grid_columnconfigure(0, weight=1)
        
        self.modelo_entry = customtkinter.CTkEntry(modelo_entry_frame, textvariable=self.modelo_var, placeholder_text="Ej: iPhone 11 Black 64GB", fg_color="#0F172A", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, placeholder_text_color=COLOR_TEXT_SECONDARY, corner_radius=8, height=32)
        self.modelo_entry.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        btn_paste_model = customtkinter.CTkButton(modelo_entry_frame, text="Pegar", width=55, command=self.pegar_modelo, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), corner_radius=8, height=32)
        btn_paste_model.grid(row=0, column=1, padx=0)

        # Campo IMEI
        customtkinter.CTkLabel(main_controls_frame, text="IMEI:", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY).grid(row=2, column=0, padx=0, pady=(0,2), sticky="w")
        
        imei_entry_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        imei_entry_frame.grid(row=3, column=0, padx=0, pady=(0,15), sticky="ew")
        imei_entry_frame.grid_columnconfigure(0, weight=1)
        
        self.imei_entry = customtkinter.CTkEntry(imei_entry_frame, textvariable=self.imei_var, placeholder_text="Ingrese IMEI...", fg_color="#0F172A", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, placeholder_text_color=COLOR_TEXT_SECONDARY, corner_radius=8, height=32)
        self.imei_entry.grid(row=0, column=0, padx=(0,5), sticky="ew")
        
        btn_paste_imei = customtkinter.CTkButton(imei_entry_frame, text="Pegar", width=55, command=self.pegar_imei, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), corner_radius=8, height=32)
        btn_paste_imei.grid(row=0, column=1, padx=0)

        # Selección de Logo (Tarjeta interna)
        logo_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="#0F172A", corner_radius=10, border_width=1, border_color=COLOR_CARD_BORDER)
        logo_frame.grid(row=4, column=0, sticky='ew', pady=(0,15))
        logo_frame.grid_columnconfigure(0, weight=1)
        
        customtkinter.CTkLabel(logo_frame, text="Ruta del Logo:", font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=COLOR_TEXT_SECONDARY).grid(row=0, column=0, columnspan=2, padx=12, pady=(8,2), sticky="w")
        
        self.logo_entry = customtkinter.CTkEntry(logo_frame, textvariable=self.logo_path_var, fg_color="#1E293B", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, corner_radius=6, height=28)
        self.logo_entry.grid(row=1, column=0, padx=(12,5), pady=(0,10), sticky='ew')
        
        btn_search_logo = customtkinter.CTkButton(logo_frame, text="Buscar...", width=70, command=self.buscar_logo, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11), corner_radius=6, height=28)
        btn_search_logo.grid(row=1, column=1, padx=(0,12), pady=(0,10))

        # Botones de Acción (PDF & Imprimir)
        pdf_buttons_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        pdf_buttons_frame.grid(row=5, column=0, sticky='ew', pady=(0,15))
        pdf_buttons_frame.grid_columnconfigure((0,1), weight=1)
        
        btn_save_pdf = customtkinter.CTkButton(pdf_buttons_frame, text="Guardar PDF 💾", command=self.generar_y_guardar_pdf, fg_color=COLOR_ACCENT_PRIMARY, hover_color=COLOR_ACCENT_PRIMARY_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), corner_radius=10, height=38)
        btn_save_pdf.grid(row=0, column=0, padx=(0,5), sticky='ew')
        
        btn_print = customtkinter.CTkButton(pdf_buttons_frame, text="Imprimir 🖨️", command=self.imprimir, fg_color=COLOR_ACCENT_SUCCESS, hover_color=COLOR_ACCENT_SUCCESS_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), corner_radius=10, height=38)
        btn_print.grid(row=0, column=1, padx=(5,0), sticky='ew')

        # Interruptores de Automatización
        switch_frame = customtkinter.CTkFrame(main_controls_frame, fg_color="transparent")
        switch_frame.grid(row=6, column=0, pady=(5,0), sticky="ew")
        
        self.sw_auto_print = customtkinter.CTkSwitch(switch_frame, text="Auto-Imprimir al Conectar", variable=self.auto_print_var, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), progress_color=COLOR_ACCENT_SUCCESS, text_color=COLOR_TEXT_SECONDARY)
        self.sw_auto_print.pack(anchor="w", pady=(0,8))
        
        self.auto_shutdown_var = tk.BooleanVar(value=False)
        self.sw_auto_shutdown = customtkinter.CTkSwitch(switch_frame, text="Apagar al Terminar", variable=self.auto_shutdown_var, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), progress_color=COLOR_ACCENT_DANGER, text_color="#EF4444")
        self.sw_auto_shutdown.pack(anchor="w")

        # Botón de SumatraPDF y Créditos en el Pie de la Tarjeta
        bottom_frame = customtkinter.CTkFrame(self.controls_frame, fg_color="transparent")
        bottom_frame.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="s")
        bottom_frame.grid_columnconfigure(0, weight=1)
        
        btn_sumatra = customtkinter.CTkButton(bottom_frame, text="⚙️ Configurar SumatraPDF", command=self.configurar_ruta_sumatra_manualmente, fg_color=COLOR_ACCENT_SECONDARY, hover_color=COLOR_ACCENT_SECONDARY_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11), corner_radius=8, height=30)
        btn_sumatra.grid(row=0, column=0, sticky="ew", pady=(0,10))
        
        lbl_author = customtkinter.CTkLabel(bottom_frame, text="Hecho por Micael  ", font=customtkinter.CTkFont(family="Segoe UI", size=10, slant="italic"), text_color="gray50")
        lbl_author.grid(row=1, column=0, sticky="w")

        # --- Tarjeta de Vista Previa (Centro) ---
        self.preview_frame.grid_rowconfigure(1, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        
        # Título de Vista Previa
        lbl_preview_title = customtkinter.CTkLabel(self.preview_frame, text="VISTA PREVIA DE ETIQUETA", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY)
        lbl_preview_title.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="w")
        
        # Subtarjeta oscura simulando una mesa de trabajo para que resalte la etiqueta blanca
        preview_container = customtkinter.CTkFrame(self.preview_frame, fg_color="#111827", corner_radius=12, border_width=1, border_color="#1F2937")
        preview_container.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")
        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)

        self.preview_image_label = customtkinter.CTkLabel(preview_container, text="La previsualización aparecerá aquí.", font=customtkinter.CTkFont(family="Segoe UI", size=11, slant="italic"), text_color="gray50")
        self.preview_image_label.grid(row=0, column=0)

        # --- Tarjeta de Base de Datos Excel (Derecha) ---
        self._setup_excel_view()

    def _setup_excel_view(self):
        self.excel_frame.grid_rowconfigure(4, weight=1) # Fila 4 es la tabla
        self.excel_frame.grid_columnconfigure(0, weight=1)
        
        # --- Encabezado de la Sección de Base de Datos ---
        db_header_frame = customtkinter.CTkFrame(self.excel_frame, fg_color="transparent")
        db_header_frame.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        db_header_frame.grid_columnconfigure(0, weight=1)
        db_header_frame.grid_columnconfigure(1, weight=0)
        
        lbl_db_title = customtkinter.CTkLabel(db_header_frame, text="HISTORIAL DE REGISTROS EXCEL", font=customtkinter.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=COLOR_TEXT_PRIMARY)
        lbl_db_title.grid(row=0, column=0, sticky="w")
        
        # Contador de Registros reubicado elegantemente aquí
        self.count_label = customtkinter.CTkLabel(db_header_frame, text=f"Total: {self.excel_manager.obtener_conteo()}", font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color="#06B6D4")
        self.count_label.grid(row=0, column=1, sticky="e")
        
        # --- Visualizador de Excel Actual (Estilo cápsula de terminal) ---
        project_frame = customtkinter.CTkFrame(self.excel_frame, fg_color="#0F172A", corner_radius=10, border_width=1, border_color=COLOR_CARD_BORDER, height=36)
        project_frame.grid(row=1, column=0, padx=15, pady=(5, 10), sticky="ew")
        project_frame.grid_propagate(False)
        project_frame.grid_columnconfigure(1, weight=1)
        
        lbl_proj = customtkinter.CTkLabel(project_frame, text=" Excel:", font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=COLOR_TEXT_SECONDARY)
        lbl_proj.grid(row=0, column=0, padx=(12, 2), pady=4, sticky="w")
        
        self.lbl_filename = customtkinter.CTkLabel(project_frame, text=os.path.basename(self.excel_manager.filepath), font=customtkinter.CTkFont(family="Consolas", size=11, weight="bold"), text_color="#10B981", anchor="w")
        self.lbl_filename.grid(row=0, column=1, padx=(2, 10), pady=4, sticky="ew")

        # Botones de Proyecto Excel (Color Coded)
        btn_box = customtkinter.CTkFrame(self.excel_frame, fg_color="transparent")
        btn_box.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")
        btn_box.grid_columnconfigure((0,1,2), weight=1)
        
        btn_new = customtkinter.CTkButton(btn_box, text="Nuevo Excel", command=self.nuevo_proyecto, fg_color=COLOR_ACCENT_SUCCESS, hover_color=COLOR_ACCENT_SUCCESS_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), corner_radius=8, height=32)
        btn_new.grid(row=0, column=0, padx=(0,4), sticky="ew")
        
        btn_load = customtkinter.CTkButton(btn_box, text="Cargar Excel", command=self.cargar_proyecto, fg_color="#3B82F6", hover_color="#2563EB", text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), corner_radius=8, height=32)
        btn_load.grid(row=0, column=1, padx=2, sticky="ew")
        
        btn_export = customtkinter.CTkButton(btn_box, text="Exportar Copia", command=self.exportar_proyecto, fg_color=COLOR_ACCENT_PRIMARY, hover_color=COLOR_ACCENT_PRIMARY_HOVER, text_color=COLOR_TEXT_PRIMARY, font=customtkinter.CTkFont(family="Segoe UI", size=11, weight="bold"), corner_radius=8, height=32)
        btn_export.grid(row=0, column=2, padx=(4,0), sticky="ew")
        
        # --- Buscador Estilizado ---
        search_frame = customtkinter.CTkFrame(self.excel_frame, fg_color="transparent")
        search_frame.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_change)
        
        search_entry = customtkinter.CTkEntry(search_frame, textvariable=self.search_var, placeholder_text="🔍  Buscar IMEI o Modelo en la tabla...", fg_color="#0F172A", border_color=COLOR_CARD_BORDER, text_color=COLOR_TEXT_PRIMARY, placeholder_text_color=COLOR_TEXT_SECONDARY, corner_radius=8, height=32)
        search_entry.grid(row=0, column=0, sticky="ew")

        # Contenedor para Treeview (Fila 4)
        table_container = customtkinter.CTkFrame(self.excel_frame, fg_color="#0F172A", corner_radius=12, border_width=1, border_color=COLOR_CARD_BORDER)
        table_container.grid(row=4, column=0, padx=15, pady=(0, 15), sticky="nsew")
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        # Estilo para Treeview (Ultra Modern Dark)
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview", 
                        background="#1E293B", 
                        foreground="#F8FAFC", 
                        fieldbackground="#1E293B", 
                        bordercolor="#334155",
                        rowheight=28,  # Fila aireada, muy moderna
                        font=('Segoe UI', 10))
        
        style.map('Treeview', background=[('selected', '#4F46E5')], foreground=[('selected', '#FFFFFF')])
        
        style.configure("Treeview.Heading",
                        background="#334155",
                        foreground="#F8FAFC",
                        borderwidth=0,
                        font=('Segoe UI', 10, 'bold'))
        
        # Treeview
        self.tree = ttk.Treeview(table_container, columns=("IMEI", "Modelo"), show="headings", style="Treeview")
        self.tree.heading("IMEI", text="IMEI")
        self.tree.heading("Modelo", text="Modelo del Dispositivo")
        self.tree.column("IMEI", width=140, anchor="center")
        self.tree.column("Modelo", width=250, anchor="w")
        
        # Zebra-striping Tags
        self.tree.tag_configure("evenrow", background="#1E293B")
        self.tree.tag_configure("oddrow", background="#161E2E")
        
        # Scrollbar Estilizada
        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(1, 0), pady=1)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 1), pady=1)

        # Context Menu
        self.tree_menu = tk.Menu(self.tree, tearoff=0, bg="#1E293B", fg="#F8FAFC", activebackground="#4F46E5", activeforeground="#FFFFFF", borderwidth=1)
        self.tree_menu.add_command(label="✏️ Editar Modelo", command=self.edit_selected_item)
        self.tree_menu.add_command(label="❌ Eliminar Registro", command=self.delete_selected_item)
        
        # Bind de click derecho
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Cargar datos iniciales
        self.refresh_excel_table()
        
    def show_context_menu(self, event):
        """Muestra el menú contextual en la fila seleccionada."""
        try:
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self.tree_menu.post(event.x_root, event.y_root)
        except: pass

    def edit_selected_item(self):
        """Permite editar el modelo del ítem seleccionado usando un diálogo simple."""
        selected_item = self.tree.selection()
        if not selected_item: return
        
        item_values = self.tree.item(selected_item[0], 'values')
        if not item_values: return
        
        imei_val, current_model = item_values[0], item_values[1]
        
        # Usar customtkinter dialog
        dialog = customtkinter.CTkInputDialog(text=f"Editar Modelo para IMEI {imei_val}:", title="Editar Modelo")
        new_model = dialog.get_input()
        
        if new_model is not None: # Si no canceló
            new_model = new_model.strip()
            if new_model:
                self.excel_manager.actualizar_registro(imei_val, new_model)
                self.refresh_excel_table()
                # Si es el actual en pantalla, actualizarlo también
                if self.imei_var.get() == imei_val:
                    self.modelo_var.set(new_model)

    def delete_selected_item(self):
        """Elimina el ítem seleccionado con confirmación."""
        selected_item = self.tree.selection()
        if not selected_item: return
        
        item_values = self.tree.item(selected_item[0], 'values')
        if not item_values: return
        
        imei_val = item_values[0]
        
        confirm = messagebox.askyesno("Confirmar Eliminación", f"¿Estás seguro de eliminar el registro?\nIMEI: {imei_val}")
        if confirm:
            if self.excel_manager.eliminar_registro(imei_val):
                self.refresh_excel_table()
            else:
                messagebox.showerror("Error", "No se pudo eliminar el registro.")

    def nuevo_proyecto(self):
        """Crea un nuevo archivo Excel."""
        filepath = filedialog.asksaveasfilename(
            title="Crear Nuevo Proyecto Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialdir=script_dir
        )
        if filepath:
            if os.path.exists(filepath):
                # Si existe, preguntar si desea sobrescribir o cargar
                 res = messagebox.askyesno("Archivo Existe", "El archivo ya existe. ¿Deseas reemplazarlo?\n(Si dices 'No', se cargará el existente)")
                 if not res:
                     self._cambiar_archivo_excel(filepath)
                     return
            
        # Crear vacío
            try:
                df = pd.DataFrame(columns=["IMEI", "Modelo"])
                df.to_excel(filepath, index=False)
                self._cambiar_archivo_excel(filepath)
                messagebox.showinfo("Proyecto Nuevo", f"Creado: {os.path.basename(filepath)}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear el archivo:\n{e}")

    def cargar_proyecto(self):
        """Carga un Excel existente."""
        filepath = filedialog.askopenfilename(
            title="Cargar Proyecto Excel",
            filetypes=[("Excel Files", "*.xlsx")],
            initialdir=script_dir
        )
        if filepath:
            self._cambiar_archivo_excel(filepath)

    def exportar_proyecto(self):
        """Guarda una copia del actual en otra ruta."""
        if not os.path.exists(self.excel_manager.filepath):
            messagebox.showwarning("Aviso", "No hay archivo actual para exportar.")
            return

        import shutil
        dest_path = filedialog.asksaveasfilename(
            title="Exportar Copia Como",
            defaultextension=".xlsx",
            initialfile=f"Copia_{os.path.basename(self.excel_manager.filepath)}",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if dest_path:
            try:
                shutil.copy2(self.excel_manager.filepath, dest_path)
                messagebox.showinfo("Exportado", f"Copia guardada en:\n{dest_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Falló la exportación:\n{e}")

    def _cambiar_archivo_excel(self, filepath):
        """Actualiza el manager y la UI."""
        self.excel_manager.set_filepath(filepath)
        guardar_excel_config(filepath) # Guardar en config
        self.lbl_filename.configure(text=os.path.basename(filepath))
        self.refresh_excel_table()
    
    def refresh_excel_table(self, filter_query=""):
        """Borra la tabla y recarga los datos desde el Excel, aplicando filtro opcional y colores cebra."""
        # Limpiar
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Cargar
        registros = self.excel_manager.obtener_registros()
        query = filter_query.lower().strip()
        
        visible_count = 0
        for reg in registros:
            val_imei = str(reg.get("IMEI", "")).replace("nan", "")
            val_mod = str(reg.get("Modelo", "")).replace("nan", "")
            
            # Filtro
            if query and (query not in val_imei.lower() and query not in val_mod.lower()):
                continue
                
            # Zebra striping con las tags configuradas
            row_tag = "evenrow" if visible_count % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=(val_imei, val_mod), tags=(row_tag,))
            visible_count += 1
            
        if self.count_label:
            if query:
                self.count_label.configure(text=f"Filtrados: {visible_count} de {len(registros)}", text_color="#06B6D4")
            else:
                self.count_label.configure(text=f"Registros: {len(registros)}", text_color="#10B981")

    def on_search_change(self, *args):
        """Callback al escribir en el buscador."""
        self.refresh_excel_table(self.search_var.get())

    def _bind_events(self):
        for var in [self.modelo_var, self.imei_var, self.logo_path_var]:
            var.trace_add("write", self.schedule_preview_update)
            # También agendar actualización de excel si cambia el modelo
        self.modelo_var.trace_add("write", self.schedule_excel_update)

    def schedule_preview_update(self, *args):
        if self._preview_update_job:
            self.after_cancel(self._preview_update_job)
        self._preview_update_job = self.after(500, self.force_preview_update) # 500ms delay

    def schedule_excel_update(self, *args):
        """Guarda los cambios en el Excel si el usuario edita el texto manualmente."""
        if self._excel_update_job:
            self.after_cancel(self._excel_update_job)
        # Esperar 1.5 segundos de inactividad antes de guardar para no saturar disco
        self._excel_update_job = self.after(1500, self.perform_excel_update)

    def perform_excel_update(self):
        imei = self.imei_var.get().strip()
        nuevo_modelo = self.modelo_var.get().strip()
        if imei and nuevo_modelo:
            print(f"--- Guardando edición manual en Excel: {nuevo_modelo} ---")
            # Solo intentamos actualizar
            self.excel_manager.actualizar_registro(imei, nuevo_modelo)
            # Refrescar la tabla visual sin bloquear
            self.refresh_excel_table()
    def force_preview_update(self):
        try:
            # Para la etiqueta (preview), limpiamos el color
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
        """Elimina el color del texto para la impresión, sin importar dónde esté."""
        if not texto: return ""
        
        texto_limpio = texto
        # Ordenar por longitud para evitar remplazos parciales (ej: "Space Gray" antes que "Gray")
        for color in sorted(COLORES_PARA_REMOVER, key=len, reverse=True):
            # Usar regex para reemplazar el color (case insensitive) respetando bordes de palabra
            # \b asegura que no reemplacemos parte de una palabra (ej: 'Red' en 'Redux' - aunque improbable aquí)
            pattern = re.compile(re.escape(color), re.IGNORECASE)
            texto_limpio = pattern.sub("", texto_limpio)
            
        # Limpiar espacios dobles o residuales que queden
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip()
        return texto_limpio

    def generar_y_guardar_pdf(self): self._procesar_generacion(guardar_permanente=True)
    def imprimir(self): self._procesar_generacion(imprimir_despues=True)

    def _procesar_generacion(self, guardar_permanente=False, imprimir_despues=False):
        if not PDF_SAVE_ENABLED:
            messagebox.showerror("Función Deshabilitada", "La librería 'ReportLab' es necesaria para esta función.")
            return
            
        # Usamos la versión limpia (sin color) para imprimir
        texto_modelo_completo = self.modelo_var.get().strip().upper()
        modelo_limpio = self._limpiar_modelo_para_impresion(texto_modelo_completo)
        
        imei = self.imei_var.get().strip().upper()
        
        temp_pdf_path = _generar_etiqueta_pdf_temporal(
            modelo_limpio, imei,
            "",
            self.logo_path_var.get().strip()
        )
        if not temp_pdf_path or not os.path.exists(temp_pdf_path):
            messagebox.showerror("Error de Generación", "No se pudo crear el archivo PDF temporal.")
            return

        if guardar_permanente:
            # Pedir donde guardar
            dest_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=f"Etiqueta_{imei}.pdf" if imei else "Etiqueta.pdf"
            )
            if dest_path:
                try:
                    import shutil
                    shutil.copy2(temp_pdf_path, dest_path)
                    messagebox.showinfo("Guardado", "Etiqueta guardada exitosamente.")
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

        # IMPRIMIR
        if imprimir_despues:
            self._enviar_a_impresora(temp_pdf_path)
            
            # AUTO-APAGAR SI CORRESPONDE
            if self.auto_shutdown_var.get():
                # Esperar un momento a que la orden de impresión salga
                self.after(2000, self.intentar_apagar_dispositivo)

    def _enviar_a_impresora(self, pdf_path):
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error de Impresión", f"El archivo a imprimir no fue encontrado: {pdf_path}")
            return
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
            else:
                messagebox.showwarning("Sistema No Soportado", f"La impresión directa no está configurada para {current_os}.")
        except Exception as e:
            print(f"Error lanzando impresión: {e}")
            messagebox.showerror("Error de Impresión", f"No se pudo imprimir:\n{e}")
    def intentar_apagar_dispositivo(self):
        """Envía el comando de apagado al dispositivo conectado."""
        if not self.current_udid:
            print("No hay dispositivo registrado para apagar.")
            return
            
        try:
            print(f"Intentando apagar dispositivo: {self.current_udid}")
            # Crear servicio de diagnóstico para apagar
            lockdown = create_using_usbmux(serial=self.current_udid)
            diag = DiagnosticsService(lockdown=lockdown)
            diag.shutdown()
            print("Comando de apagado enviado.")
        except Exception as e:
            print(f"Error al intentar apagar: {e}")
            # A veces falla si se desconecta muy rápido, pero el comando suele llegar

    def pegar_modelo(self):
        """Pega el contenido del portapapeles tal cual (limpiando espacios extras)."""
        try:
            contenido = self.clipboard_get()
            if contenido:
                # El usuario quiere poder pegar 'iPhone 13 Blue 128GB' directamente sin cortes.
                texto_limpio = contenido.strip()

                # Limpiar espacios dobles
                while "  " in texto_limpio:
                    texto_limpio = texto_limpio.replace("  ", " ")
                
                self.modelo_entry.delete(0, tk.END)
                self.modelo_entry.insert(0, texto_limpio)
        except tk.TclError:
            messagebox.showwarning("Portapapeles Vacío", "No hay contenido en el portapapeles para pegar.")

    def pegar_imei(self):
        """Pega el contenido del portapapeles en el campo de IMEI, limpiando el contenido anterior."""
        try:
            contenido = self.clipboard_get()
            if contenido:
                self.imei_entry.delete(0, tk.END)  # Limpiar todo el contenido anterior
                self.imei_entry.insert(0, contenido.strip())  # Insertar el nuevo contenido
        except tk.TclError:
            messagebox.showwarning("Portapapeles Vacío", "No hay contenido en el portapapeles para pegar.")

    def buscar_logo(self):
        filepath = filedialog.askopenfilename(title="Seleccionar archivo de logo", filetypes=[("Archivos de Imagen", "*.png *.jpg *.jpeg"), ("Todos los archivos", "*.*")])
        if filepath:
            self.logo_path_var.set(filepath)
            guardar_logo_config(filepath)  # Guardar la ruta del logo seleccionado

    def configurar_ruta_sumatra_manualmente(self):
        global SUMATRA_PDF_PATH
        if platform.system() != "Windows":
            messagebox.showinfo("Información", "Esta opción de configuración es solo para el sistema operativo Windows.")
            return
        filepath = filedialog.askopenfilename(title="Localizar el ejecutable SumatraPDF.exe", filetypes=[("Ejecutable", "SumatraPDF.exe")])
        if filepath and os.path.basename(filepath).lower() == 'sumatrapdf.exe':
            SUMATRA_PDF_PATH = filepath
            guardar_config_sumatra()
            messagebox.showinfo("Éxito", f"La ruta de SumatraPDF ha sido establecida a:\n{filepath}")
        elif filepath: messagebox.showerror("Archivo Incorrecto", "Por favor, selecciona el archivo 'SumatraPDF.exe'.")


    def actualizar_estado(self, estado):
        """Actualiza el badge de estado de conexión con colores y textos modernos."""
        if estado == "conectado":
            self.badge_frame.configure(fg_color="#065F46", border_color="#047857") # Emerald-800 background, Emerald-700 border
            self.badge_label.configure(text="Estado: Conectado ✅", text_color="#A7F3D0") # Light emerald text
        elif estado == "confiar":
            self.badge_frame.configure(fg_color="#78350F", border_color="#B45309") # Amber-800 background, Amber-700 border
            self.badge_label.configure(text="Estado: ¡Falta Confiar! ⚠️", text_color="#FDE68A") # Light amber text
        else: # desconectado
            self.badge_frame.configure(fg_color="#334155", border_color="#475569") # Slate-700 background
            self.badge_label.configure(text="Estado: Desconectado ❌", text_color="#CBD5E1") # Light gray text

    # --- Callbacks de Dispositivo ---
    def on_device_info_received(self, info):
        self.after(0, self.hide_trust_alert) # Ocultar alerta si existe
        self.after(0, lambda: self.actualizar_estado("conectado"))
        self.after(0, lambda: self._update_ui_from_device(info))

    def on_device_disconnected(self):
        self.after(0, self.hide_trust_alert)
        self.after(0, lambda: self.actualizar_estado("desconectado"))

    def on_device_trust_needed(self):
        print("--- DEBUG: on_device_trust_needed LLAMADO ---")
        self.after(0, lambda: self.actualizar_estado("confiar"))
        self.after(0, self.show_trust_alert)

    def show_trust_alert(self):
        """Muestra una ventana modal no bloqueante pidiendo confiar."""
        print("--- DEBUG: Intentando mostrar alerta de confianza ---")
        if self.trust_window is None or not self.trust_window.winfo_exists():
            self.trust_window = customtkinter.CTkToplevel(self)
            self.trust_window.title("CONFÍA EN ESTE ORDENADOR")
            self.trust_window.geometry("500x250")
            self.trust_window.attributes("-topmost", True)
            self.trust_window.lift()
            self.trust_window.focus_force()
            
            # Intentar centrar
            try:
                x = self.winfo_x() + (self.winfo_width() // 2) - 250
                y = self.winfo_y() + (self.winfo_height() // 2) - 125
                self.trust_window.geometry(f"+{x}+{y}")
            except: pass
            
            # Contenido
            customtkinter.CTkLabel(self.trust_window, text="⚠️ ACCIÓN REQUERIDA ⚠️", font=("Arial", 20, "bold"), text_color="yellow").pack(pady=(20, 10))
            
            msg = "El iPhone detectado no ha confiado en este equipo.\n\n1. Desbloquea la pantalla del iPhone.\n2. Pulsa 'Confiar' (Trust) en el mensaje que aparece en el iPhone.\n3. Si pide código, ingrésalo."
            customtkinter.CTkLabel(self.trust_window, text=msg, font=("Arial", 14), justify="center").pack(pady=10)
            
            spinner = customtkinter.CTkProgressBar(self.trust_window, width=300, mode="indeterminate")
            spinner.pack(pady=10)
            spinner.start()
        else:
            print("--- DEBUG: La ventana de confianza ya existe, subiendo... ---")
            self.trust_window.lift()
            self.trust_window.focus_force()

    def hide_trust_alert(self):
        """Cierra la ventana de alerta."""
        if self.trust_window and self.trust_window.winfo_exists():
            self.trust_window.destroy()
        self.trust_window = None

    def _update_ui_from_device(self, info):
        product_type = info.get('product_type', '')
        # Mapear a nombre legible
        model_name = IPHONE_MODEL_MAPPING.get(product_type, product_type)
        if not model_name: 
            model_name = "iPhone Desconocido"
            
        imei = info.get('imei') or info.get('serial_number') or ""
        capacidad = info.get('capacity', '')
        color = info.get('color', '') # Viene en Hex a veces
        
        self.current_udid = info.get('udid') # Guardar UDID
        
        # El modelo ya viene calculado desde el thread
        model_name = info.get('product_name', 'iPhone') # Usamos el nombre si viene, o genÃ©rico
        
        # CAMBIO: Unir Modelo, Color y Capacidad en formato 3uTools (Model Color Capacity)
        full_model_text = f"{model_name} {color} {capacidad}".strip()
        
        # Limpiar espacios dobles por si acaso falta color o cap
        while "  " in full_model_text:
            full_model_text = full_model_text.replace("  ", " ")

        self.modelo_var.set(full_model_text)
        self.imei_var.set(imei)
        
        print(f"Dispositivo procesado: {full_model_text} - {imei}")
        
        # --- REGISTRO AUTOMÁTICO EN EXCEL ---
        if imei:
            # Como modelo_var ya tiene todo el texto, usamos eso.
            # Limpiar espacios dobles
            modelo_completo_excel = self.modelo_var.get()
            while "  " in modelo_completo_excel:
                modelo_completo_excel = modelo_completo_excel.replace("  ", " ")

            exito, msg, count = self.excel_manager.registrar_dispositivo(imei, modelo_completo_excel)
            if exito:
                print(f"Excel: Registrado {imei}")
                if count >= 0:
                     self.count_label.configure(text=f"Registrados: {count}", text_color="#10B981")
                # Actualizar la tabla visual
                self.after(0, self.refresh_excel_table)
            else:
                print(f"Excel Info: {msg}")
                # Si falló porque el archivo estaba abierto, avisar al usuario
                if "abierto" in msg.lower():
                     messagebox.showwarning("Error Excel", msg)
        
        if self.auto_print_var.get():
            # Pequeño delay para asegurar que la UI se refresque visualmente antes de imprimir
            self.after(500, self.imprimir)


if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")
    
    app = AppGeneradorEtiquetas()
    app.mainloop()
