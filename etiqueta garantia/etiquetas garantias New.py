# -*- coding: utf-8 -*-
"""
Generador de Etiquetas de Garantía v2.3 (Python Edition)
Autor: Asistente de IA (basado en el trabajo de Micael)
Fecha: 2025-06-27

Descripción:
- Versión final enfocada en Python con CustomTkinter.
- Interfaz moderna y oscura.
- Diseño de etiqueta horizontal (4x3 pulgadas).
- Sistema de impresión con SumatraPDF.
- AÑADIDO: Lógica de auto-ajuste de ancho para el código de barras, evitando que se corte.
- Convierte automáticamente a mayúsculas el texto del cliente, problema e IMEI/SN.
"""
import customtkinter
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont

import barcode
from barcode.writer import ImageWriter
import io
import os
import platform
import subprocess
import tempfile
import atexit
import json
from datetime import datetime

# --- Dependencias para Guardado/Impresión en PDF ---
PDF_SAVE_ENABLED = False
try:
    from reportlab.pdfgen import canvas as reportlab_canvas
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.units import inch
    from reportlab.lib.utils import ImageReader as ReportLabImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.graphics.barcode import code128
    PDF_SAVE_ENABLED = True
except ImportError:
    print("ADVERTENCIA: Las librerías requeridas no están instaladas. El guardado en PDF y la impresión estarán deshabilitados.")
    print("Por favor, ejecute: pip install customtkinter Pillow reportlab python-barcode")
    PDF_SAVE_ENABLED = False

# --- Constantes ---
CONFIG_FILE_NAME = "etiqueta_garantia_config.json"
LABEL_WIDTH_INCHES = 4
LABEL_HEIGHT_INCHES = 3
PREVIEW_MAX_WIDTH = 400
PREVIEW_MAX_HEIGHT = int(PREVIEW_MAX_WIDTH * (LABEL_HEIGHT_INCHES / LABEL_WIDTH_INCHES))

# --- Rutas de Fuentes (Asegúrate de que estos archivos .ttf estén en la misma carpeta) ---
FONT_BOLD_PATH_TTF = "arialbd.ttf"
FONT_REGULAR_PATH_TTF = "arial.ttf"

# --- Nombres de Fuentes para ReportLab (PDF) ---
RL_FONT_BOLD_NAME = "ArialBoldRegistered"
RL_FONT_REGULAR_NAME = "ArialRegularRegistered"

# --- Variables Globales ---
SUMATRA_PDF_PATH = None
temporary_files_to_delete = []

# --- Funciones de Configuración y Limpieza (del script original) ---

def _read_config():
    if os.path.exists(CONFIG_FILE_NAME):
        with open(CONFIG_FILE_NAME, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return {}
    return {}

def _write_config(config_data):
    try:
        with open(CONFIG_FILE_NAME, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
    except Exception as e:
        print(f"Error al guardar configuración: {e}")

def cargar_config_inicial():
    global SUMATRA_PDF_PATH
    config = _read_config()
    path_guardado = config.get("sumatra_pdf_path")
    if path_guardado and os.path.exists(path_guardado) and os.path.isfile(path_guardado):
        SUMATRA_PDF_PATH = path_guardado
    else:
        detectar_sumatra_si_no_configurado()

def guardar_config_sumatra():
    if SUMATRA_PDF_PATH and platform.system() == "Windows":
        config = _read_config()
        config["sumatra_pdf_path"] = SUMATRA_PDF_PATH
        _write_config(config)

def detectar_sumatra_si_no_configurado():
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
                return
            elif os.path.exists(path_candidate) and os.path.isfile(path_candidate):
                SUMATRA_PDF_PATH = path_candidate
                return
        except Exception: pass

def cleanup_temp_files():
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

# --- Funciones de Generación de Etiquetas ---

def _generar_etiqueta_pil_image(cliente, fecha, imei_sn, problema):
    """Genera la etiqueta como una imagen PIL para la previsualización."""
    DPI = 300
    LABEL_WIDTH_PX, LABEL_HEIGHT_PX = int(LABEL_WIDTH_INCHES * DPI), int(LABEL_HEIGHT_INCHES * DPI)
    
    image = Image.new("RGB", (LABEL_WIDTH_PX, LABEL_HEIGHT_PX), "white")
    draw = ImageDraw.Draw(image)
    
    try:
        font_title = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=52)
        font_label = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=36)
        font_text = ImageFont.truetype(FONT_REGULAR_PATH_TTF, size=36)
        font_sn = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=34)
    except IOError:
        font_title, font_label, font_text, font_sn = [ImageFont.load_default()]*4

    padding = 30
    
    title_text = "RECIBO DE GARANTÍA"
    title_w = draw.textlength(title_text, font=font_title)
    draw.text(((LABEL_WIDTH_PX - title_w) / 2, padding), title_text, fill="black", font=font_title)
    draw.line([(padding, 100), (LABEL_WIDTH_PX - padding, 100)], fill="black", width=3)

    y_pos = 130
    draw.text((padding, y_pos), "Cliente:", fill="black", font=font_label)
    draw.text((padding + 160, y_pos), cliente, fill="black", font=font_text)
    y_pos += 55
    draw.text((padding, y_pos), "Fecha:", fill="black", font=font_label)
    draw.text((padding + 160, y_pos), fecha, fill="black", font=font_text)
    y_pos += 55
    draw.text((padding, y_pos), "Problema:", fill="black", font=font_label)
    
    lines = []
    if problema:
        max_chars = 45
        lines = [problema[i:i+max_chars] for i in range(0, len(problema), max_chars)]
    for i, line in enumerate(lines):
        draw.text((padding, y_pos + 55 + (i * 45)), line, fill="black", font=font_text)

    footer_y = LABEL_HEIGHT_PX - 220
    draw.line([(padding, footer_y), (LABEL_WIDTH_PX - padding, footer_y)], fill="gray", width=2)
    
    # Lado derecho del footer
    listo_area_width = 300
    listo_x = LABEL_WIDTH_PX - padding - listo_area_width
    draw.rectangle([(listo_x + 20, footer_y + 110), (listo_x + 70, footer_y + 160)], outline="black", width=4)
    draw.text((listo_x + 90, footer_y + 115), "Listo?", fill="black", font=font_label)

    # Lado izquierdo del footer
    left_footer_width = LABEL_WIDTH_PX - (padding * 2) - listo_area_width
    sn_text = f"IMEI / S/N: {imei_sn}"
    sn_w = draw.textlength(sn_text, font=font_sn)
    draw.text(((left_footer_width - sn_w) / 2 + padding, footer_y + 40), sn_text, fill="black", font=font_sn)
    
    if imei_sn and imei_sn.strip().lower() != 'n/a':
        try:
            code128_lib = barcode.get_barcode_class('code128')
            barcode_pil = code128_lib(imei_sn, writer=ImageWriter()).render({'module_height': 10.0, 'write_text': False, 'quiet_zone': 2})
            
            # --- LÓGICA DE AUTO-AJUSTE PARA PREVIEW ---
            max_bc_width = left_footer_width - 20
            if barcode_pil.width > max_bc_width:
                ratio = max_bc_width / barcode_pil.width
                new_height = int(barcode_pil.height * ratio)
                barcode_pil = barcode_pil.resize((int(max_bc_width), new_height), Image.Resampling.LANCZOS)
            
            bc_x = int((left_footer_width - barcode_pil.width) / 2) + padding
            image.paste(barcode_pil, (bc_x, footer_y + 90))
        except Exception:
            pass

    return image

def _generar_etiqueta_pdf_temporal(cliente, fecha, imei_sn, problema):
    """Genera la etiqueta como PDF para imprimir con calidad mejorada y tamaño 4x3."""
    if not PDF_SAVE_ENABLED: return None
    fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="etiqueta_")
    os.close(fd)
    temporary_files_to_delete.append(temp_pdf_path)
    
    page_size = (LABEL_WIDTH_INCHES * inch, LABEL_HEIGHT_INCHES * inch)
    c = reportlab_canvas.Canvas(temp_pdf_path, pagesize=page_size)
    width, height = page_size

    padding = 0.25 * inch
    
    c.setFont(RL_FONT_BOLD_NAME, 18)
    c.drawCentredString(width / 2, height - (padding * 1.8), "RECIBO DE GARANTÍA")
    c.line(padding, height - (padding * 2.2), width - padding, height - (padding * 2.2))

    y = height - (padding * 3.2)
    c.setFont(RL_FONT_BOLD_NAME, 11)
    c.drawString(padding, y, "Cliente:")
    c.setFont(RL_FONT_REGULAR_NAME, 11)
    c.drawString(padding + 0.9 * inch, y, cliente)
    
    y -= 0.35 * inch
    c.setFont(RL_FONT_BOLD_NAME, 11)
    c.drawString(padding, y, "Fecha:")
    c.setFont(RL_FONT_REGULAR_NAME, 11)
    c.drawString(padding + 0.9 * inch, y, fecha)

    y -= 0.35 * inch
    c.setFont(RL_FONT_BOLD_NAME, 11)
    c.drawString(padding, y, "Problema:")
    
    text_area_y_start = y - 0.25 * inch
    text_object = c.beginText(padding, text_area_y_start)
    text_object.setFont(RL_FONT_REGULAR_NAME, 11)
    
    max_width = width - (padding * 2)
    lines = problema.split('\n')
    for line in lines:
        while pdfmetrics.stringWidth(line, RL_FONT_REGULAR_NAME, 11) > max_width:
            idx = line.rfind(' ', 0, int(max_width / (pdfmetrics.stringWidth('a', RL_FONT_REGULAR_NAME, 11) or 1)))
            if idx == -1: idx = int(max_width / (pdfmetrics.stringWidth('a', RL_FONT_REGULAR_NAME, 11) or 1))
            text_object.textLine(line[:idx])
            line = line[idx+1:]
        text_object.textLine(line)
    c.drawText(text_object)
    
    footer_y_line = 1.1 * inch
    c.setDash(3, 3)
    c.line(padding, footer_y_line, width - padding, footer_y_line)
    c.setDash([])

    # Columna Derecha: Listo?
    listo_area_width = 1.2 * inch
    listo_x_pos = width - padding - listo_area_width
    c.rect(listo_x_pos, 0.45 * inch, 0.3 * inch, 0.3 * inch, stroke=1, fill=0)
    c.setFont(RL_FONT_BOLD_NAME, 12)
    c.drawString(listo_x_pos + 0.4 * inch, 0.5 * inch, "Listo?")

    # Columna Izquierda: IMEI y Barcode
    left_column_width = width - (padding * 2) - listo_area_width
    c.setFont(RL_FONT_BOLD_NAME, 11) 
    c.drawCentredString(padding + left_column_width / 2, 0.8 * inch, f"IMEI / S/N: {imei_sn}")
    
    if imei_sn and imei_sn.strip().lower() != 'n/a':
        try:
            # --- SOLUCIÓN v2.3: Lógica de auto-ajuste de ancho ---
            bar_width = 1.4 # Grosor ideal de la barra
            barcode = code128.Code128(imei_sn, barHeight=0.4*inch, barWidth=bar_width)

            # Si el código de barras es más ancho que el espacio disponible, se reduce el grosor de la barra
            if barcode.width > left_column_width:
                scale = left_column_width / barcode.width
                bar_width *= scale
                barcode = code128.Code128(imei_sn, barHeight=0.4*inch, barWidth=bar_width)
            
            x_pos = padding + (left_column_width - barcode.width) / 2
            barcode.drawOn(c, x_pos, 0.25 * inch)

        except Exception as e:
            error_msg = f"No se pudo generar el código de barras.\n\nRazón: {e}\n\nAsegúrese de que el IMEI/SN no contenga caracteres extraños."
            messagebox.showerror("Error de Código de Barras", error_msg)
            print(error_msg)
    
    c.save()
    return temp_pdf_path

# --- Clase Principal de la Aplicación ---
class AppGeneradorGarantias(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self._preview_update_job = None
        
        self.title("Generador de Etiquetas de Garantía v2.3")
        self.geometry("800x650") 
        self.minsize(750, 600)
        
        customtkinter.set_appearance_mode("dark")
        cargar_config_inicial()
        cargar_fuentes_pdf()
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.controls_frame = customtkinter.CTkFrame(self, width=320, corner_radius=0)
        self.controls_frame.grid(row=0, column=0, sticky="nsw")
        self.controls_frame.grid_rowconfigure(4, weight=1)

        self.preview_frame = customtkinter.CTkFrame(self)
        self.preview_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        self._setup_ui()
        self._bind_events()
        self.after(100, self.force_preview_update)

    def _setup_ui(self):
        self.controls_frame.grid_columnconfigure(0, weight=1)
        
        self.cliente_var = tk.StringVar()
        self.fecha_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.imei_sn_var = tk.StringVar()
        
        main_controls_frame = customtkinter.CTkFrame(self.controls_frame, fg_color="transparent")
        main_controls_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        main_controls_frame.grid_columnconfigure(0, weight=1)

        customtkinter.CTkLabel(main_controls_frame, text="Nombre del Cliente:", font=customtkinter.CTkFont(weight="bold")).pack(anchor="w")
        customtkinter.CTkEntry(main_controls_frame, textvariable=self.cliente_var).pack(fill="x", pady=(2, 10))

        customtkinter.CTkLabel(main_controls_frame, text="Fecha de Ingreso:", font=customtkinter.CTkFont(weight="bold")).pack(anchor="w")
        customtkinter.CTkEntry(main_controls_frame, textvariable=self.fecha_var).pack(fill="x", pady=(2, 10))

        customtkinter.CTkLabel(main_controls_frame, text="IMEI o S/N:", font=customtkinter.CTkFont(weight="bold")).pack(anchor="w")
        customtkinter.CTkEntry(main_controls_frame, textvariable=self.imei_sn_var).pack(fill="x", pady=(2, 10))
        
        customtkinter.CTkLabel(main_controls_frame, text="Problema del Equipo:", font=customtkinter.CTkFont(weight="bold")).pack(anchor="w")
        self.problema_textbox = customtkinter.CTkTextbox(main_controls_frame, height=120)
        self.problema_textbox.pack(fill="x", pady=(2, 10))

        customtkinter.CTkButton(self.controls_frame, text="Imprimir Etiqueta", command=self.imprimir, height=40).grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        customtkinter.CTkButton(self.controls_frame, text="Configurar SumatraPDF", command=self.configurar_ruta_sumatra_manualmente).grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        
        customtkinter.CTkLabel(self.controls_frame, text="Hecho por IA (Basado en Micael)", font=customtkinter.CTkFont(size=10, slant="italic"), text_color="gray50").grid(row=4, column=0, padx=20, pady=10, sticky="s")

        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_image_label = customtkinter.CTkLabel(self.preview_frame, text="La previsualización aparecerá aquí.", text_color="gray60")
        self.preview_image_label.grid(row=0, column=0, sticky="nsew")

    def _bind_events(self):
        self.cliente_var.trace_add("write", self.schedule_preview_update)
        self.fecha_var.trace_add("write", self.schedule_preview_update)
        self.imei_sn_var.trace_add("write", self.schedule_preview_update)
        self.problema_textbox.bind("<KeyRelease>", self.schedule_preview_update)

    def schedule_preview_update(self, *args):
        if self._preview_update_job: self.after_cancel(self._preview_update_job)
        self._preview_update_job = self.after(250, self.force_preview_update)

    def force_preview_update(self):
        try:
            pil_image = _generar_etiqueta_pil_image(
                self.cliente_var.get().upper(),
                self.fecha_var.get(),
                self.imei_sn_var.get().upper(),
                self.problema_textbox.get("1.0", "end-1c").upper()
            )
            
            self.preview_ctk_image = customtkinter.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(PREVIEW_MAX_WIDTH, PREVIEW_MAX_HEIGHT)
            )
            
            self.preview_image_label.configure(image=self.preview_ctk_image, text="")
        except Exception as e:
            self.preview_image_label.configure(image=None, text=f"Error en preview:\n{e}")

    def imprimir(self):
        if not PDF_SAVE_ENABLED:
            messagebox.showerror("Función Deshabilitada", "La librería 'ReportLab' es necesaria para esta función.")
            return
        
        cliente = self.cliente_var.get().strip()
        imei_sn = self.imei_sn_var.get().strip()
        if not cliente or not imei_sn:
            messagebox.showerror("Campos Obligatorios", "'Nombre del Cliente' y 'IMEI o S/N' son campos obligatorios.")
            return
            
        temp_pdf_path = _generar_etiqueta_pdf_temporal(
            cliente.upper(),
            self.fecha_var.get(),
            imei_sn.upper(),
            self.problema_textbox.get("1.0", "end-1c").upper()
        )
        if not temp_pdf_path or not os.path.exists(temp_pdf_path):
            if 'last_error' not in self.__dict__ or self.last_error != 'barcode':
                 messagebox.showerror("Error de Generación", "No se pudo crear el archivo PDF temporal.")
            return
        
        self.imprimir_pdf_directo(temp_pdf_path)

    def imprimir_pdf_directo(self, filepath):
        if not os.path.exists(filepath):
            messagebox.showerror("Error de Impresión", f"El archivo a imprimir no fue encontrado: {filepath}")
            return
        current_os = platform.system()
        try:
            if current_os == "Windows":
                if SUMATRA_PDF_PATH and os.path.exists(SUMATRA_PDF_PATH):
                    subprocess.Popen([SUMATRA_PDF_PATH, "-print-to-default", "-silent", filepath])
                else: 
                    messagebox.showwarning("SumatraPDF no encontrado", "SumatraPDF no está configurado. Se abrirá el diálogo de impresión de Windows.")
                    os.startfile(filepath, "print")
            elif current_os in ["Darwin", "Linux"]:
                cmd = "lpr" if current_os == "Darwin" else "lp"
                subprocess.run([cmd, filepath], check=True)
            else: messagebox.showwarning("Sistema No Soportado", f"La impresión directa no está configurada para {current_os}.")
        except FileNotFoundError: messagebox.showerror("Error de Comando", "Comando de impresión no encontrado (lpr o lp). Asegúrate de que esté instalado.")
        except Exception as e: messagebox.showerror("Error de Impresión", f"Ocurrió un error inesperado al imprimir:\n\n{e}")

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


if __name__ == "__main__":
    app = AppGeneradorGarantias()
    app.mainloop()
