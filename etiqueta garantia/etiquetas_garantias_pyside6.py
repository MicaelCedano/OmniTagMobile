# -*- coding: utf-8 -*-
"""
Generador de Etiquetas de Garantía - PySide6 Edition
Autor: Asistente de IA
Fecha: 2025-12-05

Descripción:
- Migración a PySide6 para una interfaz más moderna y fluida.
- Diseño de etiqueta modernizado.
- Mantiene funcionalidad de impresión y configuración.
"""

import sys
import os
import json
import platform
import subprocess
import tempfile
import atexit
from datetime import datetime
import io

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QLineEdit, QTextEdit, 
                               QPushButton, QMessageBox, QFileDialog, QFrame,
                               QSplitter, QSizePolicy)
from PySide6.QtCore import Qt, QTimer, QSize
from PySide6.QtGui import QFont, QIcon, QPixmap, QImage, QColor, QPalette

from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter

# --- Dependencias para Guardado/Impresión en PDF ---
PDF_SAVE_ENABLED = False
try:
    from reportlab.pdfgen import canvas as reportlab_canvas
    from reportlab.lib.pagesizes import landscape
    from reportlab.lib.units import inch
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.graphics.barcode import code128
    PDF_SAVE_ENABLED = True
except ImportError:
    print("ADVERTENCIA: ReportLab no encontrado. Funciones de PDF deshabilitadas.")

# --- Constantes ---
CONFIG_FILE_NAME = "etiqueta_garantia_config.json"
LABEL_WIDTH_INCHES = 4
LABEL_HEIGHT_INCHES = 3
PREVIEW_WIDTH = 380
PREVIEW_HEIGHT = int(PREVIEW_WIDTH * (LABEL_HEIGHT_INCHES / LABEL_WIDTH_INCHES))

# --- Rutas de Fuentes ---
# Intentaremos usar fuentes del sistema si no hay archivos locales, pero definimos fallbacks
FONT_BOLD_PATH_TTF = "arialbd.ttf"
FONT_REGULAR_PATH_TTF = "arial.ttf"
RL_FONT_BOLD_NAME = "ArialBoldRegistered"
RL_FONT_REGULAR_NAME = "ArialRegularRegistered"

# --- Variables Globales ---
SUMATRA_PDF_PATH = None
temporary_files_to_delete = []

# --- Estilos (Modern Dark Theme) ---
STYLESHEET = """
QMainWindow {
    background-color: #1e1e1e;
}
QWidget {
    color: #e0e0e0;
    font-family: 'Segoe UI', 'Roboto', sans-serif;
    font-size: 14px;
}
QFrame#ControlsFrame {
    background-color: #252526;
    border-right: 1px solid #3e3e42;
}
QLabel {
    color: #cccccc;
    font-weight: 500;
}
QLineEdit, QTextEdit {
    background-color: #3c3c3c;
    border: 1px solid #555555;
    border-radius: 5px;
    color: #ffffff;
    padding: 5px;
    selection-background-color: #264f78;
}
QLineEdit:focus, QTextEdit:focus {
    border: 1px solid #007acc;
}
QPushButton {
    background-color: #0e639c;
    color: white;
    border: none;
    border-radius: 5px;
    padding: 8px 16px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #1177bb;
}
QPushButton:pressed {
    background-color: #094771;
}
QPushButton#ConfigButton {
    background-color: #3a3d41;
}
QPushButton#ConfigButton:hover {
    background-color: #45494e;
}
QLabel#PreviewPlaceholder {
    border: 2px dashed #555555;
    border-radius: 10px;
    color: #666666;
}
"""

# --- Funciones de Configuración y Utilidades ---

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
            # Intentar buscar con 'where'
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
        else: raise IOError
    except Exception:
        RL_FONT_BOLD_NAME = 'Helvetica-Bold'
    try:
        if os.path.exists(FONT_REGULAR_PATH_TTF):
            pdfmetrics.registerFont(TTFont(RL_FONT_REGULAR_NAME, FONT_REGULAR_PATH_TTF))
        else: raise IOError
    except Exception:
        RL_FONT_REGULAR_NAME = 'Helvetica'

# --- Generación de Etiquetas (Lógica Mejorada) ---

def _generar_etiqueta_pil_image(cliente, fecha, imei_sn, problema):
    """Genera la etiqueta como imagen PIL para previsualización."""
    DPI = 300
    WIDTH_PX, HEIGHT_PX = int(LABEL_WIDTH_INCHES * DPI), int(LABEL_HEIGHT_INCHES * DPI)
    
    # Fondo blanco
    img = Image.new("RGB", (WIDTH_PX, HEIGHT_PX), "white")
    draw = ImageDraw.Draw(img)
    
    try:
        # Intentar cargar fuentes, fallback a default
        font_header = ImageFont.truetype(FONT_BOLD_PATH_TTF, 60)
        font_label = ImageFont.truetype(FONT_BOLD_PATH_TTF, 32)
        font_content = ImageFont.truetype(FONT_REGULAR_PATH_TTF, 32)
        font_small = ImageFont.truetype(FONT_REGULAR_PATH_TTF, 24)
        font_footer = ImageFont.truetype(FONT_BOLD_PATH_TTF, 28)
    except IOError:
        font_header = font_label = font_content = font_small = font_footer = ImageFont.load_default()

    padding = 40
    
    # --- Header Moderno ---
    # --- Header Moderno (Ink Saving) ---
    # Línea superior e inferior del título
    draw.line([(padding, 130), (WIDTH_PX - padding, 130)], fill="black", width=4)
    
    header_text = "RECIBO DE GARANTÍA"
    # Centrar texto
    text_bbox = draw.textbbox((0, 0), header_text, font=font_header)
    text_width = text_bbox[2] - text_bbox[0]
    draw.text(((WIDTH_PX - text_width) / 2, 45), header_text, fill="black", font=font_header)
    
    # --- Contenido ---
    current_y = 180
    line_spacing = 50
    
    # Cliente
    draw.text((padding, current_y), "CLIENTE:", fill="#333333", font=font_label)
    draw.text((padding + 180, current_y), cliente, fill="black", font=font_content)
    
    current_y += line_spacing
    # Fecha
    draw.text((padding, current_y), "FECHA:", fill="#333333", font=font_label)
    draw.text((padding + 180, current_y), fecha, fill="black", font=font_content)
    
    current_y += line_spacing + 10
    # Problema (con caja gris suave de fondo)
    draw.text((padding, current_y), "PROBLEMA REPORTADO:", fill="#333333", font=font_label)
    current_y += 40
    
    # Caja de problema
    problem_box_height = 220
    draw.rectangle([(padding, current_y), (WIDTH_PX - padding, current_y + problem_box_height)], 
                   fill="#f5f5f5", outline="#dddddd")
    
    # Texto del problema (multilinea simple)
    max_chars = 50
    lines = []
    if problema:
        words = problema.split()
        current_line = []
        for word in words:
            if len(" ".join(current_line + [word])) <= max_chars:
                current_line.append(word)
            else:
                lines.append(" ".join(current_line))
                current_line = [word]
        if current_line: lines.append(" ".join(current_line))
    
    text_y = current_y + 15
    for i, line in enumerate(lines[:5]): # Limitar a 5 líneas
        draw.text((padding + 15, text_y + (i * 35)), line, fill="#333333", font=font_content)

    # --- Footer ---
    footer_y = HEIGHT_PX - 180
    
    # Línea separadora
    draw.line([(padding, footer_y), (WIDTH_PX - padding, footer_y)], fill="#333333", width=2)
    
    # Sección IMEI/SN
    footer_text_y = footer_y + 20
    sn_text = f"IMEI / SN: {imei_sn}"
    draw.text((padding, footer_text_y), sn_text, fill="black", font=font_footer)
    
    # Código de barras
    if imei_sn and imei_sn.strip().lower() != 'n/a':
        try:
            code128_lib = barcode.get_barcode_class('code128')
            # Generar barcode
            barcode_pil = code128_lib(imei_sn, writer=ImageWriter()).render(
                {'module_height': 8.0, 'write_text': False, 'quiet_zone': 1}
            )
            
            # Redimensionar si es necesario
            max_bc_width = WIDTH_PX - padding - 350 # Espacio dejando lugar para "Listo?"
            if barcode_pil.width > max_bc_width:
                ratio = max_bc_width / barcode_pil.width
                new_h = int(barcode_pil.height * ratio)
                barcode_pil = barcode_pil.resize((int(max_bc_width), new_h), Image.Resampling.LANCZOS)
            
            img.paste(barcode_pil, (padding, footer_text_y + 40))
        except Exception:
            pass

    # Sección "¿Listo?" (Estilo Checkbox)
    checkbox_size = 50
    # Posicionar a la derecha
    check_x = WIDTH_PX - 280
    check_y = footer_y + 50
    
    # Dibujar cuadrito
    draw.rectangle([(check_x, check_y), (check_x + checkbox_size, check_y + checkbox_size)], outline="black", width=3)
    
    # Texto "Listo?" al lado
    listo_text = "¿Listo?"
    listo_font = ImageFont.truetype(FONT_BOLD_PATH_TTF, 36)
    draw.text((check_x + checkbox_size + 20, check_y + 5), listo_text, fill="black", font=listo_font)

    return img

def _generar_etiqueta_pdf_temporal(cliente, fecha, imei_sn, problema):
    """Genera PDF temporal con diseño moderno."""
    if not PDF_SAVE_ENABLED: return None
    fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="etiqueta_moderna_")
    os.close(fd)
    temporary_files_to_delete.append(temp_pdf_path)
    
    page_size = (LABEL_WIDTH_INCHES * inch, LABEL_HEIGHT_INCHES * inch)
    c = reportlab_canvas.Canvas(temp_pdf_path, pagesize=page_size)
    width, height = page_size
    
    # --- Diseño PDF ReportLab ---
    
    # Header Simple (Ink Saving)
    c.setFillColorRGB(0, 0, 0) # Negro
    c.setFont(RL_FONT_BOLD_NAME, 20)
    c.drawCentredString(width/2, height - 0.5*inch, "RECIBO DE GARANTÍA")
    
    # Línea debajo del título
    c.setLineWidth(2)
    c.line(padding, height - 0.7*inch, width - padding, height - 0.7*inch)
    
    # Reset color texto
    c.setFillColorRGB(0, 0, 0)
    
    padding = 0.25 * inch
    current_y = height - 0.9 * inch
    
    # Info Cliente
    c.setFont(RL_FONT_BOLD_NAME, 12)
    c.drawString(padding, current_y, "CLIENTE:")
    c.setFont(RL_FONT_REGULAR_NAME, 12)
    c.drawString(padding + 1.0*inch, current_y, cliente)
    
    current_y -= 0.3 * inch
    c.setFont(RL_FONT_BOLD_NAME, 12)
    c.drawString(padding, current_y, "FECHA:")
    c.setFont(RL_FONT_REGULAR_NAME, 12)
    c.drawString(padding + 1.0*inch, current_y, fecha)
    
    current_y -= 0.4 * inch
    c.setFont(RL_FONT_BOLD_NAME, 12)
    c.drawString(padding, current_y, "PROBLEMA REPORTADO:")
    
    # Caja gris para problema
    box_top = current_y - 0.1*inch
    box_height = 0.8 * inch
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(padding, box_top - box_height, width - 2*padding, box_height, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    
    # Texto Problema
    text_obj = c.beginText(padding + 0.1*inch, box_top - 0.2*inch)
    text_obj.setFont(RL_FONT_REGULAR_NAME, 10)
    
    # Wrap texto simple
    max_w = width - 2.5*padding
    words = problema.split()
    line = ""
    for word in words:
        if pdfmetrics.stringWidth(line + " " + word, RL_FONT_REGULAR_NAME, 10) < max_w:
            line += " " + word
        else:
            text_obj.textLine(line.strip())
            line = word
    text_obj.textLine(line.strip())
    c.drawText(text_obj)
    
    # Footer
    footer_y = 0.9 * inch
    c.setLineWidth(1)
    c.line(padding, footer_y, width - padding, footer_y)
    
    # IMEI
    c.setFont(RL_FONT_BOLD_NAME, 10)
    c.drawString(padding, footer_y - 0.2*inch, f"IMEI / SN: {imei_sn}")
    
    # Barcode
    if imei_sn and imei_sn.strip().lower() != 'n/a':
        try:
            bar_width = 1.2
            bc = code128.Code128(imei_sn, barHeight=0.35*inch, barWidth=bar_width)
            # Auto-scale
            available_w = width * 0.6
            if bc.width > available_w:
                scale = available_w / bc.width
                bc = code128.Code128(imei_sn, barHeight=0.35*inch, barWidth=bar_width*scale)
            
            bc.drawOn(c, padding, 0.15*inch)
        except Exception: pass
        
    # Box Listo (Checkbox Style)
    box_size = 0.3 * inch
    # Posicionar a la derecha
    box_x = width - 1.3 * inch
    box_y = 0.3 * inch
    
    c.setLineWidth(1.5)
    c.rect(box_x, box_y, box_size, box_size)
    
    c.setFont(RL_FONT_BOLD_NAME, 12)
    c.drawString(box_x + box_size + 0.1*inch, box_y + 0.1*inch, "¿Listo?")
    
    c.save()
    return temp_pdf_path

# --- Interfaz Gráfica PySide6 ---

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generador de Etiquetas de Garantía v3.0")
        if os.path.exists("micael.ico"):
            self.setWindowIcon(QIcon("micael.ico"))
        self.resize(800, 600)
        self.setMinimumSize(750, 550)
        
        # Cargar configuración
        cargar_config_inicial()
        cargar_fuentes_pdf()
        
        # Widget Central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # --- Panel de Control (Izquierda) ---
        self.controls_frame = QFrame()
        self.controls_frame.setObjectName("ControlsFrame")
        self.controls_frame.setFixedWidth(350)
        controls_layout = QVBoxLayout(self.controls_frame)
        controls_layout.setContentsMargins(20, 20, 20, 20)
        controls_layout.setSpacing(15)
        
        # Título
        title_label = QLabel("Datos de la Etiqueta")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: white; margin-bottom: 10px;")
        controls_layout.addWidget(title_label)
        
        # Campos
        self.cliente_input = self._create_input_field("Nombre del Cliente:", controls_layout)
        self.fecha_input = self._create_input_field("Fecha de Ingreso:", controls_layout, default=datetime.now().strftime("%d/%m/%Y"))
        self.imei_input = self._create_input_field("IMEI o S/N:", controls_layout)
        
        controls_layout.addWidget(QLabel("Problema del Equipo:"))
        self.problema_input = QTextEdit()
        self.problema_input.setPlaceholderText("Describa el problema...")
        self.problema_input.setFixedHeight(100)
        controls_layout.addWidget(self.problema_input)
        
        # Botones
        controls_layout.addSpacing(20)
        controls_layout.addStretch()
        
        self.print_btn = QPushButton("IMPRIMIR ETIQUETA")
        self.print_btn.setMinimumHeight(45)
        self.print_btn.setCursor(Qt.PointingHandCursor)
        self.print_btn.clicked.connect(self.imprimir)
        controls_layout.addWidget(self.print_btn)
        
        self.config_btn = QPushButton("Configurar SumatraPDF")
        self.config_btn.setObjectName("ConfigButton")
        self.config_btn.setCursor(Qt.PointingHandCursor)
        self.config_btn.clicked.connect(self.configurar_sumatra)
        controls_layout.addWidget(self.config_btn)
        
        footer_label = QLabel("v3.0 - PySide6 Edition")
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setStyleSheet("color: #666666; font-size: 12px; margin-top: 10px;")
        controls_layout.addWidget(footer_label)
        
        main_layout.addWidget(self.controls_frame)
        
        # --- Área de Previsualización (Derecha) ---
        preview_container = QWidget()
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setAlignment(Qt.AlignCenter)
        preview_layout.setContentsMargins(40, 40, 40, 40)
        
        preview_title = QLabel("Vista Previa")
        preview_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #e0e0e0; margin-bottom: 20px;")
        preview_layout.addWidget(preview_title, alignment=Qt.AlignCenter)
        
        # Label para la imagen
        self.preview_label = QLabel()
        self.preview_label.setObjectName("PreviewPlaceholder")
        self.preview_label.setFixedSize(PREVIEW_WIDTH, PREVIEW_HEIGHT)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setText("Generando vista previa...")
        
        # Efecto de sombra para la preview
        # (Nota: Las sombras complejas en Qt a veces requieren GraphicsEffect, 
        # pero un borde simple en el stylesheet funciona bien para este estilo flat)
        
        preview_layout.addWidget(self.preview_label)
        main_layout.addWidget(preview_container)
        
        # --- Eventos ---
        self.cliente_input.textChanged.connect(self.schedule_update)
        self.fecha_input.textChanged.connect(self.schedule_update)
        self.imei_input.textChanged.connect(self.schedule_update)
        self.problema_input.textChanged.connect(self.schedule_update)
        
        # Timer para debounce de preview
        self.update_timer = QTimer()
        self.update_timer.setSingleShot(True)
        self.update_timer.timeout.connect(self.update_preview)
        
        # Actualización inicial
        self.update_preview()

    def _create_input_field(self, label_text, layout, default=""):
        layout.addWidget(QLabel(label_text))
        inp = QLineEdit(default)
        layout.addWidget(inp)
        return inp

    def schedule_update(self):
        self.update_timer.start(300) # 300ms delay

    def update_preview(self):
        cliente = self.cliente_input.text().upper()
        fecha = self.fecha_input.text()
        imei = self.imei_input.text().upper()
        problema = self.problema_input.toPlainText().upper()
        
        try:
            # Generar imagen PIL
            pil_image = _generar_etiqueta_pil_image(cliente, fecha, imei, problema)
            
            # Convertir a QPixmap
            im_data = pil_image.convert("RGBA").tobytes("raw", "RGBA")
            qim = QImage(im_data, pil_image.size[0], pil_image.size[1], QImage.Format_RGBA8888)
            pixmap = QPixmap.fromImage(qim)
            
            # Escalar a la vista previa
            scaled_pixmap = pixmap.scaled(
                self.preview_label.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            
            self.preview_label.setPixmap(scaled_pixmap)
            self.preview_label.setText("") # Borrar texto placeholder
            
        except Exception as e:
            self.preview_label.setText(f"Error: {str(e)}")
            self.preview_label.setPixmap(QPixmap())

    def imprimir(self):
        if not PDF_SAVE_ENABLED:
            QMessageBox.critical(self, "Error", "Librería ReportLab no encontrada.")
            return

        cliente = self.cliente_input.text().strip()
        imei = self.imei_input.text().strip()
        
        if not cliente or not imei:
            QMessageBox.warning(self, "Datos Incompletos", "Nombre y IMEI/SN son obligatorios.")
            return
            
        temp_pdf = _generar_etiqueta_pdf_temporal(
            cliente.upper(),
            self.fecha_input.text(),
            imei.upper(),
            self.problema_input.toPlainText().upper()
        )
        
        if temp_pdf:
            self.enviar_a_impresora(temp_pdf)

    def enviar_a_impresora(self, filepath):
        current_os = platform.system()
        try:
            if current_os == "Windows":
                if SUMATRA_PDF_PATH and os.path.exists(SUMATRA_PDF_PATH):
                    subprocess.Popen([SUMATRA_PDF_PATH, "-print-to-default", "-silent", filepath])
                else:
                    QMessageBox.information(self, "Imprimiendo", "Se abrirá el diálogo de impresión predeterminado.")
                    os.startfile(filepath, "print")
            elif current_os in ["Darwin", "Linux"]:
                cmd = "lpr" if current_os == "Darwin" else "lp"
                subprocess.run([cmd, filepath], check=True)
            else:
                QMessageBox.warning(self, "Error", "Sistema operativo no soportado para impresión directa.")
        except Exception as e:
            QMessageBox.critical(self, "Error de Impresión", str(e))

    def configurar_sumatra(self):
        if platform.system() != "Windows":
            QMessageBox.information(self, "Info", "Solo disponible en Windows.")
            return
            
        fname, _ = QFileDialog.getOpenFileName(self, "Buscar SumatraPDF.exe", "", "Ejecutables (*.exe)")
        if fname:
            if os.path.basename(fname).lower() == "sumatrapdf.exe":
                global SUMATRA_PDF_PATH
                SUMATRA_PDF_PATH = fname
                guardar_config_sumatra()
                QMessageBox.information(self, "Éxito", "Ruta configurada correctamente.")
            else:
                QMessageBox.warning(self, "Error", "Por favor selecciona el archivo SumatraPDF.exe correcto.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())
