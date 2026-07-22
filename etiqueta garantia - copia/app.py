from flask import Flask, render_template, request, jsonify, send_file
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import json
import os
import tempfile
import io
from datetime import datetime

app = Flask(__name__)

# Constantes
FONT_BOLD_PATH_TTF = "arialbd.ttf"
FONT_REGULAR_PATH_TTF = "arial.ttf"
LABEL_WIDTH_INCHES = 4
LABEL_HEIGHT_INCHES = 3
DPI = 300

# Registro de fuentes para PDF
try:
    pdfmetrics.registerFont(TTFont('ArialBold', FONT_BOLD_PATH_TTF))
    pdfmetrics.registerFont(TTFont('ArialRegular', FONT_REGULAR_PATH_TTF))
except:
    print("Advertencia: No se pudieron cargar las fuentes personalizadas")

def generate_label_image(cliente, fecha, imei_sn, problema):
    """Genera la etiqueta como una imagen PIL."""
    LABEL_WIDTH_PX = int(LABEL_WIDTH_INCHES * DPI)
    LABEL_HEIGHT_PX = int(LABEL_HEIGHT_INCHES * DPI)
    
    image = Image.new("RGB", (LABEL_WIDTH_PX, LABEL_HEIGHT_PX), "white")
    draw = ImageDraw.Draw(image)
    
    try:
        font_title = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=52)
        font_label = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=36)
        font_text = ImageFont.truetype(FONT_REGULAR_PATH_TTF, size=36)
        font_sn = ImageFont.truetype(FONT_BOLD_PATH_TTF, size=34)
    except IOError:
        font_title = font_label = font_text = font_sn = ImageFont.load_default()

    padding = 30
    
    # Título
    title_text = "RECIBO DE GARANTÍA"
    title_w = draw.textlength(title_text, font=font_title)
    draw.text(((LABEL_WIDTH_PX - title_w) / 2, padding), title_text, fill="black", font=font_title)
    draw.line([(padding, 100), (LABEL_WIDTH_PX - padding, 100)], fill="black", width=3)

    # Contenido
    y_pos = 130
    draw.text((padding, y_pos), "Cliente:", fill="black", font=font_label)
    draw.text((padding + 160, y_pos), cliente.upper(), fill="black", font=font_text)
    y_pos += 55
    draw.text((padding, y_pos), "Fecha:", fill="black", font=font_label)
    draw.text((padding + 160, y_pos), fecha, fill="black", font=font_text)
    y_pos += 55
    draw.text((padding, y_pos), "Problema:", fill="black", font=font_label)
    
    # Manejo de texto largo en problema
    lines = []
    if problema:
        max_chars = 45
        problema = problema.upper()
        lines = [problema[i:i+max_chars] for i in range(0, len(problema), max_chars)]
    for i, line in enumerate(lines):
        draw.text((padding, y_pos + 55 + (i * 45)), line, fill="black", font=font_text)

    # Footer
    footer_y = LABEL_HEIGHT_PX - 220
    draw.line([(padding, footer_y), (LABEL_WIDTH_PX - padding, footer_y)], fill="gray", width=2)
    
    # Área "Listo?"
    listo_area_width = 300
    listo_x = LABEL_WIDTH_PX - padding - listo_area_width
    draw.rectangle([(listo_x + 20, footer_y + 110), (listo_x + 70, footer_y + 160)], outline="black", width=4)
    draw.text((listo_x + 90, footer_y + 115), "Listo?", fill="black", font=font_label)

    # IMEI/SN
    left_footer_width = LABEL_WIDTH_PX - (padding * 2) - listo_area_width
    sn_text = f"IMEI / S/N: {imei_sn.upper()}"
    sn_w = draw.textlength(sn_text, font=font_sn)
    draw.text(((left_footer_width - sn_w) / 2 + padding, footer_y + 40), sn_text, fill="black", font=font_sn)
    
    return image

def generate_pdf(image):
    """Genera un PDF a partir de una imagen PIL."""
    pdf_buffer = io.BytesIO()
    pdf = canvas.Canvas(pdf_buffer, pagesize=landscape(letter))
    
    # Centrar la imagen en la página
    page_width, page_height = landscape(letter)
    x = (page_width - (LABEL_WIDTH_INCHES * inch)) / 2
    y = (page_height - (LABEL_HEIGHT_INCHES * inch)) / 2
    
    # Convertir la imagen PIL a formato que ReportLab pueda usar
    img_buffer = io.BytesIO()
    image.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    
    # Insertar la imagen en el PDF
    pdf.drawImage(img_buffer, x, y, width=LABEL_WIDTH_INCHES * inch, height=LABEL_HEIGHT_INCHES * inch)
    pdf.save()
    
    pdf_buffer.seek(0)
    return pdf_buffer

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/generate-label', methods=['POST'])
def generate_label():
    try:
        data = request.json
        cliente = data.get('cliente', '')
        fecha = data.get('fecha', datetime.now().strftime('%d/%m/%Y'))
        imei_sn = data.get('imei', '')
        problema = data.get('problema', '')
        
        # Generar la imagen de la etiqueta
        image = generate_label_image(cliente, fecha, imei_sn, problema)
        
        # Convertir la imagen a PDF
        pdf_buffer = generate_pdf(image)
        
        # Guardar temporalmente el PDF
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_pdf.write(pdf_buffer.getvalue())
        temp_pdf.close()
        
        return jsonify({
            'status': 'success',
            'message': 'Etiqueta generada correctamente',
            'pdf_path': temp_pdf.name
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/print', methods=['POST'])
def print_label():
    try:
        data = request.json
        pdf_path = data.get('pdf_path')
        
        if not pdf_path or not os.path.exists(pdf_path):
            raise ValueError("PDF no encontrado")

        # En Windows, usamos SumatraPDF para imprimir
        if os.name == 'nt':
            sumatra_path = config.get('sumatra_pdf_path')
            if not sumatra_path or not os.path.exists(sumatra_path):
                # Buscar SumatraPDF en ubicaciones comunes
                possible_paths = [
                    "SumatraPDF.exe",
                    os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
                    os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
                ]
                
                for path in possible_paths:
                    if os.path.exists(path):
                        sumatra_path = path
                        break
                
                if not sumatra_path:
                    raise ValueError("SumatraPDF no encontrado. Por favor instálelo para imprimir.")

            # Imprimir usando SumatraPDF
            import subprocess
            subprocess.run([sumatra_path, "-print-to-default", pdf_path], check=True)
            
            # Eliminar el archivo temporal después de imprimir
            os.remove(pdf_path)
            
            return jsonify({
                'status': 'success',
                'message': 'Etiqueta enviada a imprimir correctamente'
            })
        else:
            # Para otros sistemas operativos
            return jsonify({
                'status': 'error',
                'message': 'La impresión directa solo está soportada en Windows por ahora'
            }), 500
            
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    # Cargar configuración al inicio
    config = {}
    if os.path.exists('etiqueta_garantia_config.json'):
        with open('etiqueta_garantia_config.json', 'r') as f:
            config = json.load(f)
    
    app.run(host='0.0.0.0', port=5000, debug=True)
