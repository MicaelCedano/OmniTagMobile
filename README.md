# OmniTag Mobile 📱🏷️

**OmniTag Mobile** es una solución integral para la lectura automática de información técnica de teléfonos móviles (iPhone, Samsung Galaxy, Google Pixel y otros Android) para la generación de etiquetas térmicas/PDF y registro automático en plantillas de Excel para procesos de RMA, garantías y compras.

## ✨ Características Principales

- **Detección Multi-Plataforma**:
  - **iOS (iPhone)**: Detección automática por USB mux con `pymobiledevice3` (IMEI, Modelo, Capacidad, Batería).
  - **Android (Samsung, Pixel, Xiaomi, etc.)**: Lectura mediante ADB con `adbutils` (Serie, Modelo, IMEI, Capacidad, Batería).
- **Generación de Etiquetas**:
  - Vista previa interactiva con `CustomTkinter`.
  - Generación de PDF e impresión térmica con código de barras en formato Code128.
- **Exportación a Excel**:
  - Inserción y actualización automática en reportes Excel.

## 🛠️ Requisitos e Instalación

### Requisitos Previos (Python)
```bash
pip install custom-tkinter pillow reportlab pandas openpyxl barcode pymobiledevice3 adbutils
```

### Configuración para Dispositivos
1. **iPhone**: Conectar por USB y aceptar "Confiar en este equipo".
2. **Android**: Activar *Opciones de Desarrollador* -> **Depuración por USB** y aceptar el permiso en la pantalla del dispositivo.

## 🚀 Ejecución

Para iniciar la aplicación:
```bash
python etiqueta_iphone_2025.py
```
O ejecutando el acceso directo:
```cmd
Iniciar_App.bat
```