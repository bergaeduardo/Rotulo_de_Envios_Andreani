# Rotulo_de_Envios_Andreani

## Descripción

Este proyecto automatiza el proceso de impresión de etiquetas de envío de Andreani. Incluye scripts para web scraping, normalización de datos y gestión de impresoras.

## Estructura

- `config/`: Archivos de configuración.
- `logs/`: Archivos de registro.
- `src/`: Código fuente de los scripts de Python.
- `Documents/`: Documentos de Excel utilizados por los scripts.

## Scripts

- `main.py`: Script principal para ejecutar la automatización de Andreani.
- `excel_manager.py`: Módulo para manejar operaciones de archivos Excel.
- `andreani_automation.py`: Módulo para automatizar las interacciones con el sitio web de Andreani.
- `printer_manager.py`: Módulo para gestionar la configuración de impresoras.
- `configuracion.py`: Módulo para cargar configuraciones.
- `logger_config.py`: Módulo para configurar el logging.

## Configuración

La configuración se gestiona a través de `config/config.ini`. Actualice este archivo con su configuración específica.

## Requisitos

- Python 3
- Librerías listadas en `requirements.txt` (aunque pip no está disponible en este entorno, estas son las dependencias)
- ChromeDriver para Selenium

pip install -r requirements.txt

## Ejecutar

Ejecute `RunEtiquetasAndreani.bat` para iniciar el script.

---
Este README proporciona una descripción general básica del proyecto.
