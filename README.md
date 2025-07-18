Transcriptor de PDF a Word con Detección de Actas y OCR
Este script en Python automatiza la conversión de documentos PDF a formato Word (.docx), con una especial atención a la gestión de actas.
Es capaz de:
 * Detectar y separar actas individuales dentro de un PDF extenso.
 * Procesar PDFs digitales (texto seleccionable) extrayendo su contenido y tablas.
 * Manejar PDFs escaneados utilizando OCR (Reconocimiento Óptico de Caracteres) para extraer el texto.
 * Guardar cada acta detectada como un documento Word editable separado.
Ideal para automatizar la gestión de grandes volúmenes de documentos administrativos y legales, facilitando la edición y organización de la información.
Tecnologías Clave
 * Python 3.x
 * PyMuPDF / pdfplumber: Para manipulación y extracción de datos de PDF.
 * pytesseract: Para la funcionalidad de OCR en PDFs escaneados.
 * python-docx: Para la generación de documentos Word.
Cómo Usarlo (Configuración Rápida)
 * Asegúrate de tener Python 3.x instalado.
 * Instala Tesseract OCR y Poppler (necesarios para el OCR y la conversión de imagen).
 * Instala las librerías Python requeridas:
   pip install PyMuPDF pdf2image Pillow pytesseract python-docx pdfplumber
 * En el script, actualiza las rutas a tus carpetas de entrada y salida, y la ruta a tu tesseract.exe.
   Configura tus rutas aquí:
   pytesseract.pytesseract.tesseract_cmd = r"C:\ruta\a\tu\tesseract.exe"
   CARPETA_ORIGEN_PDFS = r"C:\ruta\a\tus\pdfs"
   CARPETA_DESTINO_ACTAS_WORD = r"C:\ruta\de\salida\word"
 * Ejecuta el script: python tu_script_nombre.py
