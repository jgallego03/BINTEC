import os
import mimetypes
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image, ImageEnhance
import pdfplumber

# Especificar la ruta a Tesseract directamente
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Función para revisar archivos en una carpeta y clasificar el tipo
def detectar_tipo_archivo(file_path):
    mime_type, _ = mimetypes.guess_type(file_path)
    if mime_type:
        if 'pdf' in mime_type:
            return 'pdf'
        elif 'word' in mime_type:
            return 'word'
        elif 'image' in mime_type:
            return 'image'
    ext = os.path.splitext(file_path)[-1].lower()
    if ext in ['.docx']:
        return 'word'
    elif ext in ['.pdf']:
        return 'pdf'
    elif ext in ['.jpg', '.jpeg', '.png', '.tiff']:
        return 'image'
    return 'unknown'

# Función para extraer texto y tablas de PDFs estructurados con pdfplumber
def extraer_tablas_pdf(file_path):
    texto = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            # Extraer texto de toda la página
            texto += page.extract_text()

            # Extraer tablas
            tablas = page.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    # Unir los elementos de la fila con tabulaciones para mantener la estructura
                    texto += "\t".join([str(celda) for celda in fila]) + "\n"
    return texto

# Función para extraer texto de PDFs escaneados o imágenes usando OCR
def extraer_texto_ocr(file_path):
    texto = ""

    # Especificar la ruta a Poppler
    poppler_path = r'C:\poppler-23.05.0\Library\bin'

    # Convertir cada página del PDF a una imagen usando la ruta de Poppler
    try:
        print(f"Usando Poppler desde: {poppler_path}")
        pages = convert_from_path(file_path, poppler_path=poppler_path)

        # Usar OCR en cada página convertida, con más control sobre la segmentación
        for page in pages:
            # Preprocesar la imagen (opcional, si la calidad es baja)
            page = preprocesar_imagen(page)

            # Extraer texto con Tesseract usando un modo PSM adecuado para tablas
            texto += pytesseract.image_to_string(page, config='--psm 6')  # Usar PSM 6 para tablas
    except Exception as e:
        print(f"Error al convertir el archivo con OCR: {e}")
    
    return texto

# Función de preprocesamiento de imagen para mejorar el contraste en OCR (opcional)
def preprocesar_imagen(imagen):
    # Mejorar el contraste de la imagen para que las tablas sean más visibles
    enhancer = ImageEnhance.Contrast(imagen)
    imagen_mejorada = enhancer.enhance(2)  # Aumentar el contraste
    return imagen_mejorada

# Función para extraer texto de archivos Word
def extraer_texto_word(file_path):
    doc = Document(file_path)
    texto = '\n'.join([para.text for para in doc.paragraphs])
    return texto

# Función para extraer texto de imágenes
def extraer_texto_imagen(file_path):
    img = Image.open(file_path)
    texto = pytesseract.image_to_string(img, config='--psm 6')  # Usar PSM 6 para manejar imágenes con tablas
    return texto

# Función principal que revisa la carpeta y procesa cada archivo
def procesar_archivos_carpeta(carpeta):
    for file_name in os.listdir(carpeta):
        file_path = os.path.join(carpeta, file_name)
        tipo = detectar_tipo_archivo(file_path)
        texto_extraido = ""

        if tipo == 'pdf':
            print(f"Procesando PDF: {file_name}")
            # Intentar extraer texto de PDFs estructurados
            texto_extraido = extraer_tablas_pdf(file_path)
            if not texto_extraido.strip():  # Si no se extrae texto, se asume que es PDF escaneado
                print(f"El PDF {file_name} parece estar escaneado, aplicando OCR...")
                texto_extraido = extraer_texto_ocr(file_path)
        
        elif tipo == 'word':
            print(f"Procesando archivo Word: {file_name}")
            texto_extraido = extraer_texto_word(file_path)
        
        elif tipo == 'image':
            print(f"Procesando imagen: {file_name}")
            texto_extraido = extraer_texto_imagen(file_path)

        else:
            print(f"Tipo de archivo desconocido o no soportado: {file_name}")
            continue
        
        # Imprimir todo el texto extraído
        print(f"\nTexto completo extraído de '{file_name}':\n")
        print(texto_extraido)
        print('-' * 50)

# Carpeta de ejemplo
carpeta = r'C:\Users\jeisson.gallego\OneDrive - SOPHOS SOLUTIONS SAS\CARPETA PERSONAL\PRACTICAS\BINTEC\Documentos\Consecucion norma\Pruebas'
procesar_archivos_carpeta(carpeta)
