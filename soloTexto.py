import os
import mimetypes
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image

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

# Función para extraer texto de PDFs con texto estructurado
def extraer_texto_pdf(file_path):
    texto = ""
    reader = PdfReader(file_path)
    for page in reader.pages:
        texto += page.extract_text()
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
            texto += pytesseract.image_to_string(page, config='--psm 6')  # Configurar el modo PSM
    except Exception as e:
        print(f"Error al convertir el archivo con OCR: {e}")
    
    return texto

# Función para extraer texto de archivos Word
def extraer_texto_word(file_path):
    doc = Document(file_path)
    texto = '\n'.join([para.text for para in doc.paragraphs])
    return texto

# Función para extraer texto de imágenes
def extraer_texto_imagen(file_path):
    img = Image.open(file_path)
    texto = pytesseract.image_to_string(img)
    return texto

# Función para buscar todos los párrafos que contienen "Industria y Comercio"
def buscar_parrafos_industria_comercio(texto):
    parrafos = texto.split('\n')  # Dividir el texto en párrafos por saltos de línea
    parrafos_industria = [p for p in parrafos if 'industria y comercio' in p.lower()]
    return parrafos_industria

# Función principal que revisa la carpeta y procesa cada archivo
def procesar_archivos_carpeta(carpeta):
    for file_name in os.listdir(carpeta):
        file_path = os.path.join(carpeta, file_name)
        tipo = detectar_tipo_archivo(file_path)
        texto_extraido = ""

        if tipo == 'pdf':
            print(f"Procesando PDF: {file_name}")
            texto_extraido = extraer_texto_pdf(file_path)
            if not texto_extraido.strip():
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
        
        # Buscar todos los párrafos que contengan "Industria y Comercio"
        parrafos_industria = buscar_parrafos_industria_comercio(texto_extraido)
        print(f"\nPárrafos de '{file_name}' que contienen 'Industria y Comercio':\n")
        for parrafo in parrafos_industria:
            print(parrafo)
        print('-' * 50)

# Carpeta de ejemplo
carpeta = r'C:\Users\jeisson.gallego\OneDrive - SOPHOS SOLUTIONS SAS\CARPETA PERSONAL\PRACTICAS\BINTEC\Documentos\Consecucion norma\Pruebas'
procesar_archivos_carpeta(carpeta)
