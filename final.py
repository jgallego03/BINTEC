import os
import mimetypes
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import fitz  # PyMuPDF

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
    # Si no se detecta por MIME type, mirar la extensión
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
    # Si es PDF escaneado, convertimos cada página a imagen
    pages = convert_from_path(file_path)
    for page in pages:
        texto += pytesseract.image_to_string(page)
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

# Función principal que revisa la carpeta y procesa cada archivo
def procesar_archivos_carpeta(carpeta):
    for file_name in os.listdir(carpeta):
        file_path = os.path.join(carpeta, file_name)
        tipo = detectar_tipo_archivo(file_path)
        texto_extraido = ""

        if tipo == 'pdf':
            # Intentar extraer texto usando PyPDF2
            texto_extraido = extraer_texto_pdf(file_path)
            if not texto_extraido.strip():
                # Si no hay texto estructurado, usar OCR
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
        
        # Mostrar el texto extraído
        print(f"\nTexto extraído de {file_name}:\n")
        print(texto_extraido)
        print('-' * 50)

# Carpeta de ejemplo
carpeta = r'C:\Users\jeisson.gallego\OneDrive - SOPHOS SOLUTIONS SAS\CARPETA PERSONAL\PRACTICAS\BINTEC\Documentos\Consecucion norma\Pruebas'
procesar_archivos_carpeta(carpeta)