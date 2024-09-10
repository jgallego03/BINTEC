import os
import mimetypes
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image, ImageEnhance
import pdfplumber
import re
from dateutil import parser

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

# Función para extraer texto de PDFs estructurados con pdfplumber
def extraer_tablas_pdf(file_path):
    texto = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            texto += page.extract_text() + "\n"
    
    return texto

# Función para extraer texto de PDFs escaneados o imágenes usando OCR
def extraer_texto_ocr(file_path):
    texto = ""
    poppler_path = r'C:\poppler-23.05.0\Library\bin'

    try:
        print(f"Usando Poppler desde: {poppler_path}")
        pages = convert_from_path(file_path, poppler_path=poppler_path)
        for page in pages:
            page = preprocesar_imagen(page)
            texto += pytesseract.image_to_string(page, config='--psm 6')
    except Exception as e:
        print(f"Error al convertir el archivo con OCR: {e}")
    
    return texto

# Función de preprocesamiento de imagen para mejorar el contraste en OCR
def preprocesar_imagen(imagen):
    enhancer = ImageEnhance.Contrast(imagen)
    imagen_mejorada = enhancer.enhance(2)
    return imagen_mejorada

# Función para extraer texto de archivos Word
def extraer_texto_word(file_path):
    doc = Document(file_path)
    texto = '\n'.join([para.text for para in doc.paragraphs])
    return texto

# Función para extraer texto de imágenes
def extraer_texto_imagen(file_path):
    img = Image.open(file_path)
    texto = pytesseract.image_to_string(img, config='--psm 6')
    return texto

# Función para extraer las fechas de un texto dado
def extraer_fechas(texto):
    fechas = []
    patrones_fechas = [
        r'\d{1,2}\sde\s[a-zA-Z]+\sde\s\d{4}',  # Ej: "01 de abril de 2024"
        r'\d{1,2}/\d{1,2}/\d{4}',              # Ej: "01/04/2024"
        r'hasta el \d{1,2}\sde\s[a-zA-Z]+',    # Ej: "hasta el 10 de abril"
        r'\d{1,2}\s[a-zA-Z]+'                  # Ej: "01 abril"
    ]
    
    for patron in patrones_fechas:
        fechas.extend(re.findall(patron, texto))
    
    fechas_convertidas = []
    for fecha in fechas:
        try:
            fecha_convertida = parser.parse(fecha, fuzzy=True)
            # Filtrar fechas para mantener solo las relevantes (después de 2020)
            if fecha_convertida.year >= 2020:
                fechas_convertidas.append(fecha_convertida)
        except:
            pass
    
    return fechas_convertidas

# Función para procesar y extraer el NIT
def extraer_nit(texto, nit):
    ultimos_dos_digitos = nit.split('-')[0][-2:]
    ultimo_digito = nit[-1]
    patron_nit_dos_digitos = re.compile(rf"(\d{{2}})\s–\s(\d{{2}})\s.*{ultimos_dos_digitos}")
    patron_nit_un_digito = re.compile(rf"\b{ultimo_digito}\b")

    coincidencia = patron_nit_dos_digitos.search(texto)
    if coincidencia:
        return f"Coincidencia encontrada para los dos últimos dígitos del NIT: {coincidencia.group()}"
    
    coincidencia = patron_nit_un_digito.search(texto)
    if coincidencia:
        return f"Coincidencia encontrada para el último dígito del NIT: {coincidencia.group()}"
    
    return "No se encontró coincidencia para el NIT."

# Función principal que revisa la carpeta y procesa cada archivo
def procesar_archivos_carpeta(carpeta, nit):
    for file_name in os.listdir(carpeta):
        file_path = os.path.join(carpeta, file_name)
        tipo = detectar_tipo_archivo(file_path)
        texto_extraido = ""

        if tipo == 'pdf':
            print(f"Procesando PDF: {file_name}")
            texto_extraido = extraer_tablas_pdf(file_path)
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
        
        fechas_encontradas = extraer_fechas(texto_extraido)
        print(f"Fechas encontradas en '{file_name}': {fechas_encontradas}")

        nit_info = extraer_nit(texto_extraido, nit)
        print(f"Información del NIT en '{file_name}': {nit_info}")
        print('-' * 50)

# Carpeta de ejemplo y NIT proporcionado
carpeta = r'C:\Users\jeisson.gallego\OneDrive - SOPHOS SOLUTIONS SAS\CARPETA PERSONAL\PRACTICAS\BINTEC\Documentos\Consecucion norma\Pruebas'
nit = "123456789-1"
procesar_archivos_carpeta(carpeta, nit)
