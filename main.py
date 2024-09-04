# C:\Users\marco.osorio\AppData\Local\Programs\Tesseract-OCR
from PIL import Image
from pypdf import PdfReader
from pytesseract import pytesseract
import fitz

tesseract = r"C:\Users\marco.osorio\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pytesseract.tesseract_cmd = tesseract
# Abre una imagen sencilla.
image = r"C:\Users\marco.osorio\Downloads\image002.png"
img = Image.open(image)
imageText = pytesseract.image_to_string(img)
print(imageText[:-1])

# PDF de imagenes.
pdf = r"C:\Users\marco.osorio\Downloads\Medellin Acuerdo.pdf"

#Método 1
# reader = PdfReader(pdf)
# # número de páginas en el pdf
# # print(len(reader.pages))
# page = reader.pages[5]
# pdfText = page.extract_text()
# print(pdfText)

#Método 2
reader = fitz.open(pdf)