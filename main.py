# C:\Users\marco.osorio\AppData\Local\Programs\Tesseract-OCR
from PIL import Image
from pypdf import PdfReader
from pytesseract import pytesseract
import pymupdf
import codecs

# tesseract = r"C:\Users\marco.osorio\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
# pytesseract.tesseract_cmd = tesseract
# # Abre una imagen sencilla.
# # image = r"C:\Users\marco.osorio\Downloads\image002.png"
# # img = Image.open(image)
# # imageText = pytesseract.image_to_string(img)
# # print(imageText[:-1])

# PDF de imagenes.
pdf_path = r"C:\Users\marco.osorio\Downloads\Medellin Acuerdo.pdf"
pdf = pymupdf.open(pdf_path)

text=""
for page in pdf:
    text += page.get_text("text")

with codecs.open("metodo2.txt", "w","utf-8") as file:
    # Write some text to the file
    file.write(text)
    
file.close()
pdf.close()