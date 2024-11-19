import fitz  # PyMuPDF
import pdfplumber
import reportlab
import pytesseract

# Configurar o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

print("Ambiente configurado corretamente!")
