import os
os.environ['TCL_LIBRARY'] = r'C:\Program Files\Python313\tcl\tcl8.6'
import pdfplumber
import tkinter as tk
from tkinter import filedialog

def upload_pdf():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    if not file_path:
        print("Nenhum arquivo selecionado.")
        return

    # Ler o PDF e extrair os dados do cliente
    with pdfplumber.open(file_path) as pdf:
        page = pdf.pages[0]  # Supondo que os dados estejam na primeira página
        text = page.extract_text()

        # Procura pelo nome e endereço (ajuste isso conforme necessário)
        lines = text.split('\n')
        for line in lines:
            if "Cliente:" in line or "Sacado/Cliente:" in line:
                client_data = line
                print(f"Dados do cliente encontrados: {client_data}")
                break
        else:
            print("Dados do cliente não encontrados.")

if __name__ == "__main__":
    upload_pdf()
