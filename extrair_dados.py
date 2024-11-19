import os
import re

# Ajuste os caminhos conforme necessário para garantir que apontem para a instalação correta do Tcl/Tk
os.environ['TCL_LIBRARY'] = r'C:\Program Files\Python313\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Program Files\Python313\tcl\tk8.6'

import pdfplumber
import tkinter as tk
from tkinter import filedialog

def selecionar_pdf():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal do tkinter
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o PDF do boleto",
        filetypes=[("PDF files", "*.pdf")]
    )
    return caminho_arquivo

def extrair_dados_cliente(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        primeira_pagina = pdf.pages[0]
        texto = primeira_pagina.extract_text()

        # Quebra o texto em linhas
        linhas = texto.split('\n')

        # Variáveis para armazenar nome e endereço
        nome_cliente = None
        endereco_cliente = None

        # Itera pelas linhas para encontrar o bloco com os dados do cliente
        for i, linha in enumerate(linhas):
            # Verifica se a linha contém "Sacado/Cliente"
            if "Sacado/Cliente" in linha:
                # Verifica a linha seguinte para o nome do cliente
                if i + 1 < len(linhas):
                    possivel_nome = linhas[i + 1].strip()
                    # Remover valores monetários se houver
                    nome_cliente = re.sub(r'R\$\s*\d+([.,]\d{2})?', '', possivel_nome).strip()
                
                # Verifica a próxima linha para o endereço do cliente
                if i + 2 < len(linhas):
                    endereco_cliente = linhas[i + 2].strip()
                break

        # Exibe os resultados encontrados
        if nome_cliente:
            print(f"{nome_cliente}")
        else:
            print("Nome do Cliente: Não encontrado")

        if endereco_cliente:
            print(f"{endereco_cliente}")
        else:
            print("Endereço do Cliente: Não encontrado")

# Testando a função
if __name__ == "__main__":
    caminho_pdf = selecionar_pdf()  # Permite ao usuário selecionar o arquivo PDF
    if caminho_pdf:
        extrair_dados_cliente(caminho_pdf)
    else:
        print("Nenhum arquivo selecionado.")
