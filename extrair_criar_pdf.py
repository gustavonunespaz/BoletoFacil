import os
import re
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import tkinter as tk
from tkinter import filedialog

# Ajuste os caminhos conforme necessário para garantir que apontem para a instalação correta do Tcl/Tk
os.environ['TCL_LIBRARY'] = r'C:\Program Files\Python313\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Program Files\Python313\tcl\tk8.6'

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

        return nome_cliente, endereco_cliente

def criar_pdf_final(pdf_path, nome_cliente, endereco_cliente):
    # Crie a pasta "boletos" se não existir
    nova_pasta = "boletos"
    if not os.path.exists(nova_pasta):
        os.makedirs(nova_pasta)

    # Novo caminho para salvar o PDF na pasta "boletos"
    novo_pdf_path = os.path.join(nova_pasta, os.path.basename(pdf_path))

    # Carregar o PDF original usando PdfReader
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    # Criar a primeira página com os dados do cliente
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from io import BytesIO

    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.drawString(100, 700, nome_cliente)
    can.drawString(100, 680, endereco_cliente)
    can.save()
    packet.seek(0)

    # Ler a página recém-criada
    nova_pagina = PdfReader(packet).pages[0]
    writer.add_page(nova_pagina)

    # Adicionar a página original do PDF como página 2
    pagina_original = reader.pages[0]
    writer.add_page(pagina_original)

    # Salvar o novo PDF com as duas páginas
    with open(novo_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

    print(f"PDF gerado com sucesso em: {novo_pdf_path}")

# Testando a função
if __name__ == "__main__":
    caminho_pdf = selecionar_pdf()  # Permite ao usuário selecionar o arquivo PDF
    if caminho_pdf:
        nome_cliente, endereco_cliente = extrair_dados_cliente(caminho_pdf)
        if nome_cliente and endereco_cliente:
            criar_pdf_final(caminho_pdf, nome_cliente, endereco_cliente)
        else:
            print("Erro: Não foi possível extrair os dados do cliente.")
    else:
        print("Nenhum arquivo selecionado.")
