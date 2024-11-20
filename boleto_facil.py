import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import fitz  # PyMuPDF
import win32api

# Função para imprimir o PDF compilado
def imprimir_pdf(pdf_path):
    try:
        win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)
        print(f"PDF compilado enviado para impressão: {pdf_path}")
    except Exception as e:
        print(f"Erro ao imprimir {pdf_path}: {e}")

# Função para selecionar e processar arquivos PDF
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivos:
        for arquivo in arquivos:
            novo_pdf_path = processar_pdf(arquivo)
            if novo_pdf_path:
                adicionar_item_lista(novo_pdf_path)
        messagebox.showinfo("Processamento Concluído", "Todos os PDFs foram processados.")
    else:
        messagebox.showinfo("Nenhum arquivo selecionado", "Selecione pelo menos um arquivo PDF para processar.")

# Função para adicionar cada PDF processado à lista com uma caixa de seleção
def adicionar_item_lista(pdf_path):
    var = tk.BooleanVar()  # Variável associada à caixa de seleção
    checkbox = tk.Checkbutton(frame_lista, text=os.path.basename(pdf_path), variable=var, bg="#f0f0f0", fg="#000000", selectcolor="#d9d9d9", font=("Calibri", 11), wraplength=700, anchor='w', justify='left')
    checkbox.var = var
    checkbox.pdf_path = pdf_path
    checkbox.bind("<Button-3>", lambda e: abrir_pdf(pdf_path))  # Clique com botão direito para abrir PDF
    checkbox.pack(anchor='w')

# Função para abrir PDF ao clicar com botão direito
def abrir_pdf(pdf_path):
    try:
        os.startfile(pdf_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o PDF: {e}")

# Função para selecionar todos os itens
def selecionar_todos():
    for widget in frame_lista.winfo_children():
        if isinstance(widget, tk.Checkbutton):
            widget.var.set(True)

# Função para imprimir os PDFs selecionados
def imprimir_selecionados():
    # Criar um PDF temporário com todos os PDFs selecionados
    writer = PdfWriter()
    for widget in frame_lista.winfo_children():
        if isinstance(widget, tk.Checkbutton) and widget.var.get():
            reader = PdfReader(widget.pdf_path)
            for page in reader.pages:
                writer.add_page(page)

    # Salvar o PDF compilado na pasta "Compilados"
    data_compilacao = datetime.now()
    ano_compilacao = data_compilacao.strftime("%Y")
    meses_em_portugues = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_compilacao_nome = meses_em_portugues[data_compilacao.month - 1]
    pasta_compilados = os.path.join("Compilados", ano_compilacao, mes_compilacao_nome)
    if not os.path.exists(pasta_compilados):
        os.makedirs(pasta_compilados)

    numero_arquivos = len(os.listdir(pasta_compilados)) + 1
    temp_pdf_path = os.path.join(pasta_compilados, f"{numero_arquivos} - {data_compilacao.strftime('%d-%m-%Y')}.pdf")
    with open(temp_pdf_path, "wb") as temp_pdf:
        writer.write(temp_pdf)

    # Abrir a janela de impressão para o PDF compilado
    try:
        os.startfile(temp_pdf_path, "print")
        print(f"PDF compilado enviado para impressão: {temp_pdf_path}")
    except Exception as e:
        print(f"Erro ao imprimir {temp_pdf_path}: {e}")
    messagebox.showinfo("Impressão", "PDF(s) selecionado(s) enviado(s) para impressão.")

# Função para processar os PDFs e extrair dados
def processar_pdf(pdf_path):
    nome_cliente, endereco_cliente, venda_parcela, data_vencimento = extrair_dados_cliente_e_instrucoes(pdf_path)
    if nome_cliente and endereco_cliente and data_vencimento:
        return criar_pdf_final(pdf_path, nome_cliente, endereco_cliente, venda_parcela, data_vencimento)
    return None

# Função de extração de dados e criação do PDF final
def extrair_dados_cliente_e_instrucoes(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        primeira_pagina = pdf.pages[0]
        texto = primeira_pagina.extract_text()

        linhas = texto.split('\n')
        nome_cliente, endereco_cliente, venda_parcela, data_vencimento = None, None, None, None

        for i, linha in enumerate(linhas):
            if "Sacado/Cliente" in linha:
                if i + 1 < len(linhas):
                    possivel_nome = linhas[i + 1].strip()
                    nome_cliente = re.sub(r'R\$\s*\d+([.,]\d{2})?', '', possivel_nome).strip()
                if i + 2 < len(linhas):
                    endereco_cliente = linhas[i + 2].strip()
            if "Instruções" in linha and i + 1 < len(linhas):
                venda_parcela = linhas[i + 1].strip()
            if "Vencimento" in linha and i + 1 < len(linhas):
                data_vencimento = re.sub(r'[^0-9/]', '', linhas[i + 1].strip())

        return nome_cliente, endereco_cliente, venda_parcela, data_vencimento

# Função para criar o PDF final formatado e organizado nas pastas
def criar_pdf_final(pdf_path, nome_cliente, endereco_cliente, venda_parcela, data_vencimento):
    try:
        data_obj = datetime.strptime(data_vencimento, "%d/%m/%Y")
        ano_vencimento = data_obj.strftime("%Y")
        meses_em_portugues = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes_vencimento_nome = meses_em_portugues[data_obj.month - 1]
        data_formatada = data_obj.strftime("%d-%m-%Y")
    except ValueError:
        print("Erro: Formato de data de vencimento inválido.")
        return None

    nova_pasta = os.path.join("boletos", ano_vencimento, mes_vencimento_nome)
    if not os.path.exists(nova_pasta):
        os.makedirs(nova_pasta)

    novo_nome_arquivo = f"{data_formatada} - {nome_cliente}"
    if venda_parcela:
        novo_nome_arquivo += f" - {venda_parcela.replace('/', '-')}"
    novo_pdf_path = os.path.join(nova_pasta, f"{novo_nome_arquivo}.pdf")

    documento = fitz.open(pdf_path)
    x0, y0, x1, y1 = 30, 710, 550, 810

    for page_number in range(len(documento)):
        pagina = documento.load_page(page_number)
        rect = fitz.Rect(x0, y0, x1, y1)
        pagina.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

    documento.save("temp_modificado.pdf")
    documento.close()

    reader_modificado = PdfReader("temp_modificado.pdf")
    writer = PdfWriter()

    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from io import BytesIO

    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.drawString(100, 700, nome_cliente)
    can.drawString(100, 680, endereco_cliente)
    can.save()
    packet.seek(0)

    nova_pagina = PdfReader(packet).pages[0]
    writer.add_page(nova_pagina)
    writer.add_page(reader_modificado.pages[0])

    with open(novo_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

    os.remove("temp_modificado.pdf")
    return novo_pdf_path

# Função para abrir a pasta de boletos gerados
def abrir_boletos():
    os.startfile("boletos")

# Função para abrir uma janela com todas as funcionalidades do sistema
def abrir_ajuda():
    nova_janela = tk.Toplevel(janela)
    nova_janela.title("Ajuda - Funcionalidades do Sistema")
    nova_janela.geometry("600x400")
    nova_janela.configure(bg="#f0f0f0")

    texto_ajuda = (
        "Passo a passo para usar o sistema:\n"
        "1. Carregar PDF: Selecione os arquivos PDF que deseja processar.\n"
        "2. Selecionar Todos: Marque todos os PDFs carregados.\n"
        "3. Imprimir: Compile os PDFs selecionados em um único arquivo e envie para impressão.\n"
        "4. Boletos: Abra a pasta onde os boletos processados estão salvos.\n\n"
        "Funcionalidades adicionais:\n"
        "- Clique com o botão direito em um PDF na lista para abri-lo diretamente."
    )

    label_ajuda = tk.Label(nova_janela, text=texto_ajuda, font=("Calibri", 12), fg="#000000", bg="#f0f0f0", justify="left", wraplength=580)
    label_ajuda.pack(pady=20, padx=20)

# Função para excluir os PDFs selecionados
def excluir_selecionados():
    for widget in frame_lista.winfo_children():
        if isinstance(widget, tk.Checkbutton) and widget.var.get():
            try:
                os.remove(widget.pdf_path)
                widget.destroy()
                print(f"PDF excluído: {widget.pdf_path}")
            except Exception as e:
                print(f"Erro ao excluir {widget.pdf_path}: {e}")
    messagebox.showinfo("Exclusão", "PDF(s) selecionado(s) excluído(s).")

# Configuração da interface principal
janela = tk.Tk()
janela.title("Boleto Fácil - Nunes|Paz")
janela.geometry("800x600")
janela.configure(bg="#f0f0f0")

# Frame dos botões organizados lado a lado
frame_botoes = tk.Frame(janela, bg="#f0f0f0")
frame_botoes.pack(pady=10)

btn_selecionar = tk.Button(frame_botoes, text="Carregar PDF", command=selecionar_arquivos, font=("Calibri", 12), bg="#d9d9d9", fg="#000000")
btn_selecionar.grid(row=0, column=0, padx=5)

btn_selecionar_todos = tk.Button(frame_botoes, text="Selecionar Todos", command=selecionar_todos, font=("Calibri", 12), bg="#d9d9d9", fg="#000000")
btn_selecionar_todos.grid(row=0, column=1, padx=5)

btn_imprimir = tk.Button(frame_botoes, text="Imprimir", command=imprimir_selecionados, font=("Calibri", 12), bg="#d9d9d9", fg="#000000")
btn_imprimir.grid(row=0, column=2, padx=5)

btn_excluir = tk.Button(frame_botoes, text="Excluir", command=excluir_selecionados, font=("Calibri", 12), bg="#d9d9d9", fg="#000000")
btn_excluir.grid(row=0, column=3, padx=5)

btn_boletos = tk.Button(frame_botoes, text="Boletos", command=abrir_boletos, font=("Calibri", 12), bg="#d9d9d9", fg="#000000")
btn_boletos.grid(row=0, column=4, padx=5)

btn_ajuda = tk.Button(frame_botoes, text="Ajuda", command=abrir_ajuda, font=("Calibri", 12), bg="#d9d9d9", fg="#000000")
btn_ajuda.grid(row=0, column=5, padx=5)

frame_lista = tk.Frame(janela, bg="#f0f0f0", bd=1, relief="sunken")
frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

# Observação: Clique com o botão direito do mouse em um arquivo PDF para abri-lo.
label_observacao = tk.Label(janela, text="Observação: Clique com o botão direito do mouse em um arquivo PDF para abri-lo.", font=("Calibri", 10), fg="#000000", bg="#f0f0f0", wraplength=580, justify='center')
label_observacao.pack(pady=5)

janela.mainloop()
