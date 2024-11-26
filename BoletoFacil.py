import os
import re
import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import fitz  # PyMuPDF
import win32api
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO

# Função para imprimir o PDF compilado
def imprimir_pdf(pdf_path):
    """
    Envia o PDF compilado para impressão.
    """
    try:
        win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)
        print(f"PDF compilado enviado para impressão: {pdf_path}")
    except Exception as e:
        print(f"Erro ao imprimir {pdf_path}: {e}")

# Função para selecionar e processar arquivos PDF
def selecionar_arquivos():
    """
    Abre uma janela de seleção para escolher arquivos PDF, processa cada arquivo e adiciona à lista.
    """
    arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivos:
        for arquivo in arquivos:
            novo_pdf_path = processar_pdf(arquivo)
            if novo_pdf_path:
                adicionar_item_lista(novo_pdf_path)
        messagebox.showinfo("Processamento Concluído", "Todos os PDFs foram processados.")
    else:
        messagebox.showinfo("Nenhum arquivo selecionado", "Selecione pelo menos um arquivo PDF para processar.")

# Função para processar os PDFs e extrair dados
def processar_pdf(pdf_path):
    """
    Extrai os dados do cliente e cria um novo PDF formatado.
    """
    nome_cliente, endereco_cliente, venda_parcela, data_vencimento = extrair_dados_cliente_e_instrucoes(pdf_path)
    if nome_cliente and endereco_cliente and data_vencimento:
        return criar_pdf_final(pdf_path, nome_cliente, endereco_cliente, venda_parcela, data_vencimento)
    return None

# Função para extrair dados do PDF
def extrair_dados_cliente_e_instrucoes(pdf_path):
    """
    Extrai os dados do cliente, como nome, endereço e data de vencimento, a partir do PDF.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            primeira_pagina = pdf.pages[0]
            texto = primeira_pagina.extract_text()

            if not texto:
                print(f"Aviso: Nenhum texto extraído da primeira página do arquivo {pdf_path}")
                return None, None, None, None

            linhas = texto.split('\n')
            nome_cliente, endereco_cliente, venda_parcela, data_vencimento = None, None, None, None

            for i, linha in enumerate(linhas):
                if "Sacado/Cliente" in linha:
                    # Extrai o nome do cliente
                    if i + 1 < len(linhas):
                        possivel_nome = linhas[i + 1].strip()
                        nome_cliente = re.sub(r'R\$\s*\d+([.,]\d{2})?', '', possivel_nome).strip()
                    # Extrai o endereço do cliente
                    if i + 2 < len(linhas):
                        endereco_cliente = linhas[i + 2].strip()
                        # Caso o endereço seja dividido em múltiplas linhas, concatenar a próxima linha se necessário
                        if endereco_cliente.endswith('-') and (i + 3 < len(linhas)):
                            endereco_cliente = endereco_cliente[:-1].strip() + ' ' + linhas[i + 3].strip()
                # Extrai a parcela da venda
                if "Instruções" in linha and i + 1 < len(linhas):
                    venda_parcela = linhas[i + 1].strip()
                # Extrai a data de vencimento
                if "Vencimento" in linha and i + 1 < len(linhas):
                    data_vencimento = re.sub(r'[^0-9/]', '', linhas[i + 1].strip())

            print(f"Dados extraídos do arquivo {pdf_path}: Nome: {nome_cliente}, Endereço: {endereco_cliente}, Vencimento: {data_vencimento}")
            return nome_cliente, endereco_cliente, venda_parcela, data_vencimento
    except Exception as e:
        print(f"Erro ao extrair dados do arquivo {pdf_path}: {e}")
        return None, None, None, None

# Função para criar o PDF final formatado e organizado nas pastas
def criar_pdf_final(pdf_path, nome_cliente, endereco_cliente, venda_parcela, data_vencimento):
    """
    Cria o PDF final formatado e organiza-o em pastas com base na data de vencimento.
    """
    try:
        # Formata a data para organizar o PDF em pastas
        data_obj = datetime.strptime(data_vencimento, "%d/%m/%Y")
        ano_vencimento = data_obj.strftime("%Y")
        meses_em_portugues = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes_vencimento_nome = meses_em_portugues[data_obj.month - 1]
        data_formatada = data_obj.strftime("%d-%m-%Y")
    except ValueError:
        print("Erro: Formato de data de vencimento inválido.")
        return None

    # Cria a pasta para armazenar o PDF
    nova_pasta = os.path.join("Documentos", "Boletos", ano_vencimento, mes_vencimento_nome)
    if not os.path.exists(nova_pasta):
        os.makedirs(nova_pasta)

    # Define o nome do novo arquivo PDF
    novo_nome_arquivo = f"{data_formatada} - {nome_cliente}"
    if venda_parcela:
        novo_nome_arquivo += f" - {venda_parcela.replace('/', '-')}"
    novo_pdf_path = os.path.join(nova_pasta, f"{novo_nome_arquivo}.pdf")

    # Abre o PDF original e modifica a área especificada
    documento = fitz.open(pdf_path)
    x0, y0, x1, y1 = 30, 710, 550, 810

    for page_number in range(len(documento)):
        pagina = documento.load_page(page_number)
        rect = fitz.Rect(x0, y0, x1, y1)
        pagina.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

    documento.save("temp_modificado.pdf")
    documento.close()

    # Cria o novo PDF com o nome e endereço do cliente
    reader_modificado = PdfReader("temp_modificado.pdf")
    writer = PdfWriter()

    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)

    # Ajustar o texto para caber dentro do retângulo branco identificado
    x_text, y_text = 25, 500  # Coordenadas do texto
    max_width = x1 - x0 - 10  # Limitar a largura do texto ao retângulo branco
    font_size = 10  # Tamanho de fonte fixo
    can.setFont("Helvetica", font_size)

    # Desenhar o texto no PDF
    if "-" in endereco_cliente:
        partes_endereco = endereco_cliente.split("-", 1)
        endereco_cliente = f"{partes_endereco[0].strip()}\n{partes_endereco[1].strip()}"

    texto_completo = f"{nome_cliente}\n{endereco_cliente}"
    linhas_texto = texto_completo.split("\n")
    for i, linha in enumerate(linhas_texto):
        while can.stringWidth(linha, "Helvetica", font_size) > max_width:
            linha = linha[:-1]  # Reduzir o texto até caber na largura máxima
        can.drawString(x_text, y_text - (i * 15), linha)
    
    can.save()
    packet.seek(0)

    # Adiciona a nova página com os dados do cliente
    nova_pagina = PdfReader(packet).pages[0]
    writer.add_page(nova_pagina)
    writer.add_page(reader_modificado.pages[0])

    # Salva o novo PDF final
    with open(novo_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

    os.remove("temp_modificado.pdf")
    return novo_pdf_path

# Função para adicionar cada PDF processado à lista com uma caixa de seleção
def adicionar_item_lista(pdf_path):
    """
    Adiciona cada PDF processado à lista de PDFs com uma caixa de seleção.
    """
    var = tk.BooleanVar()  # Variável associada à caixa de seleção
    checkbox = ttk.Checkbutton(frame_lista, text=os.path.basename(pdf_path), variable=var)
    checkbox.var = var
    checkbox.pdf_path = pdf_path
    checkbox.pack(anchor='w', padx=5, pady=2)
    checkbox.bind('<Button-3>', lambda event, path=pdf_path: abrir_pdf(path))

# Função para abrir PDF ao clicar com botão direito
def abrir_pdf(pdf_path):
    """
    Abre o arquivo PDF selecionado ao clicar com o botão direito.
    """
    try:
        os.startfile(pdf_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o PDF: {e}")

# Função para selecionar todos os itens
def selecionar_todos(event=None):
    """
    Seleciona todos os PDFs na lista de PDFs processados.
    """
    for widget in frame_lista.winfo_children():
        if isinstance(widget, ttk.Checkbutton):
            widget.var.set(True)

# Função para salvar os dados em um arquivo Excel
def save_to_excel(data, output_file_path):
    columns = ['Nome', 'Rua', 'Número', 'Bairro', 'CEP', 'Cidade', 'Estado']
    df = pd.DataFrame(data, columns=columns)
    df.to_excel(output_file_path, index=False)

# Função para extrair dados do arquivo TXT e criar um DataFrame
def process_txt_file(txt_file_path):
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file if line.strip()]

    data = []
    i = 0
    while i < len(lines):
        # Obtendo as informações de cada bloco de 3 linhas
        nome = lines[i]
        rua_numero = lines[i + 1] if i + 1 < len(lines) else ''
        bairro_cep_cidade_estado = lines[i + 2] if i + 2 < len(lines) else ''

        # Verificando se o formato do bairro/cep/cidade/estado está diferente, como no caso da Midiã Gomes da Costa
        if re.match(r'^.*\s\d{8}\s.*$', rua_numero):
            # Caso especial: linha 2 contém o bairro, cep, cidade e estado
            bairro_cep_cidade_estado = rua_numero
            rua_numero = lines[i + 2] if i + 2 < len(lines) else ''

        # Extraindo informações da rua e número
        rua, numero = rua_numero.split(',', 1) if ',' in rua_numero else (rua_numero, '')
        rua = rua.strip()
        numero = numero.strip()

        # Extraindo bairro, cep, cidade e estado
        bairro_cep_pattern = r'^(.*)\s(\d{8})\s(.*)/(\w{2})$'
        match = re.match(bairro_cep_pattern, bairro_cep_cidade_estado)
        if match:
            bairro = match.group(1).strip()
            cep = match.group(2).strip()
            cidade = match.group(3).strip()
            estado = match.group(4).strip()
        else:
            bairro = cep = cidade = estado = ''

        # Removendo prefixo "- " da cidade, se existir
        if cidade.startswith('- '):
            cidade = cidade[2:].strip()

        # Adicionando os dados extraídos à lista
        data.append([nome, rua, numero, bairro, cep, cidade, estado])

        # Avançando para o próximo bloco de 3 linhas
        i += 3

    return data

# Função para imprimir os PDFs selecionados
def imprimir_selecionados(event=None):
    """
    Cria um PDF compilado de todos os PDFs selecionados e envia para impressão.
    """
    writer = PdfWriter()
    clientes_info = []
    for widget in frame_lista.winfo_children():
        if isinstance(widget, ttk.Checkbutton) and widget.var.get():
            reader = PdfReader(widget.pdf_path)
            for page in reader.pages:
                writer.add_page(page)
            nome_cliente, endereco_cliente, _, _ = extrair_dados_cliente_e_instrucoes(widget.pdf_path)
            if nome_cliente and endereco_cliente:
                clientes_info.append((nome_cliente, endereco_cliente))

    # Salvar o PDF compilado na pasta "Compilados"
    data_compilacao = datetime.now()
    ano_compilacao = data_compilacao.strftime("%Y")
    meses_em_portugues = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_compilacao_nome = meses_em_portugues[data_compilacao.month - 1]
    pasta_compilados = os.path.join("Documentos", "Compilados", ano_compilacao, mes_compilacao_nome)
    if not os.path.exists(pasta_compilados):
        os.makedirs(pasta_compilados)

    numero_arquivos = len(os.listdir(pasta_compilados)) + 1
    temp_pdf_path = os.path.join(pasta_compilados, f"{numero_arquivos} - {data_compilacao.strftime('%d-%m-%Y')}.pdf")
    with open(temp_pdf_path, "wb") as temp_pdf:
        writer.write(temp_pdf)

    # Criar um arquivo TXT com nome e endereço dos clientes na pasta "Correios"
    pasta_correios = os.path.join("Documentos", "Correios", ano_compilacao, mes_compilacao_nome)
    if not os.path.exists(pasta_correios):
        os.makedirs(pasta_correios)

    txt_clientes_correios_path = os.path.join(pasta_correios, f"{numero_arquivos} - {data_compilacao.strftime('%d-%m-%Y')} - clientes.txt")
    try:
        with open(txt_clientes_correios_path, "w", encoding="utf-8") as txt_file:
            for page_number in range(0, len(writer.pages), 2):  # Páginas ímpares do PDF compilado
                page = writer.pages[page_number]
                texto = page.extract_text()
                if texto:
                    txt_file.write(texto + "\n\n")
        print(f"Arquivo TXT salvo em: {txt_clientes_correios_path}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo TXT: {e}")

    # Chamar a função para processar o TXT e criar o Excel automaticamente
    data = process_txt_file(txt_clientes_correios_path)
    save_to_excel(data, f"{os.path.splitext(txt_clientes_correios_path)[0]}.xlsx")

    # Debug: Verificando se o arquivo foi realmente salvo
    if os.path.exists(txt_clientes_correios_path):
        print(f"Verificação: O arquivo TXT foi salvo corretamente em {txt_clientes_correios_path}")
    else:
        print(f"Erro: O arquivo TXT não foi salvo em {txt_clientes_correios_path}")

    # Abrir o PDF compilado para visualização
    try:
        os.startfile(temp_pdf_path)
        print(f"PDF compilado aberto: {temp_pdf_path}")
    except Exception as e:
        print(f"Erro ao abrir o arquivo {temp_pdf_path}: {e}")
    messagebox.showinfo("PDF Aberto", "O PDF compilado foi aberto para visualização e impressão.")

# Função para abrir a pasta de boletos gerados
def abrir_documentos():
    """
    Abre a pasta onde os documentos estão salvos.
    """
    os.startfile("Documentos")

# Função para abrir uma janela com todas as funcionalidades do sistema
def abrir_ajuda():
    """
    Abre uma janela de ajuda com informações sobre como usar o sistema.
    """
    nova_janela = tk.Toplevel(janela)
    nova_janela.title("Ajuda - Funcionalidades do Sistema")
    nova_janela.geometry("600x400")

    texto_ajuda = (
        "Passo a passo para usar o sistema:\n"
        "1. Carregar PDF: Selecione os arquivos PDF que deseja processar.\n"
        "2. Selecionar Todos: Marque todos os PDFs carregados.\n"
        "3. Imprimir: Compile os PDFs selecionados em um único arquivo e envie para impressão.\n"
        "4. Documentos: Abra a pasta onde os documentos estão salvos.\n\n"
        "Funcionalidades adicionais:\n"
        "- Clique com o botão direito em um PDF na lista para abri-lo diretamente."
    )

    ttk.Label(nova_janela, text=texto_ajuda, wraplength=580, justify="left").pack(pady=20, padx=20)

# Função para excluir os PDFs selecionados
def excluir_selecionados():
    """
    Exclui os PDFs selecionados da lista e do sistema de arquivos.
    """
    for widget in frame_lista.winfo_children():
        if isinstance(widget, ttk.Checkbutton) and widget.var.get():
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
janela.geometry("900x600")

# Frame dos botões organizados lado a lado
frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=10)

# Botões principais
btn_selecionar = ttk.Button(frame_botoes, text="Carregar PDF", command=selecionar_arquivos)
btn_selecionar.grid(row=0, column=0, padx=5)

btn_selecionar_todos = ttk.Button(frame_botoes, text="Selecionar Todos", command=selecionar_todos)
btn_selecionar_todos.grid(row=0, column=1, padx=5)

btn_imprimir = ttk.Button(frame_botoes, text="Imprimir", command=imprimir_selecionados)
btn_imprimir.grid(row=0, column=2, padx=5)

btn_excluir = ttk.Button(frame_botoes, text="Excluir", command=excluir_selecionados)
btn_excluir.grid(row=0, column=3, padx=5)

btn_documentos = ttk.Button(frame_botoes, text="Documentos", command=abrir_documentos)
btn_documentos.grid(row=0, column=4, padx=5)

btn_ajuda = ttk.Button(frame_botoes, text="Ajuda", command=abrir_ajuda)
btn_ajuda.grid(row=0, column=5, padx=5)

# Frame da lista de PDFs processados
frame_lista = tk.Frame(janela)
frame_lista.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

# Adicionando atalhos de teclado
janela.bind("<Control-a>", selecionar_todos)
janela.bind("<Control-p>", imprimir_selecionados)

# Inicia o loop principal da interface
janela.mainloop()
