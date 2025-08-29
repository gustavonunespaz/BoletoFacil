import os
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
from datetime import datetime

def atualizar_contador():
    total = len(frame_lista.winfo_children())
    selecionados = sum(1 for widget in frame_lista.winfo_children() if isinstance(widget, ttk.Checkbutton) and widget.var.get())
    contador_label.set(f"Selecionados: {selecionados} / Total: {total}")

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

def processar_pdf(pdf_path):
    nome_cliente, endereco_cliente, venda_parcela, data_vencimento = extrair_dados_cliente_e_instrucoes(pdf_path)
    if nome_cliente and endereco_cliente and data_vencimento:
        return criar_pdf_final(pdf_path, nome_cliente, endereco_cliente, venda_parcela, data_vencimento)
    return None

def extrair_dados_cliente_e_instrucoes(pdf_path):
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
                    if i + 1 < len(linhas):
                        possivel_nome = linhas[i + 1].strip()
                        nome_cliente = re.sub(r'R\$\s*\d+([.,]\d{2})?', '', possivel_nome).strip()
                    if i + 2 < len(linhas):
                        endereco_cliente = linhas[i + 2].strip()
                        if endereco_cliente.endswith('-') and (i + 3 < len(linhas)):
                            endereco_cliente = endereco_cliente[:-1].strip() + ' ' + linhas[i + 3].strip()
                if "Instruções" in linha and i + 1 < len(linhas):
                    venda_parcela = linhas[i + 1].strip()
                if "Vencimento" in linha and i + 1 < len(linhas):
                    data_vencimento = re.sub(r'[^0-9/]', '', linhas[i + 1].strip())

            print(f"Dados extraídos do arquivo {pdf_path}: Nome: {nome_cliente}, Endereço: {endereco_cliente}, Vencimento: {data_vencimento}")
            return nome_cliente, endereco_cliente, venda_parcela, data_vencimento
    except Exception as e:
        print(f"Erro ao extrair dados do arquivo {pdf_path}: {e}")
        return None, None, None, None

def atualizar_pdf(pdf_path, nome_cliente, endereco, numero, complemento, bairro, cep, cidade_estado):
    """
    Remove a página 1, cria uma nova com as informações atualizadas e mantém a página 2 intacta.
    """
    try:
        documento = fitz.open(pdf_path)

        pagina2 = documento[1] 

        novo_documento = fitz.open()

        nova_pagina = novo_documento.new_page(width=fitz.paper_rect("a4").width, height=fitz.paper_rect("a4").height)

        x_start = 25  
        y_start = 450 

        nova_pagina.insert_text(
            fitz.Point(x_start, y_start - 70),  
            nome_cliente,
            fontsize=10,
            fontname="helv",
            color=(0, 0, 0)
        )

        endereco_linha2 = f"{endereco}, {numero} - {complemento}".strip(", ")
        nova_pagina.insert_text(
            fitz.Point(x_start, y_start - 50),  
            endereco_linha2,
            fontsize=10,
            fontname="helv",
            color=(0, 0, 0)
        )

        endereco_linha3 = f"{bairro} {cep} - {cidade_estado}".strip()
        nova_pagina.insert_text(
            fitz.Point(x_start, y_start - 30),  
            endereco_linha3,
            fontsize=10,
            fontname="helv",
            color=(0, 0, 0)
        )

        novo_documento.insert_pdf(documento, from_page=1, to_page=1)

        documento.close()

        novo_documento.save(pdf_path)
        novo_documento.close()

        print(f"Endereço atualizado com sucesso em: {pdf_path}")
    except Exception as e:
        print(f"Erro ao atualizar o PDF: {e}")

def adicionar_complemento(pdf_path, complemento):
    try:
        documento = fitz.open(pdf_path)

        novo_documento = fitz.open()

        pagina1 = documento[0]
        pagina1_texto = pagina1.get_text("text")

        linhas = pagina1_texto.split("\n")
        endereco_original = linhas[1] 
        if "," in endereco_original: 
            endereco_complemento = f"{endereco_original} - {complemento}"
        else:
            endereco_complemento = f"{endereco_original}, {complemento}"

        linhas[1] = endereco_complemento  

        nova_pagina = novo_documento.new_page(width=pagina1.rect.width, height=pagina1.rect.height)

        x_start, y_start = 25, 400  
        for i, linha in enumerate(linhas):
            nova_pagina.insert_text(
                fitz.Point(x_start, y_start - (i * -15)),
                linha,
                fontsize=10,
                fontname="helv",
                color=(0, 0, 0),
            )

        for page_num in range(1, len(documento)):
            novo_documento.insert_pdf(documento, from_page=page_num, to_page=page_num)

        documento.close()

        novo_documento.save(pdf_path)
        novo_documento.close()

        print(f"Arquivo atualizado com sucesso: {pdf_path}")
    except Exception as e:
        print(f"Erro ao adicionar o complemento: {e}")

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

    nova_pasta = os.path.join("Documentos", "Boletos", ano_vencimento, mes_vencimento_nome)
    if not os.path.exists(nova_pasta):
        os.makedirs(nova_pasta)

    novo_nome_arquivo = f"{nome_cliente} - {data_formatada}"
    if venda_parcela:
        novo_nome_arquivo += f" - {venda_parcela.replace('/', '-')}"
    novo_pdf_path = os.path.join(nova_pasta, f"{novo_nome_arquivo}.pdf")

    documento = fitz.open(pdf_path)
    x0, y0, x1, y1 = 30, 750, 550, 810

    for page_number in range(len(documento)):
        pagina = documento.load_page(page_number)
        rect = fitz.Rect(x0, y0, x1, y1)
        pagina.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

    documento.save("temp_modificado.pdf")
    documento.close()

    reader_modificado = PdfReader("temp_modificado.pdf")
    writer = PdfWriter()

    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)

    x_text, y_text = 25, 450  
    max_width = x1 - x0 - 10 
    font_size = 10  
    can.setFont("Helvetica", font_size)

    if "-" in endereco_cliente:
        partes_endereco = endereco_cliente.split("-", 1)
        endereco_cliente = f"{partes_endereco[0].strip()}\n{partes_endereco[1].strip()}"

    texto_completo = f"{nome_cliente}\n{endereco_cliente}"
    linhas_texto = texto_completo.split("\n")
    for i, linha in enumerate(linhas_texto):
        while can.stringWidth(linha, "Helvetica", font_size) > max_width:
            linha = linha[:-1] 
        can.drawString(x_text, y_text - (i * 15), linha)
    
    can.save()
    packet.seek(0)

    nova_pagina = PdfReader(packet).pages[0]
    writer.add_page(nova_pagina)

    pagina_modificada = reader_modificado.pages[0]
    pagina_modificada.rotate(180)  
    writer.add_page(pagina_modificada)

    with open(novo_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

    os.remove("temp_modificado.pdf")
    return novo_pdf_path

def adicionar_item_lista(pdf_path):
    """
    Adiciona cada PDF processado à lista de PDFs com uma caixa de seleção.
    """
    var = tk.BooleanVar() 
    checkbox = ttk.Checkbutton(frame_lista, text=os.path.basename(pdf_path), variable=var)
    checkbox.var = var
    checkbox.pdf_path = pdf_path
    checkbox.pack(anchor='w', padx=5, pady=2)
    checkbox.bind('<Button-3>', lambda event, path=pdf_path: abrir_pdf(path))

def adicionar_item_lista(pdf_path):
    var = tk.BooleanVar()
    checkbox = ttk.Checkbutton(frame_lista, text=os.path.basename(pdf_path), variable=var)
    checkbox.var = var
    checkbox.pdf_path = pdf_path
    checkbox.pack(anchor='w', padx=5, pady=2)
    checkbox.bind('<Button-3>', lambda event, path=pdf_path: abrir_pdf(path))
    atualizar_contador()

ordem_crescente = True

def classificar_lista():

    global ordem_crescente
    itens = list(frame_lista.winfo_children())
    itens.sort(key=lambda widget: widget.cget("text"), reverse=not ordem_crescente)
    
    for widget in itens:
        widget.pack_forget()
    for widget in itens:
        widget.pack(anchor="w", padx=5, pady=2)
    
    ordem_crescente = not ordem_crescente
    
    atualizar_contador()

    frame_lista.update_idletasks()
    canvas_lista.configure(scrollregion=canvas_lista.bbox("all"))

def abrir_pdf(pdf_path):

    try:
        os.startfile(pdf_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o PDF: {e}")

    # Atualizar o contador após selecionar
    atualizar_contador()    

def abrir_documentos():

    os.startfile("Documentos")

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

    atualizar_contador()

janela = tk.Tk()
janela.title("Boleto Fácil - Nunes|Paz")
janela.geometry("900x600")

selecao_geral_var = tk.BooleanVar()

def abrir_editor_de_endereco():
    """
    Abre uma janela para editar o endereço de um PDF selecionado.
    """
    selecionados = [
        widget for widget in frame_lista.winfo_children()
        if isinstance(widget, ttk.Checkbutton) and widget.var.get()
    ]

    if len(selecionados) != 1:
        messagebox.showinfo(
            "Seleção Inválida", "Por favor, selecione apenas um PDF para editar o endereço."
        )
        return

    pdf_path = selecionados[0].pdf_path
    nome_cliente = os.path.basename(pdf_path).split(" - ")[1].strip()

    nova_janela = tk.Toplevel(janela)
    nova_janela.title("Editar Endereço")
    nova_janela.geometry("600x300")

    ttk.Label(nova_janela, text="Nome do Cliente:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    ttk.Label(nova_janela, text=nome_cliente).grid(row=0, column=1, columnspan=3, padx=10, pady=10, sticky="w")
    ttk.Label(nova_janela, text="Endereço:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    entrada_endereco = ttk.Entry(nova_janela, width=30)
    entrada_endereco.grid(row=1, column=1, padx=5, pady=5)
    ttk.Label(nova_janela, text="Número:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
    entrada_numero = ttk.Entry(nova_janela, width=10)
    entrada_numero.grid(row=1, column=3, padx=5, pady=5)
    ttk.Label(nova_janela, text="Complemento:").grid(row=1, column=4, padx=5, pady=5, sticky="w")
    entrada_complemento = ttk.Entry(nova_janela, width=15)
    entrada_complemento.grid(row=1, column=5, padx=5, pady=5)

    ttk.Label(nova_janela, text="Bairro:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    entrada_bairro = ttk.Entry(nova_janela, width=20)
    entrada_bairro.grid(row=2, column=1, padx=5, pady=5)
    ttk.Label(nova_janela, text="CEP:").grid(row=2, column=2, padx=5, pady=5, sticky="w")
    entrada_cep = ttk.Entry(nova_janela, width=15)
    entrada_cep.grid(row=2, column=3, padx=5, pady=5)
    ttk.Label(nova_janela, text="Cidade/Estado:").grid(row=2, column=4, padx=5, pady=5, sticky="w")
    entrada_cidade_estado = ttk.Entry(nova_janela, width=20)
    entrada_cidade_estado.grid(row=2, column=5, padx=5, pady=5)

    def salvar_endereco():
        endereco = entrada_endereco.get()
        numero = entrada_numero.get()
        complemento = entrada_complemento.get()
        bairro = entrada_bairro.get()
        cep = entrada_cep.get()
        cidade_estado = entrada_cidade_estado.get()

        if endereco and numero and bairro and cep and cidade_estado:
            atualizar_pdf(pdf_path, nome_cliente, endereco, numero, complemento, bairro, cep, cidade_estado)
            nova_janela.destroy()
        else:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")

    ttk.Button(nova_janela, text="Salvar", command=salvar_endereco).grid(row=4, column=0, columnspan=6, pady=20)

def abrir_editor_de_complemento():
    """
    Abre uma janela para adicionar um complemento ao endereço no PDF selecionado.
    """
    selecionados = [
        widget for widget in frame_lista.winfo_children()
        if isinstance(widget, ttk.Checkbutton) and widget.var.get()
    ]

    if len(selecionados) != 1:
        messagebox.showinfo(
            "Seleção Inválida", "Por favor, selecione apenas um PDF para editar o endereço."
        )
        return

    pdf_path = selecionados[0].pdf_path

    nova_janela = tk.Toplevel(janela)
    nova_janela.title("Adicionar Complemento")
    nova_janela.geometry("400x200")

    ttk.Label(nova_janela, text="Complemento:").pack(pady=10)
    entrada_complemento = ttk.Entry(nova_janela, width=50)
    entrada_complemento.pack(pady=10)

    def salvar_complemento():
        complemento = entrada_complemento.get().strip()
        if complemento:
            adicionar_complemento(pdf_path, complemento)
            nova_janela.destroy()
        else:
            messagebox.showerror("Erro", "O complemento não pode estar vazio!")

    ttk.Button(nova_janela, text="Salvar", command=salvar_complemento).pack(pady=20)

def alternar_selecao_geral():
    """
    Marca ou desmarca todas as caixas de seleção com base na caixa geral.
    """
    for widget in frame_lista.winfo_children():
        if isinstance(widget, ttk.Checkbutton):
            widget.var.set(selecao_geral_var.get())
    atualizar_contador()

checkbox_selecao_geral = ttk.Checkbutton(
    janela,
    text="Selecionar Todos",
    variable=selecao_geral_var,
    command=alternar_selecao_geral
)
checkbox_selecao_geral.pack(anchor="w", padx=20)

frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=10)

btn_selecionar = ttk.Button(frame_botoes, text="Carregar PDF", command=selecionar_arquivos)
btn_selecionar.grid(row=0, column=0, padx=5)

btn_excluir = ttk.Button(frame_botoes, text="Excluir", command=excluir_selecionados)
btn_excluir.grid(row=0, column=3, padx=5)

btn_documentos = ttk.Button(frame_botoes, text="Documentos", command=abrir_documentos)
btn_documentos.grid(row=0, column=4, padx=5)

btn_classificar = ttk.Button(frame_botoes, text="Classificar A-Z", command=classificar_lista)
btn_classificar.grid(row=0, column=6, padx=5)

btn_editar = ttk.Button(frame_botoes, text="Adicionar Complemento", command=abrir_editor_de_complemento)
btn_editar.grid(row=0, column=7, padx=5)

contador_label = tk.StringVar()
contador_label.set("Selecionados: 0 / Total: 0")

contador = ttk.Label(janela, textvariable=contador_label, anchor="w")
contador.pack(fill=tk.X, padx=20, pady=5)

frame_lista_container = tk.Frame(janela)
frame_lista_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

canvas_lista = tk.Canvas(frame_lista_container)
canvas_lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar_lista = ttk.Scrollbar(frame_lista_container, orient=tk.VERTICAL, command=canvas_lista.yview)
scrollbar_lista.pack(side=tk.RIGHT, fill=tk.Y)

canvas_lista.configure(yscrollcommand=scrollbar_lista.set)
canvas_lista.bind(
    "<Configure>",
    lambda e: canvas_lista.configure(scrollregion=canvas_lista.bbox("all"))
)

frame_lista = tk.Frame(canvas_lista)
canvas_lista.create_window((0, 0), window=frame_lista, anchor="nw")

def on_mousewheel(event):
    canvas_lista.yview_scroll(-1 * int(event.delta / 120), "units")

canvas_lista.bind_all("<MouseWheel>", on_mousewheel)

janela.mainloop()
