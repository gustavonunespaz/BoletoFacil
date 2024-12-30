import os
from tkinter import Tk, filedialog, messagebox
import fitz  # PyMuPDF
from datetime import datetime
import openpyxl
import re

def get_next_series_number(output_dir):
    try:
        existing_files = [f for f in os.listdir(output_dir) if f.endswith(".pdf")]
        series_numbers = [int(f.split('_')[0]) for f in existing_files if f.split('_')[0].isdigit()]
        if series_numbers:
            return str(max(series_numbers) + 1).zfill(3)
        else:
            return "001"
    except Exception as e:
        print(f"Error calculating next series number: {type(e).__name__} - {e}")
        return "001"

def process_txt_file(txt_path):
    # Lê o arquivo TXT e processa os dados
    with open(txt_path, "r", encoding="utf-8") as file:
        data = file.read().strip()
        
    # Cada cliente está separado por duas quebras de linha
    registros = data.split("\n\n")
    
    # Extrair dados estruturados de cada registro
    clientes = []
    for registro in registros:
        linhas = registro.strip().split("\n")
        if len(linhas) >= 3:
            nome = linhas[0].strip()
            
            # Separar endereço e número
            endereco_parts = linhas[1].split(",", 1)
            endereco = endereco_parts[0].strip() if len(endereco_parts) > 0 else ""
            numero = endereco_parts[1].strip() if len(endereco_parts) > 1 else ""

            # Extrair Bairro, CEP e Cidade/Estado
            linha_terceira = linhas[2].strip()

            # Encontrar o primeiro dígito que indica o início do CEP
            match_cep = re.search(r'\d', linha_terceira)
            if match_cep:
                index_inicio_cep = match_cep.start()
                bairro = linha_terceira[:index_inicio_cep].strip()
                resto = linha_terceira[index_inicio_cep:].strip()

                # Encontrar o primeiro caractere que indica o início da Cidade/Estado
                match_cidade_estado = re.search(r'[A-Za-z]', resto)
                if match_cidade_estado:
                    index_inicio_cidade = match_cidade_estado.start()
                    cep = resto[:index_inicio_cidade].strip()
                    cidade_estado = resto[index_inicio_cidade:].strip()
                else:
                    cep = resto
                    cidade_estado = ""
            else:
                bairro = ""
                cep = ""
                cidade_estado = ""

            # Remover hífen do início do Cidade/Estado, se presente
            cidade_estado = cidade_estado.lstrip('-').strip()

            # Adicionar aos clientes
            clientes.append({
                "Nome": nome,
                "Endereço": endereco,
                "Número": numero,
                "Bairro": bairro,
                "CEP": cep,
                "Cidade/Estado": cidade_estado
            })
    
    return clientes

def save_to_excel(data, excel_path):
    # Cria um novo arquivo Excel e escreve os dados dos clientes
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Clientes Correios"

    # Cabeçalhos
    sheet.append(["Nome", "Endereço", "Número", "Bairro", "CEP", "Cidade/Estado"])

    # Adiciona cada cliente à planilha
    for cliente in data:
        sheet.append([cliente["Nome"], cliente["Endereço"], cliente["Número"], cliente["Bairro"], cliente["CEP"], cliente["Cidade/Estado"]])

    # Salva o arquivo Excel
    workbook.save(excel_path)
    print(f"Arquivo Excel salvo em: {excel_path}")

def merge_pdfs():
    # Step 1: Create GUI for selecting PDF files
    root = Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF files", "*.pdf")])
    
    if not file_paths:
        print("No files selected.")
        return

    # Step 2: Create output directory if it doesn't exist
    output_dir = "Compilados"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Step 3: Create a filename with series number + compilation date
    series_number = get_next_series_number(output_dir)
    compilation_date = datetime.now().strftime("%d-%m-%Y")
    output_filename = f"{series_number}_{compilation_date}.pdf"
    output_path = os.path.join(output_dir, output_filename)
    
    # Step 4: Create a new PDF document
    merged_pdf = fitz.open()
    
    # Step 5: Read and append each selected PDF to the merged document
    for file_path in file_paths:
        try:
            print(f"Processing file: {file_path}")  # Log the file being processed
            pdf_document = fitz.open(file_path)
            merged_pdf.insert_pdf(pdf_document)
            pdf_document.close()
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
    
    # Step 6: Save the merged PDF to the output file
    merged_pdf.save(output_path)
    print(f"PDFs merged successfully and saved to {output_path}")

    # Step 7: Create a TXT file with client names and addresses in the "Compilados" folder
    txt_clientes_correios_path = os.path.join(output_dir, f"{series_number}_{compilation_date}.txt")
    try:
        with open(txt_clientes_correios_path, "w", encoding="utf-8") as txt_file:
            for page_number in range(0, len(merged_pdf), 2):  # Páginas ímpares do PDF compilado
                page = merged_pdf.load_page(page_number)
                texto = page.get_text()
                if texto:
                    txt_file.write(texto + "\n\n")
        print(f"Arquivo TXT salvo em: {txt_clientes_correios_path}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo TXT: {e}")

    # Step 8: Process the TXT and create an Excel file automatically in the "Compilados" folder
    data = process_txt_file(txt_clientes_correios_path)
    excel_clientes_correios_path = os.path.join(output_dir, f"{series_number}_{compilation_date}.xlsx")
    save_to_excel(data, excel_clientes_correios_path)

    # Debug: Check if the TXT file was really saved
    if os.path.exists(txt_clientes_correios_path):
        print(f"Verificação: O arquivo TXT foi salvo corretamente em {txt_clientes_correios_path}")
    else:
        print(f"Erro: O arquivo TXT não foi salvo em {txt_clientes_correios_path}")

    # Step 9: Close the merged PDF
    merged_pdf.close()

    # Step 10: Display success message
    messagebox.showinfo("Success", f"PDFs merged successfully and saved to:\n{output_path}")

if __name__ == "__main__":
    merge_pdfs()
