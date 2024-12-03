
# Boleto Fácil - Nunes|Paz

## Visão Geral

O projeto **Boleto Fácil - Nunes|Paz** é um software desenvolvido para facilitar o gerenciamento de boletos de clientes, incluindo a seleção, processamento, impressão e organização de arquivos PDF. O sistema foi projetado para melhorar a eficiência no controle de documentos financeiros, proporcionando uma interface intuitiva para os usuários. Esse sistema foi criado tendo em vista a dificuldade da minha empresa em fazer a impressão dos boletos. Com o sistema otimizamos tempo e recursos.

## Funcionalidades

- **Carregar PDF**: Selecione arquivos PDF para processamento e extração de informações relevantes dos boletos.
- **Selecionar Todos**: Caixa de seleção para marcar ou desmarcar todos os PDFs presentes na lista.
- **Processamento Automático**: Extração automática dos dados dos clientes e criação de novos PDFs organizados.
- **Classificar A-Z**: Permite que os arquivos na lista sejam classificados em ordem alfabética crescente ou decrescente, facilitando a organização.
- **Impressão de PDFs Selecionados**: É possível selecionar vários arquivos da lista e imprimi-los em um único arquivo PDF compilado.
- **Exclusão de PDFs**: Os PDFs selecionados podem ser excluídos tanto do sistema de arquivos quanto da interface gráfica.
- **Abrir Documentos**: Botão que abre a pasta onde estão armazenados os boletos gerados.
- **Ajuda**: Fornece instruções sobre como utilizar o sistema, garantindo que todos os usuários consigam aproveitar as funcionalidades do programa.

## Tecnologias Utilizadas

Este programa foi implementado utilizando as seguintes bibliotecas do Python:

- **tkinter**: Para a criação da interface gráfica de usuário.
- **PyPDF2**, **pdfplumber**, **PyMuPDF (fitz)**: Para manipulação, extração de dados e edição visual dos arquivos PDF.
- **reportlab**: Para gerar novas páginas em PDF que contêm informações extraídas dos boletos, como nome e endereço do cliente.
- **win32api**: Utilizada para comandos do sistema Windows, como envio de documentos diretamente para impressão.
- **pandas** e **openpyxl**: Para criar planilhas Excel a partir dos dados extraídos dos boletos, facilitando a análise e gestão.

## Fluxo do Programa

1. **Processamento de Boletos**: O usuário carrega os arquivos PDF de boletos, que são processados pelo programa. Durante esse processamento, o sistema extrai informações importantes dos boletos, como nome do cliente, endereço e data de vencimento.
2. **Geração de PDFs Finalizados**: Após a extração dos dados, um novo arquivo PDF é gerado contendo as informações formatadas, como nome e endereço do cliente.
3. **Manipulação dos Boletos na Interface**: O usuário pode escolher imprimir, excluir, classificar ou visualizar os arquivos diretamente da interface do programa.
4. **Impressão de PDFs Compilados**: Ao selecionar vários boletos, o sistema cria um arquivo PDF único compilado, que pode ser enviado diretamente para impressão, abrindo a janela de configuração de impressão.
5. **Criação de Planilhas Excel**: Um arquivo Excel é gerado a partir dos dados dos clientes extraídos dos boletos, facilitando a análise e organização das informações.

## Estrutura do Projeto

- **boleto_facil.py**: Arquivo principal com o código do sistema.
- **Documentos/**:
  - **Boletos/**: Pasta onde os boletos processados são salvos, organizados por ano e mês.
  - **Compilados/**: Pasta onde os PDFs compilados são salvos.
  - **Correios/**: Pasta onde arquivos TXT com informações dos clientes são armazenados.

## Possíveis Problemas e Soluções

- **Erro ao Imprimir PDF**:
  - Certifique-se de que o Adobe Reader ou outro software de leitura de PDF esteja instalado e configurado como aplicativo padrão para abrir arquivos .pdf.
- **Problemas de Permissão**:
  - Execute o programa com permissão de administrador, caso encontre erros ao salvar ou excluir arquivos.
- **Erro ao Abrir PDF**:
  - Verifique se o caminho do arquivo PDF é válido e se o arquivo não foi movido ou excluído.

## Requisitos de Sistema

- Python 3.8 ou superior
- Sistema operacional Windows (devido ao uso da biblioteca **win32api** para impressão)
- As seguintes bibliotecas Python:
  - tkinter
  - PyPDF2
  - pdfplumber
  - fitz (PyMuPDF)
  - reportlab
  - win32api
  - pandas
  - openpyxl

## Como Executar

1. Clone o repositório para o seu computador:
   ```
   git clone https://github.com/gustavonunespaz/BoletoFacil.git
   ```
2. Instale as dependências necessárias:
   ```
   pip install -r requirements.txt
   ```
3. Execute o script principal para abrir a interface do sistema:
   ```
   python boleto_facil.py
   ```

## Contribuição

Se desejar contribuir para o projeto, sinta-se à vontade para fazer um fork do repositório, criar uma branch com suas melhorias e enviar um pull request. Todas as contribuições são bem-vindas!

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo LICENSE para mais detalhes.

## Contato

- **Gustavo Paz**
  - Telefone: 41995953862
  - E-mail: gust.nunes.paz@gmail.com
