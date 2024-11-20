# BoletoFacil
Boleto Fácil - Nunes|Paz

## Visão Geral

Boleto Fácil - Nunes|Paz é um software desenvolvido para facilitar o gerenciamento de boletos de clientes, incluindo a seleção, processamento, impressão e organização de arquivos PDF. O sistema foi projetado para melhorar a eficiência no controle de documentos financeiros, proporcionando uma interface intuitiva para os usuários.
Esse sistema foi criado tendo em vista a dificuldade da minha empresa em fazer a impressão dos boletos. Com o sistema otimizamos tempo e recursos.

## Funcionalidades

- **Carregar PDF**: Selecione arquivos PDF para processamento e extração de informações relevantes dos boletos.
- **Selecionar Todos**: Marque todos os PDFs carregados para serem impressos ou excluídos.
- **Processamento Automático**: Extração automática dos dados dos clientes e criação de novos PDFs organizados.
- **Impressão de PDFs**: Compile os PDFs selecionados em um único arquivo e envie-o para impressão, abrindo a janela de configuração de impressão.
- **Exclusão de PDFs**: Exclua PDFs que já não são necessários diretamente pela interface do software.
- **Abrir Boletos Processados**: Acesse rapidamente a pasta onde os boletos processados estão salvos.
- **Ajuda**: Uma janela com instruções detalhadas sobre como usar o sistema.

## Tecnologias Utilizadas

- **Python**: Linguagem principal utilizada no desenvolvimento do software.
- **Tkinter**: Biblioteca para criação da interface gráfica do usuário (GUI).
- **PyPDF2, pdfplumber, PyMuPDF (fitz)**: Bibliotecas usadas para manipulação de PDFs, extração de texto e edição.
- **Win32api**: Utilizada para comandos do sistema Windows, como a abertura da janela de impressão.
- **ReportLab**: Biblioteca para geração de arquivos PDF e adição de elementos gráficos.

## Como Utilizar

### Requisitos

- **Python 3.8+**
- **Bibliotecas Python**: As bibliotecas listadas abaixo podem ser instaladas usando o comando `pip`:
  ```sh
  pip install PyPDF2 pdfplumber pymupdf pywin32 reportlab
  ```

### Executando o Programa

1. Execute o programa Python:
   ```sh
   python BoletoFacil.py
   ```

## Usando o Sistema

- **Carregar PDF**: Clique no botão "Carregar PDF" e selecione os arquivos que deseja processar.
- **Selecionar Todos**: Clique no botão "Selecionar Todos" para marcar todos os PDFs carregados.
- **Imprimir**: Clique em "Imprimir" para abrir a janela de impressão com as configurações necessárias.
- **Excluir**: Utilize o botão "Excluir" para remover os PDFs selecionados.
- **Documentos**: Clique em "Documentos" para acessar rapidamente a pasta onde os boletos estão organizados.
- **Ajuda**: Clique no botão "Ajuda" para obter informações sobre como usar o sistema.

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

## Contribuição

Se desejar contribuir para o projeto, sinta-se à vontade para fazer um fork do repositório, criar uma branch com suas melhorias e enviar um pull request. Todas as contribuições são bem-vindas!

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo LICENSE para mais detalhes.

## Contato

- **Gustavo Paz**
  - Telefone: 41995953862
  - E-mail: gust.nunes.paz@gmail.com
