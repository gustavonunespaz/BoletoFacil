# BoletoFacil

BoletoFacil é um aplicativo para manipulação e geração automatizada de boletos da Conta Azul ERP com informações personalizadas para clientes, salvando os arquivos de acordo com a data de vencimento e outras informações importantes. O projeto utiliza `pdfplumber`, `reportlab`, `PyMuPDF`, e `PyInstaller` para oferecer uma solução robusta e prática de processamento em lote de boletos.

## Funcionalidades

- **Carregamento de múltiplos PDFs:** Permite processar vários boletos em uma única execução.
- **Extração de dados do cliente e de instruções:** Identifica e organiza informações como nome do cliente, endereço, data de vencimento e dados de pagamento.
- **Ocultação de blocos específicos de texto:** Aplica uma máscara em informações indesejadas do boleto original.
- **Estrutura automática de diretórios:** Salva os arquivos gerados organizados em pastas conforme ano e mês de vencimento do boleto.
- **Geração de executável:** O projeto pode ser transformado em um executável independente, sem a necessidade de Python instalado na máquina.
