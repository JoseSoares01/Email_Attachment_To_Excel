# Projeto de Extração de Dados de E-mails e Planilhas

Este projeto é um script em Python que faz o download de anexos de e-mails, extrai dados de arquivos Excel (XLSX) e salva esses dados em uma nova planilha. Inclui uma interface gráfica para facilitar o uso.

## Funcionalidades

1. **Download de Anexos de E-mails:**
   - Conecta-se a uma conta de e-mail usando o protocolo IMAP.
   - Baixa todos os anexos dos e-mails na caixa de entrada.

2. **Extração de Dados de Arquivos XLSX:**
   - Procura arquivos Excel em uma pasta especificada.
   - Extrai dados das planilhas e os organiza em uma lista de dicionários.

3. **Salvar Dados Extraídos:**
   - Salva os dados extraídos em um novo arquivo Excel usando o Pandas.

4. **Interface Gráfica do Usuário (GUI):**
   - Interface simples para inserção de credenciais e seleção de pastas.

## Requisitos

- Python 3.x
- Bibliotecas:
  - imaplib
  - email
  - openpyxl
  - pandas
  - PySimpleGUI

## Instalação

1. Clone o repositório:
   ```sh
   git clone https://github.com/seu-usuario/seu-repositorio.git
Navegue até o diretório do projeto:
sh
Copiar código
cd seu-repositorio
Instale as dependências:
sh
Copiar código
pip install -r requirements.txt
Uso
Execute o script principal:
sh
Copiar código
python seu_script.py
Insira suas credenciais de e-mail e escolha as pastas para download e processamento de arquivos.
Clique em "Salvar" para iniciar o processo.
Estrutura do Código
download_attachments
Função que faz o download dos anexos de e-mails.

extract_data_from_excel
Função que extrai dados de arquivos XLSX em uma pasta especificada.

save_to_excel
Função que salva os dados extraídos em um novo arquivo XLSX.

Interface Gráfica
Utiliza a biblioteca PySimpleGUI para criar uma interface simples.

Licença
Este projeto está licenciado sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.

css
Copiar código

### Código de Exemplo

Aqui está o código descrito anteriormente para referência.

```python
import imaplib
import email
from email.header import decode_header
import os
import zipfile
import openpyxl
import pandas as pd
import PySimpleGUI as sg

# Função para fazer download dos anexos
def download_attachments(email_user, email_pass, download_folder):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_user, email_pass)
    mail.select("inbox")

    status, messages = mail.search(None, 'ALL')
    email_ids = messages[0].split()

    for e_id in email_ids:
        res, msg = mail.fetch(e_id, "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    subject = subject.decode()
                print("Subject:", subject)
                
                if msg.is_multipart():
                    for part in msg.walk():
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                filepath = os.path.join(download_folder, filename)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"Downloaded {filename}")
    mail.logout()

# Função para extrair dados de arquivos XLSX
def extract_data_from_excel(folder_path):
    data = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):  # Ajuste conforme necessário
                data.append({
                    "Data": row[0],
                    "Entrada": row[1],
                    "Contabilizadas": row[2],
                    "Submetidas em PS": row[3],
                    "Erro": row[4]
                })
    return data

# Função para salvar dados extraídos em um novo arquivo XLSX
def save_to_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)
    print(f"Dados salvos em {output_path}")

# Tema da Interface Gráfica
sg.theme("DarkTeal12")

# Layout da Interface Gráfica
layout = [
    [sg.Text("Email"), sg.Input(key="email")],
    [sg.Text("Senha"), sg.Input(key="senha", password_char="*")],
    [sg.FolderBrowse("Escolher Pasta de Download", target="download_folder"), sg.Input(key="download_folder")],
    [sg.FolderBrowse("Escolher Pasta Planilha", target="input_planilha"), sg.Input(key="input_planilha")],
    [sg.Button("Salvar")]
]

window = sg.Window("principal", layout)

# Loop para não fechar o sistema automaticamente
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "Salvar":
        email_user = values["email"]
        email_pass = values["senha"]
        download_folder = values["download_folder"]
        output_folder = values["input_planilha"]
        
        # Fazendo download dos anexos
        download_attachments(email_user, email_pass, download_folder)
        
        # Extraindo dados dos arquivos XLSX
        extracted_data = extract_data_from_excel(download_folder)
        
        # Salvando dados extraídos em uma nova planilha
        output_path = os.path.join(output_folder, "dados_extraidos.xlsx")
        save_to_excel(extracted_data, output_path)

window.close()
Considerações Finais
Ao criar o seu repositório no GitHub, inclua o arquivo README.md no diretório raiz. Isso ajudará os visitantes do seu repositório a entenderem rapidamente o propósito do seu projeto, como usá-lo e quais são suas funcionalidades principais. Além disso, mantenha o código organizado e documentado para facilitar a manutenção e futuras atualizações.

consoante a isto deem um nome em ingles para este repositorio
ChatGPT
Para o seu repositório no GitHub, você pode escolher um nome em inglês que reflita claramente as principais funcionalidades do projeto. Aqui estão algumas sugestões:

EmailAttachmentExtractor
ExcelDataDownloader
EmailToExcelProcessor
AttachmentDataExtractor
InboxAttachmentManager
EmailAttachmentToExcel
IMAPAttachmentDownloader
ExcelDataAggregator
EmailAttachmentHandler
ExcelDataExtractorTool
Dentre essas opções, EmailToExcelProcessor ou EmailAttachmentToExcel seriam nomes particularmente descritivos e apropriados, pois indicam claramente que o projeto envolve a extração de anexos de e-mails e o processamento de dados de Excel.

Atualização do README.md com o Nome do Projeto
markdown
Copiar código
# EmailAttachmentToExcel

This project is a Python script that downloads email attachments, extracts data from Excel files (XLSX), and saves this data into a new spreadsheet. It includes a graphical user interface to facilitate user interaction.

## Features

1. **Download Email Attachments:**
   - Connects to an email account using the IMAP protocol.
   - Downloads all attachments from emails in the inbox.

2. **Extract Data from XLSX Files:**
   - Searches a specified folder for Excel files.
   - Extracts data from the spreadsheets and organizes it into a list of dictionaries.

3. **Save Extracted Data:**
   - Saves the extracted data into a new Excel file using Pandas.

4. **Graphical User Interface (GUI):**
   - Simple interface for inputting credentials and selecting folders.

## Requirements

- Python 3.x
- Libraries:
  - imaplib
  - email
  - openpyxl
  - pandas
  - PySimpleGUI

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/your-username/EmailAttachmentToExcel.git
