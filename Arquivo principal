import pdfplumber
from openpyxl import Workbook
import re
import os
from datetime import datetime

# Obtém a lista de todos os arquivos PDF na mesma pasta
pasta_atual = os.getcwd()  # Pega o caminho da pasta atual
arquivos_pdf = [arquivo for arquivo in os.listdir(pasta_atual) if arquivo.endswith('.pdf')]

# Nome da planilha consolidada
data_atual = datetime.now().strftime('%Y-%m-%d')
nome_planilha = f"processos_de_{data_atual}.xlsx"

# Inicia a contagem de tempo
start_time = datetime.now()

# Cria a planilha
workbook = Workbook()
sheet = workbook.active
sheet.append(['Processo', 'Requerentes', 'Requeridos'])

for nome_arquivo in arquivos_pdf:
    print(f"Analisando o arquivo: {nome_arquivo}")
    
    # Abre o PDF
    with pdfplumber.open(nome_arquivo) as pdf:
        # Define a expressão regular para o número do processo
        processo_regex = re.compile(r'PROCESSO : (\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})')

        # Inicializa as variáveis
        processo = None
        classe = None
        requerentes = []
        requeridos = []

        # Loop pelas páginas do PDF
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')

            for line in lines:
                match_processo = processo_regex.search(line)
                if match_processo:
                    print("Processo encontrado:", match_processo.group(1))
                    if processo and classe == "BUSCA E APREENSÃO EM ALIENAÇÃO FIDUCIÁRIA":
                        requeridos.append("")  # Adiciona espaço em branco para REQDO/REQDA não encontrados
                        for reqte, req in zip(requerentes, requeridos):
                            sheet.append([processo, reqte, req])
                    processo = match_processo.group(1)
                    classe = None
                    requerentes = []
                    requeridos = []
                elif "CLASSE :" in line:
                    classe = line.replace("CLASSE :", "").strip()
                elif "REQTE :" in line:
                    requerentes.append(line.replace("REQTE :", "").strip())
                elif "REQDO :" in line or "REQDA :" in line:
                    requeridos.append(line.replace("REQDO :", "").replace("REQDA :", "").strip())

# Salva a planilha consolidada
workbook.save(nome_planilha)
print(f"Planilha consolidada criada: {nome_planilha}")

# Calcula o tempo de execução
end_time = datetime.now()
execution_time = end_time - start_time
print(f"Tempo de execução: {execution_time}")
