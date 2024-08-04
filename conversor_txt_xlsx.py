from platform import python_version
#pip install openpyxl #https://openpyxl.readthedocs.io/en/stable/ acesse pa entender a biblioteca  remova o '#' para instalar  apos instalar comnte novamente
#pip install xlsxwriter #https://pypi.org/project/XlsxWriter/ acesse para entender a biblioteca remova o '#' para instalar  apos instalar comnte novamente
print(f'A versão em uso é {python_version()}')

import re
import numpy as np #https://numpy.org/
import pandas as pd #https://pandas.pydata.org/
import datetime as dt #https://docs.python.org/3/library/datetime.html

#Definindo a variável content com o conteúdo do seu arquivo
with open('01-2024.txt', 'r', encoding='utf-8') as file:
    content = file.read() # lendo o aquivo de texto

# Função para extrair informações com base em padrões regex
def extract_info(section, pattern):
  match = re.search(pattern, section)
  return match.group(1).strip() if match else None

# Dividindo o conteúdo em seções com base no 'Empregador: ' para tratar várias seções separadamente
sections = content.split('Empregador: ')[1:]

# Processando cada seção em um DataFrame
dataframes = []

for section in sections:
    lines = section.split('\n')

    # Extraindo informações comuns da seção
    empregador_info = 'Empregador: ' + extract_info(section, r'(.+?) Cartão de ponto')
    cnpj = extract_info(section, r'CNPJ: (.+?)\s')
    endereco = extract_info(section, r'End: (.+?) Competência')

    # Verificando se a linha da competência existe
    competencia_line = lines[3] if len(lines) > 3 else ''
    competencia = competencia_line.split(': ')[1] if ':' in competencia_line else None

    unidade = extract_info(section, r'Unidade: (.+?) Data de admissão')
    admissao = extract_info(section, r'Data de admissão (.+)')
    funcionario_id_nome = extract_info(section, r'Funcionário: (\d+ - .+) Cargo:')
    cargo = extract_info(section, r'Cargo: (.+)')

    if funcionario_id_nome:
        funcionario_id, funcionario_nome = funcionario_id_nome.split(' - ')
    else:
        funcionario_id, funcionario_nome = None, None

    # Extraindo registros
    records_start = section.find('Registros Crédito Débitos Faltas e Atrasos') + len('Registros Crédito Débitos Faltas e Atrasos') + 1
    records_end = section.find('Crédito', records_start)
    records = section[records_start:records_end].strip().split('\n')

    # Analisando registros em uma lista de dicionários
    data = []
    for record in records:
        parts = record.split()
        if len(parts) < 14:  # Ignorar registros com elementos ausentes
            continue
        date = parts[0]
        work_times = ' '.join(parts[1:6])
        work_type = ' '.join(parts[6:-6])
        credit = parts[-6]
        debit = parts[-5]
        faltas_atrasos = parts[-4]
        data.append({
            'Empregador': empregador_info,
            'CNPJ': cnpj,
            'Endereço': endereco,
            'Competência': competencia,
            'Unidade': unidade,
            'Data de admissão': admissao,
            'Funcionário ID': funcionario_id,
            'Funcionário Nome': funcionario_nome,
            'Cargo': cargo,
            'Data': date,
            'Horários de Trabalho': work_times,
            'Tipo de Trabalho': work_type,
            'Crédito': credit,
            'Débito': debit,
            'Faltas e Atrasos': faltas_atrasos
        })

    # Convertendo a lista de dicionários em um DataFrame
    df = pd.DataFrame(data)
    dataframes.append(df)

# Combinando todos os DataFrames em um
combined_df = pd.concat(dataframes, ignore_index=True)

#dividindo a coluna Horários  de Trabalho em tres colunas
combined_df[['Sáída', 'Saida Almoço', 'Retorno Almoço']] = combined_df['Horários de Trabalho'].str.split('-', expand=True)

# Removendo o caractere "-" no início de cada linha na coluna 'Tipo de Trabalho'
combined_df['Tipo de Trabalho'] = combined_df['Tipo de Trabalho'].str.lstrip('-')
combined_df['Tipo de Trabalho'] = combined_df['Tipo de Trabalho'].str.replace('/', '')

# Dividindo a coluna Tipo de Trabalho em duas colunas
combined_df[['Entrada', 'Tipo de Trabalho']] = combined_df['Tipo de Trabalho'].str.split(' ', n=1, expand=True)

# Extrair os horários da coluna 'Tipo de Trabalho'
combined_df['Entrada'] = combined_df['Tipo de Trabalho'].str.extract(r'(\d{2}:\d{2})')

# Formatando os dados da tabela do tipo entrada para O TIPO HORAS
#combined_df['Entrada'] = pd.to_datetime(combined_df['Entrada'], format='%H:%M', errors='coerce').dt.time

# Excluido colunas Horários de Trabalho e Tipo de Trabalho
combined_df.drop(columns=['Horários de Trabalho', 'Tipo de Trabalho'], inplace=True)

# #Salvando as alterações e convertendo em um arquivo
output_file_path = 'c:/Users/adria/Desktop/laboratório python/01-2024.xlsx'
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    combined_df.to_excel(writer, index=False, sheet_name='Data')
    worksheet = writer.sheets['Data']
    for col_num, value in enumerate(combined_df.columns.values):
        worksheet.write(0, col_num, value)
        worksheet.set_column(col_num, col_num, 20)
    worksheet.autofilter(0, 0, len(combined_df), len(combined_df.columns) - 1)

print(f"Arquivo Excel salvo em: {output_file_path}")