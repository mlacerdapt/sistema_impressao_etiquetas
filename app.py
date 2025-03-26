import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Carregar o arquivo Excel
file_path = r'\\srv-pt3\groups\02-Blades\02-Process Engineering\9. Projetos\4. Dashboard\Base de Dados\All - Por Ano\All_2025.XLSX'
work_center = 'E103_CUT' 
# Lê o arquivo Excel e carrega em um DataFrame
df = pd.read_excel(file_path)

# Remove espaços extras nas colunas
df.columns = df.columns.str.strip()


# Filtra apenas as linhas com o Work Center do Corte
filtered_df = df[df['Work Center'] == work_center]

# Lista final para armazenar as etiquetas
tag_list = []

# Percorre cada linha filtrada
for index, row in filtered_df.iterrows():
    ordem = row['Order']  # Ajuste conforme o nome correto da coluna
    serie = row['Serialnumber (PO Head)']
    quantidade = row['Operation Quantity (MEINH)']
    material = row['Material']
    descricao = row['Material Description']

    # Extrai o prefixo e o número final da série
    prefix = ''.join(filter(str.isalpha, serie))
    start_number = int(''.join(filter(str.isdigit, serie)))

    # Gera as séries sequenciais com 4 dígitos fixos
    for i in range(quantidade):
        numero_serie_atual = f'{prefix}{start_number + i:04d}'
        tag_list.append({'Ordem': ordem, 'Número de Série': numero_serie_atual, 'Material': material, 'Descrição': descricao})

# Transforma a lista em um DataFrame final
final_df = pd.DataFrame(tag_list)

# Salva em um novo arquivo Excel
final_df.to_excel('etiquetas_geradas.xlsx', index=False)

print('Arquivo de etiquetas gerado com sucesso!')

