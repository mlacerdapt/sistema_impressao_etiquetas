import pandas as pd

# Carregar o arquivo Excel
file_path = r'\\srv-pt3\groups\02-Blades\02-Process Engineering\9. Projetos\4. Dashboard\Base de Dados\All - Por Ano\All_2025.XLSX'
work_center = 'E103_CUT' 

# Lê o arquivo Excel e carrega em um DataFrame
df = pd.read_excel(file_path)

# Filtra apenas as linhas com o Work Center do Corte
filtered_df = df[df['Work Center'] == work_center]

# Lista final para armazenar as etiquetas
tag_list = []

# Percorre cada linha filtrada
for index, row in filtered_df.iterrows():
    ordem = row['Order']
    serie = row['Serialnumber (PO Head)']
    quantidade = row['Operation Quantity (MEINH)']

    # Replica o número de série pela quantidade esperada
    for i in range(quantidade):
        tag_list.append({'Ordem': ordem, 'Número de Série': f'{serie}-{i+1}'})

# Transforma a lista em um DataFrame final
final_df = pd.DataFrame(tag_list)

# Salva em um novo arquivo Excel
final_df.to_excel('etiquetas_geradas.xlsx', index=False)

print('Arquivo de etiquetas gerado com sucesso!')
