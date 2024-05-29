import pandas as pd, os, datetime
from typing import Union

# Procurar:
#   -Cliente [2]
#   -Nome Cliente [3]
#   -Atribuição [4]
#   -Montante ME [9]
#   -Mont.(moeda trans) [11]
#   -Total

agora = datetime.datetime.now()
agoraData = agora.strftime('%d.%m.%Y')

try:
    os.remove('resultados.json')
except:
    pass

partidas: list[str] = [] # Lista de todos os arquivos a ser achado pelo robô
with open('Partidas.txt', 'r') as arquivo:
    for c in arquivo:
        partidas.append(c.replace('\n', ''))

list_dir = os.listdir(r'.\EntradaExcel')

customers = pd.read_excel(fr'.\EntradaExcel\{list_dir[0]}')

customers = customers.apply(lambda row: row[customers['Tipo partida indiv.'].isin(['DP', 'DZ'])])
customers = customers[customers[['Atribuição']].notnull().all(1)]

items_dir = {} # Lista excel dos items a virarem excel

item: tuple

for item in customers.itertuples():
    montante: float = abs(item[8])
    montantet: float = abs(item[9])
    chave: str = f'{item[4][:8]}'
    
    if item[4][:8] in partidas:
        if chave not in items_dir:
            items_dir[chave] = {'Empresa':item[1], 'Cliente':str(item[2]),'Nome_Empresa': item[3],'Atribuição': item[4][:8], 'Moeda_da_transação':item[10],
                                'Montante':montantet,'Desconto':montante,'Saldo em Aberto':montante-montantet,
                                'Data':agoraData,'NumPartidas':1,'Valido (Colocar 0 para não rodar/ Colocar 1 para o robô rodar)':1, 'Partidas':[item[12]]}
        else:
            if item[12] not in items_dir[chave]['Partidas']:
                items_dir[chave]['Partidas'].append(item[12])
            items_dir[chave]['Montante'] += montantet
            items_dir[chave]['Desconto'] += montante
            items_dir[chave].update({'Saldo em Aberto': f"{items_dir[chave]['Montante'] - items_dir[chave]['Desconto']:.2f}"},)
            items_dir[chave]['NumPartidas'] += 1

for item in items_dir:
    desconto, montante, saldoEmAberto = items_dir[item]['Desconto'], items_dir[item]['Montante'], items_dir[item]['Saldo em Aberto'] 
    if desconto == 0.0:
        if montante > 0:
            items_dir[item]['Saldo em Aberto'] = -abs(montante)
        else:
            items_dir[item]['Saldo em Aberto'] = abs(montante)
    elif montante > desconto:
        items_dir[item]['Saldo em Aberto'] = -abs(items_dir[item]['Saldo em Aberto'])
    elif desconto > montante:
        items_dir[item]['Saldo em Aberto'] = abs(items_dir[item]['Saldo em Aberto'])

menores: list[dict] = []

chave: str
valor: str | float | datetime.datetime | int
for chave, valor in items_dir.items():
    menores.append(valor)

excelFinal = pd.DataFrame(menores)
excelFinal = excelFinal.to_excel('Excel.xlsx',sheet_name='Partidas', index=0)

# os.remove(r'.\EntradaExcel\CustomerLineItems')