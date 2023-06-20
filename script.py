import pandas as pd
import os

cols = ['TX_NUM_EQUIPAMENTO', 'Num_Operacional', 'Ident_Equip_Def_ONS', 'Tensao_Nominal', 'Sigla_Subestacao','ONDA']
directory = r'C:\Users\Matheus\OneDrive\Nuvem\ELSC\PROJETOS\SGBDIT - CHESF - PREENCHIMENTO PLANILHAS\004-DB-CHESF'

print("------------------ | Script iniciado com sucesso | ------------------")

results = pd.DataFrame()

for i in os.listdir(directory):
    file = os.path.join(directory, i)
    df = pd.read_excel(file, usecols=cols, sheet_name=0, skiprows=1)
    df['Nome_Planilha'] = i  # Adiciona uma coluna com o nome da planilha
    results = pd.concat([results, df])
    print("Os dados da planilha |" + i + "| foram carregados com sucesso para planilha de relatório.")

results.to_excel(os.path.join(directory, "results.xlsx"))
print("Dados combinados salvos com sucesso!")