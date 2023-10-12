import pandas as pd

def xlsx_to_dict(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    data_dict = df.to_dict(orient='records')
    return data_dict

file_path = r"C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\Planilha.xlsx"  # Substitua pelo caminho do seu arquivo XLSX
sheet_name = "laginha"  # Substitua pelo nome da planilha que deseja converter

result = xlsx_to_dict(file_path, sheet_name)

#printa a primeira linha do dicionário
#print(result[0])
#printa só a área da fazenda da primeira linha do dicionário
#print(result[0]['area'])

VALOR_RL = (result[0]['reserva_legal'])*0.10

print(VALOR_RL)

AREA_APP_fl = result[0]['app']
AREA_APP = locale .format_string("%.2f", AREA_APP_fl, grouping=True)

AREA_RL_fl = float(result[0]['reserva_legal'])*0.10
AREA_RL = locale.format_string("%.2f", AREA_RL_fl, grouping=True)

AREA_BENFEITORIAS = result[0]['benfeitoria']
AREA_BENFEITORIAS_FORMATADA = locale.format_string("%.2f", float(AREA_BENFEITORIAS), grouping=True)


AREA_TRIBUTAVEL = float(AREA_TOTAL) - (AREA_APP_fl + AREA_RL_fl)

AREA_APROVEITAVEL_fl = float(AREA_TRIBUTAVEL) - float(AREA_BENFEITORIAS)
AREA_APROVEITAVEL = locale.format_string("%.2f", AREA_APROVEITAVEL_fl, grouping=True)

VALOR_IMOVEL_fl = float(AREA_TOTAL) * 10830
VALOR_IMOVEL = locale.format_string("%.2f", VALOR_IMOVEL_fl, grouping=True)

VALOR_BENFEITORIA_fl = float(AREA_BENFEITORIAS) * 10830
VALOR_BENFEITORIA = locale.format_string("%.2f", VALOR_BENFEITORIA_fl, grouping=True)

VALOR_TERRANUA_fl = float(VALOR_IMOVEL_fl) - float(VALOR_BENFEITORIA_fl)
VALOR_TERRANUA = locale.format_string("%.2f", VALOR_TERRANUA_fl, grouping=True)
