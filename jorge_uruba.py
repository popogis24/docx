import docx
import locale
from os import listdir
from os.path import isfile, join
from docxtpl import DocxTemplate
import pandas as pd

def xlsx_to_dict(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    data_dict = df.to_dict(orient='records')
    return data_dict

file_path = r"C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\Planilha.xlsx"  # Substitua pelo caminho do seu arquivo XLSX
sheet_name = "uruba"  # Substitua pelo nome da planilha que deseja converter

result = xlsx_to_dict(file_path, sheet_name)

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
#seleciona a pasta C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_Desatualizados\GUAXUMA
mypath = fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\URUBA\PDF_Desatualizados'

#lista os pdfs dessa pasta
list_pdf = [f for f in listdir(mypath) if isfile(join(mypath, f))]
#inputs
seq_global = 0
for pdf in range(len(list_pdf)):
    
    TITULO_PRINCIPAL = result[seq_global]['Fazenda'].upper()
    AREA_TOTAL = result[seq_global]['area']
    AREA_TOTAL_FORMATADA = locale.format_string("%.2f", float(AREA_TOTAL), grouping=True)
    NOME_FAZENDA1 = TITULO_PRINCIPAL
    NM_MUN = result[seq_global]['mun'].title()
    COD_INCRA = result[seq_global]['COD_INCRA']
    ITR_NIRF =  result[seq_global]['ITR']
    if NM_MUN == 'Campo Alegre' or NM_MUN == 'Coruripe':
        REF_MF = '30'
    elif NM_MUN == 'Junqueiro' or NM_MUN == 'São Sebastião' or NM_MUN == 'Teotônio Vilela':
        REF_MF = '35'
    elif NM_MUN == 'Atalaia' or NM_MUN == 'União dos Palmares' or NM_MUN == 'Branquinha' or NM_MUN == 'Murici' or NM_MUN == 'Capela':
        REF_MF = '16'
    elif NM_MUN == 'Pilar' or NM_MUN == 'Marechal Deodoro' or NM_MUN == 'Rio Largo':
        REF_MF = '12'
    MOD_FISCAL = float(AREA_TOTAL) / float(REF_MF)
    MOD_FISCAL_FORMATADO = locale.format_string("%.2f", MOD_FISCAL, grouping=True)
    if MOD_FISCAL < 4:
        TAMANHO_PROPRIEDADE = 'Pequena Propriedade Produtiva'
    elif MOD_FISCAL >= 4 and MOD_FISCAL < 15:
        TAMANHO_PROPRIEDADE = 'Média Propriedade Produtiva'
    elif MOD_FISCAL >= 15:
        TAMANHO_PROPRIEDADE = 'Grande Propriedade Produtiva'
    
    
    #DESC_ACESSO
    #LOC_ACESSO_Guaxuma = 'A partir de Maceió, segue a AL-101 em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Continuando na AL-101 após cruzar a AL-420, seguindo até chegar ao Município de Coruripe, percorrendo um total de 93,5 km.'
    #LOC_ACESSO_Uruba = 'Saindo de Maceió, a rota segue pelo litoral sul em direção ao município de Atalaia, atravessando a AL-101. Após a interseção com a BR-316, a estrada continua por essa rodovia até chegar ao mencionado município, cobrindo uma distância de 49,4 km. Após adentrar o território de Atalaia, o trajeto segue pela AL-410 e, subsequentemente, pela AL-210, percorrendo 14,7 km até alcançar a Usina Uruba.'
    #LOC_ACESSO_Laginha = 'Saindo de Maceió, a rota segue pela BR 104, passando pelos municípios de Messias, Murici e Branquinha. Após cruzar Branquinha, o percurso se estende por mais 12 km até alcançar o município de União dos Palmares, totalizando 79 km.'
    if NM_MUN == 'Coruripe':
        LOC_ACESSO = 'Saindo de Maceió pela AL-101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Após passar pelo cruzamento com a AL 420 segue ainda pela AL 101 até chegar no Município de Coruripe, percorrendo o total de 87,3 km.'
    elif NM_MUN == 'Campo Alegre':
        LOC_ACESSO = 'Saindo de Maceió pela AL-101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. E em seguida pela AL-220 até chegar no Município de Campo Alegre, percorrendo o total de 89,4 km.'
    elif NM_MUN == 'Junqueiro':
        LOC_ACESSO = 'Saindo de Maceió pela AL-101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Em seguida AL-220 até chegar na BR-101 na qual permanece até o município de Junqueiro, percorrendo o total de 112 km.'
    elif NM_MUN == 'Teotônio Vilela':
        LOC_ACESSO = 'Saindo de Maceió pela AL-101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Em seguida AL-220 até chegar na BR-101 na qual permanece até o município de Teotônio Vilela, percorrendo o total de 96,9 km.'
    elif NM_MUN == 'São Sebastião':
        LOC_ACESSO = 'Saindo de Maceió pela AL-101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Em seguida AL-220 até chegar na BR-101 na qual permanece até o município de São Sebastião, percorrendo o total de 126 km.'
    elif NM_MUN == 'Atalaia':
        LOC_ACESSO = 'Partindo de Maceió e seguindo pelo litoral sul em direção ao município de Atalaia. Em seguida pela AL-101. Após alcançar a BR-316, continua pela mesma rodovia até chegar a Atalaia, percorrendo uma distância de 49,4 km.'
    elif NM_MUN == 'Pilar':
        LOC_ACESSO = 'Saíndo de Maceió pela AL-101 segue em direção a BR-424 e em seguida a BR-316 até chegar no município de Pilar, percorrendo o total de 35,4 km.'
    elif NM_MUN == 'Marechal Deodoro':
        LOC_ACESSO = 'Saindo de Maceió pela AL-101 segue em direção sul até chegar na bifurcação com a AL-215, na qual permanece até encontrar com o município de Marechal Deodoro, percorrendo o total de 28 km.'
    elif NM_MUN == 'Rio Largo':
        LOC_ACESSO = 'Partindo de Maceió pela AL-104, segue em direção ao município de Rio Largo, percorrendo o total de 27,3 km.'
    elif NM_MUN == 'União dos Palmares':
        LOC_ACESSO = 'Partindo de Maceió pela AL-104, segue em direção ao município de Rio Largo, posteriormente passando pelo município de Murici e Branquinha, para então chegar no município de União dos Palmares, percorrendo o total de 77,8 km.'
    elif NM_MUN == 'Branquinha':
        LOC_ACESSO = 'Partindo de Maceió pela AL-104, segue em direção ao município de Rio Largo, posteriormente passando pelo município de Murici para então chegar no município de Branquinha, percorrendo o total de 65,7 km.'
    elif NM_MUN == 'Murici':
        LOC_ACESSO = 'Partindo de Maceió pela AL-104, segue em direção ao município de Rio Largo, posteriormente chegando no município de Murici, percorrendo o total de 53,2 km.'
    elif NM_MUN == 'Capela':
        LOC_ACESSO = 'Partindo de Maceió pela AL-101, segue em direção a bifurcação com a BR-424, posteriormente passando pelo município de Pilar e Atalaia utilizando a BR-316. Por fim chegando a Capela utilizando as rodovias AL-410 e AL-210, percorrendo o total de 61,2 km.'

    #DESC_CLIMA
    if NM_MUN == 'Coruripe':
        DESC_CLIMA = 'Coruripe possui um clima tropical caracterizado por temperaturas quentes ao longo do ano, com uma estação seca e uma estação chuvosa bem definidas. A chuva é mais abundante durante a estação chuvosa, e a umidade relativa do ar é geralmente alta, tornando a região propícia para a agricultura. Coruripe tem uma temperatura média de 24.4 °C. A média anual de pluviosidade é de 1372 mm.'
    elif NM_MUN == 'Campo Alegre':
        DESC_CLIMA = 'Campo Alegre possui clima quente e temperado. Com uma pluviosidade considerável durante o ano. A temperatura média é de 16.3 °C. A média anual de pluviosidade é de 1514 mm.'
    elif NM_MUN == 'Junqueiro':
        DESC_CLIMA = 'Junqueiro possui um clima tropical, com muito mais pluviosidade no inverno do que no verão. Além disso, possuí uma temperatura média de 23.7 °C. A média anual de pluviosidade é de 1034 mm.'
    elif NM_MUN == 'Teotônio Vilela':
        DESC_CLIMA = 'Teotônio Vilela possui um clima tropical, com muito mais pluviosidade no inverno do que no verão. Além disso, possuí uma temperatura média de 24.0 °C. A média anual de pluviosidade é de 1134 mm.'
    elif NM_MUN == 'São Sebastião':
        DESC_CLIMA = 'São Sebastião possui um clima tropical, com muito menos pluviosidade no inverno do que no verão. Além disso, possuí uma temperatura média de 23.8 °C. A média anual de pluviosidade é de 953 mm.'
    elif NM_MUN == 'Atalaia':
        DESC_CLIMA = 'Atalaia possui um clima tropical, com muito mais pluviosidade no inverno do que no verão. Além disso, possuí uma temperatura média de 24.7 °C. A média anual de pluviosidade é de 1428 mm.'
    elif NM_MUN == 'Pilar':
        DESC_CLIMA = 'O clima na região é caracterizado como tropical atlântico, o que é típico do litoral. Os meses mais quentes abrangem o período de novembro a março, quando as temperaturas podem atingir até 36ºC. Durante a estação menos quente, que geralmente ocorre de maio a agosto, as temperaturas mínimas raramente caem abaixo de 20 °C.'
    elif NM_MUN == 'Marechal Deodoro':
        DESC_CLIMA = 'Em Marechal Deodoro, a estação do verão se estende por um período longo, caracterizado por temperaturas elevadas e um céu parcialmente encoberto. O inverno é breve, com temperaturas amenas, ocorrência de precipitação e um céu quase sem nuvens. Ao longo do ano, a faixa de temperatura geralmente se situa entre 21 °C e 32 °C, raramente caindo abaixo de 19 °C ou ultrapassando os 33 °C.'
    elif NM_MUN == 'Rio Largo':
        DESC_CLIMA = 'Em Rio Largo, a estação do verão é prolongada, caracterizada por temperaturas elevadas e um céu parcialmente encoberto. O inverno é breve, com temperaturas amenas, a ocorrência de precipitação e um céu quase desprovido de nuvens. Durante todo o ano, a atmosfera é frequentemente abafada, acompanhada por ventos fortes. Em geral, a variação de temperatura ao longo do ano se situa entre 20 °C e 32 °C, sendo raro que as temperaturas caiam abaixo de 18 °C ou ultrapassem os 34 °C.'
    elif NM_MUN == 'União dos Palmares':
        DESC_CLIMA = 'Em União dos Palmares, o verão se estende por um período prolongado, trazendo calor intenso e um céu quase sempre com alguma cobertura de nuvens. Por outro lado, o inverno é breve, porém agradável, caracterizado por chuvas ocasionais, ventos vigorosos e um céu geralmente limpo. O clima ao longo do ano é constantemente abafado. As temperaturas na região costumam variar entre 19 °C e 33 °C, raramente caindo abaixo de 17 °C ou ultrapassando os 36 °C.'
    elif NM_MUN == 'Branquinha':
        DESC_CLIMA = 'Em Branquinha, o verão se estende por um período prolongado, trazendo calor intenso e um céu quase sempre com alguma cobertura de nuvens. Por outro lado, o inverno é breve, porém agradável, caracterizado por chuvas ocasionais, ventos vigorosos e um céu geralmente limpo. O clima ao longo do ano é constantemente abafado. As temperaturas na região costumam variar entre 19 °C e 33 °C, raramente caindo abaixo de 17 °C ou ultrapassando os 36 °C.'
    elif NM_MUN == 'Murici':
        DESC_CLIMA = 'Em Murici, a estação do verão se estende por um período prolongado, trazendo calor intenso e frequentemente apresentando um céu parcialmente encoberto. Já o inverno é breve, porém agradável, caracterizado por precipitações e um céu quase desprovido de nuvens. Ao longo de todo o ano, o clima é constantemente opressivo e marcado por ventos fortes. As temperaturas na região geralmente variam de 19 °C a 33 °C, raramente caindo abaixo de 18 °C ou ultrapassando os 35 °C.'
    elif NM_MUN == 'Capela':
        DESC_CLIMA = 'Em Capela, o verão se estende por um longo período, trazendo calor intenso e geralmente apresentando um céu parcialmente encoberto. Enquanto isso, o inverno é breve, com temperaturas amenas, ocorrência de precipitação, ventos fortes e um céu quase desprovido de nuvens. Ao longo de todo o ano, o clima é consistentemente opressivo. As temperaturas na região tendem a variar, em média, de 21 °C a 33 °C, raramente caindo abaixo de 18 °C ou ultrapassando os 35 °C.'

    if NM_MUN == 'Campo Alegre':
        DESC_MICRO = ' Microrregião São Miguel dos Campos (INCRA), composta pelos Municípios de Anadia, São Miguel dos Campos, Junqueiro, Teotônio Vilela, Jequiá da Praia, Coruripe, Campo Alegre e Taquarana.'
        POP = ' 50.831 (IBGE 2010)'
        AREA = ' 308,058 km²'
        FUNDACAO = ' 1960'
        ALTITUDE = ' 106 m'
        DIST = ' 68 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 154.813,711 (IBGE 2008)'
        RENDA = ' R$ 3.317,13 (IBGE 2008)'
        IDH = ' 0,595 (PNUD - 2000)'
        TEMP = ' 26º C'
    elif NM_MUN == 'Coruripe':
        DESC_MICRO = ' Microrregião São Miguel dos Campos (INCRA), composta pelos Municípios de Anadia, São Miguel dos Campos, Junqueiro, Teotônio Vilela, Jequiá da Praia, Coruripe, Campo Alegre e Taquarana.'
        POP = ' 52.160 (IBGE 2010)'
        AREA = ' 912,716 km²'
        FUNDACAO = ' 1850'
        ALTITUDE = ' 16 m'
        DIST = ' 85 km'
        ECONOMIA = ' Comércio, turismo e agricultura.'
        PIB = ' R$ 450.151,209 (IBGE 2008)'
        RENDA = ' R$ 8.560,61 (IBGE 2008)'
        IDH = ' 0,626 (PNUD - 2000)'
        TEMP = ' 24º C'
    elif NM_MUN == 'Junqueiro':
        DESC_MICRO = ' Microrregião São Miguel dos Campos (INCRA), composta pelos Municípios de Anadia, São Miguel dos Campos, Junqueiro, Teotônio Vilela, Jequiá da Praia, Coruripe, Campo Alegre e Taquarana.'
        POP = ' 23.854 (IBGE 2010)'
        AREA = ' 254,067 km²'
        FUNDACAO = ' 1947'
        ALTITUDE = ' 146 m'
        DIST = ' 118 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 411.588.411,00 (IBGE 2008)'
        RENDA = ' R$ 17.216,23 (IBGE 2008)'
        IDH = ' 0,615 (PNUD - 2000)'
        TEMP = ' 26º C'
    elif NM_MUN == 'Teotonio Vilela':
        DESC_MICRO = ' Microrregião São Miguel dos Campos (INCRA), composta pelos Municípios de Anadia, São Miguel dos Campos, Junqueiro, Teotônio Vilela, Jequiá da Praia, Coruripe, Campo Alegre e Taquarana.'
        POP = ' 44.666 (IBGE 2010)'
        AREA = ' 297,875 km²'
        FUNDACAO = ' 1955'
        ALTITUDE = ' 156 m'
        DIST = ' 101 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 162.502,262 (IBGE 2008)'
        RENDA = ' R$ 3.915,91 (IBGE 2008)'
        IDH = ' 0,564 (PNUD - 2000)'
        TEMP = ' 26º C'
    elif NM_MUN == 'Sao Sebastiao': 
        DESC_MICRO = ' Microrregião Arapiraca (INCRA), composta pelos Municípios de Arapiraca, Campo Grande, Coité do Nóia, Craíbas, Feira Grande, Girau do Ponciano, Lagoa da Canoa, Limoeiro de Anadia, São Sebastião e Taquarana.'
        POP = ' 32.007 (IBGE 2010)'
        AREA = ' 305,746 km²'
        FUNDACAO = ' 1755'
        ALTITUDE = ' 200 m'
        DIST = ' 100 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 113.188,037 (IBGE 2008)'
        RENDA = ' R$ 3.545,77 (IBGE 2008)'
        IDH = ' 0,655 (PNUD - 2000)'
        TEMP = ' 26º C'
    elif NM_MUN == 'Atalaia': 
        DESC_MICRO = ' Microrregião da Mata Alagoana 1, composta pelos Municípios de: Atalaia, Cajueiro, Capela, Chã Preta e Viçosa.'
        POP = ' 47.298 (IBGE 2015)'
        AREA = ' 531,983 km²'
        FUNDACAO = ' 1764'
        ALTITUDE = ' 54 m'
        DIST = ' 48 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 202.285,031 (IBGE 2008)'
        RENDA = ' R$ 3.897,37 (IBGE 2008)'
        IDH = ' 0,594 (PNUD - 2000)'
        TEMP = ' 23º a 34º C'
    elif NM_MUN == 'Pilar': 
        DESC_MICRO = ' Microrregião de Maceió entorno, composta pelos Municípios de: Messias, Rio Largo, Pilar, Coqueiro Seco, Satuba e Santa Luzia do Norte.'
        POP = ' 33.312 (IBGE 2015)'
        AREA = ' 248,975 km²'
        FUNDACAO = ' 1872'
        ALTITUDE = ' 13 m'
        DIST = ' 36 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 211.186,404 (IBGE 2008)'
        RENDA = ' R$ 6.488,86 (IBGE 2008)'
        IDH = ' 0,610 (PNUD - 2000)'
        TEMP = ' 20º a 36º C'
    elif NM_MUN == 'Marechal Deodoro': 
        DESC_MICRO = ' Microrregião de Maceió, composta pelos Municípios de: Marechal Deodoro, Maceió, Paripueira e Barra de São Miguel.'
        POP = ' 45.994 (IBGE 2015)'
        AREA = ' 333,548 km²'
        FUNDACAO = ' 1636'
        ALTITUDE = ' 5 m'
        DIST = ' 28 km'
        ECONOMIA = ' Comércio, turismo e agricultura.'
        PIB = ' R$ 975.899 (IBGE 2008)'
        RENDA = ' R$ 20.543,52 (IBGE 2008)'
        IDH = ' 0,642 (PNUD - 2000)'
        TEMP = ' 20º a 36º C'
    elif NM_MUN == 'Rio Largo': 
        DESC_MICRO = ' Microrregião de Maceió entorno, composta pelos Municípios de: Messias, Rio Largo, Pilar, Coqueiro Seco, Satuba e Santa Luzia do Norte.'
        POP = ' 75.662(IBGE 2021)'
        AREA = ' 299,110 km²'
        FUNDACAO = ' 1938'
        ALTITUDE = ' 130 m'
        DIST = ' 27 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 734,303 (IBGE 2014)'
        RENDA = ' R$ 9.755 (IBGE 2014)'
        IDH = ' 0,643 (PNUD - 2010)'
        TEMP = ' 20º a 36º C'
    elif NM_MUN == 'Uniao dos Palmares': 
        DESC_MICRO = ' Microrregião Serrana dos Quilombos (INCRA), composta pelos Municípios de União dos Palmares, São José da Lage, Ibateguara, Santana do Mundaú e Branquinha'
        POP = ' 65.790(IBGE 2021)'
        AREA = ' 427,825 km²'
        FUNDACAO = ' 1831'
        ALTITUDE = ' 155 m'
        DIST = ' 73 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 313.855,010 (IBGE 2008)'
        RENDA = ' R$ 5030,13 (IBGE 2008)'
        IDH = ' 0,593 (PNUD - 2010)'
        TEMP = ' 24º C'
    elif NM_MUN == 'Branquinha':
        DESC_MICRO = ' Microrregião Serrana dos Quilombos (INCRA), composta pelos Municípios de União dos Palmares, São José da Lage, Ibateguara, Santana do Mundaú e Branquinha'
        POP = ' 9.603 (IBGE 2021)'
        AREA = ' 168,048 km²'
        FUNDACAO = ' 1962'
        ALTITUDE = ' 90 m'
        DIST = ' 64 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 66.043,78 (IBGE 2010)'
        RENDA = ' R$ 6.132 (IBGE 2010)'
        IDH = ' 0,513 (PNUD - 2000)'
        TEMP = ' 24º C'
    elif NM_MUN == 'Murici':
        DESC_MICRO = ' Microrregião Serrana dos Quilombos (INCRA), composta pelos Municípios de União dos Palmares, São José da Lage, Ibateguara, Santana do Mundaú e Branquinha'
        POP = ' 26.706 (IBGE 2014)'
        AREA = ' 423,983 km²'
        FUNDACAO = ' 1872'
        ALTITUDE = ' 82 m'
        DIST = ' 48 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 104 283,832 (IBGE 2008)'
        RENDA = ' R$ 3.901,23 (IBGE 2010)'
        IDH = ' 0,580 (PNUD - 2000)'
        TEMP = ' 24,5º C'
    elif NM_MUN == 'Capela':
        DESC_MICRO = ' Microrregião Serrana dos Quilombos (INCRA), composta pelos Municípios de União dos Palmares, São José da Lage, Ibateguara, Santana do Mundaú e Branquinha'
        POP = ' 17.077 (IBGE 2014)'
        AREA = ' 205,283 km²'
        FUNDACAO = ' 1890'
        ALTITUDE = ' 84 m'
        DIST = ' 60,8 km'
        ECONOMIA = ' Comércio e agricultura.'
        PIB = ' R$ 67.854,563 (IBGE 2008)'
        RENDA = ' R$ 3.876,52 (IBGE 2008)'
        IDH = ' 0,569 (PNUD - 2000)'
        TEMP = ' 24º C'

    AREA_APP_fl = result[seq_global]['app']
    AREA_APP = locale .format_string("%.2f", AREA_APP_fl, grouping=True)

    AREA_RL_fl = float(AREA_TOTAL)*0.2
    AREA_RL = locale.format_string("%.2f", AREA_RL_fl, grouping=True)

    AREA_BENFEITORIAS = result[seq_global]['benfeitoria']
    AREA_BENFEITORIAS_FORMATADA = locale.format_string("%.2f", float(AREA_BENFEITORIAS), grouping=True)


    AREA_TRIBUTAVEL_fl = float(AREA_TOTAL) - (AREA_APP_fl + AREA_RL_fl)
    AREA_TRIBUTAVEL = locale.format_string("%.2f", AREA_TRIBUTAVEL_fl, grouping=True)

    AREA_APROVEITAVEL_fl = float(AREA_TRIBUTAVEL_fl) - float(AREA_BENFEITORIAS)
    AREA_APROVEITAVEL = locale.format_string("%.2f", AREA_APROVEITAVEL_fl, grouping=True)


    VALOR_BENFEITORIA_fl = result[seq_global]['valor_benfeitoria']
    VALOR_BENFEITORIA = locale.format_string("%.2f", VALOR_BENFEITORIA_fl, grouping=True)

    VALOR_CULTIVO_fl = result[seq_global]['cultivo']
    VALOR_CULTIVO = locale.format_string("%.2f", VALOR_CULTIVO_fl, grouping=True)

    VALOR_TERRANUA_fl = float(AREA_TOTAL) * 8778.62
    VALOR_TERRANUA = locale.format_string("%.2f", VALOR_TERRANUA_fl, grouping=True)
    
    VALOR_IMOVEL_fl = float(VALOR_TERRANUA_fl) + float(VALOR_BENFEITORIA_fl) + float(VALOR_CULTIVO_fl)
    VALOR_IMOVEL = locale.format_string("%.2f", VALOR_IMOVEL_fl, grouping=True)

    
    #nome do arquivo pdf sem a extensão
    nome_pdf = result[seq_global]['Fazenda']
    #CRIA UMA CÓPIA DESSE C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\Modelo.docx
    doc = docx.Document(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\Modelo_Uruba.docx')
    #salva esse arquivo com o nome do pdf
    doc.save(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\URUBA\DOC_Atualizados\{nome_pdf}.docx')
    #abre esse arquivo usando o 
    doc = DocxTemplate(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\URUBA\DOC_Atualizados\{nome_pdf}.docx')
    #substitui o texto "TITULO_PRINCIPAL" por "cumprimento" em todo o texto usando docxtpl
    context = { 'TITULO_PRINCIPAL' : TITULO_PRINCIPAL,
                'AREA_TOTAL' : AREA_TOTAL_FORMATADA,
                'NOME_FAZENDA1' : NOME_FAZENDA1,
                'NM_MUN' : NM_MUN,
                'COD_INCRA' : COD_INCRA,
                'ITR_NIRF' : ITR_NIRF,
                'REF_MF' : REF_MF,
                'MOD_FISCAL' : MOD_FISCAL_FORMATADO,
                'TAMANHO_PROPRIEDADE' : TAMANHO_PROPRIEDADE,
                'LOC_ACESSO' : LOC_ACESSO,
                'DESC_CLIMA' : DESC_CLIMA,
                'AREA_BENFEITORIAS' : AREA_BENFEITORIAS_FORMATADA,
                'AREA_APROVEITAVEL' : AREA_APROVEITAVEL,
                'VALOR_IMOVEL' : VALOR_IMOVEL,
                'VALOR_BENFEITORIA' : VALOR_BENFEITORIA,
                'VALOR_TERRANUA' : VALOR_TERRANUA,
                'AREA_APP' : AREA_APP,
                'AREA_RL' : AREA_RL,
                'AREA_TRIBUTAVEL' : AREA_TRIBUTAVEL,
                'DESC_MICRO' : DESC_MICRO,
                'POP' : POP,
                'AREA' : AREA,
                'FUNDACAO' : FUNDACAO,
                'ALTITUDE' : ALTITUDE,
                'DIST' : DIST,
                'ECONOMIA' : ECONOMIA,
                'PIB' : PIB,
                'RENDA' : RENDA,
                'IDH' : IDH,
                'TEMP' : TEMP,
                'VALOR_CULTIVO' : VALOR_CULTIVO}
    doc.render(context)
    #salva o arquivo
    doc.save(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\URUBA\DOC_Atualizados\{nome_pdf}.docx')
    seq_global += 1
    print(VALOR_TERRANUA)

print('Processo Finalizado')
