import docx
import locale
#quero listar todos os pdfs que existem em uma pasta
import os
from os import listdir
from os.path import isfile, join
from docxtpl import DocxTemplate

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
#seleciona a pasta C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_Desatualizados\GUAXUMA
mypath = fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_Desatualizados\GUAXUMA'

#lista os pdfs dessa pasta
list_pdf = [f for f in listdir(mypath) if isfile(join(mypath, f))]
#inputs
for pdf in list_pdf:
    print(pdf)

    TITULO_PRINCIPAL = input('Digite o nome da fazenda: ').upper()
    AREA_TOTAL = input('Digite a área total da fazenda: ')
    AREA_TOTAL_FORMATADA = locale.format_string("%.2f", float(AREA_TOTAL), grouping=True)
    NOME_FAZENDA1 = TITULO_PRINCIPAL.title()
    NM_MUN = input('Digite o nome do município: ').title()
    COD_INCRA = input('Digite o código INCRA do imóvel: ')
    ITR_NIRF = input('Digite o ITR/NIRF do imóvel: ')
    if NM_MUN == 'Campo Alegre' or NM_MUN == 'Coruripe':
        REF_MF = '30'
    elif NM_MUN == 'Junqueiro' or NM_MUN == 'São Sebastião' or NM_MUN == 'Teotônio Vilela':
        REF_MF = '35'
    MOD_FISCAL = float(AREA_TOTAL) / float(REF_MF)
    if MOD_FISCAL < 4:
        TAMANHO_PROPRIEDADE = 'Pequena Propriedade Produtiva'
    elif MOD_FISCAL >= 4 and MOD_FISCAL < 15:
        TAMANHO_PROPRIEDADE = 'Média Propriedade Produtiva'
    elif MOD_FISCAL >= 15:
        TAMANHO_PROPRIEDADE = 'Grande Propriedade Produtiva'
    
    #DESC_ACESSO
    if NM_MUN == 'Coruripe':
        LOC_ACESSO = 'Saindo de Maceió pela AL 101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Após passar pelo cruzamento com a AL 420 segue ainda pela AL 101 até chegar no Município de Coruripe, percorrendo o total de 87,3 KM.'
    elif NM_MUN == 'Campo Alegre':
        LOC_ACESSO = 'Saindo de Maceió pela AL 101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. E em seguida pela AL-220 até chegar no Município de Campo Alegre, percorrendo o total de 89,4 KM.'
    elif NM_MUN == 'Junqueiro':
        LOC_ACESSO = 'Saindo de Maceió pela AL 101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Em seguida AL-220 até chegar na BR-101 na qual permanece até o município de Junqueiro, percorrendo o total de 112 KM.'
    elif NM_MUN == 'Teotônio Vilela':
        LOC_ACESSO = 'Saindo de Maceió pela AL 101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Em seguida AL-220 até chegar na BR-101 na qual permanece até o município de Teotônio Vilela, percorrendo o total de 96,9 KM.'
    elif NM_MUN == 'São Sebastião':
        LOC_ACESSO = 'Saindo de Maceió pela AL 101 segue em direção aos municípios de Marechal Deodoro e Barra de São Miguel. Em seguida AL-220 até chegar na BR-101 na qual permanece até o município de São Sebastião, percorrendo o total de 126 KM.'

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

    AREA_BENFEITORIAS = input('Digite a área de benfeitorias: ')
    AREA_BENFEITORIAS_FORMATADA = locale.format_string("%.2f", float(AREA_BENFEITORIAS), grouping=True)

    AREA_APROVEITAVEL_fl = float(AREA_TOTAL) - float(AREA_BENFEITORIAS)
    AREA_APROVEITAVEL = locale.format_string("%.2f", AREA_APROVEITAVEL_fl, grouping=True)

    VALOR_IMOVEL_fl = float(AREA_TOTAL) * 10830
    VALOR_IMOVEL = locale.format_string("%.2f", VALOR_IMOVEL_fl, grouping=True)

    VALOR_BENFEITORIA_fl = float(AREA_BENFEITORIAS) * 10830
    VALOR_BENFEITORIA = locale.format_string("%.2f", VALOR_BENFEITORIA_fl, grouping=True)

    VALOR_TERRANUA_fl = float(VALOR_IMOVEL_fl) - float(VALOR_BENFEITORIA_fl)
    VALOR_TERRANUA = locale.format_string("%.2f", VALOR_TERRANUA_fl, grouping=True)


    #nome do arquivo pdf sem a extensão
    nome_pdf = pdf.split('.')[0]
    #CRIA UMA CÓPIA DESSE C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\Modelo.docx
    doc = docx.Document(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\Modelo.docx')
    #salva esse arquivo com o nome do pdf
    doc.save(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\output\{nome_pdf}.docx')
    #abre esse arquivo usando o 
    doc = DocxTemplate(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\output\{nome_pdf}.docx')
    #substitui o texto "TITULO_PRINCIPAL" por "cumprimento" em todo o texto usando docxtpl
    context = { 'TITULO_PRINCIPAL' : TITULO_PRINCIPAL,
                'AREA_TOTAL' : AREA_TOTAL_FORMATADA,
                'NOME_FAZENDA1' : NOME_FAZENDA1,
                'NM_MUN' : NM_MUN,
                'COD_INCRA' : COD_INCRA,
                'ITR_NIRF' : ITR_NIRF,
                'REF_MF' : REF_MF,
                'MOD_FISCAL' : MOD_FISCAL,
                'TAMANHO_PROPRIEDADE' : TAMANHO_PROPRIEDADE,
                'LOC_ACESSO' : LOC_ACESSO,
                'DESC_CLIMA' : DESC_CLIMA,
                'AREA_BENFEITORIAS' : AREA_BENFEITORIAS_FORMATADA,
                'AREA_APROVEITAVEL' : AREA_APROVEITAVEL,
                'VALOR_IMOVEL' : VALOR_IMOVEL,
                'VALOR_BENFEITORIA' : VALOR_BENFEITORIA,
                'VALOR_TERRANUA' : VALOR_TERRANUA}
    doc.render(context)
    #salva o arquivo
    doc.save(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\output\{nome_pdf}.docx')




