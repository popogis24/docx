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
    #AGORA QUERO PEGAR O NOME DO ARQUIVO PDF E COLOCAR NO NOME DO ARQUIVO DOCX
    #nome do arquivo pdf sem a extensão
    nome_pdf = pdf.split('.')[0]
    #CRIA UMA CÓPIA DESSE C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\Modelo.docx
    doc = docx.Document(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\Modelo.docx')
    #salva esse arquivo com o nome do pdf
    doc.save(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\output\{nome_pdf}.docx')
    #abre esse arquivo usando o 
    doc = DocxTemplate(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\output\{nome_pdf}.docx')
    #substitui o texto "TITULO_PRINCIPAL" por "cumprimento" em todo o texto usando docxtpl
    context = { 'TITULO_PRINCIPAL' : 'aaaaaa' }
    doc.render(context)
    #salva o arquivo
    doc.save(fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_Atualizados\output\{nome_pdf}.docx')


