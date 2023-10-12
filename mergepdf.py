import fitz 
from os import listdir
from os.path import isfile, join

# Função para unir dois PDFs
def merge_pdfs(pdf1_path, pdf2_path, output_path):
    pdf_document1 = fitz.open(pdf1_path)  # Abre o primeiro PDF
    pdf_document2 = fitz.open(pdf2_path)  # Abre o segundo PDF

    pdf_document1.insert_pdf(pdf_document2)  # Insere o segundo PDF no primeiro

    pdf_document1.save(output_path)  # Salva o PDF resultante
    pdf_document1.close()

# Caminhos para os PDFs de entrada e saída
mypath = fr'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_ATUALIZADOS_EXTRA'
pdf2_path = r"C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\AVALURB\Avalurb_GUAXUMA.pdf"
pasta_pdf2 = r"C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_ATUALIZADOS_2_EXTRA"

list_pdf = [f for f in listdir(mypath) if isfile(join(mypath, f))]
for pdf in list_pdf:
    pdf1_path = join(mypath, pdf)
    output_path = join(pasta_pdf2, pdf)
    merge_pdfs(pdf1_path, pdf2_path, output_path)
    print(pdf, 'unido com sucesso!')