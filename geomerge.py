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
mypath = r'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_Atualizados_2'
pasta_pdf2 = r"C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_Atualizados_3"
pasta_geo = r"C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\GEO"

list_pdf = [f for f in listdir(mypath) if isfile(join(mypath, f))]
list_pdf_geo = [f for f in listdir(pasta_geo) if isfile(join(pasta_geo, f))]
for pdf in list_pdf:
    pdf_name = pdf
    if pdf_name in list_pdf_geo:
        pdf2_path = join(pasta_geo, pdf)
        pdf1_path = join(mypath, pdf)
        output_path = join(pasta_pdf2, pdf)
        merge_pdfs(pdf1_path, pdf2_path, output_path)

    else:
        print(pdf, 'não encontrado!')
        #faz uma copia do pdf e coloca na pasta 3
        pdf1_path = join(mypath, pdf)
        output_path = join(pasta_pdf2, pdf)
        save_pdf = fitz.open(pdf1_path)
        save_pdf.save(output_path)
        save_pdf.close()

        print(pdf, 'copiado com sucesso!')
    

