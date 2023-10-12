import os
import comtypes.client

def convert_docx_to_pdf(docx_file, pdf_file):
    # Inicie uma instância do Word
    word = comtypes.client.CreateObject("Word.Application")
    doc = word.Documents.Open(docx_file)
    
    # Salve o documento como PDF
    doc.SaveAs(pdf_file, FileFormat=17)  # 17 é o código para PDF
    doc.Close()
    
    # Feche o Wordx
    word.Quit()

def batch_convert_docx_to_pdf(input_folder, output_folder):
    # Verifique se os diretórios de entrada e saída existem
    if not os.path.exists(input_folder):
        print(f'Diretório {input_folder} não encontrado.')
        return

    if not os.path.exists(output_folder):
        print(f'Diretório {output_folder} não encontrado.')
        return

    # Liste todos os arquivos DOCX no diretório de entrada
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):  
            input_docx = os.path.join(input_folder, filename)
            output_pdf = os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf")
            
            # Converta o arquivo DOCX para PDF
            convert_docx_to_pdf(input_docx, output_pdf)
            print(f'{input_docx} convertido para {output_pdf}')

if __name__ == "__main__":
    caminho_docx = r'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\DOC_ATUALIZADOS_EXTRA'
    caminho_pdf = r'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\PDF_ATUALIZADOS_EXTRA'
    
    batch_convert_docx_to_pdf(caminho_docx, caminho_pdf)
