import os

# Defina o diretório raiz onde as subpastas estão localizadas
diretorio_raiz = r'C:\Users\Ander\OneDrive\Documentos\JORGE_DOCUMENTOS\GUAXUMA\KMZpy'

# Percorra todas as subpastas e seus arquivos
for pasta_raiz, subpastas, arquivos in os.walk(diretorio_raiz):
    for arquivo in arquivos:
        # Crie o novo nome do arquivo com base no nome da subpasta
        novo_nome = os.path.join(pasta_raiz, arquivo)
        novo_nome = os.path.join(pasta_raiz, os.path.basename(pasta_raiz) + os.path.splitext(arquivo)[1])
        
        # Renomeie o arquivo
        os.rename(os.path.join(pasta_raiz, arquivo), novo_nome)
        print(f"Renomeado: {os.path.join(pasta_raiz, arquivo)} -> {novo_nome}")
