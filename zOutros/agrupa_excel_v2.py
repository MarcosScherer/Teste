import pandas as pd
import os
import glob
import shutil


def agrupaEx(root_path, final_path):
    #Roth_path é uma pasta contendo os arquivos em excel
    #final_path é o caminho para salva o arquivo excel inclusive com seu nome

    dfs = []
    # Loop through each file in the directory
    for filename in os.listdir(root_path):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            # Construct full file path
            file_path = os.path.join(root_path, filename)
            
            # Read the Excel file into a DataFrame
            df = pd.read_excel(file_path, index_col=0)
            
            # Append the DataFrame to the list
            dfs.append(df)

    # Concatenate all DataFrames in the list into a single DataFrame
    combined_df = pd.concat(dfs, ignore_index=True)

    # Display the combined DataF
    combined_df.to_excel(final_path)

def delete_all_contents(folder_path):
    # Loop through the contents of the directory
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)  # Remove the file or link
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Remove the directory and its contents
        except Exception as e:
            pass  # Handle errors silently

def plan_grades(path_origin,path_destiny):
    
    # Selecione apenas as duas colunas desejadas
    df = pd.read_excel(path_origin)
    df_selected = df[['Corretor', 'Inspetor de producao']]

    # Remova as linhas duplicadas
    df_selected = df_selected.drop_duplicates()
    df_selected = df_selected.reset_index(drop=True)
    df_selected.to_excel(path_destiny)

def ajeita_path(path1,path2):
    return os.path.join(path1, path2)

def copiar_tudo(pasta_origem, pasta_destino):
    # Verifica se a pasta de destino existe, caso contrário, cria a pasta
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    # Percorre todos os arquivos e diretórios na pasta de origem
    for item in os.listdir(pasta_origem):
        caminho_item_origem = os.path.join(pasta_origem, item)
        caminho_item_destino = os.path.join(pasta_destino, item)
        
        # Verifica se é um diretório
        if os.path.isdir(caminho_item_origem):
            # Copia o diretório recursivamente para a pasta de destino
            shutil.copytree(caminho_item_origem, caminho_item_destino)
        else:
            # Copia o arquivo para a pasta de destino
            shutil.copy2(caminho_item_origem, caminho_item_destino)

def compactar_pasta(pasta_origem, pasta_destino):
    # Verifica se a pasta de origem existe
    if not os.path.exists(pasta_origem):
        print(f"A pasta {pasta_origem} não existe.")
        return
    
    # Cria o caminho do arquivo compactado
    arquivo_zip = shutil.make_archive(pasta_destino, 'zip', pasta_origem)
    
    print(f"Pasta {pasta_origem} compactada com sucesso em {arquivo_zip}")





