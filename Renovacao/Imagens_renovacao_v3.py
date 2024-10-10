import pandas as pd
import math
import numpy as np
import os
import dataframe_image as di
from unidecode import unidecode
import re

def create_folder(path, folder_name):
    folder_path = os.path.join(path, folder_name)
    os.makedirs(folder_path)
    return folder_path


def create_img_individual(path_file, img_path, data, dt_inicio, dt_fim):
     
    df = pd.read_excel(path_file, index_col=0, engine='openpyxl')

    if data == True:
        df = df[(df['Fim Vigencia'] >= dt_inicio) & (df['Fim Vigencia'] <= dt_fim)]
        
    
    filtered_df = df[df['Segurado'].notna()]

    rel_corretor = ['Fim Vigencia','Ramo Seguro','Frota Itens', 'Segurado', 'Corretor'] #Precisa colocar ramo
    
    filtered_df = filtered_df[rel_corretor]

    filtered_df['Frota itens'] = pd.to_numeric(df['Frota Itens'], errors='coerce').astype('Int64') #Tranforma a cluna com nan value para int (para nao aparecer nan na img e 25.0000)

    unique_names = filtered_df['Corretor'].unique().tolist()

    for i in unique_names:
        a = filtered_df[filtered_df['Corretor'] == i].copy()

        a.sort_values(by='Fim Vigencia', ascending=True, inplace=True)
        # Reset index and drop the old index
        a.reset_index(drop=True, inplace=True)

        # Add 1 to all index values
        a.index += 1

        a.sort_values(by='Fim Vigencia', ascending=True, inplace=True)



        # Set the background color of the header and caption
        styled_df = a.style.set_caption('Base Renova Fácil').set_table_styles([
            {'selector': 'caption', 'props': [('background-color', '#ff3300'), ('color', 'white'), ('font-size', '23px')]},  # Caption background color
            {'selector': 'thead th', 'props': [('background-color', '#ff3300'), ('color', 'white'), ('text-align', 'left')]},  # Header background color
            {'selector': 'tbody tr:nth-child(even)', 'props': [('background-color', '#F0F0F0')]},  # Even row background color (light gray)
            {'selector': 'tbody tr:nth-child(odd)', 'props': [('background-color', 'white')]},     # Odd row background color (white)
            {'selector': 'tbody td', 'props': [('text-align', 'left')]}
        ])
        nome = unidecode(i)
        nome = re.sub(r'[\\/]+', '', nome)
        caminho_final_img = os.path.join(img_path, nome+'.png')
        di.export(styled_df, caminho_final_img)

    


def create_img_zerados(path_file, img_path):                                    #Imagens corretores sem renovação

    df = pd.read_excel(path_file, index_col=0, engine='openpyxl')
    ##df['Frota itens'] = pd.to_numeric(df['Frota itens'], errors='coerce').astype('Int64') #Tranforma a cluna com nan value para int (para nao aparecer nan na img e 25.0000)

    filtered_df = df[df['Segurado'].isnull()]                                       #Filtra todos os corretores que não tem renovação   
    rel_corretor = ['Corretor','Inspetor de producao']                         #Seleciona apenas essas duas colunas no data Frame

    nome_comercial = df['Inspetor de producao'][0]

    filtered_df = filtered_df[['Corretor']]

    
    filtered_df.reset_index(drop=True, inplace=True)                             #Reset index and drop the old index
    filtered_df.index += 1                                                   #Add 1 to all index values
    
    styled_df = filtered_df.style.set_caption('Corretores Sem Renovação \n' + " " + nome_comercial).set_table_styles([
    {'selector': 'caption', 'props': [('background-color', '#ff3300'), ('color', 'white'), ('font-size', '22px')]},  # Caption background color
    {'selector': 'thead th', 'props': [('background-color', '#ff3300'), ('color', 'white'), ('text-align', 'left')]},  # Header background color and left align text
    {'selector': 'tbody tr:nth-child(even)', 'props': [('background-color', '#F0F0F0')]},  # Even row background color (light gray)
    {'selector': 'tbody tr:nth-child(odd)', 'props': [('background-color', 'white')]},  # Odd row background color (white)
    {'selector': 'tbody td', 'props': [('text-align', 'left')]}  # Left align text in body cells
    ])
    
    caminho_final_img = os.path.join(img_path, 'sem_renovacao'+'.png')
    di.export(styled_df, caminho_final_img)



    #Tabela de PIVO
    pivot_df = df.pivot_table(index='Corretor', columns='Ramo Seguro', aggfunc='size', fill_value=0)

    pivot_df.loc['Total'] = pivot_df.sum()

    colunas_piv = ['Auto', 'Residencial sob medida', 'Equipamento', 'Equipamento agrícola', 'Empresarial', 'Frota']

    colunas_piv_existe = pivot_df.columns.tolist()

    sequencia_c = [elemento for elemento in colunas_piv if elemento in colunas_piv_existe] #deixa a lista na sequencia certa

    pivot_df = pivot_df[sequencia_c]




    styled_df = pivot_df.style.set_caption('Mapa Renovações '+nome_comercial).set_table_styles([
            {'selector': 'caption', 'props': [('background-color', '#ff3300'), ('color', 'white'), ('font-size', '22px')]},  # Caption background color
            {'selector': 'thead th', 'props': [('background-color', '#ff3300'), ('color', 'white')]},                        # Header background color
            {'selector': 'tbody tr:nth-child(even)', 'props': [('background-color', '#F0F0F0')]},                            # Even row background color (light gray)
            {'selector': 'tbody tr:nth-child(odd)', 'props': [('background-color', 'white')]},
                                                                          # Odd row background color (white)
        ])
    
    caminho_final_img = os.path.join(img_path, 'TABELA'+'.png')
    di.export(styled_df, caminho_final_img)

def create_img_gestao(path_file, path_result):                                    #Imagens corretores sem renovação

    df = pd.read_excel(path_file, index_col=0, engine='openpyxl')

    a = df[df['Ramo Seguro'] == 'Frota'].copy()
    a = a[['Frota itens', 'Segurado', 'Fim Vigencia', 'Inspetor de producao', ]]

    styled_df = a.style.set_caption('Mapa Renovações Frotas').set_table_styles([
            {'selector': 'caption', 'props': [('background-color', '#ff3300'), ('color', 'white'), ('font-size', '22px')]},  # Caption background color
            {'selector': 'thead th', 'props': [('background-color', '#ff3300'), ('color', 'white')]},                        # Header background color
            {'selector': 'tbody tr:nth-child(even)', 'props': [('background-color', '#F0F0F0')]},                            # Even row background color (light gray)
            {'selector': 'tbody tr:nth-child(odd)', 'props': [('background-color', 'white')]},
                                                                          # Odd row background color (white)
        ])
    #arrumar depois img_path
    di.export(styled_df, path_result)
    #r'C:\Users\Thomas\Downloads\Base_v2\Renovacao\Dados_Renovacao\Imagens\Gestao\FrotaEquipe.png'