import pandas as pd
import math
import numpy as np
import os
import dataframe_image as di

from unidecode import unidecode
import re


def list_files_in_folder(folder_path):
    files = []
    for file in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, file)) and not file.startswith('~'):
            files.append(os.path.join(folder_path, file))
    return files


#########################################################################################

def padronizaEmissao(caminho_original, caminho_editado, ano_mes):

    df = pd.read_excel(caminho_original, header=None, engine='openpyxl')

    #########################################################################################

    header = ['Cia', 'Ramo', 'Apolice','Item',
            'Endosso','Vigencia', 'Segurado',
            'N Contrato', 'Data de Emissao', 'Prestacao',
            'Valor Premio Liquido', 'Valor Premio Total','Hist da Apolice']# Replace with your desired header names
    df.columns = header
    # Define your desired header
    header2 = ['Companhia' , 'Assessoria', 'Sucursal', 'Inspetor de produção','Corretor', 'Ramo Seguro','Frota itens']  # Replace with your desired header names

    # Add new columns filled with NaN values
    for col_name in header2:
        df[col_name] = np.nan

    #########################################################################################

    #Version 2
    def r_to_dict(df, r_index):
        row_list = df.iloc[r_index].tolist()
        row_list_cleaned_nan = [x for x in row_list if isinstance(x, str) or not math.isnan(x)]
        # Create a dictionary using zip()
        return dict(zip(row_list_cleaned_nan[::2], row_list_cleaned_nan[1::2]))

    #########################################################################################

    list_index = df[df['Cia'] == 'Companhia'].index.to_list()

    #########################################################################################

    def update_dict(dict_col_value,row):
        dict_novo = {}
        for key, value in dict_col_value.items():
            dict_novo[key] = {row:value}
        return dict_novo

    #########################################################################################

    for i in list_index:
        d1 = r_to_dict(df,i)
        d2 = r_to_dict(df,i+1)
        d1.update(d2)
        update_df = update_dict(d1,i+3)
        df.update(pd.DataFrame(update_df))

    #########################################################################################

    i_del_list = [x for val in list_index for x in (val, val + 1, val + 2)]

    #########################################################################################

    df = df.drop(i_del_list, axis=0)

    #########################################################################################

    df.dropna(how='all', inplace=True)

    #########################################################################################

    for col in header2:
        df[col].fillna(method='ffill', inplace=True)

    #########################################################################################

    sep_col_header = {'Companhia':['Cod Companhia', 'Companhia'],
                    'Assessoria':['CNPJ Assessoria', 'Assessoria'],
                    'Sucursal':['Cod Sucursal', 'Sucursal'],
                    'Inspetor de produção':['CPF Inspetor de produção', 'Inspetor de produção'],
                    'Corretor':['CPD Corretor','Corretor']
                    }

    #########################################################################################

    for key, value in sep_col_header.items():

        if key == 'Corretor':
            df[value[0]] = df['Corretor'].str[:6]
            df[key] = df['Corretor'].str[9:]
            continue

        df[[value[0], key]] = df[key].str.split(' - ', expand=True)

    #########################################################################################

    dict_ramos  = {919: 'Condominio',
                926: 'Empresarial',
                917: 'Equipamento',
                600: 'Equipamento agrícola',
                927: 'Residencial sob medida',
                990: 'Auto'}

    for key, value in dict_ramos.items():
        df.loc[df['Ramo'] == key, 'Ramo Seguro'] = value

    #########################################################################################

    df['Vigencia'] = df['Vigencia'].str.replace('De: ', '')
    df[['Inicio Vigencia', 'Fim Vigencia']] = df['Vigencia'].str.split(' a ', expand=True)

    #########################################################################################

    #df['Cia'] = df['Cia'].str.replace('Não existem seguros emitidos.', '')

    #########################################################################################

    df.drop(columns=['Vigencia'], inplace=True)

    #########################################################################################

    df.loc[df['Cia'] == 'Não existem seguros emitidos.', 'Cia'] = pd.NA

    #########################################################################################

    df = df.rename(columns={'CPF Inspetor de produção': 'CPF Inspetor de producao',
                    'Inspetor de produção':'Inspetor de producao'
                    })

    #########################################################################################

    df.reset_index(drop=True, inplace=True)

    #########################################################################################

    # Assuming your DataFrame is named df and the columns containing "//" are 'column1', 'column2', 'column3'
    columns_to_replace = ['Inicio Vigencia', 'Fim Vigencia']

    for col in columns_to_replace:
        df[col] = df[col].str.replace('/', '-')

    #########################################################################################

    df['Inicio Vigencia'].fillna('01-01-1999', inplace=True)
    df['Fim Vigencia'].fillna('01-01-1999', inplace=True)

    #########################################################################################

    df['Inicio Vigencia'] = pd.to_datetime(df['Inicio Vigencia'], format='%d-%m-%Y')
    df['Fim Vigencia'] = pd.to_datetime(df['Fim Vigencia'], format='%d-%m-%Y')

    #########################################################################################

    df['Inicio Vigencia'] = df['Inicio Vigencia'].dt.strftime('%d-%m-%Y')
    df['Fim Vigencia'] = df['Fim Vigencia'].dt.strftime('%d-%m-%Y')

    #########################################################################################

    df.loc[df['Item'] == 0, 'Ramo Seguro'] = 'Frota'                                        #Troca o Ramo do seguro para FROTA onde o item é = 0

    frota_list = df[df['Item'] == 0]['Segurado'].to_list()                                  #Cria uma lista com o nome de segurados de Frota
    unique_list_frota = list(set(frota_list))                                               #Apaga os nomes repetidos da lista

    for i in unique_list_frota:                                                             # i é o nome de cada segurado que possui frota
        n_itens = df[df['Segurado']==i].shape[0]                                            #Pega o numero de itens medindo o shape do dataframe filtrado para o segurado especifico
        df.loc[(df['Segurado'] == i) & (df['Item'] == 0), 'Frota itens'] = n_itens          #Imputa o numero de itens onde o item é zero é o nome do segurado é igual a i
        list_frota_del = df[(df['Segurado'] == i) & ~(df['Item'] == 0)].index.to_list()     #Cria uma lista de indice onde o item é diferente de 0 e o segurado igual a i (apagar as linha da frota)
        df = df.drop(list_frota_del)   

    #########################################################################################

    new_order = ['Cia', 'Ramo', 'CPD Corretor', 'Corretor','Apolice', 'Item', 'Endosso', 'Segurado', 'N Contrato',
        'Data de Emissao','Inicio Vigencia', 'Fim Vigencia', 'Prestacao', 'Valor Premio Liquido',
        'Valor Premio Total', 'Hist da Apolice','Cod Sucursal', 'Sucursal', 'CPF Inspetor de producao','Inspetor de producao', 'CNPJ Assessoria','Assessoria','Cod Companhia','Companhia' ,
        'Ramo Seguro','Frota itens'

                    ]

    df = df[new_order]

    #########################################################################################
    #try:
    #    df['Valor Premio Liquido'] = df['Valor Premio Liquido'].str.replace('?', '')
    
    #df['Valor Premio Total'] = df['Valor Premio Total'].str.replace('?', '')


    #########################################################################################

    inspetor = df['Inspetor de producao'][0].replace(' ','_')
    file = f'{ano_mes}_{unidecode(inspetor)}_Editado.xlsx'
    df.to_excel(os.path.join(caminho_editado, file))





