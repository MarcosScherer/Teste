import pandas as pd
import numpy as np
from unidecode import unidecode
import os

class StanEmissions:

    def __init__(self, path_file, path_save=None) -> None:
        self.df = pd.read_excel(path_file) 
        self.path_save = path_save

        #Call function
        self.rename_headers()
        self.separate_cols()
        self.from_to_ramo_seguro()
        self.sep_dates()
        self.frota()
        self.new_order()
        self.transformar_data(['Inicio Vigencia', 'Fim Vigencia', 'Data de Emissao']) ####teste
        self.remover_simbolos_moeda(['Valor Premio Liquido','Valor Premio Total'])
        self.save_xlsx()
    
    def rename_headers(self):
        """Substitute all special characters in the headers of the DataFrame."""
        self.df = self.df.rename(columns={'Nº Con trato': 'N Contrato',
                                'Valor Prêmio Líquido (*)':'Valor Prêmio Líquido',
                                'Hist. da Apólice':'Hist da Apólice'
                    })

        self.df.columns = [unidecode(col).replace('ç', 'c') for col in self.df.columns]


    def separate_cols(self) -> None:
        sep_col_header = {
            'Companhia': ['Cod Companhia', 'Companhia'],
            'Assessoria': ['CNPJ Assessoria', 'Assessoria'],
            'Sucursal': ['Cod Sucursal', 'Sucursal'],
            'Inspetor de producao': ['CPF Inspetor de Producao', 'Inspetor de Producao'], 
            'Corretor': ['CPD Corretor', 'Corretor']
        }

        for key, value in sep_col_header.items():
            if key == 'Corretor':  # The 'Corretores' can have a '-' in the name, so we use a slice system.
                self.df[value[0]] = self.df['Corretor'].str[:6]
                self.df[key] = self.df['Corretor'].str[9:]
                continue

            if key in self.df.columns:
                self.df[[value[0], key]] = self.df[key].str.split(' - ', expand=True)


    def from_to_ramo_seguro(self):
        """Create a column from cod number to text code"""
        self.df['Ramo Seguro'] = np.nan

        dict_ramos  = {
                       919: 'Condominio',
                       926: 'Empresarial',
                       917: 'Equipamento',
                       600: 'Equipamento Agricola',
                       927: 'Residencial Sob Medida',
                       990: 'Auto'
                       }

        for key, value in dict_ramos.items():
            self.df.loc[self.df['Ramo'] == key, 'Ramo Seguro'] = value

    def sep_dates(self):
        """Sep the column dates and corrects the form"""
        col_dates = ['Inicio Vigencia', 'Fim Vigencia']
        #Separete cols
        self.df['Vigencia'] = self.df['Vigencia'].str.replace('De: ', '')
        self.df[col_dates] = self.df['Vigencia'].str.split(' a ', expand=True)
        self.df.drop(columns=['Vigencia'], inplace=True)
        #Substitute 
        for col in ['Inicio Vigencia', 'Fim Vigencia', 'Data de Emissao']:
            self.df[col] = self.df[col].str.replace('/', '-')

    def frota(self):
        """Count the itens in the fleet"""
        self.df['Frota Itens'] = np.nan
        self.df.loc[self.df['Item'] == 0, 'Ramo Seguro'] = 'Frota'                                        #Troca o Ramo do seguro para FROTA onde o item é =       
        frota_list = self.df[self.df['Item'] == 0]['Segurado'].to_list()                                  #Cria uma lista com o nome de segurados de Frota
        unique_list_frota = list(set(frota_list)) 

        for i in unique_list_frota:                                                                            # i é o nome de cada segurado que possui frota
            n_itens = self.df[self.df['Segurado']==i].shape[0]                                                 #Pega o numero de itens medindo o shape do dataframe filtrado para o segurado especifico
            self.df.loc[(self.df['Segurado'] == i) & (self.df['Item'] == 0), 'Frota Itens'] = n_itens          #Imputa o numero de itens onde o item é zero é o nome do segurado é igual a i
            list_frota_del = self.df[(self.df['Segurado'] == i) & ~(self.df['Item'] == 0)].index.to_list()     #Cria uma lista de indice onde o item é diferente de 0 e o segurado igual a i (apagar as linha da frota)
            self.df = self.df.drop(list_frota_del)

    def new_order(self):
        """Change the order of the columns"""
        new_order = [
            'Cia', 'Ramo', 'CPD Corretor', 'Corretor','Apolice', 'Item', 'Endosso', 'Segurado', 'N Contrato',
        'Data de Emissao','Inicio Vigencia', 'Fim Vigencia', 'Prestacao', 'Valor Premio Liquido',
        'Valor Premio Total', 'Hist da Apolice','Cod Sucursal', 'Sucursal', 'CPF Inspetor de Producao','Inspetor de producao', 'CNPJ Assessoria','Assessoria','Cod Companhia','Companhia' ,
        'Ramo Seguro','Frota Itens'
                    ]

        self.df = self.df[new_order]

    
    def transformar_data(self, colunas):
        for coluna in colunas:
            # Converte cada coluna de texto para o formato de data no pandas
            self.df[coluna] = pd.to_datetime(self.df[coluna], format='%d-%m-%Y')
            
            # Converte para o formato 'YYYY-MM-DD' (aaaa-mm-dd)
            self.df[coluna] = self.df[coluna].dt.strftime('%Y-%m-%d')
    
    def remover_simbolos_moeda(self, colunas):
        for coluna in colunas:
            # Remove os símbolos 'R$' e '$' de cada coluna e converte para numérico
            self.df[coluna] = self.df[coluna].replace({'R\$': '', '\$': ''}, regex=True)
            
            # Converte a coluna para numérico (float), caso seja necessário para análise de dados
            #self.df[coluna] = pd.to_numeric(self.df[coluna], errors='coerce')


    def save_xlsx(self):
        inspetor = self.df['Inspetor de producao'][0].replace(' ','_')
        file = f'{unidecode(inspetor)}.xlsx'
        self.df.to_excel(os.path.join(self.path_save, file))

