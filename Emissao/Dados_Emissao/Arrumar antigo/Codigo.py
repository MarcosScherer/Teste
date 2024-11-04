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
        #self.renomear_colunas()
        self.normalize_contents(['Hist da Apolice'])

        self.transformar_data(['Inicio Vigencia', 'Fim Vigencia', 'Data de Emissao']) ####teste
        #self.limpar_colunas_monetarias(['Valor Premio Liquido','Valor Premio Total'])
        self.remover_valores_especificos(['Inicio Vigencia', 'Fim Vigencia'])
        self.criar_coluna_zerado_mes()
        self.save_xlsx()
    
    def rename_headers(self):
        """Substitute all special characters in the headers of the DataFrame."""
        self.df.columns = [unidecode(col).replace('ç', 'c') for col in self.df.columns]
    
    def normalize_contents(self, colunas):
        """Normalize all special characters in the specified columns of the DataFrame."""
        for coluna in colunas:
            if coluna in self.df.columns:  # Verifica se a coluna existe no DataFrame
                self.df[coluna] = self.df[coluna].apply(lambda x: unidecode(str(x)).replace('ç', 'c') if isinstance(x, str) else x)
            else:
                print(f"Coluna '{coluna}' não encontrada no DataFrame.")
        return self.df
    
    def renomear_colunas(self):
        # Lista de novos nomes para as colunas
        novos_nomes = [
            "Cia", "Ramo", "CPD Corretor", "Corretor", "Apolice", "Item", "Endosso",
            "Segurado", "N Contrato", "Data de Emissao", "Inicio Vigencia", 
            "Fim Vigencia", "Prestacao", "Valor Premio Liquido", "Valor Premio Total", 
            "Hist da Apolice", "Cod Sucursal", "Sucursal", "CPF Inspetor de Producao", 
            "Inspetor de producao", "CNPJ Assessoria", "Assessoria", "Cod Companhia", 
            "Companhia", "Ramo Seguro", "Frota Itens"
        ]
        
        # Verifica se o número de colunas no DataFrame corresponde ao número de novos nomes
        if len(novos_nomes) != len(self.df.columns):
            raise ValueError("O número de novos nomes de colunas deve corresponder ao número de colunas do DataFrame.")
        
        # Renomeia as colunas
        self.df.columns = novos_nomes
        
        return self.df
    
    
    def transformar_data(self, colunas):
        for coluna in colunas:
            # Converte cada coluna de texto para o formato de data no pandas
            self.df[coluna] = pd.to_datetime(self.df[coluna], format='%d-%m-%Y')
            
            # Converte para o formato 'YYYY-MM-DD' (aaaa-mm-dd)
            self.df[coluna] = self.df[coluna].dt.strftime('%Y-%m-%d')

    
    def limpar_colunas_monetarias(self, colunas):
        for coluna in colunas:
            # Remove os símbolos de moeda 'R$' e '$'
            self.df[coluna] = self.df[coluna].replace({'R\$': '', '\$': ''}, regex=True)
            
            # Remove pontos e substitui vírgulas por pontos
            self.df[coluna] = self.df[coluna].str.replace('.', '', regex=False)  # Remove os pontos
            self.df[coluna] = self.df[coluna].str.replace(',', '.', regex=False) # Substitui a vírgula por ponto
            
            # Converte a coluna para numérico (float)
            self.df[coluna] = pd.to_numeric(self.df[coluna], errors='coerce')
            

    
    def remover_valores_especificos(self, colunas, valor_remover="1999-01-01"):
        for coluna in colunas:
            # Substitui o valor específico por NaN
            self.df[coluna] = self.df[coluna].replace(valor_remover, np.nan)

    
    def criar_coluna_zerado_mes(self):
        # Cria a coluna 'ZeradoMes' e preenche com "2024-01-01" se 'ramo' estiver vazio
        self.df['ZeradoMes'] = self.df['Ramo'].apply(lambda x: "2024-07-01" if pd.isna(x) or x == "" else np.nan)
        return self.df
        

    def save_xlsx(self):
        inspetor = '24_07_Agrupado_Novo'
        file = f'{unidecode(inspetor)}.xlsx'
        self.df.to_excel(os.path.join(self.path_save, file))


StanEmissions(r'C:\Users\Thomas\Downloads\Base_v2\Emissao\Dados_Emissao\Arrumar antigo\Original\24_07_Emissoes_agrupado.xlsx', r'C:\Users\Thomas\Downloads\Base_v2\Emissao\Dados_Emissao\Arrumar antigo\Novo')