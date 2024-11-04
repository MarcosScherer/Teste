import os
import pandas as pd

class ExcelToTxtConverter:
    def __init__(self, input_folder, output_folder, colunas_monetarias=None, delimiter=','):
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.delimiter = delimiter
        self.colunas_monetarias = colunas_monetarias or []

        # Processa todos os arquivos .xlsx na pasta de entrada
        self.process_all_files()

    def process_all_files(self):
        # Verifica se a pasta de saída existe; se não, cria
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        # Itera sobre todos os arquivos .xlsx na pasta de entrada
        for file_name in os.listdir(self.input_folder):
            if file_name.endswith('.xlsx'):
                input_path = os.path.join(self.input_folder, file_name)
                output_path = os.path.join(self.output_folder, file_name.replace('.xlsx', '.txt'))
                
                # Processa o arquivo individualmente
                self.df = self.read_excel(input_path)
                if self.df is not None:
                    if self.colunas_monetarias:
                        self.limpar_colunas_monetarias(self.colunas_monetarias)
                    self.force_integer_conversion()  # Nova chamada para conversão de inteiros
                    self.convert_to_txt(output_path)
    
    def read_excel(self, input_path):
        try:
            df = pd.read_excel(input_path)
            print(f"Excel file '{input_path}' successfully read.")
            return df
        except Exception as e:
            print(f"An error occurred while reading the Excel file '{input_path}': {e}")
            return None


    def limpar_colunas_monetarias(self, colunas):
        for coluna in colunas:
            if coluna in self.df.columns:
                self.df[coluna] = self.df[coluna].apply(self.ajustar_milhar_decimal)

    def ajustar_milhar_decimal(self, valor):
        # Se o valor for NaN, retorna uma string vazia imediatamente
        if pd.isna(valor):
            return ''

        # Converte o valor para string se ainda não for
        valor = str(valor)

        # Se há uma vírgula como separador decimal, remove pontos e substitui vírgula por ponto
        if ',' in valor and valor.rfind(',') > valor.rfind('.'):
            valor = valor.replace('.', '')  # Remove pontos como separadores de milhar
            valor = valor.replace(',', '.')  # Substitui vírgula por ponto decimal
        elif '.' in valor and valor.rfind('.') > valor.rfind(','):
            valor = valor.replace(',', '')  # Remove vírgulas como separadores de milhar

        # Retorna o valor ajustado
        return valor

    def force_integer_conversion(self):
        """Converte colunas numéricas que contêm apenas inteiros para o tipo Int64."""
        for coluna in self.df.columns:
            if pd.api.types.is_numeric_dtype(self.df[coluna]):
                # Verifica se todos os valores não nulos na coluna são inteiros
                if all(self.df[coluna].dropna().apply(lambda x: x.is_integer() if isinstance(x, float) else True)):
                    self.df[coluna] = self.df[coluna].astype('Int64')  # Converte para tipo Int64

    def convert_to_txt(self, output_path):
        if self.df is not None:
            try:
                self.df['Valor Premio Liquido'].fillna('', inplace=True)
                self.df['Valor Premio Total'].fillna('', inplace=True)
                self.df.to_csv(output_path, sep=self.delimiter, index=False, header=True, na_rep='')
                print(f"File successfully converted and saved to '{output_path}'")
            except Exception as e:
                print(f"An error occurred while saving the .txt file '{output_path}': {e}")

# Exemplo de uso
converter = ExcelToTxtConverter(
    input_folder=r'C:\Users\Thomas\Downloads\Base_v2\Emissao\Dados_Emissao\Arrumar antigo\Novo',
    output_folder=r'C:\Users\Thomas\Downloads\Base_v2\Emissao\Dados_Emissao\Arrumar antigo\txtfiles',
    colunas_monetarias=['Valor Premio Liquido', 'Valor Premio Total'],
    delimiter=';'
)
