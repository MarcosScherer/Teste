import pandas as pd

class ExcelToTxtConverter:
    def __init__(self, input_path, output_path, colunas_monetarias=None, delimiter=','):

        self.input_path = input_path
        self.output_path = output_path
        self.delimiter = delimiter
        self.df = None

        # Execute the steps
        self.read_excel()
        if colunas_monetarias:
            self.limpar_colunas_monetarias(colunas_monetarias)
        
        
        self.convert_to_txt()

    def read_excel(self):
        try:
            self.df = pd.read_excel(self.input_path,index_col=None)
            self.df = self.df.loc[:, ~self.df.columns.str.contains('^Unnamed')]
            print("Excel file successfully read.")
        except Exception as e:
            print(f"An error occurred while reading the Excel file: {e}")

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


    def convert_to_txt(self):
        if self.df is not None:
            try:
                # Converte colunas float sem decimais para inteiros
                for coluna in self.df.select_dtypes(include=['float']):
                    if self.df[coluna].dropna().apply(float.is_integer).all():
                        self.df[coluna] = self.df[coluna].astype('Int64')  # Usa o tipo 'Int64' que permite nulos

                self.df.to_csv(self.output_path, sep=self.delimiter, index=False, header=True)
                print(f"File successfully converted and saved to {self.output_path}")
            except Exception as e:
                print(f"An error occurred while saving the .txt file: {e}")
        else:
            print("DataFrame is empty. Make sure to read the Excel file first.")

# Usage example (all steps are executed upon instantiation)

