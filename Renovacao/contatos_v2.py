import pandas as pd


def padronizaContatos(path_contatos,path_emissoes,path_resultado):
    a = pd.read_excel(path_contatos)
    b = pd.read_excel(path_emissoes)

    a = a[['CORRETOR', 'CONTATO','telefone','E-MAIL']]  
    #b = [['Corretor','Apolice', 'Segurado','Inicio Vigencia','Fim Vigencia','Cod Sucursal','Sucursal','CPF Inspetor de producao','Inspetor de producao','Ramo Seguro','Frota itens', 'CONTATO','telefone','E-MAIL']]
    df_unique_first = a.drop_duplicates(subset='CORRETOR', keep='first')

    # Renomear coluna em b para combinar com a chave de junção
    df_unique_first.rename(columns={'CORRETOR': 'Corretor'}, inplace=True)

    # Fazer a junção (VLOOKUP)
    c = b.merge(df_unique_first, on='Corretor', how='left')

    c = c[c['Inicio Vigencia'] != '01-01-1999']
    #c = c[c['Inicio Vigencia'] == '20-06-2024']
    c = c[['Corretor','Apolice', 'Segurado','Inicio Vigencia','Fim Vigencia','Cod Sucursal','Sucursal','CPF Inspetor de producao','Inspetor de producao','Ramo Seguro','Frota itens', 'CONTATO','telefone','E-MAIL']]
    c.to_excel(path_resultado)
