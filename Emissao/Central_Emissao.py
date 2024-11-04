import Padronizar_emissao_v8 as pe
import agrupa_excel_v2 as ag
import os
import contatos_v2 as ct

import novo_v2

import converttxt

####Padronizar renovacao

path_raiz = r'C:\Users\Thomas' #Mude o caminho para o do seu computador
ano_mes = '24_08'

padroniza = True
agrupa = True
plan_grades = True

#############################

contatos = False
filtrar_data = False #Deixe True para filtrar e False para nao filtrar
data_contatos = '21-06-2024'

#############################

caminho_original_p = r'Downloads\Base_v2\Emissao\Dados_Emissao\Original' #Crie um caminho nesses padroes nem seu compudator e cole aqui entre as aspas
caminho_original = os.path.join(path_raiz, caminho_original_p)

caminho_editado_p = r'Downloads\Base_v2\Emissao\Dados_Emissao\Editado'   #Crie um caminho nesses padroes nem seu compudator e cole aqui entre as aspas
caminho_editado = os.path.join(path_raiz, caminho_editado_p)

if padroniza == True:
    ag.del_special_files(caminho_original)
    ag.delete_all_contents(caminho_editado)
    file_paths = pe.list_files_in_folder(caminho_original)

    for i in file_paths:
        
        #pe.padronizaEmissao(i,caminho_editado,ano_mes)
        novo_v2.StanEmissions(i,caminho_editado)

if agrupa == True:

    path_grupo = r'Downloads\Base_v2\Emissao\Dados_Emissao\Agrupado\Emissoes_agrupado.xlsx' #caminho igual #arrumar esta funcao por conta do caminho esta mal feita
    destino_agrupado = os.path.join(path_raiz, path_grupo)
    destino_agrupadotxt = os.path.join(path_raiz, r'Downloads\Base_v2\Emissao\Dados_Emissao\Agrupado\Emissoes_agrupado.txt')
    pasta_apagar = os.path.join(path_raiz, r'Downloads\Base_v2\Emissao\Dados_Emissao\Agrupado')
    ag.delete_all_contents(pasta_apagar)
    ag.agrupaEx(caminho_editado,destino_agrupado)
    converttxt.ExcelToTxtConverter(destino_agrupado, destino_agrupadotxt, colunas_monetarias=['Valor Premio Liquido','Valor Premio Total'], delimiter=';')

if plan_grades == True:
    path_grupo = r'Downloads\Base_v2\Emissao\Dados_Emissao\Agrupado\Emissoes_agrupado.xlsx' #caminho igual #arrumar esta funcao por conta do caminho esta mal feita
    path_grades = r'Downloads\Base_v2\Emissao\Dados_Emissao\Grades\Grades_agrupadas.xlsx'
    destino_agrupado = os.path.join(path_raiz, path_grupo)
    grade_agrupado = os.path.join(path_raiz, path_grades)
    pasta_apagar = os.path.join(path_raiz, r'Downloads\Base_v2\Emissao\Dados_Emissao\Grades')
    
    ag.delete_all_contents(pasta_apagar)
    ag.plan_grades(destino_agrupado,grade_agrupado)

if contatos == True:
    path_contatos = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Contatos\PlanContatos\Lista de corretores consolidada.xlsx')
    path_emissoes = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Emissao\Dados_Emissao\Agrupado\Emissoes_agrupado.xlsx')
    path_resultados = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Emissao\Dados_Emissao\Contatos\EmissoesAgContatos.xlsx')

    path_endosso = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Emissao\Dados_Emissao\Endosso\endossos.xlsx')


    ct.padronizaContatos(path_contatos, path_emissoes, path_resultados,filtrar_data,data_contatos,path_endosso)


