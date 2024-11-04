import Padronizar_Renovacao_v7 as pr
import Imagens_renovacao_v3 as ir
import agrupa_excel_v2 as ag
import contatos_v2 as ct
import Padronizar_Renovacao_v8

import os


#1Certifique-se de que não há nenhum excel aberto em sua máquina
#2Certifique-se de que os arquivos em excel originais estão alocados na pasta correta

path_raiz = r'C:\Users\Thomas' #Mude o caminho para o do seu computador
ano_mes = '24_10'              #Mude para o mes da renovação

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
padroniza = True              #Deixe True se quiser executar ou False se não quiser executar
############
imagem = False                 #Deixe True se quiser executar ou False se não quiser executar
date = False ##### Filtro quinzena                  #True se quiser filtrar as imagens por data   
data_inicio = '15-06-2024'
data_fim = '30-06-2024'
##################
imagem_interna = False      #Deixe True se quiser executar ou False se não quiser executar
agrupa = True                 #Deixe True se quiser executar ou False se não quiser executar
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
imagem_gestao = False
renovacoes_contatos = False




#Ao executar apagara os arquivos existentes na pasta

caminho_original_p = r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Original' #Crie um caminho nesses padroes nem seu compudator e cole aqui entre as aspas
caminho_original = os.path.join(path_raiz, caminho_original_p)

caminho_editado_p = r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Editado'   #Crie um caminho nesses padroes nem seu compudator e cole aqui entre as aspas
caminho_editado = os.path.join(path_raiz, caminho_editado_p)

arquivos_editado = pr.list_files_in_folder(caminho_editado)

####Padronizar renovacao

if padroniza == True:
    
    ag.delete_all_contents(caminho_editado)
    file_paths = pr.list_files_in_folder(caminho_original)

    for i in file_paths:
        
        #pr.padronizaRenovacao(i,caminho_editado,ano_mes)
        Padronizar_Renovacao_v8.StanRenovacoes(i,caminho_editado)



#Imagens renovacao corretor individual

###caminho_pasta_img = r'C:\Users\Thomas\Downloads\Base_v2\Renovacao\Dados_Renovacao\Imagens\Corretor_Individual\24_05'


###arquivos_editado = pr.list_files_in_folder(caminho_editado)

if imagem == True:
    
    path_igual = r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Imagens\Corretor_Individual' #caminho igual
    caminho_pasta_img = os.path.join(path_raiz, path_igual)
    ag.delete_all_contents(caminho_pasta_img)
    arquivos_editado = pr.list_files_in_folder(caminho_editado)

    for i in arquivos_editado:
        

        file_name_no_extension = os.path.splitext(os.path.basename(i))[0]
        folder_path = ir.create_folder(caminho_pasta_img,file_name_no_extension)
        ir.create_img_individual(i,folder_path, date, data_inicio, data_fim)

#Relatorio interno comercial

if imagem_interna == True:

    ###caminho_pasta_img = r'C:\Users\Thomas\Downloads\Base_v2\Renovacao\Dados_Renovacao\Imagens\Relatorio_Interno_Comercial\24_05'
    
    path_img_interna = r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Imagens\Relatorio_Interno_Comercial' #caminho igual
    caminho_pasta_int = os.path.join(path_raiz, path_img_interna)
    ag.delete_all_contents(caminho_pasta_int)  
    arquivos_editado = pr.list_files_in_folder(caminho_editado)
    for i in arquivos_editado:
        file_name_no_extension = os.path.splitext(os.path.basename(i))[0]
        folder_path = ir.create_folder(caminho_pasta_int,file_name_no_extension)
        ir.create_img_zerados(i,folder_path)

#Código responsável por agrupar não é necessário alterar nada aqui

if agrupa == True:
    
    path_grupo = r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Agrupado\Renovacoes_agrupado.xlsx' #caminho igual #arrumar esta funcao por conta do caminho esta mal feita
    destino_agrupado = os.path.join(path_raiz, path_grupo)
    pasta_apagar = os.path.join(path_raiz, r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Agrupado')
    ag.delete_all_contents(pasta_apagar)
    ag.agrupaEx(caminho_editado,destino_agrupado)

if imagem_gestao == True:

    path_grupo = r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Agrupado\Renovacoes_agrupado.xlsx' #caminho igual #arrumar esta funcao por conta do caminho esta mal feita
    destino_agrupado = os.path.join(path_raiz, path_grupo)

    
    path_resultadoss = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Imagens\Gestao\FrotaEquipe.png')

    ir.create_img_gestao(destino_agrupado,path_resultadoss)

if renovacoes_contatos == True:
    path_contatos = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Contatos\PlanContatos\Lista de corretores consolidada.xlsx')
    path_emissoes = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Agrupado\Renovacoes_agrupado.xlsx')
    path_resultados = ag.ajeita_path(path_raiz, r'Downloads\Base_v2\Renovacao\Dados_Renovacao\Contatos\RenovacoesAgContatos.xlsx')

    ct.padronizaContatos(path_contatos, path_emissoes, path_resultados)













