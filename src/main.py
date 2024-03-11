#importa pacotes
import pandas as pd
import os
import glob

#caminho para leitura os arquivos
folder = 'src\\data\\raw'

#lista todos os arquivos de excel
excel_files = glob.glob(os.path.join(folder,'*.xlsx'))

if not excel_files:
    print("Nenhum arquivo encontrado")
else:
    #dataframe
    dfs = []

    for file in excel_files:
        try:
            #ler arquivo excel
            df_temp = pd.read_excel(file)
            #pegar nome arquivo
            file_name = os.path.basename(file)
            #criar coluna location
            if 'brasil' in file_name.lower():
                df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it'

            #criar nova coluna campanha
            df_temp['campanha'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            #salva dados tratados   
            dfs.append(df_temp)
            print(dfs)

        except Exception as e:
            print(f"Erro ao ler o arquivo {file}: {e}")

    if dfs:
        #concatena todas as tabelas 
        result = pd.concat(dfs, ignore_index=True)

        #caminho de saida
        output_file = os.path.join('src','data','ready','clean.xlsx')
        
        #configurar a escrita
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        #Escrever arquivo
        result.to_excel(writer)

        writer._save()

    else:
        print("Nenhuma dado a ser salvo:")

