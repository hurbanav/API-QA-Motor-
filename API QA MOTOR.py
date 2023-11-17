
import requests 
import pandas as pd
import json
from datetime import datetime
import os
from pathlib import Path

def refazer_base(filename , produto):
    print(filename)
    oldfilepath = filename #get_latest_file(0)
    base = pd.read_excel(oldfilepath, sheet_name = 'Sheet1')
    n = max(base['TotalAnnualPremium'].count(), base['TotalNetPremium'].count() or base['Error_Code'].count())+1
    newfilepath = oldfilepath  #r'\\\Taxas '+ str(base.iloc[n , 2]) + ' - ' + processo +'.xlsx
    row_count = len(base.index)
    final = row_count-1
    print (row_count)
    for index, row in base.iterrows() :
        if base.iloc[n , 12] != "" or base.iloc[n , 13] == "":
            base.loc[n, "TotalAnnualPremium"] = base.loc[n, "TotalAnnualPremium"]                        #cola o resultado do POST
            base.loc[n, "TotalNetPremium"] = base.loc[n, "TotalNetPremium"]
            base.loc[n, "RatePerThousand"] = base.loc[n, "RatePerThousand"] 
            base.loc[n, "Error_Code"] = ""
            n = n+1
        else:
            n = row
            cm = str(base.iloc[n , 1])
            pf = str(base.iloc[n , 5])
            bd = str(base.iloc[n , 6])
            si = str(base.iloc[n , 4])
            gender = str(base.iloc[n , 3])
            prodcode = str(base.iloc[n , 2])
            ia = str(base.iloc[n,8])
            covcode = str(base.iloc[n , 0])
            url = "API URL"
            
            payload = json.dumps({
              "DatToBeUsedForTariff": datetime.today().strftime('%Y-%m-%d'), #fixo
              "LineOfBusiness": 13,#fixo
              "ProductCode": prodcode,#cód de produto
              "CoverageModule": cm, #plano
              "PaymentFrequency": pf, #pf
              "TypeOfDistributionChannel": 46,#fixo
              "Insured": {
                "BirthDate": bd, #bd
                "SmokerIndicator": si, #si
                "Gender": gender #gender
              },
              "Coverages": [
                {
                  "CoverageCode": covcode, #covcode
                  "InsuredAmount": ia, #ia
                  "AgravioMultipl1": 1, #comercialrate
                  "AgravioMultipl2": 1,#fixo
                  "AgravioAditivo": 0 #fixo
                }
              ],
              "Assistance": [
                {
                  "CoverageCode": 20
                }
              ]
            })
            headers = {
              'Content-Type': 'application/json'
            }
    
            response = requests.request("POST", url, headers=headers, data=payload)
    
            #print(response.elapsed.total_seconds())
            
            responsestatus = response.status_code
            #print(responsestatus)
            response = response.text
            data = json.loads(response) #quebra a resposta da API em umm dicionário com listas
            if responsestatus == 200:
                rpt = (data["Coverages"][0]['RatePerThousand'])  
                
                '''O comando acima está pegando apenas a lista, que é o módulo da API, Coverages, como essa 
                lista tem apenas um item com todos os subitens, pegamos a coluna zero, e dentro dessa coluna 
                pegamos apenas a chave RatePerThousand, separando apenas o que queremos da API.'''
                        
                base.loc[n, "TotalAnnualPremium"] = (data["TotalAnnualPremium"])                       #cola o resultado do POST
                base.loc[n, "TotalNetPremium"] = (data["TotalNetPremium"])
                base.loc[n, "RatePerThousand"] = (rpt)
                base.loc[n, "Error_Code"] = ""
                print(responsestatus)
            else:
                #base.loc[n, "Error_Code"] = (data[0]['Message'])
                base.loc[n, "Error_Code"] = response
            n = n + 1
        #if n%50 == 0 or n == final:
        base.to_excel(newfilepath, index = False)
        print (n)
    base.to_excel(newfilepath, index = False)  
    
def get_latest_file(ordem):




    # Save all .xlsx files paths and modification time into paths
    paths = [(p.stat().st_mtime, p) for p in Path(r'C:\Users\urbanaviciushr\Testes Novos Produtos').iterdir() if p.suffix == ".xlsx"]

    # Sort them by the modification time
    paths = sorted(paths, key=lambda x: x[0], reverse=True)

    # Get the last modified file
    last = paths[ordem][1] #O primeiro número indica qual arquivo você quer, colocando zero, pegamos o último, colocando 1 pegamos o penultimo e assim vai '
    latest_file = last
    return latest_file


    # Save all .xlsx files paths and modification time into paths
    paths = [(p.stat().st_mtime, p) for p in Path(r'\\\Respostas ' + str(produto)).iterdir() if p.suffix == ".xlsx"]

    # Sort them by the modification time
    paths = sorted(paths, key=lambda x: x[0], reverse=True)

    # Get the last modified file
    last = paths[ordem][1] #O primeiro número indica qual arquivo você quer, colocando zero, pegamos o último, colocando 1 pegamos o penultimo e assim vai '
    latest_file = last
    return latest_file
def main(planilha, produto):
    planatual = planilha
    oldfilepath =     r'FILEPATH' + str(produto) + '.xlsx'
    planatual = planatual
    #if 
    base = pd.read_excel(oldfilepath, sheet_name = planatual)
    base2 = pd.read_excel(oldfilepath, sheet_name = "Funeral Individual")
    n = 0
    newfilepath = r'\\\\Taxas '+ str(base.iloc[n , 2]) + ' - ' + planatual +'.xlsx'
    if newfilepath == str(get_latest_file(0)):
        oldfilepath = get_latest_file(0)
        planatual = planatual
        base = pd.read_excel(oldfilepath, sheet_name = 'Sheet1')
        n = max(base['TotalAnnualPremium'].count(), base['TotalNetPremium'].count() or base['Error_Code'].count())+1
    print(n)
    row_count = len(base.index)
    final = row_count-1
    print (row_count)
    print(newfilepath)
    while n <= final :
        cm = str(base.iloc[n , 1])
        pf = str(base.iloc[n , 5])
        bd = str(base.iloc[n , 6])
        si = str(base.iloc[n , 4])
        gender = str(base.iloc[n , 3])
        prodcode = str(base.iloc[n , 2])
        ia = str(base.iloc[n,8])
        covcode = str(base.iloc[n , 0])
        codcov_FuneralInd = str(base2[n, 0])
        url = "API URL"
        payload = json.dumps({
          "DatToBeUsedForTariff": datetime.today().strftime('%Y-%m-%d'), #fixo
          "LineOfBusiness": 13,#fixo
          "ProductCode": prodcode,#cód de produto
          "CoverageModule": cm, #plano
          "PaymentFrequency": pf, #pf
          "TypeOfDistributionChannel": 46,#fixo
          "Insured": {
            "BirthDate": bd, #bd
            "SmokerIndicator": si, #si
            "Gender": gender #gender
          },
          "Coverages": [
            {
              "CoverageCode": codcov_FuneralInd, # 7 covcode
              "InsuredAmount": ia, #ia
              "AgravioMultipl1": 1, #comercialrate
              "AgravioMultipl2": 1,#fixo
              "AgravioAditivo": 0 #fixo
            },
            {
              "CoverageCode": covcode, #covcode
              "InsuredAmount": ia, #ia
              "AgravioMultipl1": 1, #comercialrate
              "AgravioMultipl2": 1,#fixo
              "AgravioAditivo": 0 #fixo
            }
          ],
          "Assistance": [
            {
              "CoverageCode": 20
            }
          ]
        })
        headers = {
          'Content-Type': 'application/json'
        }
        response = requests.request("POST", url, headers=headers, data=payload)
        #print(response.elapsed.total_seconds())
        responsestatus = response.status_code
        #print(responsestatus)
        response = response.text
        data = json.loads(response) #quebra a resposta da API em umm dicionário com listas
        if responsestatus == 200:
            if planilha == 'Funeral Cônjuge':
                rpt = (data["Coverages"][1]['RatePerThousand'])  #para fun conj pegar item 1 
            else:
                rpt = (data["Coverages"][0]['RatePerThousand'])  
            
            '''O comando acima está pegando apenas a lista, que é o módulo da API, Coverages, como essa 
            lista tem apenas um item com todos os subitens, pegamos a coluna zero, e dentro dessa coluna 
            pegamos apenas a chave RatePerThousand, separando apenas o que queremos da API.'''
                    
            base.loc[n, "TotalAnnualPremium"] = (data["TotalAnnualPremium"])                       #cola o resultado do POST
            base.loc[n, "TotalNetPremium"] = (data["TotalNetPremium"])
            base.loc[n, "RatePerThousand"] = (rpt)
        else:
            #base.loc[n, "Error_Code"] = (data[0]['Message'])
            base.loc[n, "Error_Code"] = response
        n = n + 1
        if n%50 == 0 or n == final:
            base.to_excel(newfilepath, index = False)
            print ("Salvo")
    base.to_excel(newfilepath, index = False)
    
def roda_automatico(produto):
    p1 =  "Morte Temp" #"Temporario" ok
    p2 = "Morte Desc"   #ok
    p3 = "DG I" #ok
    p4 = "Morte Ac" #ok
    p5 = "Invalidez Ac"#
    p6 = "Básica" #
    p7 = "IPDF"#
    p8 = "Funeral Individual"
    p9 = "Funeral Cônjuge"
    p10 = "Invalidez Ac"
    p11 = "DIH (Falta)"
    p12 = "AmparoConj"
    p13 = "AmparoFamiliar"
    p14 = "AmparoTit"
    p15 = "BrokenBones"
    p16 = "PRIT"
    if produto == 790:
        main(p8, produto)
        main(p9, produto)
    else:
        #main(p1, produto)
        #main(p2, produto)
        #main(p3, produto)
        #main(p4, produto)
        #main(p5, produto)
        #main(p6, produto)
        #main(p7, produto)
        #main(p8, produto)
        main(p16, produto)
        #main(p10, produto)
        #main(p11, produto)
        #main(p12, produto)
        #main(p13, produto)
        #main(p14, produto)
        #main(p15, produto)
        #main(p16, produto)

def reprocessa_automatico_rapido(produto):
    import os
    from time import sleep
    arquivos = os.listdir(r'C:\Users\urbanaviciushr\Testes Novos Produtos\Respostas ' + str(produto))
    for arquivo in arquivos:
        caminho = os.path.join(r'C:\Users\urbanaviciushr\Testes Novos Produtos\Respostas ' + str(produto), arquivo)
        if checa_erro(caminho, produto) == True:
            #refazer_base(caminho, produto)
            print(arquivo)
            
def reprocessa_automatico_lento(produto):
    
    arquivos = os.listdir(r'C:\Users\urbanaviciushr\Testes Novos Produtos\Respostas ' + str(produto))
    n=0
    while n <= len(arquivos):
        arquivo = arquivos[n]
        caminho = os.path.join(r'C:\Users\urbanaviciushr\Testes Novos Produtos\Respostas ' + str(produto), arquivo)
        if checa_erro(caminho, produto) == True:
            refazer_base(caminho, produto)
            if refazer_base(caminho, produto) == True:
                n = n+1
        else:
            n = n+1

def checa_erro(filename, produto):
    
    oldfilepath = filename #get_latest_file(0)
    base = pd.read_excel(oldfilepath)
    #tem_erro = (base['Error_Code'] != 0).any()
    if 'Error_Code' in base.columns:
        return True
    else:
        return False
            
if __name__ == "__main__":
    roda_automatico(791)
    #main()
    #get_latest_file()
    #reprocessa_automatico_rapido(790)
    
    
    