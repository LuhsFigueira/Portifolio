import pandas as pd
import requests
import json

import pandas as pd
tabela = pd.read_excel(r"C:\Users\luis.oliveira\Downloads\empresas.xlsx",dtype=str)
display(tabela)


for linha in tabela.index:
    
    cnpj = tabela.loc[linha,"CNPJCPF"]

    if len(cnpj) == 14:
        link = f"https://minhareceita.org/{cnpj}"
        cnpj = requests.get(link)
        cnpj = cnpj.json()
        tabela.loc[linha, "SITUACAORECEITA"] = str(cnpj) 
        #print(cnpj)
        cnaefiscaldescricao = cnpj['cnae_fiscal_descricao'] 
        CNAE = cnpj['cnae_fiscal']
        SITUACAORECEITA = cnpj['descricao_situacao_cadastral']
        tabela.loc[linha, "DESCRICAO"] = str(cnaefiscaldescricao)
        tabela.loc[linha, "CNAE"] = str(CNAE)
        tabela.loc[linha, "SITUACAORECEITA"] = str(SITUACAORECEITA)
   
    else:
        tabela.loc[linha, "DESCRICAO"] = str("Cliente é Pessoal Fisica")
        tabela.loc[linha, "CNAE"] = str("Cliente é Pessoal Fisica")   
        tabela.loc[linha, "SITUACAORECEITA"] = str("Cliente é Pessoal Fisica")

    #print(cnpj)
    #print(razao_social)
    
    
    
display(tabela)

tabela.to_excel(r"C:\Users\luis.oliveira\Downloads\empresas_atualizado.xlsx", index=False)
