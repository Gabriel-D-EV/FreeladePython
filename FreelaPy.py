import pandas as pd
import os
import xmltodict 

def pegar_inf(arq, valores):
    with open(f'NFs/{arq}','rb') as arq_xml:
        dic_arq = xmltodict.parse(arq_xml)
    
    if 'NFe' in dic_arq:
        inf_nf = dic_arq['NFe']['infNFe']
    else:
        inf_nf = dic_arq['nfeProc']['NFe']['infNFe']
    n_nota = inf_nf['@Id']
    empresa = inf_nf['emit']['xNome']
    cliente = inf_nf['dest']['xNome'] 
    endereço = inf_nf['dest']['enderDest']
    if 'vol' in inf_nf['transp']:
        peso = inf_nf['transp']['vol']
    else:
        peso = 'não informado'
    valores.append([n_nota, empresa, cliente, endereço, peso])        
        
coluna = ['numero_nota', 'empresa', 'cliente', 'endereço', 'peso']
valores = []
lts_arq = os.listdir('NFs')
for arquivo in lts_arq:
    pegar_inf(arquivo, valores)
    
    tabela = pd.DataFrame(columns=coluna, data=valores)
    #display(tabela)
    #print(tabela)
    
tabela.to_excel('NotasFiscais.xlsx', index=False)
