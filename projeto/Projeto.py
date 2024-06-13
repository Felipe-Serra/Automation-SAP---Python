import pyautogui as pg
import pandas as pd
import openpyxl
import time
import clipboard
from SAP import getClient
from SAP import main
 
pg.PAUSE = 0.5
 
def pegarDoc():
    pg.hotkey("ctrl","home")
    pg.keyDown("shift") #N Doc - cabeçalho
    pg.press("tab", presses=8)
    pg.keyUp("shift")
def pegarNDoc():
    pg.hotkey("ctrl","y")
    pg.press("down", presses=(i+1)) #vai de 1 até o final da coluna1 (tamanho coluna SAP)
    pg.hotkey("ctrl", "c")
    pg.hotkey("ctrl", "c")
def pegarNData():
    pg.hotkey("ctrl","home")
    pg.keyDown("shift") #Vencimento liq - cabeçalho
    pg.press("tab", presses=12)
    pg.keyUp("shift")
 
    pg.hotkey("ctrl","y")
    pg.press("down", presses=(i+1)) #vai de 1 até o final da coluna1 (tamanho coluna SAP)
    pg.hotkey("ctrl", "c")
    pg.hotkey("ctrl", "c")
# def pegarFatura():
#     data = fatura.loc['Vencimento']
#     codSap = fatura.loc['Cod SAP']
#     documento = fatura.loc['Fatura']
#     razao = fatura.loc['Conta Razão']
#     valorC = fatura.loc['Valor cobrado']
#     centroC = fatura.loc['Centro de custo']
 
 
open("FaturaNA.txt", 'w').write(str('')) #Zerando o arquivo
open("FaturaData.txt", 'w').write(str('')) #Zerando o arquivo
open("FaturaBaixada.txt", 'w').write(str('')) #Zerando o arquivo
 
tabela = pd.read_excel("Base de dados.xlsm", sheet_name="DZ - Juros")
#tabela = tabela.drop(columns=["Nosso Número", "Data da Liquidação", "Pagador", "Responsável"])
#print(tabela)
 
getClient()
pg.hotkey("alt", "tab")
pg.write("FBL5N") # FBL5N - Partidas individuais de clientes / Rel CLIENTES
pg.press("enter")
time.sleep(0.5)
 
for linha in tabela.index:
 
    supbloq= 0
    supdata= 0
    permissao=0  
    # Selecionar cliente na FBL5N
    pg.write(str(tabela.loc[linha, "Cod SAP"]))
    pg.press("f8") #Rodar
    time.sleep(2) #Timer fixo
 
    # Adicionando nova coluna para comparacao de N Doc
    pegarDoc()
    pg.hotkey("ctrl", "space") #Pegar coluna de N Doc  
    pg.hotkey("ctrl","c")
   
 
    coluna1=0
    coluna = clipboard.paste()
    coluna1 = coluna.count('\n') - 3
    print("\n^^^^^^^\n",coluna,"\n**",coluna1,"**""\n^^^^^^^^^\n")
 
    i=0
    while i < coluna1:
       
        #Pegar valor de N Doc
        #time.sleep(0.3)
        pegarDoc()
        pegarNDoc()
       
        texto1 = clipboard.paste()
        compDoc = pd.DataFrame({"Comp Doc":[]})
        compDoc.loc[linha] = [texto1[1:]]
        new_tabela = pd.concat([tabela, compDoc], axis=1)
       
        # Comparando Ndoc e ref
        numDoc=str(new_tabela.loc[linha,"Fatura"])
        numRef=str(new_tabela.loc[linha,"Comp Doc"])
 
        print(new_tabela.loc[linha])
        if(numDoc == numRef): #Se Doc=ref
 
            # Adicionando nova coluna para comparacao de Data
            #time.sleep(0.3)
            pegarNData()
           
            texto2 = clipboard.paste()
            compData = pd.DataFrame({"Comp Data": []})
            compData.loc[linha] = [texto2]
            new_tabela = pd.concat([new_tabela, compData], axis=1)
 
            # Comparando data de vencimento
            dataDoc=str(new_tabela.loc[linha,"Vencimento"])
            dataRef=str(new_tabela.loc[linha,"Comp Data"])
 
            print(new_tabela.loc[linha])
 
            if(dataRef==dataDoc):            
                permissao=1 #Permissão Para rodar Script do SAP
                fatura=new_tabela.loc[linha]
 
                #tirar bloq pag
                pg.press("f2")
                pg.hotkey("shift", "f1")
                pg.press("down", presses=2)
                pg.press("backspace")
                pg.press("enter")
 
                pg.hotkey("shift", "f1")
                pg.hotkey("crtl", "s")
                pg.press("enter")
                time.sleep(2) #Timer Fixo
                pg.press("f3")
                pg.press("f3")
                supdata=1
 
 
            else:
                #colocar bloq pag
                print("Bloqueio de pagamento em: ",(i+1))
                pg.press("f2")
                pg.hotkey("shift", "f1")
                #time.sleep(0.5)
                pg.press("down", presses=2)
                pg.press("backspace")
                pg.write('a')
                pg.press("enter")
 
                pg.hotkey("shift", "f1")
                pg.hotkey("crtl", "s")
                pg.press("enter")
                time.sleep(2) #Timer Fixo
                pg.press("f3")
                pg.press("f3")
                #time.sleep(0.5) #timer Fixo
 
                supbloq=1  
        else:
            print("Doc diferentes na tabela SAP em:",(i+1))
            print(new_tabela.loc[linha])
           
        i=i+1
 
    pg.press("f3")
 
    #Rodar Script do SAP            
    if(permissao==1):
 
        print("\nRodar script em\n",fatura)
        main(fatura.loc['Vencimento'],fatura.loc['Cod SAP'],fatura.loc['Fatura'],fatura.loc['Conta Razão'],fatura.loc['Valor cobrado'],fatura.loc['Centro de custo'])
        permissao=0
 
    # LOG de dados -Tabela de retorno
    # Fatura baixada
    if (supdata ==1):
       
        lista3=tabela.iloc[[linha]]
        for index, row in lista3.iterrows():
            lista3.to_csv("FaturaBaixada.txt", sep=' ', index=False, header=False, mode='a')
            print("Baixou", fatura)
 
    # Fatura não encontrada
    if (supdata ==0 and supbloq==0):
       
        lista1=tabela.iloc[[linha]]
        for index, row in lista1.iterrows():
            lista1.to_csv("FaturaNA.txt", sep=' ', index=False, header=False, mode='a')
            print("Não Baixou", fatura)
 
    # Fatura com data divergente
    if (supdata ==0 and supbloq==1):
       
        lista2=tabela.iloc[[linha]]
        for index, row in lista2.iterrows():
            lista2.to_csv("FaturaData.txt", sep=' ', index=False, header=False, mode='a')
 

pg.hotkey("alt", "tab")