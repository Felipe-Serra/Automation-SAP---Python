import pyautogui as pg
import pandas as pd
import openpyxl
import time
import pyperclip
 
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
def pegarNData():
    pg.hotkey("ctrl","home")
    pg.keyDown("shift") #Vencimento liq - cabeçalho
    pg.press("tab", presses=12)
    pg.keyUp("shift")
 
    pg.hotkey("ctrl","y")
    pg.press("down", presses=(i+1)) #vai de 1 até o final da coluna1 (tamanho coluna SAP)
    pg.hotkey("ctrl", "c")
 
permissao=0
 
tabela = pd.read_excel("Base de dados.xltm", sheet_name="DZ - Juros")
#tabela = tabela.drop(columns=["Nosso Número", "Data da Liquidação", "Pagador", "Responsável"])
print(tabela)
 
pg.hotkey("alt", "tab")
pg.write("FBL5N") # FBL5N - Partidas individuais de clientes / Rel CLIENTES
pg.press("enter")
time.sleep(0.5)
 
for linha in tabela.index:
    supbloq= 1
    supdata= 0
 
    # Selecionar cliente na FBL5N
    pg.write(str(tabela.loc[linha, "Cod Sap"]))
    pg.press("f8") #Rodar
    time.sleep(2)
 
    # Adicionando nova coluna para comparacao de N Doc
    pegarDoc()
    pg.hotkey("ctrl", "space") #Pegar coluna de N Doc  
    pg.hotkey("ctrl","c")
 
    coluna1=0
    coluna = pyperclip.paste()
    coluna1 = coluna.count('\n') - 3
    print("\n^^^^^^^\n",coluna,"\n**",coluna1,"**""\n^^^^^^^^^\n")
 
    i=0
    while i < coluna1:
        #Pegar valor de N Doc
        time.sleep(0.3)
        pegarDoc()
        pegarNDoc()
 
        texto1 = float(pyperclip.paste())
        compDoc = pd.DataFrame({"Comp Doc":[texto1]})
        new_tabela = pd.concat([tabela, compDoc], axis=1)
       
        # Adicionando nova coluna para comparacao de Data
        time.sleep(0.3)
        pegarNData()
       
        texto2 = pyperclip.paste()
        compData = pd.DataFrame({"Comp Data":[texto2]})
        new_tabela = pd.concat([new_tabela, compData], axis=1)
        print(new_tabela)
 
        # Comparando Ndoc e ref
        numDoc=float(new_tabela.loc[linha,"Seu Número"])
        numRef=float(new_tabela.loc[linha,"Comp Doc"])
 
        if(numDoc == numRef): #Se Doc=ref
 
            # Comparando data de vencimento
            dataDoc=new_tabela.loc[linha,"Vencimento"]
            dataRef=new_tabela.loc[linha,"Comp Data"]
 
            if(dataRef==dataDoc):            
                permissao=1 #Permissão Para rodar Script do SAP
                fatura=new_tabela.loc[linha,"Seu Número"]
 
                #tirar bloq pag
                pg.press("f2")
                pg.hotkey("shift", "f1")
                pg.press("down", presses=2)
                pg.write("backspace")
                pg.hotkey("crtl", "s")  
                pg.press("f3")
                supdata=1
 
 
            else:
                #colocar bloq pag
                print("Bloqueio de pagamento em: ",(i+1))
                pg.press("f2")
                pg.hotkey("shift", "f1")
                pg.press("down", presses=2)
                pg.write("a")
                pg.hotkey("crtl", "s")  
                pg.press("f3")
     
                supbloq=1    
        else:
            print("Doc diferentes na tabela SAP em:",(i+1))
            print(new_tabela.loc[linha])
           
        i=i+1
 
    pg.press("f3")
 
    #Rodar Script do SAP            
    if(permissao==1):
 
        print("\nRodar script em\n",fatura)
        permissao=0
 
    # LOG de dados -Tabela de retorno
    # Fatura não encontrada
    if (supbloq & supdata  == 0):
        open("FaturaNA.txt", 'w').write(str('')) #Zerando o arquivo
        lista1=tabela.iloc[[linha]]
        for index, row in lista1.iterrows():
            lista1.to_csv("FaturaNA.txt", sep=' ', index=False, header=False, mode='a')
 
    # Fatura com data divergente
    if (supbloq==1 & supdata==0):
        open("FaturaData.txt", 'w').write(str('')) #Zerando o arquivo
        lista2=tabela.iloc[[linha]]
        for index, row in lista2.iterrows():
            lista2.to_csv("FaturaData.txt", sep=' ', index=False, header=False, mode='a')
 
    # Fatura baixada
    if ((supbloq==1 | supbloq==0) & supdata==1):
        open("FaturaBaixada.txt", 'w').write(str('')) #Zerando o arquivo
        lista3=tabela.iloc[[linha]]
        for index, row in lista2.iterrows():
            lista3.to_csv("FaturaBaixada.txt", sep=' ', index=False, header=False, mode='a')
 
 
 
pg.hotkey("alt", "tab")