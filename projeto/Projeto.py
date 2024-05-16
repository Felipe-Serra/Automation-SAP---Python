import pyautogui as pg
import pandas as pd
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
 
permissão=0
tabela = pd.read_excel("Projeto de automacao SAP.xlsx", sheet_name="Organizador de planilha - Sant")
tabela = tabela.drop(columns=["Nosso Número", "Data da Liquidação", "Pagador","Conta Cobrança", "Tipo Cobrança / Modalidade", "Responsável"])
print(tabela)
 
 
 
pg.hotkey("alt", "tab")
pg.write("FBL5N") # FBL5N - Partidas individuais de clientes / Rel CLIENTES
pg.press("enter")
time.sleep(0.5)
 
for linha in tabela.index:
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
    #coluna = coluna.str.counts('r')
    print("\n^^^^^^^\n",coluna,"\n**",coluna1,"**""\n^^^^^^^^^\n")
 
    i=0
    while i < coluna1:
        #Pegar valor de N Doc
        time.sleep(1)
        pegarDoc()
        pegarNDoc()
 
        texto1 = float(pyperclip.paste())
        print(pyperclip.paste())
        compDoc = pd.DataFrame({"Comp Doc":[texto1]})
        new_tabela = pd.concat([tabela, compDoc], axis=1)
        print(new_tabela)
 
        # Adicionando nova coluna para comparacao de Data
        time.sleep(1)
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
                permissão=1 #Permissão Para rodar Script do SAP
                j=i
            else:(
                #colocar bloq pag
                print("Bloqueio de pagamento em: ",(i+1))
            )
        else:(
            print("Doc diferentes na tabela SAP em:",(i+1))
            )
        i=i+1
    pg.press("f3")
    #Rodar Script do SAP            
    if(permissão==1):
        print("Rodar script em linha: ",j)
        permissão=0
 
pg.hotkey("alt", "tab")