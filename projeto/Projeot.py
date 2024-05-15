import pyautogui as pg
import pandas as pd
import time
import pyperclip
pg.PAUSE = 0.3
 
tabela = pd.read_excel("Projeto de automacao SAP.xlsx", sheet_name="Organizador de planilha - Sant")
tabela = tabela.drop(columns=["Oscilação (R$)", "Pagador", "Responsável", "Data da Liquidação", "Seu Número"])
print(tabela)
 
#pg.click(x=108, y=402, clicks=2) #cliente
#time.sleep(0.5)
for linha in tabela.index:
    # Adicionando nova coluna para comparacao de N Doc
 
    #pg.write(str(tabela.loc[linha, "Cod Sap"]))
    #pg.click(x=22, y=112) #Rodar
    #time.sleep(5)
    #pg.click(x=872, y=289) #N Doc - cabeçalho
    #pg.hotkey("ctrl", "c") #Pegar coluna de N Doc
    coluna1=0
    coluna = pyperclip.paste()
    coluna1 = coluna.count('021')
    #coluna = coluna.str.counts('r')
    print("\n^^^^^^^\n",coluna,"\n**",coluna1,"**""\n^^^^^^^^^\n")
    
    #Pegar valor de N Doc
    #pg.hotkey("down")
    #pg.hotkey("ctrl", "c")
    texto1 = pyperclip.paste()
    compDoc = pd.DataFrame({"Comp Doc":[31095]})
    new_tabela = pd.concat([tabela, compDoc], axis=1)
 
    # Adicionando nova coluna para comparacao de Data
 
    #time.sleep(2)
    #pg.click(x=872, y=289) #Data - cabeçalho
    #pg.hotkey("down")
    #pg.hotkey("ctrl", "c") #Pegar valor de Data
    texto2 = pyperclip.paste()
    compData = pd.DataFrame({"Comp Data":[r"04/05/2024"]})
    new_tabela = pd.concat([new_tabela, compData], axis=1)
 
    print(new_tabela)
    i=0
    while i <= coluna1:
        # Comparando Ndoc e ref
        numDoc=float(new_tabela.loc[linha,"Nosso Número"])
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
                print("Bloqueio de pagamento em: ",i)
            )
        i+=1
    
    #Rodar Script do SAP            
    if(permissão==1):
        print("Rodar script em: ",j)
        permissão=0

