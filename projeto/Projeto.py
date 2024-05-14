import pyautogui as pg
import pandas as pd
import time
import pyperclip
pg.PAUSE = 0.3

tabela = pd.read_excel("projeto\ProjetodeautomacaoSAP.xlsx", sheet_name="Organizador de planilha - Sant")
tabela = tabela.drop(columns=["Oscilação (R$)", "Pagador", "Responsável", "Data da Liquidação"])
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
    coluna = pyperclip.paste()
    coluna = coluna.str.counts('r')
    print("\n^^^^^^^\n",coluna,"\n^^^^^^^^^\n")

    #Pegar valor de N Doc
    #pg.hotkey("down") 
    #pg.hotkey("ctrl", "c")
    texto1 = pyperclip.paste()
    compDoc = pd.DataFrame({"Comp Doc":[154.0]})
    new_tabela = pd.concat([tabela, compDoc], axis=1)

    # Adicionando nova coluna para comparacao de Data

    #time.sleep(2)
    #pg.click(x=872, y=289) #Data - cabeçalho
    #pg.hotkey("down")
    #pg.hotkey("ctrl", "c") #Pegar valor de Data
    texto2 = pyperclip.paste()
    compData = pd.DataFrame({"Comp Data":[ 31.0 ]})
    new_tabela = pd.concat([new_tabela, compData], axis=1)

    print(new_tabela)

    for coluna in coluna.index:
    # Comparando Ndoc e ref
        numDoc=float(new_tabela.loc[linha,"Nosso Número"])
        numRef=float(new_tabela.loc[linha,"Comp Doc"])
        if(numDoc == numRef): #Se Doc=ref
            # Comparando data de vencimento
            dataDoc=float(new_tabela.loc[linha,"Vencimento"])
            dataRef=float(new_tabela.loc[linha,"Comp Data"])
            if(dataRef==dataDoc):            
                print("Dar baixa") #Rodar Script do SAP 
            else:()
        else:(print("não vamo"))