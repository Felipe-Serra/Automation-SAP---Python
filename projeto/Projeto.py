import pyautogui as pg
import pandas as pd
import time
import pyperclip
pg.PAUSE = 0.5
 
def pegarDoc():
    pg.keyDown("shift") #N Doc - cabeçalho
    pg.press("tab", presses=8)
    pg.keyUp("shift")
 
tabela = pd.read_excel("Projeto de automacao SAP.xlsx", sheet_name="Organizador de planilha - Sant")
tabela = tabela.drop(columns=["Nosso Número", "Data da Liquidação", "Pagador","Conta Cobrança", "Tipo Cobrança / Modalidade", "Responsável"])
print(tabela)
 
 
 
pg.hotkey("alt", "tab")
pg.write("FBL5N") # FBL5N - Partidas individuais de clientes / Rel CLIENTES
pg.press("enter")
time.sleep(0.5)
 
#for linha in tabela.index:
    # Adicionando nova coluna para comparacao de N Doc
 
pg.write(str(tabela.loc[0, "Cod Sap"]))
pg.press("f8") #Rodar
time.sleep(5)
 
pegarDoc()
 
pg.hotkey("ctrl", "space") #Pegar coluna de N Doc  
pg.hotkey("ctrl","c")
 
coluna1=0
coluna = pyperclip.paste()
coluna1 = coluna.count('\n') - 3
#coluna = coluna.str.counts('r')
print("\n^^^^^^^\n",coluna,"\n**",coluna1,"**""\n^^^^^^^^^\n")
   
#Pegar valor de N Doc
time.sleep(5)
pg.hotkey("ctrl","y")
pg.press("down", presses=1)
pg.hotkey("ctrl", "c")
 
texto1 = float(pyperclip.paste())
print(pyperclip.paste())
compDoc = pd.DataFrame({"Comp Doc":texto1})
new_tabela = pd.concat([tabela, compDoc], axis=1)
print(new_tabela)
#     # Adicionando nova coluna para comparacao de Data
 
#     #time.sleep(2)
#     #pg.click(x=872, y=289) #Data - cabeçalho
#     #pg.hotkey("down")
#     #pg.hotkey("ctrl", "c") #Pegar valor de Data
#     texto2 = pyperclip.paste()
#     compData = pd.DataFrame({"Comp Data":[r"04/05/2024"]})
#     new_tabela = pd.concat([new_tabela, compData], axis=1)
 
#     print(new_tabela)
#     i=0
#     while i <= coluna1:
#         # Comparando Ndoc e ref
#         numDoc=float(new_tabela.loc[linha,"Nosso Número"])
#         numRef=float(new_tabela.loc[linha,"Comp Doc"])
#         if(numDoc == numRef): #Se Doc=ref
#             # Comparando data de vencimento
#             dataDoc=new_tabela.loc[linha,"Vencimento"]
#             dataRef=new_tabela.loc[linha,"Comp Data"]
#             if(dataRef==dataDoc):            
#                 permissão=1 #Permissão Para rodar Script do SAP
#                 j=i
#             else:(
#                 #colocar bloq pag
#                 print("Bloqueio de pagamento em: ",i)
#             )
#         i+=1
   
#     #Rodar Script do SAP            
#     if(permissão==1):
#         print("Rodar script em: ",j)
#         permissão=0