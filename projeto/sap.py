import win32com.client
import subprocess
import time
import sys
import pyautogui as pg
 
#Abre o SAP
try:
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(5)
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection("# JSL -  ECC - Produção (ECP)", True)
    time.sleep(3)
    session = connection.Children(0)
except:
    print(sys.exc_info()[0])
    session = None
    connection = None
    application = None
    SapGuiAuto = None
 
# Faz o login e abre a FBL5N
def getClient():
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "30127215"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "V4mS@0fs"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
    session.findById("wnd[0]").sendVKey(0)
 
def baixa():
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "FB05"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = "06.01.2024"
    session.findById("wnd[0]/usr/ctxtBKPF-BLART").text = "dz"
    session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "6800"
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").setFocus
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").caretPosition = 3
    session.findById("wnd[0]").sendVKey (6)
   
    session.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = "7213"
    session.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]").select
    session.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]").setFocus
    session.findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]").press
    time.sleep(5)
    session.findById("wnd[0]").sendVKey (16)
    time.sleep(5)
    session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[1,0]").text = "154427298"
    time.sleep(5)
    session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[1,0]").setFocus
    session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[1,0]").caretPosition = 9
    session.findById("wnd[0]").sendVKey (0)
    time.sleep(100)
    session.findById("wnd[0]").sendVKey (16)
    session.findById("wnd[0]").sendVKey (7)
    session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "40"
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "1110162502"
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").setFocus
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = 10
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = "9.169,25"
    session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1005/ctxtCOBL-PRCTR").text = "H15602-P01"
    session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1005/ctxtCOBL-PRCTR").setFocus
    session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1005/ctxtCOBL-PRCTR").caretPosition = 10
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (16)
    session.findById("wnd[0]").sendVKey (7)
    # session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "50"
    # session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "5118203002"
    # session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").setFocus
    # session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = 10
    # session.findById("wnd[0]").sendVKey (0)
    # session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = "*"
    # session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-KOSTL").text = "H15602-P01"
    # session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-KOSTL").setFocus
    # session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-KOSTL").caretPosition = 10
    # session.findById("wnd[0]").sendVKey (0)
    # session.findById("wnd[0]").sendVKey (0)
    # session.findById("wnd[0]/tbar[1]/btn[16]").press
    # session.findById("wnd[0]/tbar[0]/btn[3]").press
 
def main():
    getClient()
    baixa()
       
main()