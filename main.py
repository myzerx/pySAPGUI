import win32com.client
import subprocess
import time
import config

def getSAPTransaction(self):
        session = self.connection.Children(0)
        #acessando a transação SE16N
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
        session.findById("wnd[0]").sendVKey(0)
        #Tabelas VBAP
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "VBAP"
        session.findById("wnd[0]").sendVKey(8)


class SapGui(object):
    def __init__(self):     
        try:
            #verifica se o SAP GUI já está aberto
            sapgui = win32com.client.GetObject('SAPGUI')           
        except:
            #se não estiver aberto, abrirá via .exe
                self.path = config.SAP_EXE
                subprocess.Popen(self.path)
                time.sleep(5)    
                            
        finally:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        #faz conexão com a connection do SAP_CONNECT
        self.connection = application.OpenConnection(config.SAP_CONNECT, True)
        session = self.connection.Children(0)    
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = config.SAP_USER
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = config.SAP_PASS
        session.findById("wnd[0]").sendVKey(0)
        getSAPTransaction(self)
       
    
if __name__ == '__main__':
 Object  =  SapGui()
