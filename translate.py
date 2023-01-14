import sys, win32com.client

def Main():

  try:

	#Obtendo conexão ativa

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      SapGuiAuto = None
      return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      application = None
      SapGuiAuto = None
      return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      return

	#Inicializando tela e transação (conexão já aberta)

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nse63"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/mbar/menu[0]/menu[0]/menu[0]").select()
    session.findById("wnd[1]/usr/lbl[2,1]").setFocus()
    session.findById("wnd[1]/usr/lbl[2,1]").caretPosition = 0
    session.findById("wnd[1]").sendVKey(2)
    session.findById("wnd[1]").sendVKey(82)
    session.findById("wnd[1]/usr/lbl[6,3]").setFocus()
    session.findById("wnd[1]/usr/lbl[6,3]").caretPosition = 2
    session.findById("wnd[1]").sendVKey(2)
    
    #Ler informações do txt (lista de transações)

    txtFile = open("trans.txt", "r", encoding = "utf-16")
    data = txtFile.read()  
    dataList = data.split("\n")
    txtFile.close()

	#Processar cada linha para tradução

    for info in dataList:
      print(info)

      session.findById("wnd[0]/usr/ctxtDYNP_2000-OBJECT1").text = info
      session.findById("wnd[0]/usr/ctxtDYNP_2000-INPT_SLANG").text = "ptBR"
      session.findById("wnd[0]/usr/ctxtDYNP_2000-INPT_SLANG").setFocus
      session.findById("wnd[0]/usr/ctxtDYNP_2000-INPT_SLANG").caretPosition = 4
      session.findById("wnd[0]").sendVKey(0)
      session.findById("wnd[0]/usr/lbl[52,2]").setFocus()
      session.findById("wnd[0]/usr/lbl[52,2]").caretPosition = 0
      session.findById("wnd[0]").sendVKey(2)
      session.findById("wnd[0]/tbar[0]/btn[11]").press()
      session.findById("wnd[0]/tbar[0]/btn[3]").press()

	#Finalização 

    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press

#   #session.findById("wnd[0]").resizeWorkingPane 173, 36, 0
#    session.findById("wnd[0]").resizeWorkingPane(173, 36, 0)
#   #session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16"
#    session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16"
#   #session.findById("wnd[0]").sendVKey 0
#    session.findById("wnd[0]").sendVKey(0)
#   #session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "TADIR"
#    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "TADIR"
#   #session.findById("wnd[0]").sendVKey 0
#    session.findById("wnd[0]").sendVKey(0)
#   #session.findById("wnd[0]/tbar[1]/btn[8]").press
#    session.findById("wnd[0]/tbar[1]/btn[8]").press()

  except:
    print(sys.exc_info()[0])

  finally:
    session = None
    connection = None
    application = None
    SapGuiAuto = None

if __name__ == "__main__":
  Main()

