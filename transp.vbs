If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse63"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[0]/menu[0]/menu[0]").select
session.findById("wnd[1]/usr/lbl[2,1]").setFocus
session.findById("wnd[1]/usr/lbl[2,1]").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[1]").sendVKey 82
session.findById("wnd[1]/usr/lbl[6,3]").setFocus
session.findById("wnd[1]/usr/lbl[6,3]").caretPosition = 2
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/ctxtDYNP_2000-OBJECT1").text = "/LKMCGER/GER_005_01"
session.findById("wnd[0]/usr/ctxtDYNP_2000-INPT_SLANG").text = "ptBR"
session.findById("wnd[0]/usr/ctxtDYNP_2000-INPT_SLANG").setFocus
session.findById("wnd[0]/usr/ctxtDYNP_2000-INPT_SLANG").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/lbl[52,2]").setFocus
session.findById("wnd[0]/usr/lbl[52,2]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
