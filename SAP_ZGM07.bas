Sub SAP_OpenSessionFromLogon_AltAcc()
Range("A3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Dim SapGui
Dim Applic
Dim connection
Dim session
Dim WSHShell
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
Set WSHShell = CreateObject("WScript.Shell")
Do Until WSHShell.AppActivate("SAP Logon ")
Application.Wait Now + TimeValue("0:00:01")
Loop
Set WSHShell = Nothing
Set SapGui = GetObject("SAPGUI")
Set Applic = SapGui.GetScriptingEngine
Set connection = Applic.OpenConnection("BTMV Production Server", True)
Set session = connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "010"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "VHACS12"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "ducanh@16119"
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nZGM07")
'MsgBox "Waiting..."
'Set session = Nothing
'connection.CloseSession ("ses[0]")
'Set connection = Nothing
'Set sap = Nothing
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtP_BUKRS").Text = "1000"
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "1000"
session.findById("wnd[0]/usr/radP_POBASE").SetFocus
session.findById("wnd[0]/usr/radP_POBASE").Select
session.findById("wnd[0]/usr/ctxtS_PO_DAT-LOW").Text = "010116"
session.findById("wnd[0]/usr/ctxtS_PO_DAT-HIGH").Text = "290324"
session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").Text = ""
session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").SetFocus
session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").caretPosition = 0
session.findById("wnd[0]/usr/btn%_S_PO_NUM_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press 'press "copy from clipboard" on input range
session.findById("wnd[1]/tbar[0]/btn[8]").press 'press F8 on input range
session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").Text = ""
session.findById("wnd[0]/usr/txtS_PO_CRE-LOW").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press

End Sub
