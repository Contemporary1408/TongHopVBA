Sub SAP_OpenSessionFromLogon()

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

'session.SendCommand ("/nBIBS")

'MsgBox "Waiting..."

'Set session = Nothing

'connection.CloseSession ("ses[0]")

'Set connection = Nothing

'Set sap = Nothing




End Sub
