Sub SAP_OpenSessionFromLogon_AltAcc()
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
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Cells(1, 1).Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Cells(2, 1).Value
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nvf01")
'MsgBox "Waiting..."
'Set session = Nothing
'connection.CloseSession ("ses[0]")
'Set connection = Nothing
'Set sap = Nothing
Dim i As Integer, k As Integer, z As Integer
Dim FR As Integer, LR As Integer, LRSAP As Integer
For i = 5 To Cells(3, 7).Value
FR = Cells(i, 8).Value
LR = Cells(i, 9).Value
LRSAP = Cells(i, 10).Value
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtRV60A-FKDAT").Text = Cells(i, 12).Value
session.findById("wnd[0]/usr/ctxtRV60A-PRSDT").Text = Cells(i, 12).Value
k = 0
For z = FR To LR
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0," & CStr(k) & "]").Text = Cells(z, 2).Value
    k = k + 1
Next z
session.findById("wnd[0]/usr/ctxtRV60A-PRSDT").caretPosition = 0
session.findById("wnd[0]").sendVKey 11
Cells(i, 11).Value = session.findById("wnd[0]/sbar").Text 'ghi lai so park
Next i

End Sub

