Sub Auto_post_MIR7()
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
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Cells(1, "B").Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Cells(2, "B").Value
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nmir7")
session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").Text = "1000"
session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").caretPosition = 4
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]").maximize
'Goi so doc MIR7
Dim lrow As Long, i As Integer
lrow = Range("A4").End(xlDown).Offset(1, 0).Row
For i = 5 To lrow
session.findById("wnd[0]/tbar[1]/btn[34]").press
session.findById("wnd[1]/usr/txtRBKPV-BELNR").Text = Cells(i, "A").Value
session.findById("wnd[1]/usr/txtRBKPV-GJAHR").Text = Cells(1, "C").Value
session.findById("wnd[1]/usr/txtRBKPV-GJAHR").SetFocus
session.findById("wnd[1]/usr/txtRBKPV-GJAHR").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[23]").press 'nhan Post MIR7
Cells(i, "B").Value = session.findById("wnd[0]/sbar").Text
Next i
End Sub
