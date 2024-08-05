Sub VL02N()
On Error Resume Next
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
Set connection = Applic.OpenConnection(Range("A3").Value, True)
Set session = connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "010"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Range("A1").Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Range("A2").Value
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/NVL02N")
lrow = Range("C1").End(xlDown).Row
For i = 2 To lrow:
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Cells(i, "C").Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select

x = 0
For q = 1 To 100
    'Do While session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-VRKME[5," + CStr(x) + "]").Text = "T"
    Qty = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-G_LFIMG[4," + CStr(x) + "]").Text
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," + CStr(x) + "]").Text = Qty
    If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-G_LFIMG[4," + CStr(x) + "]").Text <> "" Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-LGORT[3," + CStr(x) + "]").Text = "1403"
    End If
    'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-G_LFIMG[4," + CStr(x) + "]").Text = Value1
    x = x + 1
    'session.findById("wnd[0]").sendVKey 0
    If x = 8 Then
        x = 0
        session.findById("wnd[0]/tbar[0]/btn[82]").press 'next page
        Do While session.findById("wnd[0]/sbar").Text <> ""
            session.findById("wnd[0]").sendVKey 0
        Loop
    End If
Next q
session.findById("wnd[0]/tbar[1]/btn[20]").press
session.findById("wnd[1]").sendVKey 0

Next i

End Sub
