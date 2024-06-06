Sub MIGOGR()
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
Dim server As String
server = Sheets("Data").Range("D1").Value
Set connection = Applic.OpenConnection(server, True)
Set session = connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "010"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Sheets("Data").Cells(1, 1).Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Sheets("Data").Cells(2, 1).Value
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nmigo_gr")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A01"
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = Sheets("Data").Cells(1, 3).Value
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/ctxtGOHEAD-BLDAT").Text = CStr(Sheets("Data").Cells(7, "A").Value)
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/ctxtGOHEAD-BUDAT").Text = CStr(Sheets("Data").Cells(7, "A").Value)
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").Text = Sheets("Data").Cells(7, "D").Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").SetFocus
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").caretPosition = 15
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/btnBUTTON_HEADER_TOGGLE").press 'Expand bang input
lrow = Range("A6").End(xlDown).Offset(1, 0).Row
For i = 7 To lrow
'Tao vong lap:
Dim q As Integer, x As Integer
frow = Sheets("Data").Cells(i, "B").Value
erow = Sheets("Data").Cells(i, "C").Value
x = 0
For q = frow To erow
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MAKTX[2," & CStr(x) & "]").Text = Sheets("Data").Cells(q, "J").Value
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/txtGOITEM-ERFMG[4," & CStr(x) & "]").Text = Sheets("Data").Cells(q, "K").Value
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-ERFME[5," & CStr(x) & "]").Text = "T"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-LGOBE[7," & CStr(x) & "]").Text = "1403"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-NAME1[6," & CStr(x) & "]").Text = "1000"
    x = x + 1
    If x > 9 Then
        session.findById("wnd[0]").sendVKey 0
        For m = 0 To 9
            If session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MAKTX[2," & CStr(m) & "]").Text <> vbNullString Then
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-CHARG[8," & CStr(m) & "]").Text = "EX"
            End If
        Next m
    x = 1
    session.findById("wnd[0]/tbar[0]/btn[82]").press 'next page
    End If
Next q
session.findById("wnd[0]").sendVKey 0
For x = 0 To 9
    If session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MAKTX[2," & CStr(x) & "]").Text <> vbNullString Then
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-CHARG[8," & CStr(x) & "]").Text = "EX"
    End If
Next x
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_HEADER:SAPLMIGO:0102/btnOK_HEADER").press 'Collapse bang input
If Range("A4").Text = "TRUE" Then
session.findById("wnd[0]/tbar[1]/btn[23]").press 'post
End If
Next i
End Sub
