Sub SAP()
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
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Sheets(1).Cells(1, 1).Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Sheets(1).Cells(1, 2).Value
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nmm01")
' Chay code bat dau tu day:
lrow = Sheets(1).Range("C1").End(xlDown).Row
For i = 2 To lrow
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = Sheets(1).Cells(i, "C").Value
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 10
session.findById("wnd[0]").sendVKey (0)
If session.findById("wnd[0]/usr/cmbRMMG1-MTART").Key <> Sheets(1).Cells(i, "G").Value Then 'Check xem mã khi tạo đã phân loại đúng chưa (Z400, Z700,...)
    MsgBox ("Loai material chua khop voi dang ky")
    Exit For
End If
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "1000"
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
'Kiểm tra nếu mã vừa tạo input sai tên hoặc unit thì skip và báo lỗi vào 1 cell trong Excel:
If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1008/txtMAKT-MAKTX").Text = Sheets(1).Cells(i, "E").Value And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLMGD1:2801/ctxtMARA-MEINS").Text = Sheets(1).Cells(i, "H").Value Then
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").Text = Sheets(1).Cells(i, "D").Value
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").Text = "V"
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLMGD1:2801/ctxtMARA-MEINS").SetFocus
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLMGD1:2801/ctxtMARA-MEINS").caretPosition = 2
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    Sheets(1).Cells(i, "I").Value = "OK"
Else
    Sheets(1).Cells(i, "I").Value = "Kiem tra lai ten material/unit"
    Sheets(1).Cells(i, "J").Value = "Name: " & session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1008/txtMAKT-MAKTX").Text & " Unit: " & session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLMGD1:2801/ctxtMARA-MEINS").Text
End If
Next i
End Sub
