Sub Auto_MIR7()
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
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Sheets("Execute").Cells(1, "B").Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Sheets("Execute").Cells(2, "B").Value
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nmir7")
session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").Text = "1000"
session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").caretPosition = 4
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cmbRM08M-VORGANG").Key = "1" 'Chon loai chung tu: 1=Invoice
'Tab Basic data
i = 8 'SUA LAI NEU MUON CHAY VONG LAP CHO NHIEU HOA DON
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").Text = Sheets("Execute").Cells(i, "D").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BUDAT").Text = Sheets("Execute").Cells(i, "F").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BUDAT").SetFocus
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BUDAT").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR").Text = Sheets("Execute").Cells(i, "E").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").Text = Sheets("Execute").Cells(i, "G").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-WAERS").Text = Sheets("Execute").Cells(i, "H").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-SGTXT").Text = Sheets("Execute").Cells(i, "I").Value
'Tab Detail
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI").Select
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-KURSF").Text = Sheets("Execute").Cells(i, "K").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/cmbINVFO-BLART").Key = Sheets("Execute").Cells(i, "J").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-GSBER").Text = Sheets("Execute").Cells(i, "L").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR").Text = Sheets("Execute").Cells(i, "M").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-BKTXT").Text = Sheets("Execute").Cells(i, "N").Value
'Nhap PO no & PO item:
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/btnRM08M-XMSEL").press
q = Sheets("Execute").Cells(i, "B").Value
u = 0
Do Until q > Sheets("Execute").Cells(i, "C").Value
x = 0
Do Until x > 7
session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0," & CStr(x) & "]").Text = Sheets("LIV").Cells(q, "U").Value
session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1," & CStr(x) & "]").Text = Sheets("LIV").Cells(q, "AA").Value
x = x + 1
q = q + 1
Loop
u = u + 8
session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = u
Loop
session.findById("wnd[1]/tbar[0]/btn[8]").press
'Nhap amt & qty cho PO data:
session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
q = Sheets("Execute").Cells(i, "B").Value
u = 0
Do Until q > Sheets("Execute").Cells(i, "C").Value
On Error Resume Next
y = 0
Do Until y > 8
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WRBTR[1," & CStr(y) & "]").Text = Sheets("LIV").Cells(q, "Z").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-MENGE[2," & CStr(y) & "]").Text = Sheets("LIV").Cells(q, "X").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-SGTXT[94," & CStr(y) & "]").Text = Sheets("LIV").Cells(q, "G").Value
y = y + 1
q = q + 1
Loop
u = u + 9
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M").verticalScrollbar.Position = u
Loop
'Exit PO data chuyen sang tab Payment & Save:
session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY").Select
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT").Text = Sheets("Execute").Cells(i, "O").Value
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZLSCH").Text = Sheets("Execute").Cells(i, "P").Value
'Kiem tra balance truoc khi Park
If session.findById("wnd[0]/usr/txtRM08M-DIFFERENZ").Text = 0 Then
MsgBox "Chung tu OK"
session.findById("wnd[0]/tbar[0]/btn[11]").press 'Nhan SAVE de park
Sheets("Execute").Cells(i, "Q").Value = session.findById("wnd[0]/sbar").Text 'Get so park tu thong bao
Else
'MsgBox "Kiem tra lai"
session.SendCommand ("/nmir7")
End If
End Sub
