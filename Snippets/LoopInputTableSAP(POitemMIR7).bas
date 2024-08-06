'Insert snippet code above
'*************************
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/btnRM08M-XMSEL").press
q = Sheets("Execute").Cells(i, "B").Value
u = 0
    Do Until q > Sheets("Execute").Cells(i, "C").Value
    x = 0
        Do Until x > 7 'Vì tọa độ dòng trong SAP chạy từ line 0 đến 7
        session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0," & CStr(x) & "]").Text = Sheets("LIV").Cells(q, "U").Value
        session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1," & CStr(x) & "]").Text = Sheets("LIV").Cells(q, "AA").Value
        x = x + 1
        q = q + 1
        Loop
    u = u + 8 'Gán giá trị cho biến u vì cứ sang dòng sau thì giá trị này cộng thêm 8
    session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = u
    Loop
session.findById("wnd[1]/tbar[0]/btn[8]").press
'*************************
'Insert snippet code below

'Shorten version:
Do Until i > lrow
    Set tbl = session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE")
    Do Until x > 7
        tbl.GetCell(x, 1).Text = Sheets(1).Cells(i, "H").Value
        x = x + 1
        i = i + 1
    Loop
    u = u + 8
    x = 1
    tbl.verticalScrollbar.Position = u
Loop
