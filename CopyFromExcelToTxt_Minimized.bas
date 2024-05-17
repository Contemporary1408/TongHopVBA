Sub ZGF08()
erow = ThisWorkbook.Sheets("F-65 data").Range("N12").End(xlDown).Row
'MsgBox ("Luu y tao 1 file Upload.txt ngoai Desktop")
Dim ws As Worksheet
Dim txt As String
Dim lastRow As Long, i As Long
Dim fileNumber As Integer
Set ws = ThisWorkbook.Sheets("F-65 data")
txt = "C:\Users\" & Environ("username") & "\Desktop\Upload.txt"
fileNumber = FreeFile
Open txt For Output As #fileNumber
For i = 13 To erow ' For rows 13 to erow
cellData = ""
For j = 1 To 66 ' For columns A to BN
cellData = cellData & ws.Cells(i, j).Value & vbTab
Next j
' Write the concatenated string of cell values to the text file
Print #fileNumber, cellData
Next i
Close #fileNumber
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
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Range("G1").Text
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Range("G2").Text
session.findById("wnd[0]").sendVKey 0
session.SendCommand ("/nzgf08")
session.findById("wnd[0]/usr/ctxtP_BUKRS1").Text = "1000"
session.findById("wnd[0]/usr/radP_UPL").SetFocus
session.findById("wnd[0]/usr/radP_UPL").Select
session.findById("wnd[0]/usr/ctxtP_UPFILE").Text = txt
session.findById("wnd[0]/usr/ctxtP_UPFILE").SetFocus
session.findById("wnd[0]/usr/ctxtP_UPFILE").caretPosition = 37
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 8
End Sub
