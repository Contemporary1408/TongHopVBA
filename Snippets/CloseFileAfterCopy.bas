Dim wb As Workbook
Dim path As String
path = "C:\Users\" & Environ("username") & "\Desktop\"
Set wb = Workbooks.Open(path & "FB03.XLS")
lrow = wb.Sheets("FB03").Range("B4").End(xlDown).Row
With wb.Sheets("FB03")
    .Activate
    .Range(Cells(4, "B"), Cells(lrow, "O")).Copy
End With
Workbooks("!KPI 2024-DucAnh.xlsm").Activate
Workbooks("!KPI 2024-DucAnh.xlsm").Sheets("FB03").Range(Cells(5, "G"), Cells(lrow + 1, "T")).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Workbooks("FB03.xls").Close
