Dim wbSource As Workbook
Dim wsSource As Worksheet
Dim wbDest As Workbook
Dim fName As String
i = 5
strname = ThisWorkbook.Sheets("TxtCrt").Cells(i, "F").Value
strpath = "C:\Users\" & Environ("username") & "\Desktop\Helix\"
FRH = ThisWorkbook.Sheets("TxtCrt").Cells(i, "D").Value
LRH = ThisWorkbook.Sheets("TxtCrt").Cells(i, "E").Value
'References
Set wbDest = Workbooks.Add
ThisWorkbook.Sheets("HeaderUpload").Activate
ThisWorkbook.Sheets("HeaderUpload").Range(Cells(FRH, "B"), Cells(LRH, "AH")).Copy
'Save in new workbook
wbDest.Worksheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False
'Get file name and location
fName = strpath & strname
'Save new tab delimited file
wbDest.SaveAs fName, xlText
wbDest.Close SaveChanges:=True
