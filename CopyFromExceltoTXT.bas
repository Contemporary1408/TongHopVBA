Sub FromXLStoTXT()
'Ref: https://stackoverflow.com/questions/54662875/copying-the-range-and-pasting-it-directly-to-a-txt-file
Dim wbSource As Workbook, wsSource As Worksheet, wbDest As Workbook
Dim fName As String, strpath As String, strname As String
Dim lrow As Integer, FRH As Integer, LRH As Integer, FRI As Integer, LRI As Integer
lrow = ThisWorkbook.Sheets("TxtCrt").Cells(4, "C").End(xlDown).Row
strpath = "C:\Users\" & Environ("username") & "\Desktop\Helix\"
For i = 5 To lrow
Application.ScreenUpdating = False

'*******Make Header txt file*******
strname = ThisWorkbook.Sheets("TxtCrt").Cells(i, "F").Value
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
' The string to be replaced
Dim strFilePath As String
Dim strOldContent As String
Dim strNewContent As String
Dim intFileNumber As Integer
strFilePath = strpath & strname
Dim strSearchString As String
strSearchString = vbTab
' The string to replace with
Dim strReplaceString As String
strReplaceString = ""
' Assign a unique number for opening the file
intFileNumber = FreeFile
' Open the text file for reading
Open strFilePath For Input As intFileNumber
' Read the entire content of the file into a string
strOldContent = Input$(LOF(intFileNumber), intFileNumber)
Close intFileNumber
' Replace the string
strNewContent = Replace(strOldContent, strSearchString, strReplaceString)
strNewContent = Replace(strNewContent, Chr(34), vbNullString) 'This returns the double-quote character (").
strNewContent = Replace(strNewContent, Chr(10), vbNullString) 'This returns the line feed character (LF), which is used to move the cursor to the next line.
strNewContent = Replace(strNewContent, Chr(13), vbNullString) 'This returns the carriage return character (CR), which moves the cursor to the beginning of the line.
' Open the text file for writing
Open strFilePath For Output As intFileNumber
' Write the new content to the file
Print #intFileNumber, strNewContent
Close intFileNumber


'*******Make Item txt file*******
strname1 = ThisWorkbook.Sheets("TxtCrt").Cells(i, "I").Value
FRI = ThisWorkbook.Sheets("TxtCrt").Cells(i, "G").Value
LRI = ThisWorkbook.Sheets("TxtCrt").Cells(i, "H").Value
'References
Set wbDest = Workbooks.Add
ThisWorkbook.Sheets("ItemUpload").Activate
ThisWorkbook.Sheets("ItemUpload").Range(Cells(FRI, "B"), Cells(LRI, "Q")).Copy
'Save in new workbook
wbDest.Worksheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False
'Get file name and location
fName = strpath & strname1
'Save new tab delimited file
wbDest.SaveAs fName, xlText
wbDest.Close SaveChanges:=True
' The string to be replaced
strFilePath = strpath & strname1
strSearchString = vbTab
' The string to replace with
strReplaceString = ""
' Assign a unique number for opening the file
intFileNumber = FreeFile
' Open the text file for reading
Open strFilePath For Input As intFileNumber
' Read the entire content of the file into a string
strOldContent = Input$(LOF(intFileNumber), intFileNumber)
Close intFileNumber
' Replace the string
strNewContent = Replace(strOldContent, strSearchString, strReplaceString)
' Open the text file for writing
Open strFilePath For Output As intFileNumber
' Write the new content to the file
Print #intFileNumber, strNewContent
Close intFileNumber
Next i
Application.ScreenUpdating = True
MsgBox "DONE, 520!"
End Sub
