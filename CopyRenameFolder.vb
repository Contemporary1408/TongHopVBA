Sub CopyMyFolder()
Set fs = CreateObject("Scripting.FileSystemObject")
'copy folder        
fs.CopyFolder "D:\Temp", "D:\Temp2"
'rename folder        
Name "D:\Temp2" As "D:\Temp3"
'tham khao:https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'co the copy ca subfolders ben trong neu co        
End Sub
'******************************
  Private Sub UserForm_Initialize()
  'Luu y chon Userform & Initialize
   txtnam.Value = ""
   txtid.Value = ""
End Sub
'***************************
Private Sub btnok_Click()
   Dim emptyRow As Long
    Sheet1.Activate
    If txtid.Value = vbNullString Or txtnam.Value = vbNullString Then
    MsgBox "Hay nhap du thong tin"
    End If
   emptyRow = Range("B4").End(xlDown).Offset(1, 0).Row
   Cells(emptyRow, 2).Value = txtid.Value
   Cells(emptyRow, 3).Value = txtnam.Value
  Dim DuongDanFolder As String
    Dim FoldName As String
    Dim Temp As String
    Dim Temp2 As String
    'Dim i As Integer
    'i = ActiveCell.Row
    'For i = i To 2000
    DuongDanFolder = ThisWorkbook.Path
    Temp = DuongDanFolder & Application.PathSeparator & "Temp"
    Temp2 = DuongDanFolder & Application.PathSeparator & "Temp2"
    FoldName = Cells(emptyRow, 4).Value
    Dim folderPathWithName As String
    folderPathWithName = DuongDanFolder & Application.PathSeparator & FoldName
     Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CopyFolder Temp, Temp2
    Name Temp2 As folderPathWithName
             'MsgBox "Thu muc da ton tai"
       'Exit For
        'End If
    'Next i
End Sub
'*********************
Private Sub btnng_Click()
   Unload Me
End Sub
'*********************
Sub ShowUserForm()
'Trigger hien thi button
UserForm1.Show

End Sub
