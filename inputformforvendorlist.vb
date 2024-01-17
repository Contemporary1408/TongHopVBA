Private Sub UserForm_Initialize()
  'Luu y chon Userform & Initialize
'Empty all other text box fields
   txtnam.Value = ""
   txtid.Value = ""
 'Tham khao: https://www.tutorialspoint.com/vba/vba_userforms.htm
End Sub
'***************************
Private Sub btnsubmit_Click()
   Dim emptyRow As Long
  
   'Make Sheet1 active
    Sheet1.Activate
  
   'Determine emptyRow
   emptyRow = Range("B4").End(xlDown).Offset(1, 0).Row
  
   'Transfer information
   Cells(emptyRow, 2).Value = txtid.Value
   Cells(emptyRow, 3).Value = txtnam.Value
  Dim DuongDanFolder As String
    Dim FoldName As String
    Dim i As Integer
    'i = ActiveCell.Row
    'For i = i To 2000
    DuongDanFolder = ThisWorkbook.Path
    FoldName = Cells(emptyRow, 4).Value
    Dim folderPathWithName, folder_1_path, folder_2_path, folder_3_path, folder_4_path As String
    folderPathWithName = DuongDanFolder & Application.PathSeparator & FoldName
        If Dir(folderPathWithName, vbDirectory) = vbNullString Then
            MkDir (folderPathWithName)
            folder_1_path = folderPathWithName & Application.PathSeparator & Cells(1, 12).Value
            folder_2_path = folderPathWithName & Application.PathSeparator & Cells(1, 13).Value
            folder_3_path = folderPathWithName & Application.PathSeparator & Cells(1, 14).Value
            folder_4_path = folderPathWithName & Application.PathSeparator & Cells(1, 15).Value

            MkDir (folder_1_path)
            MkDir (folder_2_path)
            MkDir (folder_3_path)
            MkDir (folder_4_path)
        Else
             'MsgBox "Thu muc da ton tai"
       'Exit For
        End If
    'Next i
End Sub
'*********************
Private Sub btncancel_Click()
   Unload Me
End Sub
'*********************

