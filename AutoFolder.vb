Sub TaoFolderVendorList()

    Dim DuongDanFolder As String
    Dim FoldName As String
    Dim i As Integer
    i = ActiveCell.Row
    For i = i To 2000
    DuongDanFolder = ThisWorkbook.Path
    FoldName = Cells(i, 4).Value
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
       Exit For
        End If
    Next i
End Sub
'Created by Do Duc Anh
