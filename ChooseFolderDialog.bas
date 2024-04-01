Sub folder_names_in_a_directory_excluding_subfolder()
Application.ScreenUpdating = False
Dim fldpath
Dim fso As Object, j As Long, folder, SubFolders, SubFolder
    With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Choose the folder"
    .Show
    End With
    On Error Resume Next
    fldpath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"
    If fldpath = False Then
        MsgBox "Folder Not Selected"
    Exit Sub
    End If
    Workbooks.Add
    Cells(1, 1).Value = fldpath
    Cells(2, 1).Value = "Path"
    Cells(2, 2).Value = "Dir"
    Cells(2, 3).Value = "Name"
    Cells(2, 4).Value = "Date Created"
    Cells(2, 5).Value = "Date Last Modified"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getfolder(fldpath)
    Set SubFolders = folder.SubFolders
    For Each SubFolder In SubFolders
        j = Range("A1").End(xlDown).Row + 1
        Cells(j, 1).Value = SubFolder.Path
        Cells(j, 2).Value = Left(SubFolder.Path, InStrRev(SubFolder.Path, "\"))
        Cells(j, 3).Value = SubFolder.Name
        Cells(j, 4).Value = SubFolder.DateCreated
        Cells(j, 5).Value = SubFolder.DateLastModified
    Next SubFolder
    Set fso = Nothing
    Range("a1").Font.Size = 9
    ActiveWindow.DisplayGridlines = False
    Range("a3:e" & Range("a2").End(xlDown).Row).Font.Size = 9
    Range("a2:e2").Interior.Color = vbCyan
    Columns("c:h").AutoFit
Application.ScreenUpdating = True
End Sub
