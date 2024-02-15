Sub getfiles()

    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object, sf
    Dim i As Integer, colFolders As New Collection, ws As Worksheet
    i = 1
    Set ws = ActiveSheet
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder("\\10.118.29.7\BTMV-Data\4-ACCOUNTING\4.Budget\Automate\Final")
    
    colFolders.Add oFolder          'start with this folder
    
    Do While colFolders.Count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
    
        For Each oFile In oFolder.Files
            If oFile.DateLastModified > Now - 7 Then
                ws.Cells(i + 1, 1) = oFolder.Path
                ws.Cells(i + 1, 2) = oFile.Name
                ws.Cells(i + 1, 3) = oFile.DateLastModified
            
                i = i + 1
            End If
        Next oFile

        'add any subfolders to the collection for processing
        For Each sf In oFolder.subfolders
            colFolders.Add sf
        Next sf
    Loop

End Sub
