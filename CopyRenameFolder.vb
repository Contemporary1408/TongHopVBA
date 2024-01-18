Option Explicit

Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'https://stackoverflow.com/questions/18966991/how-can-i-use-the-filesystemobject-to-copy-and-rename 
Const FO_COPY = &H2 '~~> Copy File/Folder
Const FOF_SILENT = &H4 '~~> Silent Copy

Private Type SHFILEOPSTRUCT
    hwnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Sub Sample()
    Dim lresult  As Long, lFlags   As Long
    Dim SHFileOp As SHFILEOPSTRUCT

    With SHFileOp
        '~~> For Copy
        .wFunc = FO_COPY
        .pFrom = "C:\Temp"
        .pTo = "C:\Temp2\"
        '~~> For Silent Copy
        '.fFlags = FOF_SILENT
    End With
    lresult = SHFileOperation(SHFileOp)

    '~~> SHFileOp.fAborted will be true if user presses cancel during operation
    If lresult <> 0 Or SHFileOp.fAborted Then Exit Sub

    MsgBox "Operation Complete", vbInformation, "File Operations"
End Sub
For renaming a folder, here is a one liner

Sub Sample()
    Name "C:\Temp2" As "C:\Temp3"
End Sub
