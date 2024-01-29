Sub ShowUserForm()
'Trigger hien thi button
UserForm1.Show

End Sub
'************************
Private Sub UserForm_Initialize()
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Var")
Set Rng = ws.Range("A2:A13")
For i = 1 To Rng.Rows.Count
UserForm1.txtmon.AddItem (Rng.Cells(i, 1))
Next i
  'Luu y chon Userform & Initialize
   txtmon.Value = ""
   caprt.Value = False
   exob.Value = False
   capob.Value = False
   exrt.Value = False
End Sub
'***********************
Private Sub btnok_Click()
On Error Resume Next
Sheet1.Activate
If txtmon.Value = vbNullString Then
    MsgBox "Hay nhap du thong tin"
    Exit Sub
End If
Cells(2, 2).Value = txtmon.Value
Cells(2, 4).Value = txty.Value
Call LoopThroughFiles
'populate variables
Dim strTo As String
Dim strSubject As String
Dim strBody As String
Dim strCC As String
Dim AttExpOB As String
Dim AttCapOB As String
Dim AttExpRT As String
Dim AttCapRT As String
Dim PathAttExpOB As String
Dim PathAttCapOB As String
Dim PathAttExpRT As String
Dim PathAttCapRT As String
Dim DuongDanFolder As String
Dim Dept As String
Dim Dept1 As String
DuongDanFolder = ThisWorkbook.Path
Dim i As Integer
Dim k As Integer
Dim Att As String
Dim DirAtt As String
    For i = 4 To 5
        
'populate variables
      strTo = ThisWorkbook.Worksheets("Email Information").Cells(i, 3).Value
      strCC = ThisWorkbook.Worksheets("Email Information").Cells(i, 4).Value
      strSubject = ThisWorkbook.Worksheets("Email Information").Cells(i, 5).Value
      strBody = ThisWorkbook.Worksheets("Email Information").Cells(i, 6).Value & vbNewLine & _
                ThisWorkbook.Worksheets("Email Information").Cells(i, 7).Value & vbNewLine & _
                ThisWorkbook.Worksheets("Email Information").Cells(i, 8).Value & vbNewLine & _
                ThisWorkbook.Worksheets("Email Information").Cells(i, 9).Value & vbNewLine & _
                " " & vbNewLine & _
                "******************" & vbNewLine & _
                ThisWorkbook.Worksheets("Email Information").Cells(i, 10).Value & vbNewLine & _
                ThisWorkbook.Worksheets("Email Information").Cells(i, 11).Value
      Dept = ThisWorkbook.Worksheets("Email Information").Cells(i, 1).Value
    Dim appOutlook As Object
    Dim mItem As Object
    Set appOutlook = CreateObject("Outlook.Application")
      Set mItem = appOutlook.CreateItem(0)
    With mItem
     .To = strTo
     .CC = strCC
     .Subject = strSubject
     .body = strBody
     For k = 4 To 90
        Att = Cells(k, 13).Value
        Dept1 = Cells(k, 15).Value
        DirAtt = DuongDanFolder & Application.PathSeparator & Att
        If Dept1 = Dept Then
        .Attachments.Add DirAtt
        End If
     Next k
'use send to send immediately or display to show on the screen
     .Display
     '.Send
    End With
'clean up objects
    Set mItem = Nothing
    Set appOutlook = Nothing
    Next i
End Sub
'***********************
Private Sub btnng_Click()
   Unload Me
End Sub
'********************
Sub LoopThroughFiles()

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
i = 3
Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFolder = oFSO.GetFolder(ThisWorkbook.Path)

For Each oFile In oFolder.Files

    ThisWorkbook.Sheets("Email Information").Cells(i + 1, 17) = oFile.Name

    i = i + 1

Next oFile

End Sub


