Private Sub UserForm_Initialize()
  'Luu y chon Userform & Initialize
'Empty all other text box fields
   txtname.Value = ""
   txtdept.Value = ""
 'Tham khao: https://www.tutorialspoint.com/vba/vba_userforms.htm  
End Sub
'***************************
Private Sub btnsubmit_Click()
   Dim emptyRow As Long
  
   'Make Sheet1 active
   Sheet1.Activate
  
   'Determine emptyRow
   emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
  
   'Transfer information
   Cells(emptyRow, 1).Value = txtname.Value
   Cells(emptyRow, 2).Value = txtdept.Value
  
End Sub
'*********************
Private Sub btncancel_Click()
   Unload Me
End Sub
'*********************
Sub ShowUserForm()
'Trigger hien thi button
UserForm1.Show

End Sub
