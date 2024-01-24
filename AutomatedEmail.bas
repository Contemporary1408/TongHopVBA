Sub SendMail()
'thamkhao:https://www.automateexcel.com/vba/send-emails-outlook/
   Dim strTo As String
   Dim strSubject As String
   Dim strBody As String
   Dim strCC As String
   Dim i As Integer
   
For i = 1 To 2
'populate variables
   strTo = Cells(i, 1).Value
   strCC = Cells(i, 2).Value
   strSubject = "Please find finance file attached"
   strBody = "Dear " & Cells(i, 3).Value
'call the function to send the email
   Call SendActiveWorkbook(strTo, strSubject, strCC, strBody)
      'MsgBox "Email creation Success"
   'Else
      'MsgBox "Email creation failed!"
   'End If
Next i
End Sub
'***********************
Function SendActiveWorkbook(strTo As String, strSubject As String, Optional strCC As String, Optional strBody As String) As Boolean
   On Error Resume Next
   Dim appOutlook As Object
   Dim mItem As Object
'create a new instance of Outlook
   Set appOutlook = CreateObject("Outlook.Application")
   Set mItem = appOutlook.CreateItem(0)
   With mItem
     .To = strTo
     .CC = strCC
     .Subject = strSubject
     .Body = strBody
     .Attachments.Add "C:\Users\anh.doduc\Desktop\Upload.txt"
'use send to send immediately or display to show on the screen
    .Display
    .Send
   End With
'clean up objects
  Set mItem = Nothing
  Set appOutlook = Nothing
End Function
