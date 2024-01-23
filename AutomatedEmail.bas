Private Sub CommandButton22_Click()
    Dim a As Integer
    Dim objOutlook As Object
    Dim objMail As Object
    Dim rngTo As Range
    Dim rngSubject As Range
    Dim rngBody As Range
    Dim rngAttach As Range

    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)

    a = ActiveCell.Row

    With ActiveSheet
        Set rngTo = .Cells(a, "C")
        Set rngSubject = .Cells(a, "E")
        'Set rngBody = .Range("B3")
        'Set rngAttach = .Range("B4")
    End With

    With objMail
        .To = rngTo.Value
        .Subject = rngSubject.Value
        '.Body = rngBody.Value
       '.Attachments.Add rngAttach.Value
        .Display 'Instead of .Display, you can use .Send to send the email _
                    or .Save to save a copy in the drafts folder
    End With

    Set objOutlook = Nothing
    Set objMail = Nothing
    Set rngTo = Nothing
    Set rngSubject = Nothing
    Set rngBody = Nothing
    Set rngAttach = Nothing
End Sub
