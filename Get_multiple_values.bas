Sub getnum()
Application.ScreenUpdating = False
Dim list1 As String
Dim lrow1 As Integer, lrow2 As Integer
Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer
lrow1 = Range("D2").End(xlDown).Row
lrow2 = Range("J2").End(xlDown).Row
For k = 3 To lrow2
    list1 = ""
    For e = 3 To lrow1
        If Cells(k, "J").Value = Cells(e, "D").Value Then
            list1 = list1 & Cells(e, "E").Value & ";"
        End If
    Next e
    Cells(k, "K").Value = list1
Next k
SecondsElapsed = Round(Timer - StartTime, 2)
Application.ScreenUpdating = True
'MsgBox ("Done!")
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub
