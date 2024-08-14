#If VBA7 Then 
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr 
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long 
#Else 
Private Declare Function GetActiveWindow Lib "user32" () As Long 
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long 
#End If 
Function MsgBoxUni(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant = vbNullString) As VbMsgBoxResult 
   'BStrMsg, BStrTitle : La chuoi Unicode
   Dim BStrMsg, BStrTitle 
   'Hàm StrConv Chuyen chuoi ve ma Unicode
   BStrMsg = StrConv(PromptUni, vbUnicode) 
   BStrTitle = StrConv(TitleUni, vbUnicode) 
   MsgBoxUni = MessageBoxW(GetActiveWindow, BStrMsg, BStrTitle, Buttons) 
End Function 
Sub TestMsgBoxUni() 
   'Test trong Excel
   'O B3 chua noi dung thong bao
   'O B4 chua tieu de cua so
   MsgBoxUni Range("B3").Value, vbInformation, Range("B4").Value 
End Sub 
