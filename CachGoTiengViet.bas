'Để gõ tiếng Việt trong VBA cần 2 thứ: Cụm từ tiếng Việt đã được mã hóa Unicode và Function để giải mã Unicode về tiếng Việt.
'Mã hóa tiếng Việt sang Unicode tại web: https://www.vietnamesetools.com/vi/vietnamese-to-unicode
'>> Lưu ý mã hóa Unicode của tiếng Việt để bỏ vào VBA cần bỏ các ký tự 00,e ví dụ chữ Tổng trên web trả về "\u0074\u1ed5\u006e\u0067" nhưng bỏ vào VBA phải là ";74;1ed5;6e;67"
'Function VBA để decode:
Function UnicodeChar(UniCharCode As String) As String
  'Luu y dat toan bo code nay vao Module
On Error GoTo Loi
Dim str
Dim desStr As String
Dim I
If Mid(UniCharCode, 1, 1) = ";" Then
UniCharCode = Mid(UniCharCode, 2)
End If
If Right(UniCharCode, 1) = ";" Then
UniCharCode = Mid(UniCharCode, 1, Len(UniCharCode) - 1)
End If
str = UniCharCode
str = Split(str, ";")
For I = LBound(str) To UBound(str)
desStr = desStr & ChrW$("&H" & str(I))
Next
UnicodeChar = desStr
Loi:
Exit Function
End Function
'Sau đó đặt đoạn code function chính xuống dưới:
Function tv()
Dim a As String
a = ";74;1ed5;6e;67"
Cells(1, 1).Value = UnicodeChar(a)
End Function
