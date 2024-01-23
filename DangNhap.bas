Sub Login_confirm()

Dim username As String
Dim password As String
  username = Sheets("Sheet1").Range("B3").Value 'Gán biến UserName tại vị trí người dùng nhập UserName là ô B3
  password = Sheets("Sheet1").Range("B6").Value 'Gán biến Password tại vị trí người dùng nhập Password là ô B6
Dim i As Integer
  i = 3         'Biến i đại diện cho dòng bắt đầu của danh sách tài khoản
  
If username = "" Or password = "" Then  'Xét trường hợp thiếu dữ kiện (không nhập 1 trong 2 ô)
    MsgBox "Vui long nhap du thong tin UserName va Password"
Else
    Do While Sheets("Sheet1").Range("E" & i).Value <> "" And i <  5 'Vòng lặp Do với điều kiện tại cột E phải có dữ liệu và chỉ xét tới dòng 4
      If username = Sheets("Sheet1").Range("E" & i).Value And _
        password = Sheets("Sheet1").Range("F" & i).Value Then    'Nếu thông tin đăng nhập UserName và Password giống với vị trí dòng i đang xét ở Danh sách tài khoản
            MsgBox "Dang nhap chinh xac"
            Exit Sub  'Kết thúc ngay Sub mà không xét các lệnh tiếp theo
        Else
          i = i + 1   'Trường hợp không chính xác thì xét tiếp dòng tiếp theo nhưng phải thỏa mãn điều kiện ban đầu của vòng lặp Do
        End If  'Kết thúc IF
    Loop        'Kết thúc vòng lặp Do
    MsgBox "Thong tin dang nhap khong chinh xac"  'Các trường hợp không thỏa mãn thì đều báo thông tin đăng nhập không chính xác
End If

End Sub
