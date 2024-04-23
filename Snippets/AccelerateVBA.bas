'Trong VBA, việc đình chỉ cập nhật màn hình sẽ cho tốc độ rất nhanh. Khi chúng ta xử lý dữ liệu trên excel, chúng ta thấy màn hình cứ giật giật.
'Việc thiết định hủy cập nhật màn hình sẽ tăng tốc xử lý của excel.
Application.ScreenUpdating = False
'đoạn code xử lý dữ liệu
Application.ScreenUpdating = True


'Tuy nhiên nó cũng có nhược điểm, người dùng không nhìn thấy tiến trình xử lý của VBA (vì màn hình đứng yên) thì có vẻ sẽ sốt ruột.
'Do đó chúng ta cân nhắc, nếu việc thực thi quá mất thời gian, thì vẫn nên dừng cập nhật màn hình để tăng tốc xử lý.
'Ngoài ra chúng ta còn có kìm chế sự kiện.
'Khi macro xử lý dữ liệu, thay đổi giá trị trên cells, thì worksheet sẽ phải phát hiện sự kiện đó và xử lý.
'Việc kiềm chế sự kiện tức là không cho worksheet bắt sự kiện, sẽ làm cho macro thực thi nhanh hơn.
'Tuy nhiên cũng có vấn đề là, trong khi macro thực hiện xử lý dữ liệu thì các sự kiện nếu phát sinh sẽ không được xử lý.
'Nhưng thông thường, khi xử lý dữ liệu, chúng ta không mấy quan tâm tới việc này. Đây chỉ là điều chú ý nho nhỏ cho các bạn biết mà thôi.
Application.EnableEvents = False
'đoạn code xử lý dữ liệu
Application.EnableEvents = True


'Tiếp theo chúng ta nói tới việc tự động tính toán.
'Chẳng hạn như trên bảng tính có chứa các công thức chẳng hạn.
'Nhưng chú ý rằng, dù bảng tính không có chứa công thức, thì tính năng tự động tính toán vẫn hoạt động dò tìm xem có công thức nào hay không.
'Do đó, để tăng tốc độ xử lý của macro, chúng ta cũng nên tắt nó đi.
Application.Calculation = xlCalculationManual
'đoạn code xử lý dữ liệu
Application.Calculation = xlCalculationAutomatic
