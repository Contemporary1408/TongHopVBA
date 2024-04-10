Sub Open_Folder()
'Range F6 nay gia tri phai la link day du, khong nam trong func Hyperlink:
Shell "C:\WINDOWS\explorer.exe """ & Range("F6").Value & "", vbNormalFocus 'or VbMaximizedFocus
End Sub
