Sub CopyMyFolder()
Set fs = CreateObject("Scripting.FileSystemObject")
'copy folder        
fs.CopyFolder "D:\Temp", "D:\Temp2"
'rename folder        
Name "D:\Temp2" As "D:\Temp3"
'tham khao:https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'co the copy ca subfolders ben trong neu co        
End Sub
