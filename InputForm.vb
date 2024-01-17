Sub OpenDataEntryForm()
  Dim nName As Name

  Worksheets("Sheet1").Activate

  Range("C5").CurrentRegion.Name = "database"
  ActiveSheet.ShowDataForm
'Luu y phai define name cho bang o vi tri "database"

  For Each nName In ActiveWorkbook.Names
    If "database" = nName.Name Then nName.Delete
  Next nName
End Sub
