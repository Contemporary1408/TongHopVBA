'ZGF08 is Table Name, Unit is Column Header
Range("ZGF08[#Headers]").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("F-65 data").ListObjects("ZGF08").Sort.SortFields. _
        clear
    ActiveWorkbook.Worksheets("F-65 data").ListObjects("ZGF08").Sort.SortFields. _
        Add2 Key:=Range("ZGF08[Unit]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F-65 data").ListObjects("ZGF08").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
