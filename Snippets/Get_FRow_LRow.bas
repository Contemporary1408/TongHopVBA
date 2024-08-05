'Source: https://stackoverflow.com/questions/29019237/find-first-and-last-row-containing-specific-text
Dim batStartRow As Long, batEndRow As Long
With Sheets("Sheet1")
    batStartRow = .Range("A:A").Find(what:="bats", after:=.Range("A1")).Row
    batEndRow = .Range("A:A").Find(what:="bats", after:=.Range("A1"), searchdirection:=xlPrevious).Row
End With
