emptyRow = Range("B4").End(xlDown).Offset(1, 0).Row 'Cell B4 is header
Cells(emptyRow, 2).Value = txtid.Value
Cells(emptyRow, 3).Value = txtnam.Value
