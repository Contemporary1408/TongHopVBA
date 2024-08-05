Set tbl = session.findById("wnd[0]/usr/tblDEMO_DYNPRO_TABCONT_LOOPFLIGHTS")
' Make the first row visible (show the top of the list) -> that calls the back-end system and screen is reloaded.
' ATTENTION: when the back-end is called, to continue working with screen elements, they must be re-instantiated.
tbl.VerticalScrollbar.Position = 0

TextsOfAllCellsInColumnZero = ""
Do While True
  ' Re-instantiate the Table Control element (mandatory each time the back-end is called)
  Set tbl = session.findById("wnd[0]/usr/tblDEMO_DYNPRO_TABCONT_LOOPFLIGHTS")
  visibleRow = 0
  currentScrollbarPosition = tbl.VerticalScrollbar.Position
  While visibleRow < tbl.VisibleRowCount And currentScrollbarPosition <= tbl.VerticalScrollbar.Maximum
    TextsOfAllCellsInColumnZero = TextsOfAllCellsInColumnZero & tbl.GetCell(visibleRow,0).Text & Chr(10)
    visibleRow = visibleRow + 1
    currentScrollbarPosition = currentScrollbarPosition + 1
  Wend
  If currentScrollbarPosition > tbl.VerticalScrollbar.Maximum Then
    Exit Do
  End If
  tbl.VerticalScrollbar.Position = currentScrollbarPosition
Loop
MsgBox TextsOfAllCellsInColumnZero
