Sub ExportSheetValuesToWorkbookAndPDF()
    Dim ws As Worksheet
    Dim wbOrig As Workbook
    Dim wbNew As Workbook
    Dim shNew As Worksheet
    Dim shtName As String
    Dim saveFolder As String
    Dim wbFilename As String
    Dim pdfFilename As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbOrig = ThisWorkbook               ' change to ActiveWorkbook if preferred
    Set ws = ActiveSheet
    shtName = ws.Name
    
    ' Save folder: same folder as original workbook; change if needed
    If wbOrig.Path <> "" Then
        saveFolder = wbOrig.Path & Application.PathSeparator
    Else
        saveFolder = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator
    End If
    
    ' filenames (safe for file system)
    wbFilename = saveFolder & CleanFileName(shtName) & " - Values.xlsx"
    pdfFilename = saveFolder & CleanFileName(shtName) & " - Values.pdf"
    
    ' create new workbook with a single sheet
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set shNew = wbNew.Worksheets(1)
    shNew.Name = shtName
    
    ' copy used range values and column widths, and basic page setup
    ws.UsedRange.Copy
    With shNew.Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats   ' optional: remove if you want values only
    End With
    Application.CutCopyMode = False
    
    ' copy column widths
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count
        shNew.Columns(c).ColumnWidth = ws.Columns(ws.UsedRange.Columns(c).Column).ColumnWidth
    Next c
    
    ' copy page setup (orientation, margins, paper size, print area)
    With shNew.PageSetup
        .Orientation = ws.PageSetup.Orientation
        .PaperSize = ws.PageSetup.PaperSize
        .Zoom = ws.PageSetup.Zoom
        .FitToPagesWide = ws.PageSetup.FitToPagesWide
        .FitToPagesTall = ws.PageSetup.FitToPagesTall
        .LeftMargin = ws.PageSetup.LeftMargin
        .RightMargin = ws.PageSetup.RightMargin
        .TopMargin = ws.PageSetup.TopMargin
        .BottomMargin = ws.PageSetup.BottomMargin
        .CenterHorizontally = ws.PageSetup.CenterHorizontally
        .CenterVertically = ws.PageSetup.CenterVertically
        .PrintArea = ws.PageSetup.PrintArea
    End With
    
    ' Save new workbook (values only)
    wbNew.SaveAs Filename:=wbFilename, FileFormat:=xlOpenXMLWorkbook ' .xlsx
    ' Export to PDF
    shNew.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilename, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Close new workbook
    wbNew.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Saved workbook: " & vbCrLf & wbFilename & vbCrLf & vbCrLf & "Saved PDF: " & vbCrLf & pdfFilename, vbInformation, "Export Complete"
End Sub

' Helper to make file-safe name
Private Function CleanFileName(s As String) As String
    Dim invalidChars As Variant, ch As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In invalidChars
        s = Replace(s, ch, "_")
    Next
    CleanFileName = Trim(s)
    If CleanFileName = "" Then CleanFileName = "Sheet"
End Function
