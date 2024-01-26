VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWMDatePicker 
   Caption         =   "ShortDate"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2805
   OleObjectBlob   =   "frmWMDatePicker.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "frmWMDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' frmWMDatePicker ©2021 by wolff68@yahoo.com Version 1.1
' Lightweight date picker control within just one form to add to your project. No extra class module or addin needed!
' Simple code that can be understood even by novices.

' GetDate(Optional StartDate As Date = 0,               Preselected date. If not given = actual date
'         Optional MinDate as Date = -657403,           Dates not allowed before MinDate. If not given = -657403 (Feb  1st 0100)
'         Optional MaxDate as date = 2958465,           Dates not allowed after  MaxDate. If not given = 2958465 (Dec 31th 2999)
'         Optional Color As Long = &HD0D0D0,            Background color. May also be given as RGB(r,g,b). Standard = grey
'         Optional pLeft As Long = 0,                   L-Position for form. 0 = center
'         Optional pTop As Long = 0)                    T-Position for form. 0 = center
'         As Date

' Examples:
' Sub PickDate()
'   ActiveCell = frmWMDatePicker.GetDate
' End Sub
'
' Sub PickDate()
'   ActiveCell = frmWMDatePicker.GetDate(Date, MinDate, MaxDate, RGB(150,200,150), ActiveWindow.Left, ActiveWindow.Top + CommandBars("Ribbon").Height)
' End Sub
'
' Sub PickDate()
'  ActiveCell = frmWMDatePicker.GetDate(pLeft:=Application.ActiveWindow.Left, pTop:=ActiveWindow.Top + CommandBars("Ribbon").Height)
' End Sub

' Please do NOT open the form like frmWMDatePicker.Show. This will open without values and will place the result into the active cell.
' If StartDate is not given, it will be replaced by the actual date.
' You may limit the pickable dates by setting MinDate and/or MaxDate value.
' If MinDate is not given, it will be set to Feb  1st 0100.
' If MaxDate is not given, it will be set to Dec 31st 9999.
' If Color is not given, standard color is grey.
' If pLeft and pTop are not given, the form will show in the middle of the screen (may not work in multi screen environments).
' Nice values are pLeft = Application.ActiveWindow.Left and pTop = ActiveWindow.Top + CommandBars("Ribbon").Height
' If user hits ENTER or doubleclick on a date, the form will close and will respond with the selected date.
' If user closes the form with X or ESC, the form will respond with date = 0. You may change this by setting the const Return0 to false.

' Datepicker may be controlled by mouse (but no mousewheel) or keybord.
' User may lookup the keyboard controls by pressing <F1>

Private Const Return0 As Boolean = True       'Defines if form returns 0 or date if closed by X or ESC.
Private Const LimitWarning As Boolean = False 'Defines if to show a message when limits are exceeded.

Private MyDate As Date         ' Inner main date var = Actual selected date
Private MyMinDate As Date      ' MinDate
Private MyMaxDate As Date      ' Max Date
Private MyDay As Integer       ' Day of MyDate
Private MyMonth As Integer     ' Month of MyDate
Private MyYear As Integer      ' Year of MyDate
Private MyColor As Long        ' Background color of the form
Private FirstDayOfWeek As Long ' FirstDayOfWeek from users system (Sunday=1 or Monday=2)
Private DirectCall As Boolean  ' Hold info if form was opened via .Show (true) or via .GetDate (false)


Public Function GetDate(Optional StartDate As Date = 0, _
                        Optional MinDate As Date = -657403, _
                        Optional MaxDate As Date = 2958465, _
                        Optional Color As Long = &HD0D0D0, _
                        Optional pLeft As Long = 0, _
                        Optional pTop As Long = 0) _
                        As Date
 If StartDate = 0 Or StartDate < -657403 Or StartDate > 3958465 _
  Then MyDate = Date Else MyDate = StartDate                        'No or wrong StartDate > Set actual Date
 If MinDate = 0 Then MyMinDate = -657403 Else MyMinDate = MinDate   'No MinDate > -657434 Set to Feb 1st 0100
 If MaxDate = 0 Then MyMaxDate = 2958465 Else MyMaxDate = MaxDate   'No MaxDate > Set to Dec 31th 9999
 If MyDate < MyMinDate Then MyDate = MyMinDate + 1                  'MyDate may not be below MinDate
 MyColor = Color                                                    'Save Color for later use
 If pLeft = 0 And pTop = 0 Then                                     'No Position > Set StartUpPosition to ScreenCenter
  frmWMDatePicker.StartUpPosition = 1
 Else
  frmWMDatePicker.StartUpPosition = 0
  frmWMDatePicker.Left = pLeft
  frmWMDatePicker.Top = pTop
 End If
 DirectCall = False                                                 'Mind it's no direct call (no frmWMDatePicker.Show)
 frmWMDatePicker.Show                                               'Show the form
 GetDate = MyDate                                                   'Once form is closed, set result of function to MyDate
End Function


Private Sub UserForm_Activate()
Dim i As Integer
 If MyColor = 0 Then 'Direct call without GetDate(...)
  MyDate = Date
  MyMinDate = -657403
  MyMaxDate = 2958465
  MyColor = &HD0D0D0
  frmWMDatePicker.StartUpPosition = 1
  DirectCall = True
 End If
 Me.BackColor = MyColor 'Set the back color
 For i = 1 To 7 'Write weekday names into grid header
  With frmWMDatePicker.Controls("lb" & Format(i, "00"))
   .Caption = WeekdayName(i, True, 0)
   If Weekday(i, 0) = 1 Then FirstDayOfWeek = i 'Find the system preset
  End With
 Next i
 DateChanged 'call sub to refresh form
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 If DirectCall Then ActiveCell = MyDate
 If Return0 And (CloseMode <> vbFormCode) Then MyDate = 0 'set MyDate to 0 if form closed by X
End Sub

Private Sub ChangeDate(nYear As Integer, nMonth As Integer, nDay As Integer)
 If nYear < 1900 Then MyDate = 0 Else MyDate = DateSerial(nYear, nMonth, nDay)
 If Month(MyDate) <> nMonth Then MyDate = LastDayOfMonth(nMonth, nYear) ' if date was not valid (as Feb 30th), use the last of given month
 DateChanged
End Sub

Private Sub DateChanged()
Dim res As Integer
'Update all views
 If MyDate < MyMinDate Then ' MyDate below MyMinDate
  MyDate = MyMinDate
  If LimitWarning Then MsgBox "No date allowed before " & MyMinDate, vbExclamation, "Date limit reached"
 End If
 If MyDate > MyMaxDate Then ' MyDate after MyMaxDate
  MyDate = MyMaxDate
  If LimitWarning Then MsgBox "No date allowed after " & MyMaxDate, vbExclamation, "Date limit reached"
 End If
 MyDay = Day(MyDate)
 MyMonth = Month(MyDate)
 MyYear = Year(MyDate)
 Me.Caption = Format(MyDate, "Short Date")
 lbMonthYear.Caption = MonthName(MyMonth)
 lbYear.Caption = MyYear
 UpdateGrid
End Sub

Private Sub UpdateGrid()
Dim d As Date
Dim i As Integer
 'Find grid start date
 d = DateSerial(Year(MyDate), Month(MyDate), 2)
 If Weekday(d - 1) = FirstDayOfWeek Then d = d - 7 'Show additional Week before
 Do
  d = d - 1
 Loop Until Weekday(d) = FirstDayOfWeek 'FirstDayOfWeek is preset by user system
 'Write labels and write representing date to .tag
 For i = 1 To 42
  With frmWMDatePicker.Controls("bt" & Format(i, "00"))
   .Caption = Day(d) 'Write day number
   .Tag = d 'Set date to control.tag
   .BorderStyle = fmBorderStyleNone 'Reset Border
   .BackColor = vbWhite 'Reset Color
   If d = MyDate Then .BorderStyle = fmBorderStyleSingle 'Red border if active day
   If Month(d) <> MyMonth Then .BackColor = RGB(230, 230, 230) 'Grey background if not in active month
   If Weekday(d) = vbSaturday Then .ForeColor = RGB(150, 150, 150) 'Grey text if Saturday
   If Weekday(d) = vbSunday Then .ForeColor = RGB(180, 0, 0) 'Red text if Sunday
   .Enabled = ((d >= MyMinDate) And (d <= MyMaxDate))
   d = d + 1
  End With
 Next i
End Sub

Private Function LastDayOfMonth(nMonth As Integer, nYear As Integer) As Date
'Determine the last day of given month
 If nMonth = 12 Then
  nMonth = 1
  nYear = nYear + 1
 Else
  nMonth = nMonth + 1
 End If
 LastDayOfMonth = DateSerial(nYear, nMonth, 1) - 1
End Function

Private Sub MonthUp()
'Move +1 month
Dim M As Integer
Dim Y As Integer
 If MyMonth = 12 Then
  Y = MyYear + 1
  M = 1
 Else
  Y = MyYear
  M = MyMonth + 1
 End If
 Call ChangeDate(Y, M, MyDay)
End Sub

Private Sub MonthDown()
'Move -1 month
Dim M As Integer
Dim Y As Integer
 If MyMonth = 1 Then
  Y = MyYear - 1
  M = 12
 Else
  Y = MyYear
  M = MyMonth - 1
 End If
 Call ChangeDate(Y, M, MyDay)
End Sub

Private Sub YearUp()
'Move +1 year
 Call ChangeDate(MyYear + 1, MyMonth, MyDay)
End Sub

Private Sub YearDown()
'Move -1 year
 Call ChangeDate(MyYear - 1, MyMonth, MyDay)
End Sub

Private Sub lbPrevMonth_Click()
 MonthDown
End Sub

Private Sub lbNextMonth_Click()
 MonthUp
End Sub

Private Sub lbNextYear_Click()
 YearUp
End Sub

Private Sub lbPrevYear_Click()
 YearDown
End Sub

Private Sub lbToday_Click()
 MyDate = Date
 DateChanged
End Sub

Private Sub Date_Clicked(ClickedDate As String) ' is called by bt01_Click to bt42_Click
 MyDate = DateValue(ClickedDate)
 DateChanged
End Sub

Private Sub Date_DblClicked(ClickedDate As String) ' is called by bt01_DblClick to bt42_DblClick
 MyDate = DateValue(ClickedDate)
 Unload frmWMDatePicker
End Sub

Private Sub btOK_Click() 'btOK invisible by width=0, but is default button
 Unload Me
End Sub

Private Sub btOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 Select Case KeyCode
  Case 27 'ESC
   If Return0 Then MyDate = 0
   Unload frmWMDatePicker
  Case 33 'PageUp
   If Shift = 0 Then MonthDown Else YearDown
  Case 34 'PageDown
   If Shift = 0 Then MonthUp Else YearUp
  Case 36 'POS1
   lbToday_Click
  Case 37 'Left
   MyDate = MyDate - 1: DateChanged
  Case 38 'Up
   MyDate = MyDate - 7: DateChanged
  Case 39 'Right
   MyDate = MyDate + 1: DateChanged
  Case 40 'Down
   MyDate = MyDate + 7: DateChanged
  Case 112
   lbInfo_Click
 End Select
End Sub

Private Sub lbInfo_Click()
Dim res As Integer
 res = MsgBox("WMDatePicker" & vbLf _
            & "©2021 Wolfgang.Marder@gmail.com" & vbLf & vbLf _
            & "Keyboard controls:" & vbLf _
            & "<POS1>            : Today" & vbLf _
            & "Arrow keys        : Move selection in grid" & vbLf _
            & "PgUp/Down      : Change Month" & vbLf _
            & "<CTRL>&PgUp/Down: Change Year" & vbLf _
            & "<ENTER> or DlbClick  : Close and use selected date", vbInformation, "About")
End Sub


'Click-Events for all grid labels:
Private Sub bt01_Click(): Call Date_Clicked(bt01.Tag): End Sub
Private Sub bt02_Click(): Call Date_Clicked(bt02.Tag): End Sub
Private Sub bt03_Click(): Call Date_Clicked(bt03.Tag): End Sub
Private Sub bt04_Click(): Call Date_Clicked(bt04.Tag): End Sub
Private Sub bt05_Click(): Call Date_Clicked(bt05.Tag): End Sub
Private Sub bt06_Click(): Call Date_Clicked(bt06.Tag): End Sub
Private Sub bt07_Click(): Call Date_Clicked(bt07.Tag): End Sub
Private Sub bt08_Click(): Call Date_Clicked(bt08.Tag): End Sub
Private Sub bt09_Click(): Call Date_Clicked(bt09.Tag): End Sub
Private Sub bt10_Click(): Call Date_Clicked(bt10.Tag): End Sub
Private Sub bt11_Click(): Call Date_Clicked(bt11.Tag): End Sub
Private Sub bt12_Click(): Call Date_Clicked(bt12.Tag): End Sub
Private Sub bt13_Click(): Call Date_Clicked(bt13.Tag): End Sub
Private Sub bt14_Click(): Call Date_Clicked(bt14.Tag): End Sub
Private Sub bt15_Click(): Call Date_Clicked(bt15.Tag): End Sub
Private Sub bt16_Click(): Call Date_Clicked(bt16.Tag): End Sub
Private Sub bt17_Click(): Call Date_Clicked(bt17.Tag): End Sub
Private Sub bt18_Click(): Call Date_Clicked(bt18.Tag): End Sub
Private Sub bt19_Click(): Call Date_Clicked(bt19.Tag): End Sub
Private Sub bt20_Click(): Call Date_Clicked(bt20.Tag): End Sub
Private Sub bt21_Click(): Call Date_Clicked(bt21.Tag): End Sub
Private Sub bt22_Click(): Call Date_Clicked(bt22.Tag): End Sub
Private Sub bt23_Click(): Call Date_Clicked(bt23.Tag): End Sub
Private Sub bt24_Click(): Call Date_Clicked(bt24.Tag): End Sub
Private Sub bt25_Click(): Call Date_Clicked(bt25.Tag): End Sub
Private Sub bt26_Click(): Call Date_Clicked(bt26.Tag): End Sub
Private Sub bt27_Click(): Call Date_Clicked(bt27.Tag): End Sub
Private Sub bt28_Click(): Call Date_Clicked(bt28.Tag): End Sub
Private Sub bt29_Click(): Call Date_Clicked(bt29.Tag): End Sub
Private Sub bt30_Click(): Call Date_Clicked(bt30.Tag): End Sub
Private Sub bt31_Click(): Call Date_Clicked(bt31.Tag): End Sub
Private Sub bt32_Click(): Call Date_Clicked(bt32.Tag): End Sub
Private Sub bt33_Click(): Call Date_Clicked(bt33.Tag): End Sub
Private Sub bt34_Click(): Call Date_Clicked(bt34.Tag): End Sub
Private Sub bt35_Click(): Call Date_Clicked(bt35.Tag): End Sub
Private Sub bt36_Click(): Call Date_Clicked(bt36.Tag): End Sub
Private Sub bt37_Click(): Call Date_Clicked(bt37.Tag): End Sub
Private Sub bt38_Click(): Call Date_Clicked(bt38.Tag): End Sub
Private Sub bt39_Click(): Call Date_Clicked(bt39.Tag): End Sub
Private Sub bt40_Click(): Call Date_Clicked(bt40.Tag): End Sub
Private Sub bt41_Click(): Call Date_Clicked(bt41.Tag): End Sub
Private Sub bt42_Click(): Call Date_Clicked(bt42.Tag): End Sub
'DblClick-Events for all grid labels:
Private Sub bt01_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt01.Tag): End Sub
Private Sub bt02_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt02.Tag): End Sub
Private Sub bt03_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt03.Tag): End Sub
Private Sub bt04_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt04.Tag): End Sub
Private Sub bt05_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt05.Tag): End Sub
Private Sub bt06_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt06.Tag): End Sub
Private Sub bt07_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt07.Tag): End Sub
Private Sub bt08_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt08.Tag): End Sub
Private Sub bt09_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt09.Tag): End Sub
Private Sub bt10_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt10.Tag): End Sub
Private Sub bt11_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt11.Tag): End Sub
Private Sub bt12_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt12.Tag): End Sub
Private Sub bt13_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt13.Tag): End Sub
Private Sub bt14_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt14.Tag): End Sub
Private Sub bt15_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt15.Tag): End Sub
Private Sub bt16_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt16.Tag): End Sub
Private Sub bt17_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt17.Tag): End Sub
Private Sub bt18_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt18.Tag): End Sub
Private Sub bt19_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt19.Tag): End Sub
Private Sub bt20_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt20.Tag): End Sub
Private Sub bt21_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt21.Tag): End Sub
Private Sub bt22_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt22.Tag): End Sub
Private Sub bt23_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt23.Tag): End Sub
Private Sub bt24_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt24.Tag): End Sub
Private Sub bt25_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt25.Tag): End Sub
Private Sub bt26_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt26.Tag): End Sub
Private Sub bt27_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt27.Tag): End Sub
Private Sub bt28_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt28.Tag): End Sub
Private Sub bt29_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt29.Tag): End Sub
Private Sub bt30_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt30.Tag): End Sub
Private Sub bt31_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt31.Tag): End Sub
Private Sub bt32_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt32.Tag): End Sub
Private Sub bt33_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt33.Tag): End Sub
Private Sub bt34_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt34.Tag): End Sub
Private Sub bt35_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt35.Tag): End Sub
Private Sub bt36_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt36.Tag): End Sub
Private Sub bt37_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt37.Tag): End Sub
Private Sub bt38_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt38.Tag): End Sub
Private Sub bt39_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt39.Tag): End Sub
Private Sub bt40_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt40.Tag): End Sub
Private Sub bt41_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt41.Tag): End Sub
Private Sub bt42_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Call Date_DblClicked(bt42.Tag): End Sub
