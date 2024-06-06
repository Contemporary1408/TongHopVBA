'Launching SAP and logging into the main screen if the program is not already open
Sub test()
If IsProcessRunning("saplogon.EXE") = False Then
Dim SapGui, Applic, connection, session, WSHShell

Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus

Set WSHShell = CreateObject("WScript.Shell")

Do Until WSHShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set WSHShell = Nothing
Set SapGui = GetObject("SAPGUI")
Set Applic = SapGui.GetScriptingEngine
Set connection = Applic.OpenConnection("BTMV Production Server", True)
Set session = connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "010"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "VHACS12"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "btmv@140894"
session.findById("wnd[0]").sendVKey 0
End If

'Executing main screen, this part here is just setting variables for the  logged in interface. So once this is done you can enter the transaction.

'-----------new--------------------------
On Error Resume Next
'-----------new--------------------------
If Not IsObject(XXX) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set XXX = SapGuiAuto.GetScriptingEngine
   '-----------new--------------------------
   myError = Err.Number
   '-----------new--------------------------
End If
If Not IsObject(connection) Then
   Set connection = XXX.Children(0)
   '-----------new--------------------------
   myError = Err.Number
   '-----------new--------------------------
End If

If Not IsObject(session) Then
   Set session = connection.Children(0)
   '-----------new--------------------------
   myError = Err.Number
   '-----------new--------------------------
End If
'-----------new--------------------------
On Error GoTo 0
If myError <> 0 Then
  Set connection = XXX.OpenConnection("BTMV Production Server", True)
  Set session = connection.Children(0)
  session.findById("wnd[0]").maximize
End If
'-----------new--------------------------

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject XXX, "on"
End If

'Transaction comes here
session.SendCommand ("/oZGM07")
session.findById("wnd[0]").maximize
End Sub

'#############################################
Function IsProcessRunning(process As String)
Dim objList As Object

Set objList = GetObject("winmgmts:") _
    .ExecQuery("select * from win32_process where name='" & process & "'")

IsProcessRunning = objList.Count > 0

End Function
