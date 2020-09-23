Attribute VB_Name = "Main"
Public Const GWL_HWNDPARENT = -8&
Public Const MF_BYCOMMAND = 0&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_MOVE = &HF010&
Public Const SC_SIZE = &HF000&
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Global r%
Global entry$
Global iniPath$
'************************
Sub Center(frmform As Form)
'************************

	frmform.Left = (Screen.Width - frmform.Width) / 2
	frmform.Top = (Screen.Height - frmform.Height) / 2

'*******
End Sub
'*******

'*************************
Sub CenterX(frmform As Form)
'*************************

	frmform.Top = (Screen.Height - frmform.Height) / 2

'*******
End Sub
'*******

'*****************************
Function FileExists(filename As String) As Integer
'*****************************

	On Error Resume Next
	X% = Len(Dir$(filename))
	If Err Or X% = 0 Then FileExists = False Else FileExists = True

'*******
End Function
'*******

'**************************
Public Sub FormMove(theform As Form)
'**************************

	ReleaseCapture
	Call SendMessage(theform.hwnd, &HA1, 2, 0&)

'*******
End Sub
'*******

'******************************
Function GetFromINI(AppName$, KeyName$, filename$) As String
'******************************

	Dim RetStr As String
	RetStr = String(255, Chr$(0))
	GetFromINI = Left$(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, vbNullString, RetStr, Len(RetStr), filename$))

'*******
End Function
'*******

'*********************
Sub Systrayme()
'*********************

	Mainfrm.Hide
	Hidefrm.Hide
	If Systemtrayfrm.SystemTray1.IsIconLoaded = False Then
		Systemtrayfrm.SystemTray1.Icon = Val(Systemtrayfrm.Icon)
		Systemtrayfrm.SystemTray1.SysTrayText = "Your Weather v4.0"
		Systemtrayfrm.SystemTray1.Action = sys_Add
	End If

'*******
End Sub
'*******

'***************************
Sub Timeout(duration)
'***************************

	DoEvents
	starttime = Timer
 
	Do While Timer - starttime < duration    
		DoEvents
	Loop    
 

'*******
End Sub
'*******

'****************
Sub main()
'****************

	 ' show the main form
	Mainfrm.Show
	 ' and show the owner.  The owner sets itself up as the parent of frmmain.  We could do all that here though.
	Form1.Show

'*******
End Sub
'*******

