VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SysTray.ocx"
Begin VB.Form Systemtrayfrm
   ClientHeight    =   570
   ClientLeft      =   165
   ClientTop       =   5865
   ClientWidth     =   1560
   ClipControls    =   0   'False
   Icon            =   "Systrayfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin SysTray.SystemTray SystemTray1
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
   Begin VB.Menu main
      Caption         =   "Menu"
      Begin VB.Menu update
         Caption         =   "Update Now"
      End
      Begin VB.Menu auto
         Caption         =   "Auto Update"
      End
      Begin VB.Menu BAR1
         Caption         =   "-"
      End
      Begin VB.Menu scroll
         Caption         =   "Local Detail Scroll"
      End
   End
End
Attribute VB_Name = "Systemtrayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const sys_Add = 0
Private Const sys_Delete = 2
'***************************************
Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
'***************************************

	 ' Removes icon from systray and shows main form '
	Hidefrm.Show
	Mainfrm.Show
	SystemTray1.Action = sys_Delete
	Systemtrayfrm.Hide

'*******
End Sub
'*******

'**********************
Private Sub auto_Click()
'**********************

	 ' Input box for interval between auto updates '
	If Mainfrm.Timer1.Enabled = True Then
		status = "On"
		int_time = Mainfrm.Timer1.Tag
	Else
		status = "Off"
		int_time = "0"
	End If
	Interval = InputBox("Update Interval (In Minutes)" & Chr$(13) & "Type 'Off' To Turn Off Automatic Updater" & Chr$(13) & Chr$(13) & "Current Status: " & status & Chr$(13) & "Interval: " & int_time & " minutes" & Chr$(13) & Chr$(13) & "* Only Zip Code Locations May Be Automatically Updated. International Weather Is Not Updated As Often By Weather.com, Therefore The Auto Updater Wouldn't Be Of Much Use." & Chr$(13) & Chr$(13) & "Enter In The Number Of Minutes Between Update Interval Below", "Automatic Update")
	If LCase(Interval) = "off" Then ' Turns off auto updater '
		Mainfrm.Timer1.Enabled = False
		Mainfrm.Timer1.Tag = 0
	ElseIf IsNumeric(Interval) Then
		Mainfrm.Timer1.Tag = Interval
		Mainfrm.Timer1.Enabled = True
	Else
	End If

'*******
End Sub
'*******

'************************
Private Sub scroll_Click()
'************************

	If scroll.Checked = False Then
		Open (App.Path & "\detail.dat") For Input As #1
		Mainfrm.detail_txt.Text = Input(LOF(1), 1)
		Close #1
		Mainfrm.tmrScroll.Enabled = True
		Mainfrm.detail_txt.Visible = True
		Mainfrm.fast.Visible = True
		Mainfrm.slow.Visible = True
		scroll.Checked = True
		entry$ = "Yes"
		r% = WritePrivateProfileString("Report", "DetailCheck", entry$, iniPath$)
	ElseIf scroll.Checked = True Then
		Mainfrm.detail_txt.Text = vbNullString
		Mainfrm.tmrScroll.Enabled = False
		Mainfrm.detail_txt.Visible = False
		Mainfrm.fast.Visible = False
		Mainfrm.slow.Visible = False
		scroll.Checked = False
		entry$ = "No"
		r% = WritePrivateProfileString("Report", "DetailCheck", entry$, iniPath$)
	End If

'*******
End Sub
'*******

'************************
Private Sub update_Click()
'************************

	 ' Checks for correct format and calls the main logic '
	If IsNumeric(Mainfrm.Ziptxt) And Len(Mainfrm.Ziptxt) = 5 Then
		Call Mainfrm.Disable_Me
		Call Mainfrm.Load_ZIP_Weather
		Call Mainfrm.Enable_Me
	Else
		MsgBox "Please Enter A Valid Zip Code", vbOKOnly + vbInformation, "Error"
	End If

'*******
End Sub
'*******

