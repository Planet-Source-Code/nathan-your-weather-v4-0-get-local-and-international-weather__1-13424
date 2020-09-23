VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form LMapfrm
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7590
   ClientLeft      =   495
   ClientTop       =   780
   ClientWidth     =   10860
   ControlBox      =   0   'False
   ForeColor       =   &H80000006&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1
      BorderStyle     =   1  'Fixed Single
      Height          =   7350
      Left            =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   10860
   End
   Begin InetCtlsObjects.Inet Inet1
      Left            =   9000
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1
      BackStyle       =   0  'Transparent
      Caption         =   "Your Maps v4.0"
      BeginProperty Font
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   25
      TabIndex        =   2
      Top             =   0
      Width           =   10560
   End
   Begin VB.Label Label12
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   43
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Image.... Please Wait"
      BeginProperty Font
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.Label Label4
      BackStyle       =   0  'Transparent
      Height          =   160
      Left            =   10680
      TabIndex        =   3
      Top             =   30
      Width           =   165
   End
   Begin VB.Line Line1
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   10920
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line Line2
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   10920
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Shape Shape1
      BorderColor     =   &H80000001&
      Height          =   135
      Left            =   10700
      Top             =   25
      Width           =   135
   End
End
Attribute VB_Name = "LMapfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************
Private Sub Form_Load()
'*********************

	Center Me
	LMapfrm.MousePointer = 99
	LMapfrm.MouseIcon = Mainfrm.ImageList1.ListImages(1).Picture
	Me.Show
	DoEvents
	Call get_largepic

'*******
End Sub
'*******

'*********************************
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*********************************

	FormMove Me

'*******
End Sub
'*******

'************************
Private Sub Label4_Click()
'************************

	Unload Me

'*******
End Sub
'*******

'*********************************
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*********************************

	Shape1.Left = 10710
	Shape1.Top = 40

'*******
End Sub
'*******

'*******************************
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************

	Shape1.Left = 10700
	Shape1.Top = 25
	Timeout (0.3)

'*******
End Sub
'*******

'************************
Sub get_largepic()
'************************

	 ' Goes out to weather.com and gets larger image 720x486 '
	If Mainfrm.Map_lst.ListIndex > 0 And Mainfrm.Map_lst.ListIndex < 39 Then
		Label3.Visible = True
		Image1.Visible = False
		Dim file() As Byte
		sel_d = Mainfrm.Map_lst.ListIndex
		If sel_d = 2 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_cen_720x486.jpg", icByteArray)
		ElseIf sel_d = 3 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_ec_720x486.jpg", icByteArray)
		ElseIf sel_d = 4 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_mw_720x486.jpg", icByteArray)
		ElseIf sel_d = 5 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_nc_720x486.jpg", icByteArray)
		ElseIf sel_d = 6 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_ne_720x486.jpg", icByteArray)
		ElseIf sel_d = 7 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_nw_720x486.jpg", icByteArray)
		ElseIf sel_d = 8 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_sc_720x486.jpg", icByteArray)
		ElseIf sel_d = 9 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_se_720x486.jpg", icByteArray)
		ElseIf sel_d = 10 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_sw_720x486.jpg", icByteArray)
		ElseIf sel_d = 11 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_wc_720x486.jpg", icByteArray)
		ElseIf sel_d = 12 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/maps/current/cur_w_720x486.jpg", icByteArray)
		ElseIf sel_d = 16 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/africasat_720x486.jpg", icByteArray)
		ElseIf sel_d = 17 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/asiasat_720x486.jpg", icByteArray)
		ElseIf sel_d = 18 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/aussiesat_720x486.jpg", icByteArray)
		ElseIf sel_d = 19 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/cenamersat_720x486.jpg", icByteArray)
		ElseIf sel_d = 20 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/europesat_720x486.jpg", icByteArray)
		ElseIf sel_d = 21 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/canadasat_720x486.jpg", icByteArray)
		ElseIf sel_d = 22 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/sasat_720x486.jpg", icByteArray)
		ElseIf sel_d = 26 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/atl_oce_sat_720x486.jpg", icByteArray)
		ElseIf sel_d = 27 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/caribsat_720x486.jpg", icByteArray)
		ElseIf sel_d = 28 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/gomex_sat_720x486.jpg", icByteArray)
		ElseIf sel_d = 29 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/indian_oce_sat_720x486.jpg", icByteArray)
		ElseIf sel_d = 30 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/pac_oce_sat_720x486.jpg", icByteArray)
		ElseIf sel_d = 31 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/tropsat_720x486.jpg", icByteArray)
		ElseIf sel_d = 35 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/euro-africasat_720x486.jpg", icByteArray)
		ElseIf sel_d = 36 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/pacglobsat_720x486.jpg", icByteArray)
		ElseIf sel_d = 37 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/wpacglobsat_720x486.jpg", icByteArray)
		ElseIf sel_d = 38 Then
			file() = Inet1.OpenURL("http://maps.weather.com/images/sat/goesdisk_720x486.jpg", icByteArray)
		End If
		 ' Saves image to data.jpg '
		Label3.Visible = False
		Image1.Visible = True
		Open App.Path & "\data2.jpg" For Binary As #1
		Put #1, , file
		Close #1
		Image1.Picture = LoadPicture(App.Path & "\data2.jpg")
	End If

'*******
End Sub
'*******

