VERSION 5.00
Begin VB.Form splashfrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1860
   ClientLeft      =   1680
   ClientTop       =   1545
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splash.frx":0000
   ScaleHeight     =   1860
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "splashfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Center Me
    Me.Show
    DoEvents
    Label1.Caption = App.Major & "." & App.Minor & "." & App.Revision
    DoEvents
    Timeout (3)
    Me.Hide
    Unload Me
    Mainfrm.Show
End Sub


