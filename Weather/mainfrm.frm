VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Mainfrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6945
   ClientLeft      =   2160
   ClientTop       =   1485
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "US City/Territory Weather"
      ForeColor       =   &H80000010&
      Height          =   1935
      Left            =   0
      TabIndex        =   53
      Top             =   4800
      Width           =   3495
      Begin VB.ListBox Alpha_city_lst 
         ForeColor       =   &H8000000C&
         Height          =   885
         ItemData        =   "mainfrm.frx":0BCA
         Left            =   120
         List            =   "mainfrm.frx":0BCC
         TabIndex        =   54
         ToolTipText     =   "Double Click To Select"
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox Map_lst 
         ForeColor       =   &H8000000C&
         Height          =   285
         ItemData        =   "mainfrm.frx":0BCE
         Left            =   120
         List            =   "mainfrm.frx":0C56
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1560
         Width           =   3255
      End
      Begin VB.ComboBox US_combo 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000C&
         Height          =   285
         ItemData        =   "mainfrm.frx":1006
         Left            =   120
         List            =   "mainfrm.frx":10B5
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   240
         Width           =   3255
      End
      Begin VB.ListBox us_city_lst 
         ForeColor       =   &H8000000C&
         Height          =   885
         ItemData        =   "mainfrm.frx":1395
         Left            =   600
         List            =   "mainfrm.frx":1397
         TabIndex        =   55
         ToolTipText     =   "Double Click To Select"
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      Caption         =   "International City Weather"
      ForeColor       =   &H8000000C&
      Height          =   1935
      Left            =   4080
      TabIndex        =   59
      Top             =   4800
      Width           =   3495
      Begin VB.ListBox Country_lst 
         ForeColor       =   &H8000000C&
         Height          =   555
         ItemData        =   "mainfrm.frx":1399
         Left            =   120
         List            =   "mainfrm.frx":139B
         TabIndex        =   61
         ToolTipText     =   "Double Click To Select"
         Top             =   600
         Width           =   3255
      End
      Begin VB.ListBox Int_city_lst 
         ForeColor       =   &H8000000C&
         Height          =   555
         ItemData        =   "mainfrm.frx":139D
         Left            =   120
         List            =   "mainfrm.frx":139F
         TabIndex        =   60
         ToolTipText     =   "Double Click To Select"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Int_combo 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000C&
         Height          =   285
         ItemData        =   "mainfrm.frx":13A1
         Left            =   120
         List            =   "mainfrm.frx":13EA
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   240
         Width           =   3255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":1587
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":2261
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":2A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":31D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":39D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":40E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":490E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":519F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":5A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":6307
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":6B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":746B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":7C15
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":855F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":8DBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":961D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":9ECD
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":A732
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":AF04
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":B5FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":BEA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":C6FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":C88A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":D126
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":DA06
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":E7FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   8160
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   90
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   240
      Width           =   7575
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Show Extended Weather Options"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   30
         TabIndex        =   71
         Top             =   4160
         Width           =   2400
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4680
         TabIndex        =   97
         Top             =   1320
         Width           =   2895
         Begin VB.OptionButton ASP 
            BackColor       =   &H8000000E&
            Caption         =   "ASP"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   255
            Left            =   1080
            TabIndex        =   103
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton MSNET 
            BackColor       =   &H8000000E&
            Caption         =   "MSINET"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   255
            Left            =   2040
            TabIndex        =   98
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton XML 
            BackColor       =   &H8000000E&
            Caption         =   "XML"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   255
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Fahrenheit"
         ForeColor       =   &H8000000C&
         Height          =   245
         Left            =   4680
         TabIndex        =   69
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000009&
         Caption         =   "Celsius"
         ForeColor       =   &H80000010&
         Height          =   245
         Left            =   5760
         TabIndex        =   68
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000005&
         Caption         =   "Kelvin"
         ForeColor       =   &H8000000C&
         Height          =   245
         Left            =   6720
         TabIndex        =   67
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   120
         Left            =   7230
         Picture         =   "mainfrm.frx":EEA7
         ScaleHeight     =   120
         ScaleWidth      =   135
         TabIndex        =   102
         Top             =   680
         Width           =   135
      End
      Begin VB.TextBox Ziptxt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   225
         Left            =   6660
         MaxLength       =   5
         MousePointer    =   3  'I-Beam
         TabIndex        =   66
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox detail_txt 
         ForeColor       =   &H80000006&
         Height          =   240
         Left            =   2520
         TabIndex        =   94
         Top             =   4140
         Width           =   4455
      End
      Begin VB.Label Barometer 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   44
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label Conditions 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Date_1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   0
         TabIndex        =   85
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Date_10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   6600
         TabIndex        =   93
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   840
         TabIndex        =   84
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   1560
         TabIndex        =   86
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   2280
         TabIndex        =   92
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   3000
         TabIndex        =   91
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   3720
         TabIndex        =   90
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   4440
         TabIndex        =   89
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   5160
         TabIndex        =   88
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Date_9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   135
         Left            =   5880
         TabIndex        =   87
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Day1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Day10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   6615
         TabIndex        =   83
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day10_Weather 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6600
         TabIndex        =   74
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Day10_hi 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6600
         TabIndex        =   77
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day10_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6600
         TabIndex        =   80
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day1_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Day1_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Day1_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Day2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   855
         TabIndex        =   10
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day2_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Day2_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day2_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1575
         TabIndex        =   9
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day3_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Day3_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day3_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   2295
         TabIndex        =   8
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day4_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   2280
         TabIndex        =   20
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Day4_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day4_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   3015
         TabIndex        =   7
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day5_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3000
         TabIndex        =   23
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Day5_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day5_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   3735
         TabIndex        =   6
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day6_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3720
         TabIndex        =   26
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Day6_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day6_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day7 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   4455
         TabIndex        =   5
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day7_Weather 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   4440
         TabIndex        =   29
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Day7_hi 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day7_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day8 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   5175
         TabIndex        =   81
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day8_Weather 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5160
         TabIndex        =   72
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Day8_hi 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   75
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day8_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   78
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Day9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   5895
         TabIndex        =   82
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Day9_Weather 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5880
         TabIndex        =   73
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Day9_hi 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   76
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Day9_lo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5880
         TabIndex        =   79
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Dewpoint 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Humidity 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   46
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature Measurements                                      "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   4680
         TabIndex        =   101
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Wind:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1920
         TabIndex        =   34
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Humidity:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Dewpoint:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Visibility:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Barometer:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunrise:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunset:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1680
         TabIndex        =   40
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Retrieval Method                                                "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   4680
         TabIndex        =   100
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code              "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   6660
         TabIndex        =   63
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Conditions:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   1560
         TabIndex        =   52
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line10 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   7920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000004&
         X1              =   0
         X2              =   7920
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   7200
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   7200
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000A&
         X1              =   1320
         X2              =   1320
         Y1              =   2040
         Y2              =   240
      End
      Begin VB.Label Report 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   60
         Width           =   7575
      End
      Begin VB.Label Sunrise 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   43
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Sunset 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Temperature 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   49
         Top             =   540
         Width           =   480
      End
      Begin VB.Label Visibility 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   45
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Wind 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   720
         Width           =   3120
      End
      Begin VB.Image dayi1 
         Height          =   420
         Left            =   480
         Stretch         =   -1  'True
         Top             =   720
         Width           =   420
      End
      Begin VB.Image dayi10 
         Height          =   420
         Left            =   6720
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi2 
         Height          =   420
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi3 
         Height          =   420
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi4 
         Height          =   420
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi5 
         Height          =   420
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi6 
         Height          =   420
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi7 
         Height          =   420
         Left            =   4560
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi8 
         Height          =   420
         Left            =   5280
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image dayi9 
         Height          =   420
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label fast 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   6960
         TabIndex        =   95
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label slow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   7140
         TabIndex        =   96
         Top             =   4080
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   344
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8211
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "1/9/01"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   2520
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   3000
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3540
      Picture         =   "mainfrm.frx":F1CD
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   7200
      TabIndex        =   70
      Top             =   60
      Width           =   150
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Weather 4.0"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   15
      TabIndex        =   64
      Top             =   15
      Width           =   7095
   End
   Begin VB.Label Label17 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Weather v3.0 "
      ForeColor       =   &H8000000C&
      Height          =   180
      Left            =   -4200
      TabIndex        =   58
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   7380
      TabIndex        =   65
      Top             =   15
      Width           =   180
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -720
      X2              =   7560
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   0
      X2              =   7560
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000001&
      Height          =   135
      Left            =   7410
      Top             =   30
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   60
      Left            =   7215
      Top             =   105
      Width           =   135
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim weatherxml As New MSXML.XMLHTTPRequest
' Ok, if you get an error message at this line of code, you need to reference msxml.dll
' Goto the Menu and 'Project' - 'Refeferences', goto browse and select the msxml.dll file
' which should be in your \system32 folder under your windows folder.
'
' then make sure the Microsoft XML, Version 2.0
' is checked in the references box and you are all set.
Dim date_str As String
Dim day_hi As String
Dim day_lo As String
Dim day_weather As String
Dim int_value As Integer
Dim pos1_a As Long
Dim str_a As String
Dim str_b As String
Dim str_local As String
Dim str_weather As String
Dim temp_day As String
Dim temp_hold_pos As Long


Sub Fix_Data()
      
        iniPath$ = App.Path + "\weather.dat"
        Data1 = Trim$(GetFromINI("Report", "Area", iniPath$))
        Data2 = Trim$(GetFromINI("Report", "Time", iniPath$))
        Data3 = Trim$(GetFromINI("Report", "Zip", iniPath$))
        Data4 = Trim$(GetFromINI("Report", "Degrees", iniPath$))
        Data5 = Trim$(GetFromINI("Report", "Position", iniPath$))
        Data6 = Trim$(GetFromINI("Report", "DetailCheck", iniPath$))
        Data7 = Trim$(GetFromINI("Report", "Interval", iniPath$))
        Data8 = Trim$(GetFromINI("Report", "Method", iniPath$))
        
        Data9 = Trim$(GetFromINI("Current", "Conditions", iniPath$))
        Data10 = Trim$(GetFromINI("Current", "Temperature", iniPath$))
        Data11 = Trim$(GetFromINI("Current", "Wind", iniPath$))
        Data12 = Trim$(GetFromINI("Current", "Humidity", iniPath$))
        Data13 = Trim$(GetFromINI("Current", "Barometer", iniPath$))
        Data14 = Trim$(GetFromINI("Current", "Dewpoint", iniPath$))
        Data15 = Trim$(GetFromINI("Current", "Visibility", iniPath$))
        Data16 = Trim$(GetFromINI("Current", "Sunrise", iniPath$))
        Data17 = Trim$(GetFromINI("Current", "Sunset", iniPath$))
        
        Data18 = Trim$(GetFromINI("Weekday", "Day1", iniPath$))
        Data19 = Trim$(GetFromINI("Weekday", "Day2", iniPath$))
        Data20 = Trim$(GetFromINI("Weekday", "Day3", iniPath$))
        Data21 = Trim$(GetFromINI("Weekday", "Day4", iniPath$))
        Data22 = Trim$(GetFromINI("Weekday", "Day5", iniPath$))
        Data23 = Trim$(GetFromINI("Weekday", "Day6", iniPath$))
        Data24 = Trim$(GetFromINI("Weekday", "Day7", iniPath$))
        Data25 = Trim$(GetFromINI("Weekday", "Day8", iniPath$))
        Data26 = Trim$(GetFromINI("Weekday", "Day9", iniPath$))
        Data27 = Trim$(GetFromINI("Weekday", "Day10", iniPath$))
        
        Data28 = Trim$(GetFromINI("High Temp", "Day1", iniPath$))
        Data29 = Trim$(GetFromINI("High Temp", "Day2", iniPath$))
        Data30 = Trim$(GetFromINI("High Temp", "Day3", iniPath$))
        Data31 = Trim$(GetFromINI("High Temp", "Day4", iniPath$))
        Data32 = Trim$(GetFromINI("High Temp", "Day5", iniPath$))
        Data33 = Trim$(GetFromINI("High Temp", "Day6", iniPath$))
        Data34 = Trim$(GetFromINI("High Temp", "Day7", iniPath$))
        Data35 = Trim$(GetFromINI("High Temp", "Day8", iniPath$))
        Data36 = Trim$(GetFromINI("High Temp", "Day9", iniPath$))
        Data37 = Trim$(GetFromINI("High Temp", "Day10", iniPath$))
        
        Data38 = Trim$(GetFromINI("Low Temp", "Day1", iniPath$))
        Data39 = Trim$(GetFromINI("Low Temp", "Day2", iniPath$))
        Data40 = Trim$(GetFromINI("Low Temp", "Day3", iniPath$))
        Data41 = Trim$(GetFromINI("Low Temp", "Day4", iniPath$))
        Data42 = Trim$(GetFromINI("Low Temp", "Day5", iniPath$))
        Data43 = Trim$(GetFromINI("Low Temp", "Day6", iniPath$))
        Data44 = Trim$(GetFromINI("Low Temp", "Day7", iniPath$))
        Data45 = Trim$(GetFromINI("Low Temp", "Day8", iniPath$))
        Data46 = Trim$(GetFromINI("Low Temp", "Day9", iniPath$))
        Data47 = Trim$(GetFromINI("Low Temp", "Day10", iniPath$))

        Data48 = Trim$(GetFromINI("Weather", "Day1", iniPath$))
        Data49 = Trim$(GetFromINI("Weather", "Day2", iniPath$))
        Data50 = Trim$(GetFromINI("Weather", "Day3", iniPath$))
        Data51 = Trim$(GetFromINI("Weather", "Day4", iniPath$))
        Data52 = Trim$(GetFromINI("Weather", "Day5", iniPath$))
        Data53 = Trim$(GetFromINI("Weather", "Day6", iniPath$))
        Data54 = Trim$(GetFromINI("Weather", "Day7", iniPath$))
        Data55 = Trim$(GetFromINI("Weather", "Day8", iniPath$))
        Data56 = Trim$(GetFromINI("Weather", "Day9", iniPath$))
        Data57 = Trim$(GetFromINI("Weather", "Day10", iniPath$))
        
        Data58 = Trim$(GetFromINI("Dates", "Day1", iniPath$))
        Data59 = Trim$(GetFromINI("Dates", "Day2", iniPath$))
        Data60 = Trim$(GetFromINI("Dates", "Day3", iniPath$))
        Data61 = Trim$(GetFromINI("Dates", "Day4", iniPath$))
        Data62 = Trim$(GetFromINI("Dates", "Day5", iniPath$))
        Data63 = Trim$(GetFromINI("Dates", "Day6", iniPath$))
        Data64 = Trim$(GetFromINI("Dates", "Day7", iniPath$))
        Data65 = Trim$(GetFromINI("Dates", "Day8", iniPath$))
        Data66 = Trim$(GetFromINI("Dates", "Day9", iniPath$))
        Data67 = Trim$(GetFromINI("Dates", "Day10", iniPath$))

        Kill App.Path & "\weather.dat"
        
        entry$ = Data1
        r% = WritePrivateProfileString("Report", "Area", entry$, iniPath$)
        entry$ = Data2
        r% = WritePrivateProfileString("Report", "Time", entry$, iniPath$)
        entry$ = Data3
        r% = WritePrivateProfileString("Report", "Zip", entry$, iniPath$)
        entry$ = Data4
        r% = WritePrivateProfileString("Report", "Degrees", entry$, iniPath$)
        entry$ = Data5
        r% = WritePrivateProfileString("Report", "Position", entry$, iniPath$)
        entry$ = Data6
        r% = WritePrivateProfileString("Report", "DetailCheck", entry$, iniPath$)
        entry$ = Data7
        r% = WritePrivateProfileString("Report", "Interval", entry$, iniPath$)
        entry$ = Data8
        r% = WritePrivateProfileString("Report", "Method", entry$, iniPath$)
        entry$ = Data9
        r% = WritePrivateProfileString("Current", "Conditions", entry$, iniPath$)
        entry$ = Data10
        r% = WritePrivateProfileString("Current", "Temperature", entry$, iniPath$)
        entry$ = Data11
        r% = WritePrivateProfileString("Current", "Wind", entry$, iniPath$)
        entry$ = Data12
        r% = WritePrivateProfileString("Current", "Humidity", entry$, iniPath$)
        entry$ = Data13
        r% = WritePrivateProfileString("Current", "Barometer", entry$, iniPath$)
        entry$ = Data14
        r% = WritePrivateProfileString("Current", "Dewpoint", entry$, iniPath$)
        entry$ = Data15
        r% = WritePrivateProfileString("Current", "Visibility", entry$, iniPath$)
        entry$ = Data16
        r% = WritePrivateProfileString("Current", "Sunrise", entry$, iniPath$)
        entry$ = Data17
        r% = WritePrivateProfileString("Current", "Sunset", entry$, iniPath$)
        entry$ = Data18
        r% = WritePrivateProfileString("Weekday", "Day1", entry$, iniPath$)
        entry$ = Data19
        r% = WritePrivateProfileString("Weekday", "Day2", entry$, iniPath$)
        entry$ = Data20
        r% = WritePrivateProfileString("Weekday", "Day3", entry$, iniPath$)
        entry$ = Data21
        r% = WritePrivateProfileString("Weekday", "Day4", entry$, iniPath$)
        entry$ = Data22
        r% = WritePrivateProfileString("Weekday", "Day5", entry$, iniPath$)
        entry$ = Data23
        r% = WritePrivateProfileString("Weekday", "Day6", entry$, iniPath$)
        entry$ = Data24
        r% = WritePrivateProfileString("Weekday", "Day7", entry$, iniPath$)
        entry$ = Data25
        r% = WritePrivateProfileString("Weekday", "Day8", entry$, iniPath$)
        entry$ = Data26
        r% = WritePrivateProfileString("Weekday", "Day9", entry$, iniPath$)
        entry$ = Data27
        r% = WritePrivateProfileString("Weekday", "Day10", entry$, iniPath$)
        entry$ = Data28
        r% = WritePrivateProfileString("High Temp", "Day1", entry$, iniPath$)
        entry$ = Data29
        r% = WritePrivateProfileString("High Temp", "Day2", entry$, iniPath$)
        entry$ = Data30
        r% = WritePrivateProfileString("High Temp", "Day3", entry$, iniPath$)
        entry$ = Data31
        r% = WritePrivateProfileString("High Temp", "Day4", entry$, iniPath$)
        entry$ = Data32
        r% = WritePrivateProfileString("High Temp", "Day5", entry$, iniPath$)
        entry$ = Data33
        r% = WritePrivateProfileString("High Temp", "Day6", entry$, iniPath$)
        entry$ = Data34
        r% = WritePrivateProfileString("High Temp", "Day7", entry$, iniPath$)
        entry$ = Data35
        r% = WritePrivateProfileString("High Temp", "Day8", entry$, iniPath$)
        entry$ = Data36
        r% = WritePrivateProfileString("High Temp", "Day9", entry$, iniPath$)
        entry$ = Data37
        r% = WritePrivateProfileString("High Temp", "Day10", entry$, iniPath$)
        entry$ = Data38
        r% = WritePrivateProfileString("Low Temp", "Day1", entry$, iniPath$)
        entry$ = Data39
        r% = WritePrivateProfileString("Low Temp", "Day2", entry$, iniPath$)
        entry$ = Data40
        r% = WritePrivateProfileString("Low Temp", "Day3", entry$, iniPath$)
        entry$ = Data41
        r% = WritePrivateProfileString("Low Temp", "Day4", entry$, iniPath$)
        entry$ = Data42
        r% = WritePrivateProfileString("Low Temp", "Day5", entry$, iniPath$)
        entry$ = Data43
        r% = WritePrivateProfileString("Low Temp", "Day6", entry$, iniPath$)
        entry$ = Data44
        r% = WritePrivateProfileString("Low Temp", "Day7", entry$, iniPath$)
        entry$ = Data45
        r% = WritePrivateProfileString("Low Temp", "Day8", entry$, iniPath$)
        entry$ = Data46
        r% = WritePrivateProfileString("Low Temp", "Day9", entry$, iniPath$)
        entry$ = Data47
        r% = WritePrivateProfileString("Low Temp", "Day10", entry$, iniPath$)
        entry$ = Data48
        r% = WritePrivateProfileString("Weather", "Day1", entry$, iniPath$)
        entry$ = Data49
        r% = WritePrivateProfileString("Weather", "Day2", entry$, iniPath$)
        entry$ = Data50
        r% = WritePrivateProfileString("Weather", "Day3", entry$, iniPath$)
        entry$ = Data51
        r% = WritePrivateProfileString("Weather", "Day4", entry$, iniPath$)
        entry$ = Data52
        r% = WritePrivateProfileString("Weather", "Day5", entry$, iniPath$)
        entry$ = Data53
        r% = WritePrivateProfileString("Weather", "Day6", entry$, iniPath$)
        entry$ = Data54
        r% = WritePrivateProfileString("Weather", "Day7", entry$, iniPath$)
        entry$ = Data55
        r% = WritePrivateProfileString("Weather", "Day8", entry$, iniPath$)
        entry$ = Data56
        r% = WritePrivateProfileString("Weather", "Day9", entry$, iniPath$)
        entry$ = Data57
        r% = WritePrivateProfileString("Weather", "Day10", entry$, iniPath$)
        entry$ = Data58
        r% = WritePrivateProfileString("Dates", "Day1", entry$, iniPath$)
        entry$ = Data59
        r% = WritePrivateProfileString("Dates", "Day2", entry$, iniPath$)
        entry$ = Data60
        r% = WritePrivateProfileString("Dates", "Day3", entry$, iniPath$)
        entry$ = Data61
        r% = WritePrivateProfileString("Dates", "Day4", entry$, iniPath$)
        entry$ = Data62
        r% = WritePrivateProfileString("Dates", "Day5", entry$, iniPath$)
        entry$ = Data63
        r% = WritePrivateProfileString("Dates", "Day6", entry$, iniPath$)
        entry$ = Data64
        r% = WritePrivateProfileString("Dates", "Day7", entry$, iniPath$)
        entry$ = Data65
        r% = WritePrivateProfileString("Dates", "Day8", entry$, iniPath$)
        entry$ = Data66
        r% = WritePrivateProfileString("Dates", "Day9", entry$, iniPath$)
        entry$ = Data67
        r% = WritePrivateProfileString("Dates", "Day10", entry$, iniPath$)
     
End Sub


Private Sub ASP_Click()
'*********************
'Uses the asptear.dll to extract HTML source

        iniPath$ = App.Path + "\weather.dat"
        entry$ = "ASP"
        r% = WritePrivateProfileString("Report", "Method", entry$, iniPath$)

'*******
End Sub
'*******

'***********************************
Private Sub Alpha_city_lst_DblClick()
'***********************************

         ' *********************************** '
         ' ***********************************                                                '
         ' This listbox collects a list of cities alphabetically,  since there is such a vast '
         ' ammount of cities for each state, they need to be broken down this way.            '
        On Error GoTo Weather_Error
        Call Disable_Me
        lstpos = Alpha_city_lst.ListIndex
        If MSNET.Value = True Then
                str_weather = Inet.OpenURL("http://www.weather.com/weather/us/states/" & UCase(str_b) & "-" & Alpha_city_lst.List(lstpos) & ".html")
        End If
        If XML.Value = True Then
                str_data = "http://www.weather.com/weather/us/states/" & UCase(str_b) & "-" & Alpha_city_lst.List(lstpos) & ".html"
                weatherxml.open "GET", str_data, False
                weatherxml.send
                str_weather = weatherxml.responseText
        End If
        If ASP.Value = True Then
                Set xObj = CreateObject("Softwing.aspTear")
                str_weather = xObj.Retrieve("http://www.weather.com/weather/us/states/" & UCase(str_b) & "-" & Alpha_city_lst.List(lstpos) & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
        End If
        us_city_lst.Clear
 
        Do
                str1 = "<A HREF=" & Chr$(34) & "/weather/cities/us_"
                If last_pos = vbNullString Then last_pos = 1
                pos1 = InStr(last_pos, str_weather, str1) + Len(str1)
                If pos1 - Len(str1) = 0 Then Exit Do
                str2 = ".html" & Chr$(34) & ">"
                pos2 = InStr(pos1, str_weather, str2)
                diff = pos2 - pos1
                mainstr = Mid(str_weather, pos1, diff)
                last_pos = pos2
                 ' Formats City, gets rid of spaces and forwards slashes, replaces with "_" '
 
                Do
                        frmt = InStr(mainstr, "_")
                        If frmt <> 0 Then
                                Mid(mainstr, frmt, 1) = " "
                        End If
                Loop Until InStr(mainstr, "_") = 0
 
                str_b = Left$(mainstr, 2)
                mainstr = Right$(mainstr, Len(mainstr) - 3)
                us_city_lst.AddItem UCase(mainstr)
        Loop
 
        Call Enable_Me
 
Weather_Error:
        Exit Sub
        MsgBox "Possible Causes For Error" & Chr$(13) & vbNullString & Chr$(13) & "- Not Connected To Internet" & Chr$(13) & "- No Weather Currently Exists For Location" & Chr$(13) & "- Data Is Corrupt Or Not In Proper Format" & Chr$(13) & vbNullString & Chr$(13) & "* Connect To The Internet" & Chr$(13) & "* Select Another City Within The Same Region" & Chr$(13) & "* Try To Update Later", vbInformation + vbOKOnly, "Weather Error"
        Enable_Me
        StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
        Exit Sub
         ' ******* '

'*******
End Sub
'*******

'*******************
Sub Celsius()
'*******************
                    
        ' Celsius = Degress In Fahrenheit - 32 * (5 / 9) '
        If Temperature.Caption <> "-" Then
                Temperature.Caption = Format((Val(Temperature) - 32) * (5 / 9), "#.0")
        End If
        If Dewpoint.Caption <> "-" Then
                Dewpoint.Caption = Format((Val(Dewpoint) - 32) * (5 / 9), "#.0")
        End If
        If Day1_hi.Caption <> "-" Then
                Day1_hi.Caption = Format((Val(Day1_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day2_hi.Caption <> "-" Then
                Day2_hi.Caption = Format((Val(Day2_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day3_hi.Caption <> "-" Then
                Day3_hi.Caption = Format((Val(Day3_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day4_hi.Caption <> "-" Then
                Day4_hi.Caption = Format((Val(Day4_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day5_hi.Caption <> "-" Then
                Day5_hi.Caption = Format((Val(Day5_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day6_hi.Caption <> "-" Then
                Day6_hi.Caption = Format((Val(Day6_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day7_hi.Caption <> "-" Then
                Day7_hi.Caption = Format((Val(Day7_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day8_hi.Caption <> "-" Then
                Day8_hi.Caption = Format((Val(Day8_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day9_hi.Caption <> "-" Then
                Day9_hi.Caption = Format((Val(Day9_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day10_hi.Caption <> "-" Then
                Day10_hi.Caption = Format((Val(Day10_hi) - 32) * (5 / 9), "#.0")
        End If
        If Day1_lo.Caption <> "-" Then
                Day1_lo.Caption = Format((Val(Day1_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day2_lo.Caption <> "-" Then
                Day2_lo.Caption = Format((Val(Day2_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day3_lo.Caption <> "-" Then
                Day3_lo.Caption = Format((Val(Day3_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day4_lo.Caption <> "-" Then
                Day4_lo.Caption = Format((Val(Day4_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day5_lo.Caption <> "-" Then
                Day5_lo.Caption = Format((Val(Day5_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day6_lo.Caption <> "-" Then
                Day6_lo.Caption = Format((Val(Day6_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day7_lo.Caption <> "-" Then
                Day7_lo.Caption = Format((Val(Day7_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day8_lo.Caption <> "-" Then
                Day8_lo.Caption = Format((Val(Day8_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day9_lo.Caption <> "-" Then
                Day9_lo.Caption = Format((Val(Day9_lo) - 32) * (5 / 9), "#.0")
        End If
        If Day10_lo.Caption <> "-" Then
                Day10_lo.Caption = Format((Val(Day10_lo) - 32) * (5 / 9), "#.0")
        End If
      

'*******
End Sub
'*******

'************************
Private Sub Check1_Click()
'************************

        Call Check_Pos

'*******
End Sub
'*******

'*********************
Sub Check_Pos()
'*********************

         ' Checks to see what size screen the user selected '
        iniPath$ = App.Path + "\weather.dat"
        If Mainfrm.Check1.Value = 0 Then
                If Mainfrm.Height > 4980 Then
                        Mainfrm.Image1.Visible = False
                        Mainfrm.Frame1.Visible = False
                        Mainfrm.Frame2.Visible = False
 
                        For X = 1 To 20
                                Mainfrm.Height = Mainfrm.Height - 102
                                CenterX Me
                                If Mainfrm.Height <= 4980 Then Exit For
                                DoEvents
                        Next X
 
                End If
                entry$ = "0"
                r% = WritePrivateProfileString("Report", "Position", entry$, iniPath$)
                Check1.Value = 0
                CenterX Me
        ElseIf Check1.Value = 1 Then
                If Mainfrm.Height < 7020 Then
 
                        For X = 1 To 20
                                Mainfrm.Height = Mainfrm.Height + 102
                                CenterX Me
                                If Mainfrm.Height >= 7020 Then Exit For
                                DoEvents
                        Next X
 
                End If
                Mainfrm.Image1.Visible = True
                Mainfrm.Frame1.Visible = True
                Mainfrm.Frame2.Visible = True
                entry$ = "1"
                r% = WritePrivateProfileString("Report", "Position", entry$, iniPath$)
                Mainfrm.Check1.Value = 1
                CenterX Me
        End If

'*******
End Sub
'*******

'*********************
Sub Check_Str()
'*********************

         ' If some data is not found, it collects large strings which also fit the criteria. '
         ' This clears the data out that we don't need.                                      '
        If Len(Day1_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day1_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day2_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day2_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day3_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day3_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day4_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day4_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day5_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day5_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day6_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day6_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day7_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day7_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day8_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day8_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day9_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day9_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Day10_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day10_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Sunrise.Caption) > 8 Then Sunrise.Caption = "-"
        If Len(Sunset.Caption) > 8 Then Sunrise.Caption = "-"

'*******
End Sub
'*******

'**********************
Sub Clear_Data()
'**********************

         ' Sets all values to "-" and clears all images
        str_weather = vbNullString
        tmrScroll.Enabled = False
        detail_txt.Visible = False
        fast.Visible = False
        slow.Visible = False
        Report.Caption = "-"
        Conditions.Caption = "-"
        Temperature.Caption = "-"
        Wind.Caption = "-"
        Humidity.Caption = "-"
        Barometer.Caption = "-"
        Dewpoint.Caption = "-"
        Visibility.Caption = "-"
        Sunrise.Caption = "-"
        Sunset.Caption = "-"
        Day1_hi.Caption = "-"
        Day1_lo.Caption = "-"
        Day1_Weather.Caption = "-"
        Day2_hi.Caption = "-"
        Day2_lo.Caption = "-"
        Day2_Weather.Caption = "-"
        Day3_hi.Caption = "-"
        Day3_lo.Caption = "-"
        Day3_Weather.Caption = "-"
        Day4_hi.Caption = "-"
        Day4_lo.Caption = "-"
        Day4_Weather.Caption = "-"
        Day5_hi.Caption = "-"
        Day5_lo.Caption = "-"
        Day5_Weather.Caption = "-"
        Day6_hi.Caption = "-"
        Day6_lo.Caption = "-"
        Day6_Weather.Caption = "-"
        Day7_hi.Caption = "-"
        Day7_lo.Caption = "-"
        Day7_Weather.Caption = "-"
        Day8_hi.Caption = "-"
        Day8_lo.Caption = "-"
        Day8_Weather.Caption = "-"
        Day9_hi.Caption = "-"
        Day9_lo.Caption = "-"
        Day9_Weather.Caption = "-"
        Day10_hi.Caption = "-"
        Day10_lo.Caption = "-"
        Day10_Weather.Caption = "-"
        Day1.Caption = "-"
        Day2.Caption = "-"
        Day3.Caption = "-"
        Day4.Caption = "-"
        Day5.Caption = "-"
        Day6.Caption = "-"
        Day7.Caption = "-"
        Day8.Caption = "-"
        Day9.Caption = "-"
        Day10.Caption = "-"
        Date_1.Caption = "-"
        Date_2.Caption = "-"
        Date_3.Caption = "-"
        Date_4.Caption = "-"
        Date_5.Caption = "-"
        Date_6.Caption = "-"
        Date_7.Caption = "-"
        Date_8.Caption = "-"
        Date_9.Caption = "-"
        Date_10.Caption = "-"
        Set dayi1.Picture = Nothing
        Set dayi2.Picture = Nothing
        Set dayi3.Picture = Nothing
        Set dayi4.Picture = Nothing
        Set dayi5.Picture = Nothing
        Set dayi6.Picture = Nothing
        Set dayi7.Picture = Nothing
        Set dayi8.Picture = Nothing
        Set dayi9.Picture = Nothing
        Set dayi10.Picture = Nothing
        Open (App.Path & "\detail.dat") For Output As #1
        Print #1, vbNullString
        Close #1

'*******
End Sub
'*******

'********************************
Private Sub Country_lst_DblClick()
'********************************

         ' The main function of this routine is to determine the cities for the selection. '
        On Error GoTo Weather_Error
        Call Disable_Me
        get_city = Country_lst.ListIndex
        Int_city_lst.Clear
        name_x = Country_lst.List(get_city)
         ' Formats City, gets rid of spaces and forwards slashes, replaces with "_" '
 
        Do
                frmt = InStr(name_x, " ")
                If frmt <> 0 Then
                        Mid(name_x, frmt, 1) = "_"
                End If
        Loop Until InStr(name_x, " ") = 0
 
 
        Do
                frmt = InStr(name_x, "/")
                If frmt <> 0 Then
                        Mid(name_x, frmt, 1) = "_"
                End If
        Loop Until InStr(name_x, "/") = 0
 
        If MSNET.Value = True Then
                str_weather = Inet.OpenURL("http://www.weather.com/ins/countries_index/" & name_x & ".html")
        End If
        If XML.Value = True Then
                str_data = "http://www.weather.com/ins/countries_index/" & name_x & ".html"
                weatherxml.open "GET", str_data, False
                weatherxml.send
                str_weather = weatherxml.responseText
        End If
        If ASP.Value = True Then
                Set xObj = CreateObject("Softwing.aspTear")
                str_weather = xObj.Retrieve("http://www.weather.com/ins/countries_index/" & name_x & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
        End If
 
        Do
                str1 = "<A HREF=" & Chr$(34) & "/weather/cities/"
                If last_pos = vbNullString Then last_pos = 1
                pos1 = InStr(last_pos, str_weather, str1) + Len(str1)
                If pos1 - Len(str1) = 0 Then Exit Do
                str2 = ".html" & Chr$(34) & ">"
                pos2 = InStr(pos1, str_weather, str2)
                diff = pos2 - pos1
                mainstr = Mid(str_weather, pos1, diff)
                last_pos = pos2
                 ' Formats City, gets rid of spaces and forwards slashes, replaces with "_" '
 
                Do
                        frmt = InStr(mainstr, "_")
                        If frmt <> 0 Then
                                Mid(mainstr, frmt, 1) = " "
                        End If
                Loop Until InStr(mainstr, "_") = 0
 
                str_a = Left$(mainstr, 2)
                mainstr = Right$(mainstr, Len(mainstr) - 4)
                If name_x = "Virgin_Islands" Then mainstr = Right$(mainstr, Len(mainstr) - 2)
                Int_city_lst.AddItem UCase(mainstr)
        Loop
 
        Call Enable_Me
 
Weather_Error:
        Exit Sub
        MsgBox "Possible Causes For Error" & Chr$(13) & vbNullString & Chr$(13) & "- Not Connected To Internet" & Chr$(13) & "- No Weather Currently Exists For Location" & Chr$(13) & "- Data Is Corrupt Or Not In Proper Format" & Chr$(13) & vbNullString & Chr$(13) & "* Connect To The Internet" & Chr$(13) & "* Select Another City Within The Same Region" & Chr$(13) & "* Try To Update Later", vbInformation + vbOKOnly, "Weather Error"
        Enable_Me
        StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
        Exit Sub

'*******
End Sub
'*******

' *********************** '
Sub Detail_List()
' *********************** '

    On Error GoTo Weather_Error
    temp_string = "<B>National Weather Service Local Forecast</B><BR>"
    error_string = "Forecast temporarily unavailable."
    If InStr(str_local, error_string) <> 0 Then GoTo Weather_Error
    pos1 = InStr(str_local, temp_string)
    pos2 = pos1 + Len(temp_string)
    pos3 = InStr(pos2, str_local, "<BR><B>") + 7
    pos4 = InStr(pos3, str_local, "</FONT>")
    pos4a = InStr(pos3, str_local, "<BR><B>Extended forecast:")
    If pos4 < pos4a Then
        tempo = Trim$(Mid(str_local, pos3, (pos4 - pos3)))
    Else
        tempo = Trim$(Mid(str_local, pos3, (pos4a - pos3)))
    End If
    temp_string2 = "<BR><B>Extended forecast:"
    pos5 = InStr(pos3, str_local, temp_string2)
    pos6 = InStr(pos5, str_local, "<BR><B>") + 7
    pos7 = InStr(pos6, str_local, "<BR CLEAR=")
    tempo2 = Trim$(Mid(str_local, pos6, (pos7 - pos6)))
    Do
        frmt = InStr(tempo, "<BR CLEAR=" & Chr$(34) & "left$" & Chr$(34) & ">")
        If frmt <> 0 Then
            Mid(tempo, frmt, 17) = "                 "
        End If
    Loop Until InStr(tempo, "<BR CLEAR=" & Chr$(34) & "left$" & Chr$(34) & ">") = 0
    Do
        frmt = InStr(tempo, "<B>")
        If frmt <> 0 Then
            Mid(tempo, frmt, 3) = "   "
        End If
    Loop Until InStr(tempo, "<B>") = 0
    Do
        frmt = InStr(tempo, "<BR>")
        If frmt <> 0 Then
            Mid(tempo, frmt, 4) = "    "
        End If
    Loop Until InStr(tempo, "<BR>") = 0
    Do
        frmt = InStr(tempo, "</B>")
        If frmt <> 0 Then
            Mid(tempo, frmt, 4) = "     "
        End If
    Loop Until InStr(tempo, "</B>") = 0
    For X = 1 To Len(tempo)
        str_s = Mid(tempo, X, 1)
        If Asc(str_s) <> 10 And Asc(str_s) <> 13 Then
            tempox = tempox + str_s
        End If
    Next X
    Do
        s1 = InStr(tempox, "   ")
        If s1 <> 0 Then
            tempox = Left$(tempox, s1) + Right$(tempox, Len(tempox) - (s1 + 1))
        End If
    Loop Until s1 = 0
    Do
        frmt = InStr(tempo2, "</FONT>")
        If frmt <> 0 Then
            Mid(tempo2, frmt, 7) = "       "
        End If
    Loop Until InStr(tempo2, "</FONT>") = 0
    Do
        frmt = InStr(tempo2, "<B>")
        If frmt <> 0 Then
            Mid(tempo2, frmt, 3) = "   "
        End If
    Loop Until InStr(tempo2, "<B>") = 0
    Do
        frmt = InStr(tempo2, "<BR>")
        If frmt <> 0 Then
            Mid(tempo2, frmt, 4) = "    "
        End If
    Loop Until InStr(tempo2, "<BR>") = 0
    Do
        frmt = InStr(tempo2, "</B>")
        If frmt <> 0 Then
            Mid(tempo2, frmt, 4) = "     "
        End If
    Loop Until InStr(tempo2, "</B>") = 0
    For X = 1 To Len(tempo2)
        str_s = Mid(tempo2, X, 1)
        If Asc(str_s) <> 10 And Asc(str_s) <> 13 Then
            tempox2 = tempox2 + str_s
        End If
    Next X
    Do
        s1 = InStr(tempox2, "   ")
        If s1 <> 0 Then
            tempox2 = Left$(tempox2, s1) + Right$(tempox2, Len(tempox2) - (s1 + 1))
        End If
    Loop Until s1 = 0
    Open (App.Path & "\detail.dat") For Output As #1
    Print #1, "         " & Trim$(tempox) & " " & Trim$(tempox2)
    Close #1
    If Systemtrayfrm.scroll.Checked = True Then
        Open (App.Path & "\detail.dat") For Input As #1
        detail_txt.Text = Input(LOF(1), 1)
        Close #1
        tmrScroll.Enabled = True
        detail_txt.Visible = True
        fast.Visible = True
        slow.Visible = True
    End If
    Exit Sub
Weather_Error:
    tempox = "No Detailed Forecast Data Currently Available For Location"
    Open (App.Path & "\detail.dat") For Output As #1
    Print #1, "         " & tempox & " " & Trim$(tempox2)
    Close #1
    If Systemtrayfrm.scroll.Checked = True Then
        Open (App.Path & "\detail.dat") For Input As #1
        detail_txt.Text = Input(LOF(1), 1)
        Close #1
        tmrScroll.Enabled = True
    End If
    Exit Sub

' ******* '
End Sub
' ******* '


'**********************
Sub Disable_Me()
'**********************

         ' Disables all functions while loading to keep people from     '
         ' trying to access 2 functions at a time and causing an error  '
        StatusBar1.Panels.Item(1).Text = "Loading Data - Please Wait..."
        Mainfrm.Enabled = False

'*******
End Sub
'*******

'*********************
Sub Enable_Me()
'*********************

         ' Enables all functions '
        StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
        Mainfrm.Enabled = True

'*******
End Sub
'*******

'**********************
Sub Fahrenheit()
'**********************

         ' Instead of converting Celsius or Kelvin to Fahrenheit, we just go out to the
         ' weather.dat file and get original data that was saved previously
        Temperature.Caption = GetFromINI("Current", "Temperature", iniPath$)
        Dewpoint.Caption = GetFromINI("Current", "Dewpoint", iniPath$)
        Day1_hi.Caption = GetFromINI("High Temp", "Day1", iniPath$)
        Day2_hi.Caption = GetFromINI("High Temp", "Day2", iniPath$)
        Day3_hi.Caption = GetFromINI("High Temp", "Day3", iniPath$)
        Day4_hi.Caption = GetFromINI("High Temp", "Day4", iniPath$)
        Day5_hi.Caption = GetFromINI("High Temp", "Day5", iniPath$)
        Day6_hi.Caption = GetFromINI("High Temp", "Day6", iniPath$)
        Day7_hi.Caption = GetFromINI("High Temp", "Day7", iniPath$)
        Day8_hi.Caption = GetFromINI("High Temp", "Day8", iniPath$)
        Day9_hi.Caption = GetFromINI("High Temp", "Day9", iniPath$)
        Day10_hi.Caption = GetFromINI("High Temp", "Day10", iniPath$)
        Day1_lo.Caption = GetFromINI("Low Temp", "Day1", iniPath$)
        Day2_lo.Caption = GetFromINI("Low Temp", "Day2", iniPath$)
        Day3_lo.Caption = GetFromINI("Low Temp", "Day3", iniPath$)
        Day4_lo.Caption = GetFromINI("Low Temp", "Day4", iniPath$)
        Day5_lo.Caption = GetFromINI("Low Temp", "Day5", iniPath$)
        Day6_lo.Caption = GetFromINI("Low Temp", "Day6", iniPath$)
        Day7_lo.Caption = GetFromINI("Low Temp", "Day7", iniPath$)
        Day8_lo.Caption = GetFromINI("Low Temp", "Day8", iniPath$)
        Day9_lo.Caption = GetFromINI("Low Temp", "Day9", iniPath$)
        Day10_lo.Caption = GetFromINI("Low Temp", "Day10", iniPath$)

'*******
End Sub
'*******

' ********************* '
Private Sub Form_Load()
' ********************* '
    
    If App.PrevInstance Then       ' Checks to see if application is already running '
        End
    End If
    Mainfrm.MousePointer = 99
    Mainfrm.MouseIcon = Mainfrm.ImageList1.ListImages(1).Picture   ' Sets Mousepointer '
    Mainfrm.detail_txt.Locked = True
    Call Mainfrm.Load_Data

' ******* '
End Sub
' ******* '


'********************
Sub Get_Icon()
'********************

         ' Loads icons for weather conditions
        X1 = LCase(Trim$(Day1_Weather.Caption))
        X2 = LCase(Trim$(Day2_Weather.Caption))
        X3 = LCase(Trim$(Day3_Weather.Caption))
        X4 = LCase(Trim$(Day4_Weather.Caption))
        X5 = LCase(Trim$(Day5_Weather.Caption))
        X6 = LCase(Trim$(Day6_Weather.Caption))
        X7 = LCase(Trim$(Day7_Weather.Caption))
        X8 = LCase(Trim$(Day8_Weather.Caption))
        X9 = LCase(Trim$(Day9_Weather.Caption))
        X10 = LCase(Trim$(Day10_Weather.Caption))
        If X1 = "t-storms" Then dayi1.Picture = ImageList1.ListImages(8).Picture
        If X2 = "t-storms" Then dayi2.Picture = ImageList1.ListImages(8).Picture
        If X3 = "t-storms" Then dayi3.Picture = ImageList1.ListImages(8).Picture
        If X4 = "t-storms" Then dayi4.Picture = ImageList1.ListImages(8).Picture
        If X5 = "t-storms" Then dayi5.Picture = ImageList1.ListImages(8).Picture
        If X6 = "t-storms" Then dayi6.Picture = ImageList1.ListImages(8).Picture
        If X7 = "t-storms" Then dayi7.Picture = ImageList1.ListImages(8).Picture
        If X8 = "t-storms" Then dayi8.Picture = ImageList1.ListImages(8).Picture
        If X9 = "t-storms" Then dayi9.Picture = ImageList1.ListImages(8).Picture
        If X10 = "t-storms" Then dayi10.Picture = ImageList1.ListImages(8).Picture
        If X1 = "cloudy" Or X1 = "cloudy/ windy" Then dayi1.Picture = ImageList1.ListImages(5).Picture
        If X2 = "cloudy" Or X2 = "cloudy/ windy" Then dayi2.Picture = ImageList1.ListImages(5).Picture
        If X3 = "cloudy" Or X3 = "cloudy/ windy" Then dayi3.Picture = ImageList1.ListImages(5).Picture
        If X4 = "cloudy" Or X4 = "cloudy/ windy" Then dayi4.Picture = ImageList1.ListImages(5).Picture
        If X5 = "cloudy" Or X5 = "cloudy/ windy" Then dayi5.Picture = ImageList1.ListImages(5).Picture
        If X6 = "cloudy" Or X6 = "cloudy/ windy" Then dayi6.Picture = ImageList1.ListImages(5).Picture
        If X7 = "cloudy" Or X7 = "cloudy/ windy" Then dayi7.Picture = ImageList1.ListImages(5).Picture
        If X8 = "cloudy" Or X8 = "cloudy/ windy" Then dayi8.Picture = ImageList1.ListImages(5).Picture
        If X9 = "cloudy" Or X9 = "cloudy/ windy" Then dayi9.Picture = ImageList1.ListImages(5).Picture
        If X10 = "cloudy" Or X10 = "cloudy/ windy" Then dayi10.Picture = ImageList1.ListImages(5).Picture
        If X1 = "sunny" Then dayi1.Picture = ImageList1.ListImages(2).Picture
        If X2 = "sunny" Then dayi2.Picture = ImageList1.ListImages(2).Picture
        If X3 = "sunny" Then dayi3.Picture = ImageList1.ListImages(2).Picture
        If X4 = "sunny" Then dayi4.Picture = ImageList1.ListImages(2).Picture
        If X5 = "sunny" Then dayi5.Picture = ImageList1.ListImages(2).Picture
        If X6 = "sunny" Then dayi6.Picture = ImageList1.ListImages(2).Picture
        If X7 = "sunny" Then dayi7.Picture = ImageList1.ListImages(2).Picture
        If X8 = "sunny" Then dayi8.Picture = ImageList1.ListImages(2).Picture
        If X9 = "sunny" Then dayi9.Picture = ImageList1.ListImages(2).Picture
        If X10 = "sunny" Then dayi10.Picture = ImageList1.ListImages(2).Picture
        If X1 = "mostly sunny" Or X1 = "fair" Then dayi1.Picture = ImageList1.ListImages(3).Picture
        If X2 = "mostly sunny" Or X2 = "fair" Then dayi2.Picture = ImageList1.ListImages(3).Picture
        If X3 = "mostly sunny" Or X3 = "fair" Then dayi3.Picture = ImageList1.ListImages(3).Picture
        If X4 = "mostly sunny" Or X4 = "fair" Then dayi4.Picture = ImageList1.ListImages(3).Picture
        If X5 = "mostly sunny" Or X5 = "fair" Then dayi5.Picture = ImageList1.ListImages(3).Picture
        If X6 = "mostly sunny" Or X6 = "fair" Then dayi6.Picture = ImageList1.ListImages(3).Picture
        If X7 = "mostly sunny" Or X7 = "fair" Then dayi7.Picture = ImageList1.ListImages(3).Picture
        If X8 = "mostly sunny" Or X8 = "fair" Then dayi8.Picture = ImageList1.ListImages(3).Picture
        If X9 = "mostly sunny" Or X9 = "fair" Then dayi9.Picture = ImageList1.ListImages(3).Picture
        If X10 = "mostly sunny" Or X10 = "Fair" Then dayi10.Picture = ImageList1.ListImages(3).Picture
        If X1 = "partly cloudy" Then dayi1.Picture = ImageList1.ListImages(4).Picture
        If X2 = "partly cloudy" Then dayi2.Picture = ImageList1.ListImages(4).Picture
        If X3 = "partly cloudy" Then dayi3.Picture = ImageList1.ListImages(4).Picture
        If X4 = "partly cloudy" Then dayi4.Picture = ImageList1.ListImages(4).Picture
        If X5 = "partly cloudy" Then dayi5.Picture = ImageList1.ListImages(4).Picture
        If X6 = "partly cloudy" Then dayi6.Picture = ImageList1.ListImages(4).Picture
        If X7 = "partly cloudy" Then dayi7.Picture = ImageList1.ListImages(4).Picture
        If X8 = "partly cloudy" Then dayi8.Picture = ImageList1.ListImages(4).Picture
        If X9 = "partly cloudy" Then dayi9.Picture = ImageList1.ListImages(4).Picture
        If X10 = "partly cloudy" Then dayi10.Picture = ImageList1.ListImages(4).Picture
        If X1 = "rain" Or X1 = "windy/ rain" Or X1 = "windy/ showers" Or X1 = "showers" Or X1 = "breezy/ rain" Or X1 = "breezy/ showers" Then dayi1.Picture = ImageList1.ListImages(6).Picture
        If X2 = "rain" Or X2 = "windy/ rain" Or X2 = "windy/ showers" Or X2 = "showers" Or X2 = "breezy/ rain" Or X2 = "breezy/ showers" Then dayi2.Picture = ImageList1.ListImages(6).Picture
        If X3 = "rain" Or X3 = "windy/ rain" Or X3 = "windy/ showers" Or X3 = "showers" Or X3 = "breezy/ rain" Or X3 = "breezy/ showers" Then dayi3.Picture = ImageList1.ListImages(6).Picture
        If X4 = "rain" Or X4 = "windy/ rain" Or X4 = "windy/ showers" Or X4 = "showers" Or X4 = "breezy/ rain" Or X4 = "breezy/ showers" Then dayi4.Picture = ImageList1.ListImages(6).Picture
        If X5 = "rain" Or X5 = "windy/ rain" Or X5 = "windy/ showers" Or X5 = "showers" Or X5 = "breezy/ rain" Or X5 = "breezy/ showers" Then dayi5.Picture = ImageList1.ListImages(6).Picture
        If X6 = "rain" Or X6 = "windy/ rain" Or X6 = "windy/ showers" Or X6 = "showers" Or X6 = "breezy/ rain" Or X6 = "breezy/ showers" Then dayi6.Picture = ImageList1.ListImages(6).Picture
        If X7 = "rain" Or X7 = "windy/ rain" Or X7 = "windy/ showers" Or X7 = "showers" Or X7 = "breezy/ rain" Or X7 = "breezy/ showers" Then dayi7.Picture = ImageList1.ListImages(6).Picture
        If X8 = "rain" Or X8 = "windy/ rain" Or X8 = "windy/ showers" Or X8 = "showers" Or X8 = "breezy/ rain" Or X8 = "breezy/ showers" Then dayi8.Picture = ImageList1.ListImages(6).Picture
        If X9 = "rain" Or X9 = "windy/ rain" Or X9 = "windy/ showers" Or X9 = "showers" Or X9 = "breezy/ rain" Or X9 = "breezy/ showers" Then dayi9.Picture = ImageList1.ListImages(6).Picture
        If X10 = "rain" Or X10 = "windy/ rain" Or X10 = "windy/ showers" Or X10 = "showers" Or X10 = "breezy/ rain" Or X10 = "breezy/ showers" Then dayi10.Picture = ImageList1.ListImages(6).Picture
        If X1 = "isolated t-storms" Or X1 = "scattered t-storms" Then dayi1.Picture = ImageList1.ListImages(10).Picture
        If X2 = "isolated t-storms" Or X2 = "scattered t-storms" Then dayi2.Picture = ImageList1.ListImages(10).Picture
        If X3 = "isolated t-storms" Or X3 = "scattered t-storms" Then dayi3.Picture = ImageList1.ListImages(10).Picture
        If X4 = "isolated t-storms" Or X4 = "scattered t-storms" Then dayi4.Picture = ImageList1.ListImages(10).Picture
        If X5 = "isolated t-storms" Or X5 = "scattered t-storms" Then dayi5.Picture = ImageList1.ListImages(10).Picture
        If X6 = "isolated t-storms" Or X6 = "scattered t-storms" Then dayi6.Picture = ImageList1.ListImages(10).Picture
        If X7 = "isolated t-storms" Or X7 = "scattered t-storms" Then dayi7.Picture = ImageList1.ListImages(10).Picture
        If X8 = "isolated t-storms" Or X8 = "scattered t-storms" Then dayi8.Picture = ImageList1.ListImages(10).Picture
        If X9 = "isolated t-storms" Or X9 = "scattered t-storms" Then dayi9.Picture = ImageList1.ListImages(10).Picture
        If X10 = "isolated t-storms" Or X10 = "scattered t-storms" Then dayi10.Picture = ImageList1.ListImages(10).Picture
        If X1 = "mostly cloudy" Then dayi1.Picture = ImageList1.ListImages(20).Picture
        If X2 = "mostly cloudy" Then dayi2.Picture = ImageList1.ListImages(20).Picture
        If X3 = "mostly cloudy" Then dayi3.Picture = ImageList1.ListImages(20).Picture
        If X4 = "mostly cloudy" Then dayi4.Picture = ImageList1.ListImages(20).Picture
        If X5 = "mostly cloudy" Then dayi5.Picture = ImageList1.ListImages(20).Picture
        If X6 = "mostly cloudy" Then dayi6.Picture = ImageList1.ListImages(20).Picture
        If X7 = "mostly cloudy" Then dayi7.Picture = ImageList1.ListImages(20).Picture
        If X8 = "mostly cloudy" Then dayi8.Picture = ImageList1.ListImages(20).Picture
        If X9 = "mostly cloudy" Then dayi9.Picture = ImageList1.ListImages(20).Picture
        If X10 = "mostly cloudy" Then dayi10.Picture = ImageList1.ListImages(20).Picture
        If X1 = "rain and snow" Or X1 = "rain/ snow/ breezy" Or X1 = "mixed rain and snow" Then dayi1.Picture = ImageList1.ListImages(21).Picture
        If X2 = "rain and snow" Or X2 = "rain/ snow/ breezy" Or X2 = "mixed rain and snow" Then dayi2.Picture = ImageList1.ListImages(21).Picture
        If X3 = "rain and snow" Or X3 = "rain/ snow/ breezy" Or X3 = "mixed rain and snow" Then dayi3.Picture = ImageList1.ListImages(21).Picture
        If X4 = "rain and snow" Or X4 = "rain/ snow/ breezy" Or X4 = "mixed rain and snow" Then dayi4.Picture = ImageList1.ListImages(21).Picture
        If X5 = "rain and snow" Or X5 = "rain/ snow/ breezy" Or X5 = "mixed rain and snow" Then dayi5.Picture = ImageList1.ListImages(21).Picture
        If X6 = "rain and snow" Or X6 = "rain/ snow/ breezy" Or X6 = "mixed rain and snow" Then dayi6.Picture = ImageList1.ListImages(21).Picture
        If X7 = "rain and snow" Or X7 = "rain/ snow/ breezy" Or X7 = "mixed rain and snow" Then dayi7.Picture = ImageList1.ListImages(21).Picture
        If X8 = "rain and snow" Or X8 = "rain/ snow/ breezy" Or X8 = "mixed rain and snow" Then dayi8.Picture = ImageList1.ListImages(21).Picture
        If X9 = "rain and snow" Or X9 = "rain/ snow/ breezy" Or X9 = "mixed rain and snow" Then dayi9.Picture = ImageList1.ListImages(21).Picture
        If X10 = "rain and snow" Or X10 = "rain/ snow/ breezy" Or X10 = "mixed rain and snow" Then dayi10.Picture = ImageList1.ListImages(21).Picture
        If X1 = "scattered showers" Or X1 = "breezy/ scattered showers" Then dayi1.Picture = ImageList1.ListImages(7).Picture
        If X2 = "scattered showers" Or X2 = "breezy/ scattered showers" Then dayi2.Picture = ImageList1.ListImages(7).Picture
        If X3 = "scattered showers" Or X3 = "breezy/ scattered showers" Then dayi3.Picture = ImageList1.ListImages(7).Picture
        If X4 = "scattered showers" Or X4 = "breezy/ scattered showers" Then dayi4.Picture = ImageList1.ListImages(7).Picture
        If X5 = "scattered showers" Or X5 = "breezy/ scattered showers" Then dayi5.Picture = ImageList1.ListImages(7).Picture
        If X6 = "scattered showers" Or X6 = "breezy/ scattered showers" Then dayi6.Picture = ImageList1.ListImages(7).Picture
        If X7 = "scattered showers" Or X7 = "breezy/ scattered showers" Then dayi7.Picture = ImageList1.ListImages(7).Picture
        If X8 = "scattered showers" Or X8 = "breezy/ scattered showers" Then dayi8.Picture = ImageList1.ListImages(7).Picture
        If X9 = "scattered showers" Or X9 = "breezy/ scattered showers" Then dayi9.Picture = ImageList1.ListImages(7).Picture
        If X10 = "scattered showers" Or X10 = "breezy/ scattered showers" Then dayi10.Picture = ImageList1.ListImages(7).Picture
        If X1 = "snow" Or X1 = "snow showers" Or X1 = "windy/ snow showers" Or X1 = "breezy/ snow showers" Or X1 = "breezy/ snow" Then dayi1.Picture = ImageList1.ListImages(17).Picture
        If X2 = "snow" Or X2 = "snow showers" Or X2 = "windy/ snow showers" Or X2 = "breezy/ snow showers" Or X2 = "breezy/ snow" Then dayi2.Picture = ImageList1.ListImages(17).Picture
        If X3 = "snow" Or X3 = "snow showers" Or X3 = "windy/ snow showers" Or X3 = "breezy/ snow showers" Or X3 = "breezy/ snow" Then dayi3.Picture = ImageList1.ListImages(17).Picture
        If X4 = "snow" Or X4 = "snow showers" Or X4 = "windy/ snow showers" Or X4 = "breezy/ snow showers" Or X4 = "breezy/ snow" Then dayi4.Picture = ImageList1.ListImages(17).Picture
        If X5 = "snow" Or X5 = "snow showers" Or X5 = "windy/ snow showers" Or X5 = "breezy/ snow showers" Or X5 = "breezy/ snow" Then dayi5.Picture = ImageList1.ListImages(17).Picture
        If X6 = "snow" Or X6 = "snow showers" Or X6 = "windy/ snow showers" Or X6 = "breezy/ snow showers" Or X6 = "breezy/ snow" Then dayi6.Picture = ImageList1.ListImages(17).Picture
        If X7 = "snow" Or X7 = "snow showers" Or X7 = "windy/ snow showers" Or X7 = "breezy/ snow showers" Or X7 = "breezy/ snow" Then dayi7.Picture = ImageList1.ListImages(17).Picture
        If X8 = "snow" Or X8 = "snow showers" Or X8 = "windy/ snow showers" Or X8 = "breezy/ snow showers" Or X8 = "breezy/ snow" Then dayi8.Picture = ImageList1.ListImages(17).Picture
        If X9 = "snow" Or X9 = "snow showers" Or X9 = "windy/ snow showers" Or X9 = "breezy/ snow showers" Or X9 = "breezy/ snow" Then dayi9.Picture = ImageList1.ListImages(17).Picture
        If X10 = "snow" Or X10 = "snow showers" Or X10 = "windy/ snow showers" Or X10 = "breezy/ snow showers" Or X10 = "breezy/ snow" Then dayi10.Picture = ImageList1.ListImages(17).Picture
        If X1 = "scattered snow showers" Or X1 = "breezy/ scattered snow showers" Then dayi1.Picture = ImageList1.ListImages(14).Picture
        If X2 = "scattered snow showers" Or X2 = "breezy/ scattered snow showers" Then dayi2.Picture = ImageList1.ListImages(14).Picture
        If X3 = "scattered snow showers" Or X3 = "breezy/ scattered snow showers" Then dayi3.Picture = ImageList1.ListImages(14).Picture
        If X4 = "scattered snow showers" Or X4 = "breezy/ scattered snow showers" Then dayi4.Picture = ImageList1.ListImages(14).Picture
        If X5 = "scattered snow showers" Or X5 = "breezy/ scattered snow showers" Then dayi5.Picture = ImageList1.ListImages(14).Picture
        If X6 = "scattered snow showers" Or X6 = "breezy/ scattered snow showers" Then dayi6.Picture = ImageList1.ListImages(14).Picture
        If X7 = "scattered snow showers" Or X7 = "breezy/ scattered snow showers" Then dayi7.Picture = ImageList1.ListImages(14).Picture
        If X8 = "scattered snow showers" Or X8 = "breezy/ scattered snow showers" Then dayi8.Picture = ImageList1.ListImages(14).Picture
        If X9 = "scattered snow showers" Or X9 = "breezy/ scattered snow showers" Then dayi9.Picture = ImageList1.ListImages(14).Picture
        If X10 = "scattered snow showers" Or x19 = "breezy/ scattered snow showers" Then dayi10.Picture = ImageList1.ListImages(14).Picture
        If X1 = "wintry mix" Or X1 = "freezing rain" Then dayi1.Picture = ImageList1.ListImages(16).Picture
        If X2 = "wintry mix" Or X2 = "freezing rain" Then dayi2.Picture = ImageList1.ListImages(16).Picture
        If X3 = "wintry mix" Or X3 = "freezing rain" Then dayi3.Picture = ImageList1.ListImages(16).Picture
        If X4 = "wintry mix" Or X4 = "freezing rain" Then dayi4.Picture = ImageList1.ListImages(16).Picture
        If X5 = "wintry mix" Or X5 = "freezing rain" Then dayi5.Picture = ImageList1.ListImages(16).Picture
        If X6 = "wintry mix" Or X6 = "freezing rain" Then dayi6.Picture = ImageList1.ListImages(16).Picture
        If X7 = "wintry mix" Or X7 = "freezing rain" Then dayi7.Picture = ImageList1.ListImages(16).Picture
        If X8 = "wintry mix" Or X8 = "freezing rain" Then dayi8.Picture = ImageList1.ListImages(16).Picture
        If X9 = "wintry mix" Or X9 = "freezing rain" Then dayi9.Picture = ImageList1.ListImages(16).Picture
        If X10 = "wintry mix" Or X10 = "freezing rain" Then dayi10.Picture = ImageList1.ListImages(16).Picture
        If X1 = "windy" Then dayi1.Picture = ImageList1.ListImages(26).Picture
        If X2 = "windy" Then dayi2.Picture = ImageList1.ListImages(26).Picture
        If X3 = "windy" Then dayi3.Picture = ImageList1.ListImages(26).Picture
        If X4 = "windy" Then dayi4.Picture = ImageList1.ListImages(26).Picture
        If X5 = "windy" Then dayi5.Picture = ImageList1.ListImages(26).Picture
        If X6 = "windy" Then dayi6.Picture = ImageList1.ListImages(26).Picture
        If X7 = "windy" Then dayi7.Picture = ImageList1.ListImages(26).Picture
        If X8 = "windy" Then dayi8.Picture = ImageList1.ListImages(26).Picture
        If X9 = "windy" Then dayi9.Picture = ImageList1.ListImages(26).Picture
        If X10 = "windy" Then dayi10.Picture = ImageList1.ListImages(26).Picture
        If LCase(Day1.Caption) = "tonight" Then
                If X1 = "clear" Then dayi1.Picture = ImageList1.ListImages(18).Picture
                If X1 = "fair" Then dayi1.Picture = ImageList1.ListImages(19).Picture
                If X1 = "snow" Or X1 = "snow showers" Or X1 = "breezy/ snow showers" Or X1 = "breezy/ snow" Then dayi1.Picture = ImageList1.ListImages(24).Picture
                If X1 = "mostly cloudy" Or X1 = "cloudy" Then dayi1.Picture = ImageList1.ListImages(11).Picture
                If X1 = "partly cloudy" Then dayi1.Picture = ImageList1.ListImages(12).Picture
                If X1 = "rain" Or X1 = "showers" Or X1 = "breezy/ rain" Or X1 = "breezy/ showers" Then dayi1.Picture = ImageList1.ListImages(13).Picture
                If X1 = "t-storms" Then dayi1.Picture = ImageList1.ListImages(24).Picture
        End If
        If X1 = "(no report)" Then
                dayi1.Picture = ImageList1.ListImages(25).Picture
                Day1_hi.Caption = "-"
                Day1_lo.Caption = "-"
        End If
        If X2 = "(no report)" Then
                dayi2.Picture = ImageList1.ListImages(25).Picture
                Day2_hi.Caption = "-"
                Day2_lo.Caption = "-"
        End If
        If X3 = "(no report)" Then
                dayi3.Picture = ImageList1.ListImages(25).Picture
                Day3_hi.Caption = "-"
                Day3_lo.Caption = "-"
        End If
        If X4 = "(no report)" Then
                dayi4.Picture = ImageList1.ListImages(25).Picture
                Day4_hi.Caption = "-"
                Day4_lo.Caption = "-"
        End If
        If X5 = "(no report)" Then
                dayi5.Picture = ImageList1.ListImages(25).Picture
                Day5_hi.Caption = "-"
                Day5_lo.Caption = "-"
        End If
        If X6 = "(no report)" Then
                dayi6.Picture = ImageList1.ListImages(25).Picture
                Day6_hi.Caption = "-"
                Day6_lo.Caption = "-"
        End If
        If X7 = "(no report)" Then
                dayi7.Picture = ImageList1.ListImages(25).Picture
                Day7_hi.Caption = "-"
                Day7_lo.Caption = "-"
        End If
        If X8 = "(no report)" Then
                dayi8.Picture = ImageList1.ListImages(25).Picture
                Day8_hi.Caption = "-"
                Day8_lo.Caption = "-"
        End If
        If X9 = "(no report)" Then
                dayi9.Picture = ImageList1.ListImages(25).Picture
                Day9_hi.Caption = "-"
                Day9_lo.Caption = "-"
        End If
        If X10 = "(no report)" Then
                dayi10.Picture = ImageList1.ListImages(25).Picture
                Day10_hi.Caption = "-"
                Day10_lo.Caption = "-"
        End If

'*******
End Sub
'*******

'*******************
Sub History()
'*******************

         ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         ' Your Weather v4.0                                                           '
         ' By: Nathan Snyder                                                           '
         ' Programmed in Visual Basic 5.0 Enterprise                                   '
         '                                                                             '
         '                                                                             '
         '                                                                             '
         ' This is by far the best weather program out there..  Sounds conceited I know.'
         ' Take a look at the features and you will see what I mean.  This isn't one of '
         ' those enter a zip code and get a few minor details about your weather, type  '
         ' programs - this gives you everything.  This gives you a 10 day forecast,     '
         ' graphics to indicate conditions, current conditions, international city      '
         ' weather, US City weather, and weather radar maps for almost all regions of   '
         ' the world.  I spent a long time on this.  Apparently this program was good   '
         ' enough to spark a few comments from Weather.com telling me that they didn't  '
         ' like me distributing this as it could potentially dramatically decrease the  '
         ' ammount of traffic that passes through their site.  It takes on average 8    '
         ' seconds to retrieve weather information through my program and can take up to'
         ' 5 minutes just to find what you want on their site.  This cuts out the middle'
         ' man and gives you what you want.  Just make sure you're connected to the     '
         ' internet and you're all set.  Best of all it 's FREE.                        '
         '                                                                             '
         '                                                                             '
         ' Your Weather                                                                '
         ' Designed & Programmed                                                       '
         '  _____    _____                 BY                    _____    _____        '
         ' /     \  /     \                                     /     \  /     \       '
         ' |      \ |     | _____  __________  ___    ___  _____|      \ |     |       '
         ' |       \|     |/     \/          \/   \  /   \/     \       \|     |       '
         ' |        \     |    o  \___     __/\   /  \   /    o  \       \     |       '
         ' |              |        \  |   |    |  \__/  |         \            |       '
         ' |     |\       |    _    \  \   \   |   ___  |     _    \  |\       |       '
         ' |     | \      |   / \    \  \   \  /   \ /   \   / \    \ | \      |       '
         ' \_____/  \_____/__/   \___/   \__/  \___/ \___/__/   \___/_/  \_____/       '
         '                                                                             '
         '                                                                             '
         ' If you use this with great frequency, and www.weather.com changes format,   '
         ' let me know and I will edit this program to keep up with www.weather.com.   '
         ' email - LosTNaTe@aol.com                                                    '
         '                                                                             '
         ' **History**                                                                 '
         '                                                                             '
         ' Revision 9/1/2000  v2.0                                                     '
         ' *---------------------------------------------------------------------*     '
         ' |New GUI                                                              |     '
         ' |Docking To Systray                                                   |     '
         ' |Auto Updater                                                         |     '
         ' *---------------------------------------------------------------------*     '
         '                                                                             '
         ' Revision 9/15/2000 - 10/01/2000  v3.0                                       '
         ' *---------------------------------------------------------------------*     '
         ' |International And US City Weather*                                   |     '
         ' |Day/Night BUG Fix                                                    |     '
         ' |Several -Week Format Changes                                         |     '
         ' |New GUI                                                              |     '
         ' |Added WAV Sounds (Embedded) *deleted*                                |     '
         ' |More Efficient Error Handling                                        |     '
         ' |Fahrenheit/Celsius/Kelvin Conversion                                 |     '
         ' |Weather Maps*                                                        |     '
         ' |Move Form Without Titlebar                                           |     '
         ' *---------------------------------------------------------------------*     '
         '                                                                             '
         ' Revision 10/01/2000 - 10/20/2000  v4.0                                      '
         ' *---------------------------------------------------------------------*     '
         ' |Fixed - Weather.com changed from 6 to 10 day format* (those bastards)|     '
         ' |                                                                     |     '
         ' |When they changed, my program was confused with the days and dates   |     '
         ' |----ALL Better Now----   *Making Program Even Better Just To Piss*   |     '
         ' |                         *Weather.Com Off*                           |     '
         ' *---------------------------------------------------------------------*     '
         '                                                                             '
         '  Revision 10/24/2000                                                        '
         ' *---------------------------------------------------------------------*     '
         ' |Added the scrolling Local Detail Thing...  Fixed a few little bugs.  |     '
         ' |                                                                     |     '
         ' |Optimized code..  Cut out about 45 KB of code that could was just    |     '
         ' |replicated code.  Now just use one function as oppose to seven.      |     '
         ' |Function in question is the Main_Load function..                     |     '
         ' |                                                                     |     '
         ' |Getting kinda lazy with recording the updates..  Because I am        |     '
         ' |changing stuff everyday...                                           |     '
         ' |                                                                     |     '
         ' |Thinking about uploading to PSC...  Keep getting damn error message  |     '
         ' |when trying to upload.. Oh well...                                   |     '
         ' *---------------------------------------------------------------------*     '
         '                                                                             '
         '  Revision 12/20/2000                                                        '
         ' *---------------------------------------------------------------------*     '
         ' |Added XML/ASP Options, just a new way to retreive data               |     '
         ' |                                                                     |     '
         ' *---------------------------------------------------------------------*     '
         ' "Those you dream by day are cognizant of many things which escape those     '
         ' who only dream by night"                                                    '
         ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'*******
End Sub
'*******

'*********************************
Private Sub Int_city_lst_DblClick()
'*********************************

        Call Disable_Me
        Call Load_Int_Weather
        Call Enable_Me

'*******
End Sub
'*******

'***************************
Private Sub Int_combo_Click()
'***************************

        If Int_combo.ListIndex < 2 Then
                Int_combo.ListIndex = 0
                Country_lst.Clear
                Int_city_lst.Clear
        End If
        If Int_combo.ListIndex > 1 Then
                On Error GoTo Weather_Error
                 ' The main purpose of this routine is to determine the countries for the selection. '
                Call Disable_Me
                Country_lst.Clear
                Int_city_lst.Clear
                name_x = Int_combo.Text
                 ' Formats City, gets rid of spaces and forwards slashes, replaces with "_" '
 
                Do
                        frmt = InStr(name_x, " ")
                        If frmt <> 0 Then
                                Mid(name_x, frmt, 1) = "_"
                        End If
                Loop Until InStr(name_x, " ") = 0
 
 
                Do
                        frmt = InStr(name_x, "/")
                        If frmt <> 0 Then
                                Mid(name_x, frmt, 1) = "_"
                        End If
                Loop Until InStr(name_x, "/") = 0
 
                If name_x = "Central_America_&_Caribbean" Then
                        name_x = "Caribbean_Central_America"
                End If
                 ' Goes out to weather.com and gets countries '
                If MSNET.Value = True Then
                        str_weather = Inet.OpenURL("http://www.weather.com/intl/regions_index/" & name_x & "_region.html")
                End If
                If XML.Value = True Then
                        str_data = "http://www.weather.com/intl/regions_index/" & name_x & "_region.html"
                        weatherxml.open "GET", str_data, False
                        weatherxml.send
                        str_weather = weatherxml.responseText
                End If
                If ASP.Value = True Then
                        Set xObj = CreateObject("Softwing.aspTear")
                        str_weather = xObj.Retrieve("http://www.weather.com/intl/regions_index/" & name_x & "_region.html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
                End If
 
                Do
                        list_pos = Int_combo.ListIndex
                        str0 = "<B>Region: </B>" & Int_combo.List(list_pos)
                        pos0 = InStr(str_weather, str0)
                        str1 = "<A HREF=" & Chr$(34) & "/intl/countries_index/"
                        If last_pos = vbNullString Then last_pos = 1
                        pos1 = InStr(last_pos, str_weather, str1) + Len(str1)
                        If pos1 - Len(str1) = 0 Then Exit Do
                        str2 = ".html" & Chr$(34) & ">"
                        pos2 = InStr(pos1, str_weather, str2)
                        diff = pos2 - pos1
                        mainstr = Mid(str_weather, pos1, diff)
                        last_pos = pos2
                         ' Formats City, gets rid of "_", replaces with " " '
 
                        Do
                                frmt = InStr(mainstr, "_")
                                If frmt <> 0 Then
                                        Mid(mainstr, frmt, 1) = " "
                                End If
                        Loop Until InStr(mainstr, "_") = 0
 
                        If mainstr <> "United States" Then
                                Country_lst.AddItem mainstr
                        End If
                Loop
 
                Call Enable_Me
                Exit Sub
 
Weather_Error:
                MsgBox "Possible Causes For Error" & Chr$(13) & vbNullString & Chr$(13) & "- Not Connected To Internet" & Chr$(13) & "- No Weather Currently Exists For Location" & Chr$(13) & "- Data Is Corrupt Or Not In Proper Format" & Chr$(13) & vbNullString & Chr$(13) & "* Connect To The Internet" & Chr$(13) & "* Select Another City Within The Same Region" & Chr$(13) & "* Try To Update Later", vbInformation + vbOKOnly, "Weather Error"
                StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
                Exit Sub
        End If

'*******
End Sub
'*******

'******************
Sub Kelvin()
'******************

         ' Kelvin = Degress In Celsius + 273.15
        If Temperature.Caption <> "-" Then
                Temperature.Caption = Format((Val(Temperature) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Dewpoint.Caption <> "-" Then
                Dewpoint.Caption = Format((Val(Dewpoint) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day1_hi.Caption <> "-" Then
                Day1_hi.Caption = Format((Val(Day1_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day2_hi.Caption <> "-" Then
                Day2_hi.Caption = Format((Val(Day2_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day3_hi.Caption <> "-" Then
                Day3_hi.Caption = Format((Val(Day3_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day4_hi.Caption <> "-" Then
                Day4_hi.Caption = Format((Val(Day4_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day5_hi.Caption <> "-" Then
                Day5_hi.Caption = Format((Val(Day5_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day6_hi.Caption <> "-" Then
                Day6_hi.Caption = Format((Val(Day6_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day7_hi.Caption <> "-" Then
                Day7_hi.Caption = Format((Val(Day7_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day8_hi.Caption <> "-" Then
                Day8_hi.Caption = Format((Val(Day8_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day9_hi.Caption <> "-" Then
                Day9_hi.Caption = Format((Val(Day9_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day10_hi.Caption <> "-" Then
                Day10_hi.Caption = Format((Val(Day10_hi) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day1_lo.Caption <> "-" Then
                Day1_lo.Caption = Format((Val(Day1_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day2_lo.Caption <> "-" Then
                Day2_lo.Caption = Format((Val(Day2_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day3_lo.Caption <> "-" Then
                Day3_lo.Caption = Format((Val(Day3_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day4_lo.Caption <> "-" Then
                Day4_lo.Caption = Format((Val(Day4_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day5_lo.Caption <> "-" Then
                Day5_lo.Caption = Format((Val(Day5_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day6_lo.Caption <> "-" Then
                Day6_lo.Caption = Format((Val(Day6_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day7_lo.Caption <> "-" Then
                Day7_lo.Caption = Format((Val(Day7_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day8_lo.Caption <> "-" Then
                Day8_lo.Caption = Format((Val(Day8_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day9_lo.Caption <> "-" Then
                Day9_lo.Caption = Format((Val(Day9_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If
        If Day10_lo.Caption <> "-" Then
                Day10_lo.Caption = Format((Val(Day10_lo) - 32) * (5 / 9) + 273.15, "#.0")
        End If

'*******
End Sub
'*******

'********************
Sub Kill_pic()
'********************

         ' This kills the picture files if you have downloaded any weather maps.
        If FileExists(App.Path & "\data.jpg") Then
                Kill App.Path & "\data.jpg"
                DoEvents
        End If
        If FileExists(App.Path & "\data2.jpg") Then
                Kill App.Path & "\data2.jpg"
                DoEvents
        End If

'*******
End Sub
'*******

'**********************************
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'**********************************

        FormMove Me ' Calls function to move form '

'*******
End Sub
'*******

'************************
Private Sub Label1_Click()
'************************

        Call Systrayme

'*******
End Sub
'*******

'*********************************
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*********************************

        Shape3.Left = 7225
        Shape3.Top = 115

'*******
End Sub
'*******

'*******************************
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************

        Shape3.Left = 7215
        Shape3.Top = 105
        Timeout (0.3)

'*******
End Sub
'*******

'************************
Private Sub Label4_Click()
'************************

        Me.Hide
        Call Kill_pic
        Call Fix_Data
        End

'*******
End Sub
'*******

'*********************************
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*********************************

        Shape2.Left = 7420
        Shape2.Top = 40

'*******
End Sub
'*******

'*******************************
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************

        Shape2.Left = 7410
        Shape2.Top = 30
        Timeout (0.3)

'*******
End Sub
'*******

' ********************* '
Sub Load_Data()
    ' *********************                                              '
    ' Loads data from weather.dat, this data is automatically saved each '
    ' time you update.                                                   '
    iniPath$ = App.Path + "\weather.dat"
    checkedx = GetFromINI("Report", "DetailCheck", iniPath$)
    If checkedx = "Yes" Then
        Open (App.Path & "\detail.dat") For Input As #1
        detail_txt.Text = Input(LOF(1), 1)
        Close #1
        tmrScroll.Enabled = True
        detail_txt.Visible = True
        fast.Visible = True
        slow.Visible = True
        Systemtrayfrm.scroll.Checked = True
    ElseIf checkedx = "No" Then
        Systemtrayfrm.scroll.Checked = False
        detail_txt.Visible = False
        fast.Visible = False
        slow.Visible = False
    End If
    tmrScroll.Interval = GetFromINI("Report", "Interval", iniPath$)
    If tmrScroll.Interval = "0" Then
        tmrScroll.Interval = 100
    End If
    Report.Caption = GetFromINI("Report", "Area", iniPath$)
    Ziptxt.Text = GetFromINI("Report", "Zip", iniPath$)
    Conditions.Caption = GetFromINI("Current", "Conditions", iniPath$)
    Temperature.Caption = GetFromINI("Current", "Temperature", iniPath$)
    Wind.Caption = GetFromINI("Current", "Wind", iniPath$)
    Humidity.Caption = GetFromINI("Current", "Humidity", iniPath$)
    Barometer.Caption = GetFromINI("Current", "Barometer", iniPath$)
    Dewpoint.Caption = GetFromINI("Current", "Dewpoint", iniPath$)
    Visibility.Caption = GetFromINI("Current", "Visibility", iniPath$)
    Sunrise.Caption = GetFromINI("Current", "Sunrise", iniPath$)
    Sunset.Caption = GetFromINI("Current", "Sunset", iniPath$)
    Day1.Caption = GetFromINI("Weekday", "Day1", iniPath$)
    Day2.Caption = GetFromINI("Weekday", "Day2", iniPath$)
    Day3.Caption = GetFromINI("Weekday", "Day3", iniPath$)
    Day4.Caption = GetFromINI("Weekday", "Day4", iniPath$)
    Day5.Caption = GetFromINI("Weekday", "Day5", iniPath$)
    Day6.Caption = GetFromINI("Weekday", "Day6", iniPath$)
    Day7.Caption = GetFromINI("Weekday", "Day7", iniPath$)
    Day8.Caption = GetFromINI("Weekday", "Day8", iniPath$)
    Day9.Caption = GetFromINI("Weekday", "Day9", iniPath$)
    Day10.Caption = GetFromINI("Weekday", "Day10", iniPath$)
    Day1_hi.Caption = GetFromINI("High Temp", "Day1", iniPath$)
    Day2_hi.Caption = GetFromINI("High Temp", "Day2", iniPath$)
    Day3_hi.Caption = GetFromINI("High Temp", "Day3", iniPath$)
    Day4_hi.Caption = GetFromINI("High Temp", "Day4", iniPath$)
    Day5_hi.Caption = GetFromINI("High Temp", "Day5", iniPath$)
    Day6_hi.Caption = GetFromINI("High Temp", "Day6", iniPath$)
    Day7_hi.Caption = GetFromINI("High Temp", "Day7", iniPath$)
    Day8_hi.Caption = GetFromINI("High Temp", "Day8", iniPath$)
    Day9_hi.Caption = GetFromINI("High Temp", "Day9", iniPath$)
    Day10_hi.Caption = GetFromINI("High Temp", "Day10", iniPath$)
    Day1_lo.Caption = GetFromINI("Low Temp", "Day1", iniPath$)
    Day2_lo.Caption = GetFromINI("Low Temp", "Day2", iniPath$)
    Day3_lo.Caption = GetFromINI("Low Temp", "Day3", iniPath$)
    Day4_lo.Caption = GetFromINI("Low Temp", "Day4", iniPath$)
    Day5_lo.Caption = GetFromINI("Low Temp", "Day5", iniPath$)
    Day6_lo.Caption = GetFromINI("Low Temp", "Day6", iniPath$)
    Day7_lo.Caption = GetFromINI("Low Temp", "Day7", iniPath$)
    Day8_lo.Caption = GetFromINI("Low Temp", "Day8", iniPath$)
    Day9_lo.Caption = GetFromINI("Low Temp", "Day9", iniPath$)
    Day10_lo.Caption = GetFromINI("Low Temp", "Day10", iniPath$)
    Day1_Weather.Caption = GetFromINI("Weather", "Day1", iniPath$)
    Day2_Weather.Caption = GetFromINI("Weather", "Day2", iniPath$)
    Day3_Weather.Caption = GetFromINI("Weather", "Day3", iniPath$)
    Day4_Weather.Caption = GetFromINI("Weather", "Day4", iniPath$)
    Day5_Weather.Caption = GetFromINI("Weather", "Day5", iniPath$)
    Day6_Weather.Caption = GetFromINI("Weather", "Day6", iniPath$)
    Day7_Weather.Caption = GetFromINI("Weather", "Day7", iniPath$)
    Day8_Weather.Caption = GetFromINI("Weather", "Day8", iniPath$)
    Day9_Weather.Caption = GetFromINI("Weather", "Day9", iniPath$)
    Day10_Weather.Caption = GetFromINI("Weather", "Day10", iniPath$)
    Date_1.Caption = GetFromINI("Dates", "Day1", iniPath$)
    Date_2.Caption = GetFromINI("Dates", "Day2", iniPath$)
    Date_3.Caption = GetFromINI("Dates", "Day3", iniPath$)
    Date_4.Caption = GetFromINI("Dates", "Day4", iniPath$)
    Date_5.Caption = GetFromINI("Dates", "Day5", iniPath$)
    Date_6.Caption = GetFromINI("Dates", "Day6", iniPath$)
    Date_7.Caption = GetFromINI("Dates", "Day7", iniPath$)
    Date_8.Caption = GetFromINI("Dates", "Day8", iniPath$)
    Date_9.Caption = GetFromINI("Dates", "Day9", iniPath$)
    Date_10.Caption = GetFromINI("Dates", "Day10", iniPath$)
    StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
    get_pos = GetFromINI("Report", "Position", iniPath$)
    If get_pos = vbNullString Then get_pos = 0
    Mainfrm.Check1.Value = get_pos
    If Mainfrm.Check1.Value = 1 Then
        Mainfrm.Height = 7020
    ElseIf Mainfrm.Check1.Value = 0 Then
        Mainfrm.Height = 4980
        Mainfrm.Frame1.Visible = False
        Mainfrm.Frame2.Visible = False
    End If
    Center Me
    degree = GetFromINI("Report", "Degrees", iniPath$)
    If degree = "Celsius" Then
        Call Celsius
        Option2.Value = True
    ElseIf degree = "Kelvin" Then
        Call Kelvin
        Option3.Value = True
    Else
        Option1.Value = True
    End If
    method = GetFromINI("Report", "Method", iniPath$)
    If method = "XML" Then
        XML.Value = True
    ElseIf method = "INET" Then
        MSNET.Value = True
    ElseIf method = "ASP" Then
        ASP.Value = True
    End If
    Call Get_Icon
    Int_combo.ListIndex = 0
    US_combo.ListIndex = 0
' ******* '
End Sub
' ******* '


'****************************
Sub Load_Int_Weather()
'****************************

         ' **************************** '
         ' ****************************                           '
         ' This is the preload for the International weather load '
        Call Clear_Data
        DoEvents
        pos_d = Int_city_lst.ListIndex
        name_x = LCase(Int_city_lst.List(pos_d))
        pos_e = Country_lst.ListIndex
        name_y = Country_lst.List(pos_e)
        detail_txt.Visible = False
        tmrScroll.Enabled = False
 
        Do
                frmt = InStr(name_x, " ")
                If frmt <> 0 Then
                        Mid(name_x, frmt, 1) = "_"
                End If
        Loop Until InStr(name_x, " ") = 0
 
 
        Do
                frmt = InStr(name_x, "/")
                If frmt <> 0 Then
                        Mid(name_x, frmt, 1) = "_"
                End If
        Loop Until InStr(name_x, "/") = 0
 
        startx = Timer ' Starts timer to calculate load time '
        If MSNET.Value = True Then
 
                Do
                        num = num + 1
                        If name_y = "Virgin Islands" Then
                                str_weather = Inet.OpenURL("http://www.weather.com/weather/cities/" & str_a & "_vi_" & name_x & ".html")
                        Else
                                str_weather = Inet.OpenURL("http://www.weather.com/weather/cities/" & str_a & "__" & name_x & ".html") '
                        End If
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
        End If
        If XML.Value = True Then
 
                Do
                        num = num + 1
                        If name_y = "Virgin Islands" Then
                                str_data = "http://www.weather.com/weather/cities/" & str_a & "_vi_" & name_x & ".html"
                                weatherxml.open "GET", str_data, False
                                weatherxml.send
                                str_weather = weatherxml.responseText
                        Else
                                str_data = "http://www.weather.com/weather/cities/" & str_a & "__" & name_x & ".html"
                                weatherxml.open "GET", str_data, False
                                weatherxml.send
                                str_weather = weatherxml.responseText
                        End If
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
        End If
        If ASP.Value = True Then
 
                Do
                        num = num + 1
                        If name_y = "Virgin Islands" Then
                                Set xObj = CreateObject("Softwing.aspTear")
                                str_weather = xObj.Retrieve("http://www.weather.com/weather/cities/" & str_a & "_vi_" & name_x & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
                        Else
                                Set xObj = CreateObject("Softwing.aspTear")
                                str_weather = xObj.Retrieve("http://www.weather.com/weather/cities/" & str_a & "__" & name_x & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
                        End If
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
        End If
        StatusBar1.Panels.Item(3).Text = Format(Timer - startx, "#.#0") & " sec."
        DoEvents
        Call RunLoad
         ' ******* '

'*******
End Sub
'*******

'***************************
Sub Load_US_Weather()
'***************************

        ' This is the preload for the US City weather Load '
        Call Clear_Data
        DoEvents
        pos_d = us_city_lst.ListIndex
        name_x = us_city_lst.List(pos_d)
 
        Do
                frmt = InStr(name_x, " ")
                If frmt <> 0 Then
                        Mid(name_x, frmt, 1) = "_"
                End If
        Loop Until InStr(name_x, " ") = 0
 
 
        Do
                frmt = InStr(name_x, "/")
                If frmt <> 0 Then
                        Mid(name_x, frmt, 1) = "_"
                End If
        Loop Until InStr(name_x, "/") = 0
 
        str_weather = vbNullString
        startx = Timer ' Starts timer to calculate load time '
        If MSNET.Value = True Then
 
                Do
                        num = num + 1
                        str_weather = Inet.OpenURL("http://www.weather.com/weather/cities/us_" & str_b & "_" & LCase(name_x) & ".html") '
                        DoEvents
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
                str_local = Inet.OpenURL("http://www.weather.com/weather/36hr/us_" & str_b & "_" & LCase(name_x) & ".html")
        End If
        If XML.Value = True Then
 
                Do
                        num = num + 1
                        str_data = "http://www.weather.com/weather/cities/us_" & str_b & "_" & LCase(name_x) & ".html"
                        weatherxml.open "GET", str_data, False
                        weatherxml.send
                        str_weather = weatherxml.responseText
                        DoEvents
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
                str_data = "http://www.weather.com/weather/36hr/us_" & str_b & "_" & LCase(name_x) & ".html"
                weatherxml.open "GET", str_data, False
                weatherxml.send
                str_local = weatherxml.responseText
        End If
        If ASP.Value = True Then
 
                Do
                        num = num + 1
                        Set xObj = CreateObject("Softwing.aspTear")
                        str_weather = xObj.Retrieve("http://www.weather.com/weather/cities/us_" & str_b & "_" & LCase(name_x) & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
                        DoEvents
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
                str_local = xObj.Retrieve("http://www.weather.com/weather/36hr/us_" & str_b & "_" & LCase(name_x) & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
        End If
        StatusBar1.Panels.Item(3).Text = Format(Timer - startx, "#.#0") & " sec."
        DoEvents
        Call RunLoad
        Call Detail_List
        If Systemtrayfrm.scroll.Checked = True Then
                tmrScroll.Enabled = True
                Systemtrayfrm.scroll.Checked = True
                detail_txt.Visible = True
                fast.Visible = True
                slow.Visible = True
        End If
         ' ******* '

'*******
End Sub
'*******

'****************************
Sub Load_ZIP_Weather()
'****************************

         ' This is the preload for the zip code weather load.     '
         ' http://www.weather.com/weather/us/zips/36hr/24541.html '
        Call Clear_Data
        DoEvents
        startx = Timer
        If MSNET.Value = True Then
 
                Do
                        num = num + 1
                        str_weather = Inet.OpenURL("http://www.weather.com/weather/us/zips/" & Ziptxt.Text & ".html")
                        DoEvents
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
                str_local = Inet.OpenURL("http://www.weather.com/weather/us/zips/36hr/" & Ziptxt.Text & ".html")
        End If
        If XML.Value = True Then
                Dim weatherxml As New XMLHTTPRequest
 
                Do
                        num = num + 1
                        str_data = "http://www.weather.com/weather/us/zips/" & Ziptxt.Text & ".html"
                        weatherxml.open "GET", str_data, False
                        Call weatherxml.setRequestHeader("pragma", "no-cache")
                        weatherxml.send
                        str_weather = weatherxml.responseText
                        DoEvents
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
                str_data = "http://www.weather.com/weather/us/zips/36hr/" & Ziptxt.Text & ".html"
                weatherxml.open "GET", str_data, False
                weatherxml.send
                str_local = weatherxml.responseText
        End If
        If ASP.Value = True Then
 
                Do
                        num = num + 1
                        Set xObj = CreateObject("Softwing.aspTear")
                        str_weather = xObj.Retrieve("http://www.weather.com/weather/us/zips/" & Ziptxt.Text & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
                        DoEvents
                        If num = 5 Then Exit Do
                Loop Until str_weather <> vbNullString
 
                str_local = xObj.Retrieve("http://www.weather.com/weather/us/zips/36hr/" & Ziptxt.Text & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
        End If
        StatusBar1.Panels.Item(3).Text = Format(Timer - startx, "#.#0") & " sec."
        DoEvents
        Call RunLoad
        Call Detail_List
        iniPath$ = App.Path + "\weather.dat"
        checkedx = GetFromINI("Report", "DetailCheck", iniPath$)
        If checkedx = "Yes" Then
                tmrScroll.Enabled = True
                Systemtrayfrm.scroll.Checked = True
                detail_txt.Visible = True
                fast.Visible = True
                slow.Visible = True
        End If


'*******
End Sub
'*******

'***********************
Private Sub MSNET_Click()
'***********************
'Uses the msinet.ocx to extract HTML source

        iniPath$ = App.Path + "\weather.dat"
        entry$ = "INET"
        r% = WritePrivateProfileString("Report", "Method", entry$, iniPath$)

'*******
End Sub
'*******

'*********************
Sub Main_Load()
'*********************

         ' This load gets cycled through for each of th 10 days
        hold_str1 = "<TD ALIGN=" & Chr$(34) & "center" & Chr$(34) & " VALIGN=" & Chr$(34) & "middle" & Chr$(34) & " WIDTH=" & Chr$(34) & "65" & Chr$(34) & " BGCOLOR=" & Chr$(34) & "#E4ECF4" & Chr$(34) & "><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=" & Chr$(34) & "2" & Chr$(34) & ">"
        hold_str2 = "<NOBR>hi&nbsp;"
        hold_str3 = "<NOBR>lo&nbsp;"
        If temp_hold_pos = 0 Then temp_hold_pos = 1
        pos1_a = InStr(temp_hold_pos, str_weather, temp_day)
        temp_hold_pos = pos1_a + Len(temp_day)
        date_pos = pos1_a + Len(temp_day)
        date_pos2 = InStr(date_pos, str_weather, "</B>")
        date_str = Mid(str_weather, date_pos, date_pos2 - date_pos)
        tempo_string = hold_str1
        pos2 = InStr(pos1_a, str_weather, tempo_string)
        day_weather = Mid(str_weather, pos2 + Len(tempo_string))
        tempo_string = InStr(day_weather, "<")
        day_weather = Mid(day_weather, 1, tempo_string - 1)
        temp_hi = hold_str2
        pos3 = InStr(pos1_a, str_weather, temp_hi)
        day_hi = Mid(str_weather, pos3 + Len(temp_hi))
        temp_hi = InStr(day_hi, "&")
        day_hi = Mid(day_hi, 1, temp_hi - 1)
        temp_lo = hold_str3
        pos3 = InStr(pos1_a, str_weather, temp_lo)
        day_lo = Mid(str_weather, pos3 + Len(temp_lo))
        temp_lo = InStr(day_lo, "&")
        day_lo = Mid(day_lo, 1, temp_lo - 1)

'*******
End Sub
'*******

'*************************
Private Sub Map_lst_Click()
'*************************

        If Map_lst.ListIndex < 2 Then
                Map_lst.ListIndex = 0
        End If
        If Map_lst.ListIndex > 1 Then
                Call Mapfrm.Get_Picture
        End If

'*******
End Sub
'*******

'*************************
Private Sub Option1_Click()
'*************************

        Call Fahrenheit
        iniPath$ = App.Path + "\weather.dat"
        entry$ = "Fahrenheit"
        r% = WritePrivateProfileString("Report", "Degrees", entry$, iniPath$)

'*******
End Sub
'*******

'*************************
Private Sub Option2_Click()
'*************************

        Call Fahrenheit
        Call Celsius
        iniPath$ = App.Path + "\weather.dat"
        entry$ = "Celsius"
        r% = WritePrivateProfileString("Report", "Degrees", entry$, iniPath$)

'*******
End Sub
'*******

'*************************
Private Sub Option3_Click()
'*************************

        Call Fahrenheit
        Call Kelvin
        iniPath$ = App.Path + "\weather.dat"
        entry$ = "Kelvin"
        r% = WritePrivateProfileString("Report", "Degrees", entry$, iniPath$)

'*******
End Sub
'*******

'**************************
Private Sub Picture2_Click()
'**************************

        PopupMenu Systemtrayfrm.Main, , 5540, 1075

'*******
End Sub
'*******

'***********************************
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***********************************

        Picture2.MouseIcon = ImageList1.ListImages(1).Picture

'*******
End Sub
'*******

'*******************
Sub RunLoad()
'*******************

        On Error GoTo Weather_Error
         ' This is the universal load for all weather formats '
        If str_weather = vbNullString Then GoTo Weather_Error
        DoEvents
        tempo = "<TD HEIGHT=20 ALIGN=" & Chr$(34) & "center" & Chr$(34) & " VALIGN=" & Chr$(34) & "top" & Chr$(34) & "><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=1>"
        pos1 = InStr(1, str_weather, tempo)
        report_info = Mid(str_weather, pos1 + Len(tempo))
        report_info = Mid$(report_info, 1, InStr(report_info, "<BR>") - 1)
        Mainfrm.Report = "Conditions " & report_info
         ' Gets weather for today '
        If InStr(str_weather, "<B>TODAY<BR>") > 0 Then
                temp_day = "<B>TODAY<BR>"
                pos1 = InStr(str_weather, temp_day)
                date_pos = pos1 + Len(temp_day)
                date_pos2 = InStr(date_pos, str_weather, "</B>")
                date_str = Mid(str_weather, date_pos, date_pos2 - date_pos)
                Mainfrm.Date_1 = date_str
                tempo_string = "<TD ALIGN=" & Chr$(34) & "center" & Chr$(34) & " VALIGN=" & Chr$(34) & "middle" & Chr$(34) & " WIDTH=" & Chr$(34) & "65" & Chr$(34) & " BGCOLOR=" & Chr$(34) & "#E4ECF4" & Chr$(34) & "><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=" & Chr$(34) & "2" & Chr$(34) & ">"
                pos2 = InStr(pos1, str_weather, tempo_string)
                today_weather = Mid(str_weather, pos2 + Len(tempo_string))
                tempo_string = InStr(today_weather, "<")
                today_weather = Mid(today_weather, 1, tempo_string - 1)
        End If
        If InStr(str_weather, "<B>TONIGHT<BR>") > 0 Then
                temp_day = "<B>TONIGHT<BR>"         '
                pos1 = InStr(str_weather, temp_day) '
                date_pos = pos1 + Len(temp_day)
                date_pos2 = InStr(date_pos, str_weather, "</B>")
                date_str = Mid(str_weather, date_pos, date_pos2 - date_pos)
                Mainfrm.Date_1 = date_str
                tempo_string = "<TD ALIGN=" & Chr$(34) & "center" & Chr$(34) & " VALIGN=" & Chr$(34) & "middle" & Chr$(34) & " WIDTH=" & Chr$(34) & "65" & Chr$(34) & " BGCOLOR=" & Chr$(34) & "#E4ECF4" & Chr$(34) & "><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=" & Chr$(34) & "2" & Chr$(34) & ">" '
                pos2 = InStr(pos1, str_weather, tempo_string)                                                                                                                                                                                                                                                                                     '
                today_weather = Mid(str_weather, pos2 + Len(tempo_string))                                                                                                                                                                                                                                                                        '
                tempo_string = InStr(today_weather, "<")                                                                                                                                                                                                                                                                                          '
                today_weather = Mid(today_weather, 1, tempo_string - 1)                                                                                                                                                                                                                                                                           '
        End If                                                                                                                                                                                                                                                                                                                             '
        If InStr(str_weather, "<B>TODAY<BR>") > 0 Then
                temp_hi = "<NOBR>hi&nbsp;"
                pos3 = InStr(pos1, str_weather, temp_hi)
                today_hi = Mid(str_weather, pos3 + Len(temp_hi))
                temp_hi = InStr(today_hi, "&")
                today_hi = Mid(today_hi, 1, temp_hi - 1)
                Mainfrm.Day1.Caption = "Today"
        Else
                Mainfrm.Day1.Caption = "Tonight"
                today_hi = "-"
        End If
        temp_lo = "<NOBR>lo&nbsp;"
        pos3 = InStr(pos1, str_weather, temp_lo)
        today_lo = Mid(str_weather, pos3 + Len(temp_lo))
        temp_lo = InStr(today_lo, "&")
        today_lo = Mid(today_lo, 1, temp_lo - 1)
        Mainfrm.Day1_hi = today_hi
        Mainfrm.Day1_lo = today_lo
        Mainfrm.Day1_Weather = Trim$(today_weather)
         ' Gets current conditions '
        tempo = "<FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=3><B>"
        pos1 = InStr(1, str_weather, tempo)
        current_condition = Mid(str_weather, pos1 + Len(tempo))
        current_condition = Mid$(current_condition, 1, InStr(current_condition, "</B>") - 1)
        Mainfrm.Conditions = Trim$(current_condition)
         ' Gets current temperature '
        tempo = "Temp:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        current_temp = Mid(str_weather, pos1 + Len(tempo))
        current_temp = Mid(current_temp, 1, InStr(current_temp, "&") - 1)
        Mainfrm.Temperature = Trim$(current_temp)
         ' Gets current wind speed '
        tempo = "Wind:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        current_wind = Mid(str_weather, pos1 + Len(tempo))
        tempo = InStr(current_wind, "<")
        current_wind = Mid(current_wind, 1, tempo - 1)
        Mainfrm.Wind = Trim$(current_wind)
         ' Gets current dewpoint '
        tempo = "Dewpoint:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        current_dewpoint = Mid(str_weather, pos1 + Len(tempo))
        tempo = InStr(current_dewpoint, "&")
        current_dewpoint = Mid(current_dewpoint, 1, tempo - 1)
        Mainfrm.Dewpoint = Trim$(current_dewpoint)
         ' Gets current humidity '
        tempo = "Rel. Humidity:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        current_humidity = Mid(str_weather, pos1 + Len(tempo))
        tempo = InStr(current_humidity, "<")
        current_humidity = Mid(current_humidity, 1, tempo - 1)
        Mainfrm.Humidity = Trim$(current_humidity)
         ' Gets current visibility '
        tempo = "Visibility:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        current_visibility = Mid(str_weather, pos1 + Len(tempo))
        tempo = InStr(current_visibility, "<")
        current_visibility = Mid(current_visibility, 1, tempo - 1)
        Mainfrm.Visibility = Trim$(current_visibility)
         ' Gets current barometric pressure '
        tempo = "Barometer:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        current_barometer = Mid(str_weather, pos1 + Len(tempo))
        tempo = InStr(current_barometer, "<")
        current_barometer = Mid(current_barometer, 1, tempo - 1)
        Mainfrm.Barometer = current_barometer
         ' Checks to see if a sunrise data is available, if so extract data '
        tempo = "Sunrise:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        If pos1 > 0 Then
                current_sunrise = Mid(str_weather, pos1 + Len(tempo))
                tempo = InStr(current_sunrise, "<")
                current_sunrise = Mid(current_sunrise, 1, tempo - 1)
                Mainfrm.Sunrise = current_sunrise
        End If
         ' Checks to see if a sunset data is available, if so extract data '
        tempo = "Sunset:</B></FONT></TD>" & Chr$(10) & "<TD WIDTH=5><IMG WIDTH=5 HEIGHT=1 SRC=" & Chr$(34) & "http://image.weather.com/pics/blank.gif" & Chr$(34) & " ALT=" & Chr$(34) & Chr$(34) & "></TD>" & Chr$(10) & "          <TD WIDTH=90><FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=2>"
        pos1 = InStr(str_weather, tempo)
        If pos1 > 0 Then
                current_sunset = Mid(str_weather, pos1 + Len(tempo))
                tempo = InStr(current_sunset, "<")
                current_sunset = Mid(current_sunset, 1, tempo - 1)
                Mainfrm.Sunset = current_sunset
        End If
         ' This determines which day it is according the weather for whatever area you are trying to find. '
         ' If could be Monday where you are and Tuesday or Sunday in a different area of the world.        '
         ' This information is crucial in loading data without errors.                                     '
        stri = "SIZE=" & Chr$(34) & "2" & Chr$(34) & " COLOR=" & Chr$(34) & "#FFFFFF" & Chr$(34) & "><B>"
        findi = InStr(str_weather, stri) + Len(stri)
        findi2 = InStr(findi, str_weather, stri)
        findday = Mid(str_weather, findi2 + Len(stri))
        stri2 = InStr(findday, "<")
        Mainday = Mid(findday, 1, stri2 - 1)
        pos1_a = 0
        temp_hold_pos = 0
        If Mainday = "MON" Then
                Mainfrm.Day2.Caption = "Mon"
                Mainfrm.Day3.Caption = "Tue"
                Mainfrm.Day4.Caption = "Wed"
                Mainfrm.Day5.Caption = "Thu"
                Mainfrm.Day6.Caption = "Fri"
                Mainfrm.Day7.Caption = "Sat"
                Mainfrm.Day8.Caption = "Sun"
                Mainfrm.Day9.Caption = "Mon"
                Mainfrm.Day10.Caption = "Tue"
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        ElseIf Mainday = "TUE" Then
                Mainfrm.Day2.Caption = "Tue"
                Mainfrm.Day3.Caption = "Wed"
                Mainfrm.Day4.Caption = "Thu"
                Mainfrm.Day5.Caption = "Fri"
                Mainfrm.Day6.Caption = "Sat"
                Mainfrm.Day7.Caption = "Sun"
                Mainfrm.Day8.Caption = "Mon"
                Mainfrm.Day9.Caption = "Tue"
                Mainfrm.Day10.Caption = "Wed"
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        ElseIf Mainday = "WED" Then
                Mainfrm.Day2.Caption = "Wed"
                Mainfrm.Day3.Caption = "Thu"
                Mainfrm.Day4.Caption = "Fri"
                Mainfrm.Day5.Caption = "Sat"
                Mainfrm.Day6.Caption = "Sun"
                Mainfrm.Day7.Caption = "Mon"
                Mainfrm.Day8.Caption = "Tue"
                Mainfrm.Day9.Caption = "Wed"
                Mainfrm.Day10.Caption = "Thu"
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        ElseIf Mainday = "THU" Then
                Mainfrm.Day2.Caption = "Thu"
                Mainfrm.Day3.Caption = "Fri"
                Mainfrm.Day4.Caption = "Sat"
                Mainfrm.Day5.Caption = "Sun"
                Mainfrm.Day6.Caption = "Mon"
                Mainfrm.Day7.Caption = "Tue"
                Mainfrm.Day8.Caption = "Wed"
                Mainfrm.Day9.Caption = "Thu"
                Mainfrm.Day10.Caption = "Fri"
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        ElseIf Mainday = "FRI" Then
                Mainfrm.Day2.Caption = "Fri"
                Mainfrm.Day3.Caption = "Sat"
                Mainfrm.Day4.Caption = "Sun"
                Mainfrm.Day5.Caption = "Mon"
                Mainfrm.Day6.Caption = "Tue"
                Mainfrm.Day7.Caption = "Wed"
                Mainfrm.Day8.Caption = "Thu"
                Mainfrm.Day9.Caption = "Fri"
                Mainfrm.Day10.Caption = "Sat"
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        ElseIf Mainday = "SAT" Then
                Mainfrm.Day2.Caption = "Sat"
                Mainfrm.Day3.Caption = "Sun"
                Mainfrm.Day4.Caption = "Mon"
                Mainfrm.Day5.Caption = "Tue"
                Mainfrm.Day6.Caption = "Wed"
                Mainfrm.Day7.Caption = "Thu"
                Mainfrm.Day8.Caption = "Fri"
                Mainfrm.Day9.Caption = "Sat"
                Mainfrm.Day10.Caption = "Sun"
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        ElseIf Mainday = "SUN" Then
                Mainfrm.Day2.Caption = "Sun"
                Mainfrm.Day3.Caption = "Mon"
                Mainfrm.Day4.Caption = "Tue"
                Mainfrm.Day5.Caption = "Wed"
                Mainfrm.Day6.Caption = "Thu"
                Mainfrm.Day7.Caption = "Fri"
                Mainfrm.Day8.Caption = "Sat"
                Mainfrm.Day9.Caption = "Sun"
                Mainfrm.Day10.Caption = "Mon"
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day2_hi = day_hi
                Mainfrm.Day2_lo = day_lo
                Mainfrm.Day2_Weather = day_weather
                Mainfrm.Date_2 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day3_hi = day_hi
                Mainfrm.Day3_lo = day_lo
                Mainfrm.Day3_Weather = day_weather
                Mainfrm.Date_3 = date_str
                temp_day = "<B>TUE<BR>"
                Call Main_Load
                Mainfrm.Day4_hi = day_hi
                Mainfrm.Day4_lo = day_lo
                Mainfrm.Day4_Weather = day_weather
                Mainfrm.Date_4 = date_str
                temp_day = "<B>WED<BR>"
                Call Main_Load
                Mainfrm.Day5_hi = day_hi
                Mainfrm.Day5_lo = day_lo
                Mainfrm.Day5_Weather = day_weather
                Mainfrm.Date_5 = date_str
                temp_day = "<B>THU<BR>"
                Call Main_Load
                Mainfrm.Day6_hi = day_hi
                Mainfrm.Day6_lo = day_lo
                Mainfrm.Day6_Weather = day_weather
                Mainfrm.Date_6 = date_str
                temp_day = "<B>FRI<BR>"
                Call Main_Load
                Mainfrm.Day7_hi = day_hi
                Mainfrm.Day7_lo = day_lo
                Mainfrm.Day7_Weather = day_weather
                Mainfrm.Date_7 = date_str
                temp_day = "<B>SAT<BR>"
                Call Main_Load
                Mainfrm.Day8_hi = day_hi
                Mainfrm.Day8_lo = day_lo
                Mainfrm.Day8_Weather = day_weather
                Mainfrm.Date_8 = date_str
                temp_day = "<B>SUN<BR>"
                Call Main_Load
                Mainfrm.Day9_hi = day_hi
                Mainfrm.Day9_lo = day_lo
                Mainfrm.Day9_Weather = day_weather
                Mainfrm.Date_9 = date_str
                temp_day = "<B>MON<BR>"
                Call Main_Load
                Mainfrm.Day10_hi = day_hi
                Mainfrm.Day10_lo = day_lo
                Mainfrm.Day10_Weather = day_weather
                Mainfrm.Date_10 = date_str
        End If
         ' If some data is not found, it collects large strings which also fit the criteria. '
         ' This clears the data out that we don't need.                                      '
        If Len(Day10_lo.Caption) > 3 Then Day10_lo.Caption = "-"
        If Len(Day10_hi.Caption) > 3 Then Day10_hi.Caption = "-"
        If Len(Sunrise.Caption) > 8 Then Sunrise.Caption = "-"
        If Len(Sunset.Caption) > 8 Then Sunrise.Caption = "-"
         ' Saves all information that we just loaded '
        Call Check_Str
        Call Save_Data
        iniPath$ = App.Path + "\weather.dat"
        degree = GetFromINI("Report", "Degrees", iniPath$)
         ' Determines which measurement to use. '
        If degree = "Celsius" Then
                Call Celsius
        ElseIf degree = "Kelvin" Then
                Call Kelvin
        End If
         ' Calls the icon loading function to load in the graphics. '
        Call Get_Icon
        Exit Sub
 
Weather_Error:
        MsgBox "An Error Has Occurred While Extracting Weather Data" & Chr$(13) & vbNullString & Chr$(13) & "Possible Causes:" & Chr$(13) & "  -Not Connected To The Internet" & Chr$(13) & "  -Location Does Not Exist" & Chr$(13) & "  -No Data Currently Exists For Location" & Chr$(13) & "  -Data Is Corrupt Or In Invalid Format" & Chr$(13) & "  -Problems With Weather.Com" & Chr$(13) & vbNullString & Chr$(13) & "Solutions:" & Chr$(13) & "  -Connect To The Internet" & Chr$(13) & "  -Enter A Valid Location" & Chr$(13) & "  -Select Another Location" & Chr$(13) & "  -Try And Update Later" & Chr$(13) & vbNullString & Chr$(13) & "Any Data That Was Retrieved Has Been Saved", vbInformation + vbOKOnly, "Weather Error"
        Enable_Me
        StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
         ' Saves all information that we just loaded '
        Call Save_Data
        iniPath$ = App.Path + "\weather.dat"
        degree = GetFromINI("Report", "Degrees", iniPath$)
         ' Determines which measurement to use. '
        If degree = "Celsius" Then
                Call Celsius
        ElseIf degree = "Kelvin" Then
                Call Kelvin
        End If
         ' Calls the icon loading function to load in the graphics. '
        Call Get_Icon
        Exit Sub

'*******
End Sub
'*******

'*********************
Sub Save_Data()
'*********************

         ' Saves all information to weather.dat to be loaded later.  This is   '
         ' useful so that you don't have to update each time just to view the  '
         ' same data.  Although you must udpate at least once a day to get the '
         ' most recent and accurate data.                                     '
        iniPath$ = App.Path + "\weather.dat"
        entry$ = Date & " at " & Time
        r% = WritePrivateProfileString("Report", "Time", entry$, iniPath$)
        entry$ = Report.Caption
        r% = WritePrivateProfileString("Report", "Area", entry$, iniPath$)
        entry$ = Ziptxt.Text
        r% = WritePrivateProfileString("Report", "Zip", entry$, iniPath$)
        entry$ = Conditions.Caption
        r% = WritePrivateProfileString("Current", "Conditions", entry$, iniPath$)
        entry$ = Temperature.Caption
        r% = WritePrivateProfileString("Current", "Temperature", entry$, iniPath$)
        entry$ = Wind.Caption
        r% = WritePrivateProfileString("Current", "Wind", entry$, iniPath$)
        entry$ = Humidity.Caption
        r% = WritePrivateProfileString("Current", "Humidity", entry$, iniPath$)
        entry$ = Barometer.Caption
        r% = WritePrivateProfileString("Current", "Barometer", entry$, iniPath$)
        entry$ = Dewpoint.Caption
        r% = WritePrivateProfileString("Current", "Dewpoint", entry$, iniPath$)
        entry$ = Visibility.Caption
        r% = WritePrivateProfileString("Current", "Visibility", entry$, iniPath$)
        entry$ = Sunrise.Caption
        r% = WritePrivateProfileString("Current", "Sunrise", entry$, iniPath$)
        entry$ = Sunset.Caption
        r% = WritePrivateProfileString("Current", "Sunset", entry$, iniPath$)
        entry$ = Day1.Caption
        r% = WritePrivateProfileString("Weekday", "Day1", entry$, iniPath$)
        entry$ = Day2.Caption
        r% = WritePrivateProfileString("Weekday", "Day2", entry$, iniPath$)
        entry$ = Day3.Caption
        r% = WritePrivateProfileString("Weekday", "Day3", entry$, iniPath$)
        entry$ = Day4.Caption
        r% = WritePrivateProfileString("Weekday", "Day4", entry$, iniPath$)
        entry$ = Day5.Caption
        r% = WritePrivateProfileString("Weekday", "Day5", entry$, iniPath$)
        entry$ = Day6.Caption
        r% = WritePrivateProfileString("Weekday", "Day6", entry$, iniPath$)
        entry$ = Day7.Caption
        r% = WritePrivateProfileString("Weekday", "Day7", entry$, iniPath$)
        entry$ = Day8.Caption
        r% = WritePrivateProfileString("Weekday", "Day8", entry$, iniPath$)
        entry$ = Day9.Caption
        r% = WritePrivateProfileString("Weekday", "Day9", entry$, iniPath$)
        entry$ = Day10.Caption
        r% = WritePrivateProfileString("Weekday", "Day10", entry$, iniPath$)
        entry$ = Day1_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day1", entry$, iniPath$)
        entry$ = Day2_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day2", entry$, iniPath$)
        entry$ = Day3_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day3", entry$, iniPath$)
        entry$ = Day4_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day4", entry$, iniPath$)
        entry$ = Day5_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day5", entry$, iniPath$)
        entry$ = Day6_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day6", entry$, iniPath$)
        entry$ = Day7_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day7", entry$, iniPath$)
        entry$ = Day8_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day8", entry$, iniPath$)
        entry$ = Day9_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day9", entry$, iniPath$)
        entry$ = Day10_hi.Caption
        r% = WritePrivateProfileString("High Temp", "Day10", entry$, iniPath$)
        entry$ = Day1_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day1", entry$, iniPath$)
        entry$ = Day2_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day2", entry$, iniPath$)
        entry$ = Day3_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day3", entry$, iniPath$)
        entry$ = Day4_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day4", entry$, iniPath$)
        entry$ = Day5_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day5", entry$, iniPath$)
        entry$ = Day6_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day6", entry$, iniPath$)
        entry$ = Day7_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day7", entry$, iniPath$)
        entry$ = Day8_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day8", entry$, iniPath$)
        entry$ = Day9_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day9", entry$, iniPath$)
        entry$ = Day10_lo.Caption
        r% = WritePrivateProfileString("Low Temp", "Day10", entry$, iniPath$)
        entry$ = Day1_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day1", entry$, iniPath$)
        entry$ = Day2_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day2", entry$, iniPath$)
        entry$ = Day3_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day3", entry$, iniPath$)
        entry$ = Day4_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day4", entry$, iniPath$)
        entry$ = Day5_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day5", entry$, iniPath$)
        entry$ = Day6_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day6", entry$, iniPath$)
        entry$ = Day7_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day7", entry$, iniPath$)
        entry$ = Day8_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day8", entry$, iniPath$)
        entry$ = Day9_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day9", entry$, iniPath$)
        entry$ = Day10_Weather.Caption
        r% = WritePrivateProfileString("Weather", "Day10", entry$, iniPath$)
        entry$ = Date_1.Caption
        r% = WritePrivateProfileString("Dates", "Day1", entry$, iniPath$)
        entry$ = Date_2.Caption
        r% = WritePrivateProfileString("Dates", "Day2", entry$, iniPath$)
        entry$ = Date_3.Caption
        r% = WritePrivateProfileString("Dates", "Day3", entry$, iniPath$)
        entry$ = Date_4.Caption
        r% = WritePrivateProfileString("Dates", "Day4", entry$, iniPath$)
        entry$ = Date_5.Caption
        r% = WritePrivateProfileString("Dates", "Day5", entry$, iniPath$)
        entry$ = Date_6.Caption
        r% = WritePrivateProfileString("Dates", "Day6", entry$, iniPath$)
        entry$ = Date_7.Caption
        r% = WritePrivateProfileString("Dates", "Day7", entry$, iniPath$)
        entry$ = Date_8.Caption
        r% = WritePrivateProfileString("Dates", "Day8", entry$, iniPath$)
        entry$ = Date_9.Caption
        r% = WritePrivateProfileString("Dates", "Day9", entry$, iniPath$)
        entry$ = Date_10.Caption
        r% = WritePrivateProfileString("Dates", "Day10", entry$, iniPath$)
        StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
        Systemtrayfrm.update.Enabled = True

'*******
End Sub
'*******

'************************
Private Sub Timer1_Timer()
'************************

         ' A simple counter which is used when setting the interval '
        int_value = Val(int_value) + 1
        If int_value = Timer1.Tag Then
                Call Disable_Me
                Call Load_ZIP_Weather
                Call Enable_Me
                int_value = 0
        Else
        End If

'*******
End Sub
'*******

'**************************
Private Sub US_combo_Click()
'**************************

        If US_combo.ListIndex < 2 Then
                US_combo.ListIndex = 0
                us_city_lst.Clear
                Alpha_city_lst.Clear
        End If
        If US_combo.ListIndex > 1 Then
                 ' The main function of this routine is to determine the city index for the selection. '
                On Error GoTo Weather_Error
                Call Disable_Me
                us_city_lst.Clear
                Alpha_city_lst.Clear
                name_x = US_combo.Text
                 ' Gets rid of "/" and spaces and replaces with "_" '
 
                Do
                        frmt = InStr(name_x, " ")
                        If frmt <> 0 Then
                                Mid(name_x, frmt, 1) = "_"
                        End If
                Loop Until InStr(name_x, " ") = 0
 
 
                Do
                        frmt = InStr(name_x, "/")
                        If frmt <> 0 Then
                                Mid(name_x, frmt, 1) = "_"
                        End If
                Loop Until InStr(name_x, "/") = 0
 
                If MSNET.Value = True Then
                        str_weather = Inet.OpenURL("http://www.weather.com/weather/us/states/" & name_x & ".html")
                End If
                If XML.Value = True Then
                        str_data = "http://www.weather.com/weather/us/states/" & name_x & ".html"
                        weatherxml.open "GET", str_data, False
                        weatherxml.send
                        str_weather = weatherxml.responseText
                End If
                If ASP.Value = True Then
                        Set xObj = CreateObject("Softwing.aspTear")
                        str_weather = xObj.Retrieve("http://www.weather.com/weather/us/states/" & name_x & ".html", 2, "FORCEDRELOAD", vbNullString, vbNullString)
                End If
                str1 = "<FONT FACE=" & Chr$(34) & "Arial, Helvetica, Chicago, Sans Serif" & Chr$(34) & " SIZE=" & Chr$(34) & "2" & Chr$(34) & ">"
                pos1 = InStr(str_weather, str1)
 
                Do
                        strmain = "&nbsp;<A HREF=" & Chr$(34)
                        If pos1 = vbNullString Then pos1 = 1
                        posmain = InStr(pos1, str_weather, strmain) + Len(strmain)
                        If posmain - Len(strmain) = 0 Then Exit Do
                        str2 = ".html" & Chr$(34) & ">"
                        pos2 = InStr(posmain, str_weather, str2)
                        diff = pos2 - posmain
                        mainstr = Mid(str_weather, posmain, diff)
                        pos1 = pos2
                        Alpha_city_lst.AddItem Right$(mainstr, 1)
                Loop
 
 
                Do
                        str1 = "<A HREF=" & Chr$(34) & "/weather/cities/us_"
                        If last_pos = vbNullString Then last_pos = 1
                        pos1 = InStr(last_pos, str_weather, str1) + Len(str1)
                        If pos1 - Len(str1) = 0 Then Exit Do
                        str2 = ".html" & Chr$(34) & ">"
                        pos2 = InStr(pos1, str_weather, str2)
                        diff = pos2 - pos1
                        mainstr = Mid(str_weather, pos1, diff)
                        last_pos = pos2
 
                        Do
                                frmt = InStr(mainstr, "_")
                                If frmt <> 0 Then
                                        Mid(mainstr, frmt, 1) = " "
                                End If
                        Loop Until InStr(mainstr, "_") = 0
 
                        str_b = Left$(mainstr, 2)
                        mainstr = Right$(mainstr, Len(mainstr) - 3)
                        us_city_lst.AddItem UCase(mainstr)
                Loop
 
                Alpha_city_lst.AddItem UCase(Left$(us_city_lst.List(0), 1)), 0
                Call Enable_Me
                DoEvents
                Alpha_city_lst.ListIndex = 0
                Exit Sub
 
Weather_Error:
                MsgBox "Possible Causes For Error" & Chr$(13) & vbNullString & Chr$(13) & "- Not Connected To Internet" & Chr$(13) & "- No Weather Currently Exists For Location" & Chr$(13) & "- Data Is Corrupt Or Not In Proper Format" & Chr$(13) & vbNullString & Chr$(13) & "* Connect To The Internet" & Chr$(13) & "* Select Another City Within The Same Region" & Chr$(13) & "* Try To Update Later", vbInformation + vbOKOnly, "Weather Error"
                Enable_Me
                StatusBar1.Panels.Item(1).Text = "Updated on " & GetFromINI("Report", "Time", iniPath$)
                Exit Sub
        End If

'*******
End Sub
'*******

'*********************
Private Sub XML_Click()
'*********************
'XML is something I had been working with to access the data at a very high speed.
'XML is useful when you want to access different areas to gather weather info, and
'if you want to go back and view a previous area that you had loaded before.  Simply
'put..  XML caches the information and uses it as oppose to trying to extract the data
'again from the website..  why is this not a GREAT thing you ask?  XML caches the info
'and uses each time, until you unload Your Weather v4.0..  Therefore the first time
'you get the weather for a particular location, the next time..  XML will just use the
'data from the previous load.  So you're not getting UP TO DATE weather..
'
'Example:
'
'Say you get the weather for Chicago using XML.  Then you get weather for New York City.
'
'These load times will be average...  Then if you access either of these 2 cities again before
'unloading Your Weather v 4.0, the cached info will be used and it will take .01 - .05 seconds

        iniPath$ = App.Path + "\weather.dat"
        entry$ = "XML"
        r% = WritePrivateProfileString("Report", "Method", entry$, iniPath$)

'*******
End Sub
'*******

'**********************************
Private Sub Ziptxt_KeyPress(KeyAscii As Integer)
'**********************************

        If KeyAscii = 13 Then
                 ' Checks for correct format and calls the main logic '
                If IsNumeric(Ziptxt) And Len(Ziptxt) = 5 Then
                        Call Disable_Me
                        Call Load_ZIP_Weather
                        Call Enable_Me
                        KeyAscii = 0
                Else
                        MsgBox "Please Enter A Valid Zip Code", vbOKOnly + vbInformation, "Error"
                End If
        End If

'*******
End Sub
'*******

'**********************
Private Sub fast_Click()
'**********************

        If tmrScroll.Interval > 10 Then
                tmrScroll.Interval = tmrScroll.Interval - 30
                iniPath$ = App.Path + "\weather.dat"
                entry$ = tmrScroll.Interval
                r% = WritePrivateProfileString("Report", "Interval", entry$, iniPath$)
        End If

'*******
End Sub
'*******

'*******************************
Private Sub fast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************

        fast.Left = 6970
        fast.Top = 4090

'*******
End Sub
'*******

'*****************************
Private Sub fast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************

        fast.Left = 6960
        fast.Top = 4080

'*******
End Sub
'*******

'**********************
Private Sub slow_Click()
'**********************

        If tmrScroll.Interval < 190 Then
                tmrScroll.Interval = tmrScroll.Interval + 30
                iniPath$ = App.Path + "\weather.dat"
                entry$ = tmrScroll.Interval
                r% = WritePrivateProfileString("Report", "Interval", entry$, iniPath$)
        End If

'*******
End Sub
'*******

'*******************************
Private Sub slow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************

        slow.Left = 7150
        slow.Top = 4090

'*******
End Sub
'*******

'*****************************
Private Sub slow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************

        slow.Left = 7140
        slow.Top = 4080

'*******
End Sub
'*******

'***************************
Private Sub tmrScroll_Timer()
'***************************

        detail_txt.Text = Right$(detail_txt.Text, Len(detail_txt.Text) - 1) + Left$(detail_txt.Text, 1)

'*******
End Sub
'*******

'********************************
Private Sub us_city_lst_DblClick()
'********************************

        Call Disable_Me
        Call Load_US_Weather
        Call Enable_Me

'*******
End Sub
'*******

