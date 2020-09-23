VERSION 5.00
Begin VB.Form Hidefrm 
   Caption         =   "Your Weather v4.0"
   ClientHeight    =   240
   ClientLeft      =   1680
   ClientTop       =   -9660
   ClientWidth     =   1560
   Icon            =   "Hidefrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleWidth      =   1560
End
Attribute VB_Name = "Hidefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_hWndParent As Long
'*********************
Private Sub Form_Load()
'*********************

        Dim hMenu& 'the handle of the system menu
                   ' Get the handle to the system menu
        hMenu = GetSystemMenu(Me.hwnd, 0&)
         ' delete inapropriate members on the system menu
        Call DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)
        Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
         ' we COULD use the move, but it is a lot more complicated and requires subclassing
         ' because the MOVE moves this form and not the one the user expects to move.
        Call DeleteMenu(hMenu, SC_MOVE, MF_BYCOMMAND)
         ' make this form the parent of the main form and get the old parent at the same time
        m_hWndParent = SetWindowLong(Mainfrm.hwnd, GWL_HWNDPARENT, Me.hwnd)

'*******
End Sub
'*******

