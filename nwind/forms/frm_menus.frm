VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_menus 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin osenxpsuite2005.OsenXPHookMenu OsenXPHookMenu1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   661
      BmpCount        =   3
      Bmp:1           =   "frm_menus.frx":0000
      Key:1           =   "#mnuLogIn"
      Bmp:2           =   "frm_menus.frx":0428
      Key:2           =   "#mnuChangePassword"
      Bmp:3           =   "frm_menus.frx":0850
      Key:3           =   "#mnuExit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripperLeft     =   8
      GradientMode    =   0
      ShowBottomLine  =   0   'False
      MCountMenu      =   2
      XMenuA1         =   "Sample "
      XMenuACS1       =   ""
      XMenuC1         =   "Mnu_1"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
      XMenuA2         =   "Main Menu "
      XMenuACS2       =   ""
      XMenuC2         =   "mnu_Main"
      XMenuE2         =   -1  'True
      XMenuH2         =   0   'False
   End
   Begin VB.Menu Mnu_1 
      Caption         =   "Sample"
      Begin VB.Menu mnuLogIn 
         Caption         =   "Log In"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnusprt0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu_Main 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnu_Action 
         Caption         =   ""
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
'    On Error Resume Next
'
'    'ChangeWindowStyles hWnd
'
'
'    ' Now Prepared Dynamic menus as same as toolbar button on frmmain
'    Dim I As Integer
'
'    For I = 3 To 16
'        Load mnu_Action(I)
'        mnu_Action(I).Caption = LoadResString(98 + I)
'        OsenXPHookMenu1.SetBitmap mnu_Action(I), frmMain.OsenXPToolBar1.ButtonPicture(I)
'    Next
'
'    mnu_Action(2).Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMain

End Sub

Private Sub mnu_Action_Click(Index As Integer)
    frmMain.GetMenuAction Index
End Sub

Private Sub mnuChangePassword_Click()
    frm_change_password.Show 1
End Sub

Private Sub mnuExit_Click()

    CloseProgram
    
End Sub

Private Sub mnuLogIn_Click()
    frm_login.Show 1
End Sub



 
 
 
 
 
 
 
 
 
 
 
 
 
 
