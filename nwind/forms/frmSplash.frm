VERSION 5.00
Object = "{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   1020
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   2610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   174
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1920
      Top             =   1680
   End
   Begin osenxpsuite2005.OsenXPProgressBar pBar 
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   600
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   2871848
      Value           =   100
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
      Height          =   450
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NORTHWIND TRADERS"
      ForeColor       =   8388608
      BackStyle       =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    FadeIn Me.hWnd
    pBar.StartSearch
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    GradientForm Me
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    pBar.StopSearch
    FadeOut Me.hWnd
    
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub




