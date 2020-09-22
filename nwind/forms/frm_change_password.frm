VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_change_password 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Change Password ..."
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   Icon            =   "frm_change_password.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4470
      TabIndex        =   1
      Top             =   450
      Width           =   4470
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3630
         Picture         =   "frm_change_password.frx":038A
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change your password ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   180
         Width           =   2355
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type the password you want to use."
         Height          =   195
         Left            =   630
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Change Password ..."
      TitleTop        =   7
      icon            =   "frm_change_password.frx":0C54
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin osenxpsuite2005.OsenXPButton CmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   2340
      TabIndex        =   4
      Top             =   2850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      Caption         =   "&OK"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frm_change_password.frx":0FEE
      PICN            =   "frm_change_password.frx":100A
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   345
      Left            =   810
      TabIndex        =   5
      Top             =   2850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      Caption         =   "&Cancel"
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MCOL            =   16711935
      MPTR            =   0
      MICON           =   "frm_change_password.frx":15A4
      PICN            =   "frm_change_password.frx":15C0
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPTextBox TxtUser 
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   1500
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      PasswordChar    =   "•"
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
      Height          =   255
      Left            =   270
      TabIndex        =   7
      Top             =   1530
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Old Password:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel2 
      Height          =   255
      Left            =   270
      TabIndex        =   8
      Top             =   1950
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "New Password:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPTextBox TxtPwd 
      Height          =   330
      Left            =   1800
      TabIndex        =   9
      Top             =   1920
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      PasswordChar    =   "•"
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
      Height          =   255
      Left            =   270
      TabIndex        =   10
      Top             =   2370
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Confrim Password:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPTextBox OsenXPTextBox1 
      Height          =   330
      Left            =   1800
      TabIndex        =   11
      Top             =   2340
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      PasswordChar    =   "•"
   End
End
Attribute VB_Name = "frm_change_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Do Nothing
Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
On Error Resume Next

    ' Xp Form initialize
    Me.OsenXPForm1.Init Me
    
    ' Draw gradient color for Pic1
    DrawGradient4Pic Picture1
    
    
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
 
 
 
