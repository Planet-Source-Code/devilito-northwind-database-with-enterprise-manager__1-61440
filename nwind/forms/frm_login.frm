VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_login 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   Icon            =   "frm_login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   1  'CenterOwner
   Begin osenxpsuite2005.OsenXPButton CmdLogin 
      Default         =   -1  'True
      Height          =   345
      Left            =   2430
      TabIndex        =   2
      Top             =   2640
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      Caption         =   "&Log In"
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
      MICON           =   "frm_login.frx":058A
      PICN            =   "frm_login.frx":05A6
      UMCOL           =   -1  'True
      XColorScheme    =   1
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   8632490
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   345
      Left            =   900
      TabIndex        =   3
      Top             =   2640
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
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
      MICON           =   "frm_login.frx":0B40
      PICN            =   "frm_login.frx":0B5C
      UMCOL           =   -1  'True
      XColorScheme    =   1
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   8632490
   End
   Begin osenxpsuite2005.OsenXPTextBox TxtUser 
      Height          =   330
      Left            =   1890
      TabIndex        =   0
      Top             =   1620
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Admin"
      BorderColor     =   8370596
      ColorScheme     =   1
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
      Height          =   255
      Left            =   420
      TabIndex        =   8
      Top             =   1650
      Width           =   960
      _ExtentX        =   1693
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
      Caption         =   "User Name:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   794
      ColorScheme     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login"
      TitleTop        =   7
      icon            =   "frm_login.frx":10F6
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin VB.PictureBox Pic1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   4590
      TabIndex        =   5
      Top             =   450
      Width           =   4590
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Information"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   90
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter user name and password to connect to the server ..."
         Height          =   435
         Left            =   480
         TabIndex        =   6
         Top             =   390
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   3540
         Picture         =   "frm_login.frx":1690
         Top             =   120
         Width           =   720
      End
   End
   Begin osenxpsuite2005.OsenXPLabel OsenXPLabel2 
      Height          =   255
      Left            =   420
      TabIndex        =   9
      Top             =   2100
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "Password:"
      ForeColor       =   0
      BackStyle       =   0
   End
   Begin osenxpsuite2005.OsenXPTextBox TxtPwd 
      Height          =   330
      Left            =   1890
      TabIndex        =   1
      Top             =   2070
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
      Text            =   "vb"
      PasswordChar    =   "•"
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ncount As Integer

Private Sub cmdCancel_Click()
    On Error Resume Next
    
    Unload Me
    
End Sub

Private Sub CmdLogin_Click()
    
    StrUserID = TxtUser
    Ncount = Ncount + 1
    ' Prepared Query
    mStrSQL = "select * from users where userid='" & StrUserID & "' and password='" & TxtPwd & "'"
    
    ' Execute current query
    If GetADORecordset.RecordCount Then ' user validation
        ' valid user
        StrUserName = ADO_SQL_RESULT(mStrSQL, , 1)
        
        Unload Me
        
        If AlreadyExist Then
            frmMain.CreateNode
        Else
            Load frmMain
            frmMain.Show
        End If
    Else
    
        MsgBoxGT "Access denied for user " & TxtUser, vbCritical, "Login Failed", 5
        If Ncount = 3 Then
            Unload Me
            CloseProgram
        End If
    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next

    ' Xp Form initialize
    Me.OsenXPForm1.Init Me
    
    ' Draw gradient color for Pic1
    DrawGradient4Pic Pic1
    
    
End Sub













 
 
 
 
