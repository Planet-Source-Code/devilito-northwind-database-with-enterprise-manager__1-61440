VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Northwind Traders"
   ClientHeight    =   8370
   ClientLeft      =   5610
   ClientTop       =   3765
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   StartUpPosition =   2  'CenterScreen
   Begin osenxpsuite2005.OsenXPTreeView OsenXPTreeView1 
      Height          =   5205
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9181
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectedColor   =   16777215
      BorderStyle     =   0
      HeaderCaption   =   "OsenXPTreeView1"
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowHeader      =   -1  'True
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPListBox OsenXPListBox1 
      Height          =   5445
      Left            =   4740
      TabIndex        =   2
      Top             =   2040
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   9604
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      AllowEdit       =   0   'False
      WordWrap        =   0   'False
      ItemHeight      =   20
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IMGLIST         =   ""
      HeaderAlignment =   1
   End
   Begin osenxpsuite2005.OsenXPStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      Top             =   7965
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   714
      BackColor       =   14936810
      ForeColor       =   -2147483630
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   3
      PWidth1         =   100
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Osen Kusnadi"
      pTextAlignment1 =   0
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   200
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   "Capsule Corporation"
      pTextAlignment2 =   0
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   120
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Personal License"
      pTextAlignment3 =   0
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
      DrawMode        =   1
      HaveXPForm      =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   655
      TabIndex        =   1
      Top             =   450
      Width           =   9825
      Begin VB.Image Image1 
         Height          =   720
         Left            =   8790
         Picture         =   "frmMain.frx":0ECA
         Top             =   270
         Width           =   720
      End
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
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
      Caption         =   "Northwind Traders"
      TitleTop        =   7
      icon            =   "frmMain.frx":1D94
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()

    ' Set The Default Color scheme for All forms in this projects
    DefaultXPTheme = xpOliveGreen
    
    ' Initialize XP Form
    Me.OsenXPForm1.Init Me

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
 
 
 
 
 
 
 
