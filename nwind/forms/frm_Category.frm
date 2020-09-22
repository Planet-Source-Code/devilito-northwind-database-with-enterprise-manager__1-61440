VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_Category 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Categories"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   Icon            =   "frm_Category.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   StartUpPosition =   1  'CenterOwner
   Begin osenxpsuite2005.OsenXPTab OsenXPTab1 
      Height          =   2625
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4630
      TabHeight       =   22
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FrameColor      =   7385970
      MaskColor       =   16711935
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   1
      ColorScheme     =   1
      BackColorParent =   14215660
      TabWidth1       =   76
      TabText1        =   "Category Info"
      TabEnabled1     =   -1  'True
      ScaleHeight     =   175
      ScaleMode       =   0
      ScaleWidth      =   441
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   510
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         Text            =   ""
         BorderColor     =   8370596
         Enabled         =   0   'False
         Locked          =   -1  'True
         ColorScheme     =   1
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   510
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "Category ID:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   930
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "Category Name:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   930
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         Text            =   ""
         BorderColor     =   8370596
         ColorScheme     =   1
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1320
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
         Caption         =   "Description:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   825
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1650
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1455
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         BorderColor     =   8370596
         ColorScheme     =   1
         MultiLine       =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   1995
         Left            =   3720
         Top             =   480
         Width           =   2715
      End
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
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
      Caption         =   "Categories"
      TitleTop        =   7
      icon            =   "frm_Category.frx":038A
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin osenxpsuite2005.OsenXPButton CmdChangeImage 
      Height          =   345
      Left            =   180
      TabIndex        =   8
      Top             =   3450
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      Caption         =   "Change Picture"
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
      MICON           =   "frm_Category.frx":0724
      PICN            =   "frm_Category.frx":0740
      UMCOL           =   -1  'True
      XColorScheme    =   1
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   8632490
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   345
      Left            =   5520
      TabIndex        =   9
      Top             =   3420
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frm_Category.frx":0ADA
      PICN            =   "frm_Category.frx":0AF6
      UMCOL           =   -1  'True
      XColorScheme    =   1
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   8632490
   End
   Begin osenxpsuite2005.OsenXPButton cmdOK 
      Height          =   345
      Left            =   4110
      TabIndex        =   10
      Top             =   3420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
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
      MICON           =   "frm_Category.frx":1090
      PICN            =   "frm_Category.frx":10AC
      UMCOL           =   -1  'True
      XColorScheme    =   1
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   8632490
   End
End
Attribute VB_Name = "frm_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rs As CLS_ADODB_Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdChangeImage_Click()
    Dim Sfile As String
    Sfile = ShowOpenDialog("Open Picture Files", "All Picture (*.BMP;*.JPG;*.JPEG;*.PNG)|*.BMP;*.JPG;*.JPEG;*.PNG", frmMain.hWnd)
    If Sfile <> "" Then
        Image1.Picture = LoadPicture(Sfile)
    End If
End Sub

Private Sub cmdOK_Click()

    SaveData
    Unload Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Rs = Nothing
End Sub

Private Sub Form_Load()
    
    Me.OsenXPForm1.Init Me
    
    If IsNew Then
        mStrSQL = "select * from categories where false" ' get empty records
    Else
        mStrSQL = "select * from categories where categoryid=" & KeyValue
    End If
    
    Set Rs = New CLS_ADODB_Recordset
    
    ' Open recordset
    Rs.RsOpen AdoCN, mStrSQL
    
    FillRecords

    
End Sub

'Purpose: Fill record(s) into txtdata
Private Sub FillRecords()
    On Error Resume Next
    
    ' initialize
    ClearMyObject txtData
    
    If Rs.Have_Records Then
    
        Dim txt As OsenXPTextBox
        
        ' Fill in txtdata ...
        For Each txt In txtData
            txt = Rs.sField(txt.Index)
        Next
        
        Rs.LoadPictureFromDB 3, Image1
    
    End If
    
End Sub

' Purpose: Save Data
Private Sub SaveData()

    On Error Resume Next
    
    ' Check Rs status' AddNew or Edit
    If IsNew Then Rs.AddNew
    
    Rs.sField(1) = txtData(1).Text
    
    Rs.sField(2) = txtData(2).Text
    
    ' Save Picture
    Rs.sField(3) = Image1.Picture
    
    ' Now Updating records
    Rs.Update
    
End Sub











 
 
 
 
