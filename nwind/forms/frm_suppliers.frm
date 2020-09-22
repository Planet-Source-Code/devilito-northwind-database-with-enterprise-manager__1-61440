VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_suppliers 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Supplier Properties"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   Icon            =   "frm_suppliers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   StartUpPosition =   1  'CenterOwner
   Begin osenxpsuite2005.OsenXPTab OsenXPTab1 
      Height          =   5685
      Left            =   120
      TabIndex        =   15
      Top             =   540
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   10028
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
      TabWidth1       =   50
      TabText1        =   "General"
      TabEnabled1     =   -1  'True
      ScaleHeight     =   379
      ScaleMode       =   0
      ScaleWidth      =   359
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   0
         Left            =   1470
         TabIndex        =   0
         Top             =   570
         Width           =   3705
         _ExtentX        =   6535
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
         Left            =   180
         TabIndex        =   16
         Top             =   570
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "Supplier ID:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   960
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Company Name:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   1
         Left            =   1470
         TabIndex        =   1
         Top             =   960
         Width           =   3705
         _ExtentX        =   6535
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
         Left            =   180
         TabIndex        =   18
         Top             =   1350
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "Contact Name:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   2
         Left            =   1470
         TabIndex        =   2
         Top             =   1350
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   1740
         Width           =   1065
         _ExtentX        =   1879
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
         Caption         =   "Contact Title:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   3
         Left            =   1470
         TabIndex        =   3
         Top             =   1740
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   4
         Left            =   180
         TabIndex        =   20
         Top             =   2130
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "Address:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   645
         Index           =   4
         Left            =   1470
         TabIndex        =   4
         Top             =   2130
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   1138
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
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   21
         Top             =   2850
         Width           =   420
         _ExtentX        =   741
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
         Caption         =   "City:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   5
         Left            =   1470
         TabIndex        =   5
         Top             =   2850
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   6
         Left            =   180
         TabIndex        =   22
         Top             =   3270
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "Region:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   6
         Left            =   1470
         TabIndex        =   6
         Top             =   3270
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   7
         Left            =   180
         TabIndex        =   23
         Top             =   3660
         Width           =   1020
         _ExtentX        =   1799
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
         Caption         =   "Postal Code:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   7
         Left            =   1470
         TabIndex        =   7
         Top             =   3660
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   8
         Left            =   180
         TabIndex        =   24
         Top             =   4050
         Width           =   705
         _ExtentX        =   1244
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
         Caption         =   "Country:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   8
         Left            =   1470
         TabIndex        =   8
         Top             =   4050
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   9
         Left            =   180
         TabIndex        =   25
         Top             =   4440
         Width           =   630
         _ExtentX        =   1111
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
         Caption         =   "Phone:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   9
         Left            =   1470
         TabIndex        =   9
         Top             =   4440
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   10
         Left            =   180
         TabIndex        =   26
         Top             =   4830
         Width           =   420
         _ExtentX        =   741
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
         Caption         =   "Fax:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   10
         Left            =   1470
         TabIndex        =   10
         Top             =   4830
         Width           =   3705
         _ExtentX        =   6535
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
         Index           =   11
         Left            =   180
         TabIndex        =   27
         Top             =   5220
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "Homepage:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox txtData 
         Height          =   315
         Index           =   11
         Left            =   1470
         TabIndex        =   11
         Top             =   5220
         Width           =   3705
         _ExtentX        =   6535
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
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
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
      Caption         =   "Supplier Properties"
      TitleTop        =   7
      icon            =   "frm_suppliers.frx":038A
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   345
      Left            =   4230
      TabIndex        =   13
      Top             =   6390
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frm_suppliers.frx":0724
      PICN            =   "frm_suppliers.frx":0740
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdOK 
      Height          =   345
      Left            =   2820
      TabIndex        =   12
      Top             =   6390
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frm_suppliers.frx":0CDA
      PICN            =   "frm_suppliers.frx":0CF6
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
End
Attribute VB_Name = "frm_suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rs As CLS_ADODB_Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    SaveData
    Unload Me

End Sub

Private Sub Form_Load()
    On Error Resume Next

    Me.OsenXPForm1.Init Me

    If IsNew Then
        mStrSQL = "select * from suppliers where false"
    Else
        mStrSQL = "select * from suppliers where supplierid=" & KeyValue
    End If

    Set Rs = New CLS_ADODB_Recordset

    Rs.RsOpen AdoCN, mStrSQL

    FillData

End Sub

Private Sub FillData()
    On Error Resume Next
    If Rs.Have_Records Then

        Dim txt As OsenXPTextBox

        ' Fill in txtdata ...
        For Each txt In txtData
            txt = Rs.sField(txt.Index)
        Next

    End If
End Sub


' Purpose: Save Data
Private Sub SaveData()

    On Error Resume Next

    ' Check Rs status' AddNew or Edit
    If IsNew Then Rs.AddNew

    ' Get New Data from current textbox
    Dim txt As OsenXPTextBox

    ' Fill in txtdata ...
    For Each txt In txtData
        If txt.Text <> "" And txt.Index > 0 Then
            Rs.sField(txt.Index) = txt.Text
        End If
    Next

    ' Now Updating records
    Rs.Update

End Sub


Private Sub txtData_OnEnter(Index As Integer)
    SendKeys "{TAB}"
End Sub
 
 
 
 
