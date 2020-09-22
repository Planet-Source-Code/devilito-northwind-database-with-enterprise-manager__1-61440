VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_products 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Products"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   Icon            =   "frm_products.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   1  'CenterOwner
   Begin osenxpsuite2005.OsenXPTab OsenXPTab1 
      Height          =   4425
      Left            =   150
      TabIndex        =   13
      Top             =   570
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7805
      TabHeight       =   22
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FrameColor      =   12164479
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
      BackColorParent =   14215660
      TabWidth1       =   71
      TabText1        =   "Product Info"
      TabEnabled1     =   -1  'True
      ScaleHeight     =   295
      ScaleMode       =   0
      ScaleWidth      =   329
      Begin osenxpsuite2005.OsenXPCheckBox chkDiscount 
         Height          =   315
         Left            =   1410
         TabIndex        =   9
         Top             =   3990
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   16250871
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Caption         =   "Discontinued"
         AutoChangeBackColor=   0   'False
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel2 
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   1680
         Width           =   795
         _ExtentX        =   1402
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
         Caption         =   "Category:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPComboBox CboCtg 
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         Top             =   1680
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSH             =   -1  'True
         LSGL            =   -1  'True
         LIH             =   18
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   -2147483630
         ASURC           =   0   'False
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   0
         Top             =   510
         Width           =   3345
         _ExtentX        =   5900
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
         Enabled         =   0   'False
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   510
         Width           =   885
         _ExtentX        =   1561
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
         Caption         =   "ProductID:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   930
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
         Caption         =   "Product Name:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   1
         Left            =   1410
         TabIndex        =   1
         Top             =   900
         Width           =   3345
         _ExtentX        =   5900
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
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   1320
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
         Caption         =   "Supplier:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   4
         Left            =   1410
         TabIndex        =   4
         Top             =   2070
         Width           =   3345
         _ExtentX        =   5900
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
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   2070
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
         Caption         =   "Qty Per Unit:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   5
         Left            =   1410
         TabIndex        =   5
         Top             =   2460
         Width           =   3345
         _ExtentX        =   5900
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
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   18
         Top             =   2460
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
         Caption         =   "Unit Price:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   6
         Left            =   1410
         TabIndex        =   6
         Top             =   2850
         Width           =   3345
         _ExtentX        =   5900
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
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   2880
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "Units in Stock:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   7
         Left            =   1410
         TabIndex        =   7
         Top             =   3240
         Width           =   3345
         _ExtentX        =   5900
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
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   20
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Units On Order:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   8
         Left            =   1410
         TabIndex        =   8
         Top             =   3630
         Width           =   3345
         _ExtentX        =   5900
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
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   21
         Top             =   3630
         Width           =   1170
         _ExtentX        =   2064
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
         Caption         =   "Reorder Level:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPComboBox cboSP 
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   1290
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
         LSH             =   -1  'True
         LSGL            =   -1  'True
         LIH             =   18
         LIO             =   2
         LITL            =   2
         IMGLIST         =   ""
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFontColor =   -2147483630
         ASURC           =   0   'False
      End
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
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
      Caption         =   "Products"
      TitleTop        =   7
      icon            =   "frm_products.frx":038A
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   345
      Left            =   3780
      TabIndex        =   11
      Top             =   5130
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
      MICON           =   "frm_products.frx":0724
      PICN            =   "frm_products.frx":0740
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdOK 
      Height          =   345
      Left            =   2370
      TabIndex        =   10
      Top             =   5130
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
      MICON           =   "frm_products.frx":0CDA
      PICN            =   "frm_products.frx":0CF6
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
End
Attribute VB_Name = "frm_products"
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

    cboSP.InsertItemBySQL AdoCN, "select SupplierID,CompanyName from suppliers order by supplierid", True, True
    cboSP.TextColumn = 1 ' Display >> CompanyName

    CboCtg.InsertItemBySQL AdoCN, "select CategoryID,CategoryName from categories order by categoryid", True, 1
    CboCtg.TextColumn = 1 ' Display Category NAme

    If IsNew Then
        mStrSQL = "select * from products where false"
    Else
        mStrSQL = "select * from products where productid=" & KeyValue
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

        cboSP.KeyValue = Rs.sField(2)
        CboCtg.KeyValue = Rs.sField(3)
        chkDiscount.Value = Rs.sField(9)
        

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

    Rs.sField(2) = cboSP.GetKeyValue

    Rs.sField(3) = CboCtg.GetKeyValue

    Rs.sField(9) = chkDiscount.Value

    ' Now Updating records
    Rs.Update

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Rs = Nothing
End Sub














 
 
 
 
