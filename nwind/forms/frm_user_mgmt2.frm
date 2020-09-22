VERSION 5.00
Object = "{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_user_mgmt 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "User Management"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   Icon            =   "frm_user_mgmt2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin osenxpsuite2005.OsenXPHookMenu OsenXPHookMenu1 
      Height          =   375
      Left            =   6510
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BmpCount        =   0
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
      MCountMenu      =   1
      XMenuA1         =   "Check All "
      XMenuACS1       =   ""
      XMenuC1         =   "mnuCheckAll"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
   End
   Begin osenxpsuite2005.OsenXPTreeView tvwMain 
      Height          =   5355
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   9446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectedColor   =   16777215
      CheckBoxes      =   -1  'True
      BorderStyle     =   0
      HeaderCaption   =   "Northwind Traders"
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowHeader      =   -1  'True
      Gradient        =   -1  'True
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPTab MyTab 
      Height          =   5385
      Left            =   3030
      TabIndex        =   2
      Top             =   1380
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9499
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
      TabWidth1       =   76
      TabText1        =   "User Info"
      TabEnabled1     =   -1  'True
      TabBack21       =   16777215
      TabCountCtls1   =   11
      TabNo1CtlID1    =   "LblGetNodes"
      TabNo1CtlIX1    =   1
      TabNo1CtlIT1    =   -1  'True
      TabNo1CtlID2    =   "cmdOK"
      TabNo1CtlIX2    =   1
      TabNo1CtlIT2    =   -1  'True
      TabNo1CtlID3    =   "cmdCancel"
      TabNo1CtlIX3    =   1
      TabNo1CtlIT3    =   -1  'True
      TabNo1CtlID4    =   "txtID(0)"
      TabNo1CtlIX4    =   1
      TabNo1CtlIT4    =   -1  'True
      TabNo1CtlID5    =   "OsenXPLabel1(0)"
      TabNo1CtlIX5    =   1
      TabNo1CtlIT5    =   -1  'True
      TabNo1CtlID6    =   "Picture2"
      TabNo1CtlIX6    =   1
      TabNo1CtlIT6    =   -1  'True
      TabNo1CtlID7    =   "OsenXPLabel1(1)"
      TabNo1CtlIX7    =   1
      TabNo1CtlIT7    =   -1  'True
      TabNo1CtlID8    =   "TxtName(1)"
      TabNo1CtlIX8    =   1
      TabNo1CtlIT8    =   -1  'True
      TabNo1CtlID9    =   "OsenXPLabel1(2)"
      TabNo1CtlIX9    =   1
      TabNo1CtlIT9    =   -1  'True
      TabNo1CtlID10   =   "TxtPwd(2)"
      TabNo1CtlIX10   =   1
      TabNo1CtlIT10   =   -1  'True
      TabNo1CtlID11   =   "OsenXPLabel5"
      TabNo1CtlIX11   =   1
      TabNo1CtlIT11   =   -1  'True
      ScaleHeight     =   359
      ScaleMode       =   0
      ScaleWidth      =   409
      Begin osenxpsuite2005.OsenXPLabel LblGetNodes 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   2130
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ForeColorOver   =   12582912
         ForeColorDown   =   33023
         Caption         =   "Get All Selected Node(s)"
         FontUnderline   =   -1  'True
         GradientOnOver  =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   16777215
         ForeColor       =   0
         GradientColor3  =   16777215
         BorderColor     =   16777215
         AutoSize        =   0   'False
      End
      Begin osenxpsuite2005.OsenXPButton cmdOK 
         Height          =   285
         Left            =   1740
         TabIndex        =   6
         Top             =   1680
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         MPTR            =   99
         MICON           =   "frm_user_mgmt2.frx":058A
         UMCOL           =   -1  'True
         XPBlendPicture  =   -1  'True
         GradientColor   =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   14854529
      End
      Begin osenxpsuite2005.OsenXPButton cmdCancel 
         Height          =   285
         Left            =   2910
         TabIndex        =   1
         Top             =   1680
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         MPTR            =   99
         MICON           =   "frm_user_mgmt2.frx":05A6
         UMCOL           =   -1  'True
         XPBlendPicture  =   -1  'True
         GradientColor   =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   14854529
      End
      Begin osenxpsuite2005.OsenXPTextBox txtID 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
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
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   570
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
         Caption         =   "User ID:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4110
         ScaleHeight     =   1815
         ScaleWidth      =   1815
         TabIndex        =   10
         Top             =   540
         Width           =   1815
         Begin osenxpsuite2005.OsenXPLabel OsenXPLabel2 
            Height          =   435
            Left            =   630
            TabIndex        =   14
            ToolTipText     =   "Double click here to insert or remove picture"
            Top             =   660
            Visible         =   0   'False
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   767
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            Caption         =   "Photo 3x4"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin VB.Image ImgPhoto 
            Height          =   1635
            Left            =   90
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            ToolTipText     =   "Double click here to insert or remove picture"
            Top             =   90
            Width           =   1635
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00D5860E&
            Height          =   1815
            Left            =   0
            Top             =   0
            Width           =   1815
         End
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   930
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
      Begin osenxpsuite2005.OsenXPTextBox TxtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   930
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
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
         TabIndex        =   13
         Top             =   1290
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
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Top             =   1290
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         PasswordChar    =   "•"
      End
      Begin osenxpsuite2005.OsenXPLabel LblDetail 
         Height          =   315
         Left            =   2280
         TabIndex        =   18
         Top             =   2130
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ForeColorOver   =   12582912
         ForeColorDown   =   33023
         Caption         =   "&Detail Privileges >>"
         FontUnderline   =   -1  'True
         GradientOnOver  =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   16777215
         ForeColor       =   0
         GradientColor3  =   16777215
         BorderColor     =   16777215
         AutoSize        =   0   'False
      End
      Begin osenxpsuite2005.OsenXPTreeView tvwTest 
         Height          =   2745
         Left            =   150
         TabIndex        =   19
         Top             =   2520
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   4842
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SelectedColor   =   16777215
         BorderStyle     =   0
         HeaderCaption   =   "Northwind Traders  >> User Privileges"
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   End
   Begin osenxpsuite2005.OsenXPListBox lvwPrivileges 
      Height          =   2745
      Left            =   120
      TabIndex        =   17
      Top             =   3990
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4842
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      WordWrap        =   0   'False
      ItemHeight      =   20
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      ItemTextLeft    =   20
      SelectModeStyle =   2
      ShowHeader      =   -1  'True
      HeaderFormatString=   "Key;70;0;0;|Caption;250;0;0;|Add;50;1;1;|Edit;50;1;1;|Delete;50;1;1;|Preview;57;1;1;|Export To Excel;60;1;1;"
      Columns         =   7
      ShowGridLines   =   -1  'True
      XPAlphaBlend    =   0   'False
      AlternateRowColors=   -1  'True
      MaxAllColumnWidth=   587
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ASURC           =   -1  'True
      IMGLIST         =   ""
      AllowSortItem   =   0   'False
   End
   Begin osenxpsuite2005.MyImageList SmallIcons 
      Left            =   2790
      Top             =   5940
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   614
      TabIndex        =   8
      Top             =   450
      Width           =   9210
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel4 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         Caption         =   "Add/Edit User ..."
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel3 
         Height          =   255
         Left            =   450
         TabIndex        =   15
         Top             =   390
         Width           =   4965
         _ExtentX        =   8758
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
         Caption         =   "This dialog allow you to create a new user and set global privilege(s)."
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   8280
         Picture         =   "frm_user_mgmt2.frx":05C2
         Top             =   90
         Width           =   720
      End
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
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
      Caption         =   "User Management"
      TitleTop        =   7
      icon            =   "frm_user_mgmt2.frx":148C
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin VB.Menu mnuCheckAll 
      Caption         =   "Check All"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Child_All 
         Caption         =   "&Select All"
         Index           =   1
      End
      Begin VB.Menu Mnu_Child_All 
         Caption         =   "&Deselect All"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_user_mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrValidation   As String
Private Rs As CLS_ADODB_Recordset
Private Flags   As Integer

Public Sub SetRs(Rst As CLS_ADODB_Recordset)
    Set Rs = Rst
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ' On Error Resume Next
    If Not Rs Is Nothing Then
        If Rs.State Then
            If IsNew Then
                Rs.AddNew
            End If
            Rs.sField(0) = txtID
            Rs.sField(1) = TxtName
            Rs.sField(2) = TxtPwd
            Rs.sField(3) = lvwPrivileges.GetPrivileges
            Rs.sField(4) = ImgPhoto.Picture
            Rs.Update
        End If
    End If

    Unload Me

End Sub

Private Sub Form_Load()
    On Error Resume Next


    ' Make sure OsenXPForm Work Fine with this method
    Me.OsenXPForm1.Init Me

    ' Draw gradient Color for Picture1
    DrawGradient4Pic Picture1

    ' Change the Gradient Color of Label and OsenXPTab
    '    LblGetNodes.GradientColor3 = tvwMain.GradientColor2
    '    MyTab.TabBackColor2(1) = tvwMain.GradientColor2
    '    Shape1.BorderColor = tvwMain.GradientColor2

    ' Create new recordset
    Set Rs = New CLS_ADODB_Recordset

    ' Prepared Query
    If IsNew Then
        mStrSQL = "select * from users where false"
    Else
        mStrSQL = "select * from users where userid='" & KeyValue & "'"
    End If

    ' Open Recordset
    Rs.RsOpen AdoCN, mStrSQL
    FillData

    ' Open Image Collection
    SmallIcons.OpenImagesCollection App.Path & "\resources\smallicons.img"

    ' Prepared and set Imagelist handle
    tvwMain.ImageList = SmallIcons.hIml
    tvwTest.ImageList = SmallIcons.hIml
    lvwPrivileges.SmallIcons = SmallIcons.hIml

    ' Insert Nodes into tvwMain
    CreateNode


End Sub

' Purpose : Insert Node item by recordset >> From table "Nodes" <<
Public Sub CreateNode()
    On Error Resume Next

    With tvwMain
        ' Clean Up
        .Clear

        .LockUpdate = True

        ' Prepared Query to get recordset ...
        mStrSQL = "select nodekey,caption,normalicon,selectedicon from nodes where parent='NWIND' order by nodekey"

        ' Create node by recordset ...
        ' GetADORecordset is a resultset from Query on the Above
        .AddNodeByRecordset GetADORecordset, 0, 1, 2, 3, , , 1, 0

        ' Check Count of nodes to expand
        If .Nodes.Count Then
            .Nodes(1).Expanded = True
        End If

        ' Unlock , and draw all nodes
        .LockUpdate = False

    End With


End Sub

' Purpose : Insert Node item by recordset >> From table "Nodes" <<
Public Sub CreateNodeDemo()
    On Error Resume Next

    With tvwTest
        ' Clean Up
        .Clear

        .LockUpdate = True

        ' Prepared Query to get recordset ...
        mStrSQL = "select nodekey,caption,normalicon,selectedicon from nodes where parent='NWIND' order by nodekey"

        ' Create node by recordset ...
        ' GetADORecordset is a resultset from Query on the Above
        .AddNodeByRecordset GetADORecordset, 0, 1, 2, 3, , , 1, 0

        ' Check Count of nodes to expand
        If .Nodes.Count Then
            .Nodes(1).Expanded = True
        End If

        ' Unlock , and draw all nodes
        .LockUpdate = False

    End With


End Sub

Private Sub ImgPhoto_dblClick()
    OpenPicture
End Sub

' Purpose: Show or Hide lvwPrivileges
Private Sub LblDetail_Click()
    On Error Resume Next
    If tvwMain.Height <> 170 Then
        lvwPrivileges.Visible = True
        tvwMain.Height = 170
        MyTab.Height = 170
        LblDetail.Caption = "&Detail Privileges <<"
    Else
        lvwPrivileges.Visible = 0
        tvwMain.Height = 359
        MyTab.Height = 359
        LblDetail.Caption = "&Detail Privileges >>"
    End If

End Sub

Private Sub LblGetNodes_Click()
    On Error Resume Next

    lvwPrivileges.InsertItemFromNodes tvwMain.Nodes
    StrValidation = lvwPrivileges.GetPrivileges
    CreateNodeDemo

End Sub

Private Sub lvwPrivileges_HeaderRightClick(Index As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Index < 2 Then Exit Sub

    Flags = 1
    PopupMenu mnuCheckAll, , X, Y
    If Flags <> 1 Then
        lvwPrivileges.CheckAll Index, Abs(Flags)
    End If

End Sub

Private Sub Mnu_Child_All_Click(Index As Integer)
    Flags = Index - 2
End Sub

Private Sub OsenXPLabel2_Click()
    OpenPicture
End Sub

Private Sub tvwMain_AllowAddNode(NewNodeKey As String, Allow As Boolean, Data As String, iStart As Long)

    Allow = True
    Data = "00000"

End Sub

' Purpose : Recursive function for AddNodeByRecordset methode
Private Sub tvwMain_NodeAdded(NewParentKey As String)
    On Error Resume Next

    ' Prepared Query to get recordset ...
    mStrSQL = "select nodekey,caption,normalicon,selectedicon from nodes where parent='" & NewParentKey & "' order by nodekey"

    ' Create node by recordset ...
    ' GetADORecordset is a resultset from Query on the Above
    tvwMain.AddNodeByRecordset GetADORecordset, 0, 1, 2, 3, NewParentKey, , 1, 0

End Sub

' Purpose : Recursive function for AddNodeByRecordset methode
Private Sub tvwTest_NodeAdded(NewParentKey As String)
    On Error Resume Next

    ' Prepared Query to get recordset ...
    mStrSQL = "select nodekey,caption,normalicon,selectedicon from nodes where parent='" & NewParentKey & "' order by nodekey"

    ' Create node by recordset ...
    ' GetADORecordset is a resultset from Query on the Above
    tvwTest.AddNodeByRecordset GetADORecordset, 0, 1, 2, 3, NewParentKey, , 1, 0

End Sub

Private Sub tvwtest_AllowAddNode(NewNodeKey As String, Allow As Boolean, Data As String, iStart As Long)

    Dim X As Long
    X = InStr(1, StrValidation, Valid_Node_Key(NewNodeKey))

    Data = Mid(StrValidation, X + Len(NewNodeKey) + 5, 5)

    Allow = X

End Sub


Private Sub OpenPicture()
    On Error Resume Next

    Dim Sfile As String
    Sfile = ShowOpenDialog("Open Picture Files", "All Picture (*.BMP;*.JPG;*.JPEG;*.PNG)|*.BMP;*.JPG;*.JPEG;*.PNG", frmMain.hWnd)

    If Sfile <> "" Then

        ImgPhoto.Picture = LoadPicture(Sfile)

    End If

End Sub

Private Sub txtID_OnEnter()
    mStrSQL = "select * from users where userid='" & txtID.Text & "'"
    If GetADORecordset.RecordCount Then
        Rs.RsOpen AdoCN, mStrSQL
        FillData
        IsNew = False
    Else
        IsNew = True
    End If
End Sub


' Purpose: Show data from recordset
Private Sub FillData()
    On Error Resume Next

    If Rs.Have_Records Then

        txtID = Rs.sField(0)
        TxtName = Rs.sField(1)
        TxtPwd = Rs.sField(2)
        StrValidation = Rs.sField(3)
        Rs.LoadPictureFromDB 4, ImgPhoto

        CreateNodeDemo
        lvwPrivileges.InsertItemFromNodes tvwTest.Nodes, False

    End If

End Sub


















