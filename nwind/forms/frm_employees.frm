VERSION 5.00
Object="{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frm_employees 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Employees"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   Icon            =   "frm_employees.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   1  'CenterOwner
   Begin osenxpsuite2005.OsenXPButton CmdChangeImage 
      Height          =   345
      Left            =   150
      TabIndex        =   36
      Top             =   4560
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
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
      MICON           =   "frm_employees.frx":038A
      PICN            =   "frm_employees.frx":03A6
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdCancel 
      Height          =   345
      Left            =   5190
      TabIndex        =   35
      Top             =   4560
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
      MICON           =   "frm_employees.frx":0740
      PICN            =   "frm_employees.frx":075C
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPButton cmdOK 
      Height          =   345
      Left            =   3780
      TabIndex        =   34
      Top             =   4560
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
      MICON           =   "frm_employees.frx":0CF6
      PICN            =   "frm_employees.frx":0D12
      UMCOL           =   -1  'True
      XPBlendPicture  =   -1  'True
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
   End
   Begin osenxpsuite2005.OsenXPTab OsenXPTab1 
      Height          =   3855
      Left            =   150
      TabIndex        =   16
      Top             =   600
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6800
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
      NumberOfTabs    =   2
      BackColorParent =   14215660
      TabWidth1       =   78
      TabText1        =   "Company Info"
      TabEnabled1     =   -1  'True
      TabCountCtls1   =   15
      TabNo1CtlID1    =   "cboReportTo"
      TabNo1CtlIX1    =   1
      TabNo1CtlIT1    =   0   'False
      TabNo1CtlID2    =   "dtHire"
      TabNo1CtlIX2    =   1
      TabNo1CtlIT2    =   -1  'True
      TabNo1CtlID3    =   "TxtData(0)"
      TabNo1CtlIX3    =   1
      TabNo1CtlIT3    =   -1  'True
      TabNo1CtlID4    =   "OsenXPLabel1(0)"
      TabNo1CtlIX4    =   1
      TabNo1CtlIT4    =   -1  'True
      TabNo1CtlID5    =   "OsenXPLabel1(1)"
      TabNo1CtlIX5    =   1
      TabNo1CtlIT5    =   -1  'True
      TabNo1CtlID6    =   "TxtData(1)"
      TabNo1CtlIX6    =   1
      TabNo1CtlIT6    =   -1  'True
      TabNo1CtlID7    =   "OsenXPLabel1(2)"
      TabNo1CtlIX7    =   1
      TabNo1CtlIT7    =   -1  'True
      TabNo1CtlID8    =   "TxtData(2)"
      TabNo1CtlIX8    =   1
      TabNo1CtlIT8    =   -1  'True
      TabNo1CtlID9    =   "OsenXPLabel1(3)"
      TabNo1CtlIX9    =   1
      TabNo1CtlIT9    =   -1  'True
      TabNo1CtlID10   =   "TxtData(3)"
      TabNo1CtlIX10   =   1
      TabNo1CtlIT10   =   -1  'True
      TabNo1CtlID11   =   "OsenXPLabel1(4)"
      TabNo1CtlIX11   =   1
      TabNo1CtlIT11   =   -1  'True
      TabNo1CtlID12   =   "TxtData(13)"
      TabNo1CtlIX12   =   1
      TabNo1CtlIT12   =   -1  'True
      TabNo1CtlID13   =   "OsenXPLabel1(5)"
      TabNo1CtlIX13   =   1
      TabNo1CtlIT13   =   -1  'True
      TabNo1CtlID14   =   "OsenXPLabel1(6)"
      TabNo1CtlIX14   =   1
      TabNo1CtlIT14   =   -1  'True
      TabNo1CtlID15   =   "Image1"
      TabNo1CtlIX15   =   1
      TabNo1CtlIT15   =   -1  'True
      TabWidth2       =   75
      TabText2        =   "Personal Info"
      TabEnabled2     =   -1  'True
      TabCountCtls2   =   18
      TabNo2CtlID1    =   "dtBirth"
      TabNo2CtlIX1    =   2
      TabNo2CtlIT1    =   -1  'True
      TabNo2CtlID2    =   "OsenXPLabel1(7)"
      TabNo2CtlIX2    =   2
      TabNo2CtlIT2    =   -1  'True
      TabNo2CtlID3    =   "TxtData(7)"
      TabNo2CtlIX3    =   2
      TabNo2CtlIT3    =   -1  'True
      TabNo2CtlID4    =   "OsenXPLabel1(8)"
      TabNo2CtlIX4    =   2
      TabNo2CtlIT4    =   -1  'True
      TabNo2CtlID5    =   "TxtData(9)"
      TabNo2CtlIX5    =   2
      TabNo2CtlIT5    =   -1  'True
      TabNo2CtlID6    =   "OsenXPLabel1(9)"
      TabNo2CtlIX6    =   2
      TabNo2CtlIT6    =   -1  'True
      TabNo2CtlID7    =   "TxtData(8)"
      TabNo2CtlIX7    =   2
      TabNo2CtlIT7    =   -1  'True
      TabNo2CtlID8    =   "OsenXPLabel1(10)"
      TabNo2CtlIX8    =   2
      TabNo2CtlIT8    =   -1  'True
      TabNo2CtlID9    =   "TxtData(10)"
      TabNo2CtlIX9    =   2
      TabNo2CtlIT9    =   -1  'True
      TabNo2CtlID10   =   "TxtData(11)"
      TabNo2CtlIX10   =   2
      TabNo2CtlIT10   =   -1  'True
      TabNo2CtlID11   =   "OsenXPLabel1(11)"
      TabNo2CtlIX11   =   2
      TabNo2CtlIT11   =   -1  'True
      TabNo2CtlID12   =   "OsenXPLabel1(12)"
      TabNo2CtlIX12   =   2
      TabNo2CtlIT12   =   -1  'True
      TabNo2CtlID13   =   "TxtData(15)"
      TabNo2CtlIX13   =   2
      TabNo2CtlIT13   =   -1  'True
      TabNo2CtlID14   =   "OsenXPLabel1(13)"
      TabNo2CtlIX14   =   2
      TabNo2CtlIT14   =   -1  'True
      TabNo2CtlID15   =   "TxtData(12)"
      TabNo2CtlIX15   =   2
      TabNo2CtlIT15   =   -1  'True
      TabNo2CtlID16   =   "OsenXPLabel1(14)"
      TabNo2CtlIX16   =   2
      TabNo2CtlIT16   =   -1  'True
      TabNo2CtlID17   =   "TxtData(4)"
      TabNo2CtlIX17   =   2
      TabNo2CtlIT17   =   -1  'True
      TabNo2CtlID18   =   "OsenXPLabel1(15)"
      TabNo2CtlIX18   =   2
      TabNo2CtlIT18   =   -1  'True
      ScaleHeight     =   257
      ScaleMode       =   0
      ScaleWidth      =   421
      Begin osenxpsuite2005.OsenXPDTPicker dtBirth 
         Height          =   315
         Left            =   76320
         TabIndex        =   13
         Top             =   3120
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         Text            =   "2005-06-30"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatDate      =   "dd.mm.yyyy"
         YEAR            =   0
         MONTH           =   0
         MYDATE          =   0
         thisdate        =   38533
      End
      Begin osenxpsuite2005.OsenXPComboBox cboReportTo 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   2310
         Width           =   1905
         _ExtentX        =   3360
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
      Begin osenxpsuite2005.OsenXPDTPicker dtHire 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   2790
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Text            =   "2005-06-30"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatDate      =   "mmm dd, yyyy"
         YEAR            =   0
         MONTH           =   0
         MYDATE          =   0
         thisdate        =   38533
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   0
         Left            =   1350
         TabIndex        =   0
         Top             =   510
         Width           =   1905
         _ExtentX        =   3360
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
         Locked          =   -1  'True
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Employee ID:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "First Name:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   1
         Top             =   960
         Width           =   1905
         _ExtentX        =   3360
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
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   1410
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Last Name:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   2
         Left            =   1350
         TabIndex        =   2
         Top             =   1410
         Width           =   1905
         _ExtentX        =   3360
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
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   20
         Top             =   1860
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Title:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   3
         Left            =   1350
         TabIndex        =   3
         Top             =   1860
         Width           =   1905
         _ExtentX        =   3360
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
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   2340
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Report To:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   13
         Left            =   1350
         TabIndex        =   6
         Top             =   3240
         Width           =   1905
         _ExtentX        =   3360
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
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   2790
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Hire Date:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   23
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Extention:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   7
         Left            =   75150
         TabIndex        =   24
         Top             =   450
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
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   645
         Index           =   7
         Left            =   76320
         TabIndex        =   7
         Top             =   450
         Width           =   4815
         _ExtentX        =   8493
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
         MultiLine       =   -1  'True
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Index           =   8
         Left            =   78120
         TabIndex        =   25
         Top             =   1200
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
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   9
         Left            =   79230
         TabIndex        =   9
         Top             =   1200
         Width           =   1905
         _ExtentX        =   3360
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
         Index           =   9
         Left            =   75150
         TabIndex        =   26
         Top             =   1230
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
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   8
         Left            =   76320
         TabIndex        =   8
         Top             =   1230
         Width           =   1635
         _ExtentX        =   2884
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
         Index           =   10
         Left            =   75150
         TabIndex        =   27
         Top             =   1680
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
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   10
         Left            =   76320
         TabIndex        =   10
         Top             =   1680
         Width           =   1635
         _ExtentX        =   2884
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
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   11
         Left            =   79230
         TabIndex        =   11
         Top             =   1680
         Width           =   1905
         _ExtentX        =   3360
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
         Index           =   11
         Left            =   78120
         TabIndex        =   28
         Top             =   1680
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
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   285
         Index           =   12
         Left            =   78120
         TabIndex        =   29
         Top             =   2100
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Notes:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   1245
         Index           =   15
         Left            =   78180
         TabIndex        =   14
         Top             =   2400
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   2196
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
         MultiLine       =   -1  'True
      End
      Begin osenxpsuite2005.OsenXPLabel OsenXPLabel1 
         Height          =   285
         Index           =   13
         Left            =   75150
         TabIndex        =   30
         Top             =   2070
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Home Phone:"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   12
         Left            =   76320
         TabIndex        =   12
         Top             =   2130
         Width           =   1635
         _ExtentX        =   2884
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
         Index           =   14
         Left            =   75150
         TabIndex        =   31
         Top             =   2610
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "Title of Courtesy:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin osenxpsuite2005.OsenXPTextBox TxtData 
         Height          =   315
         Index           =   4
         Left            =   76530
         TabIndex        =   32
         Top             =   2610
         Width           =   1425
         _ExtentX        =   2514
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
         Index           =   15
         Left            =   75150
         TabIndex        =   33
         Top             =   3120
         Width           =   870
         _ExtentX        =   1535
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
         Caption         =   "Birth Date:"
         ForeColor       =   0
         BackStyle       =   0
      End
      Begin VB.Image Image1 
         Height          =   3015
         Left            =   3390
         Top             =   510
         Width           =   2775
      End
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
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
      Caption         =   "Employees"
      TitleTop        =   7
      icon            =   "frm_employees.frx":12AC
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
End
Attribute VB_Name = "frm_employees"
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
    
    ' Make sure this form work fine ...
    Me.OsenXPForm1.Init Me
    
    ' Fill Employee Name into Cbo Reportto
    mStrSQL = "SELECT EmployeeID, titleofcourtesy+"" ""+LastName+"", ""+FirstName AS [Employee Name] from employees order by employeeid"
    cboReportTo.InsertItemBySQL AdoCN, mStrSQL, True, True
    cboReportTo.TextColumn = 1
    
    Set Rs = New CLS_ADODB_Recordset
    
    If IsNew Then
        mStrSQL = "select * from employees where false" ' to get no records
    Else
        mStrSQL = "select * from employees where employeeid=" & KeyValue
    End If
        
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
        
        Rs.LoadPictureFromDB 14, Image1
        
        dtHire.Value = Rs.sField(6)
        
        dtBirth.Value = Rs.sField(5)
        
        cboReportTo.Text = Rs.sField(16)
    
    End If
    
End Sub

' Purpose: Save Data
Private Sub SaveData()

   ' On Error Resume Next
    
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

    Rs.sField(6) = dtHire.Value
    
    Rs.sField(5) = dtBirth.Value
    
    ' Get EmployeeID from cboreportto
    Rs.sField(16) = cboReportTo.GetKeyValue
    
    ' Save Picture
    Rs.sField(14) = Image1.Picture
    
    ' Now Updating records
    Rs.Update
    
End Sub




 
 
 
 
