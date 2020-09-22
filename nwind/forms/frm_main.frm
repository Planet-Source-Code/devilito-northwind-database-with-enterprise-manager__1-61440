VERSION 5.00
Object = "{81B2046F-0AFA-4E96-AC85-90F2434F97E0}#1.0#0"; "osenxpsuite2005.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Northwind Traders"
   ClientHeight    =   7185
   ClientLeft      =   2940
   ClientTop       =   1770
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   669
   StartUpPosition =   2  'CenterScreen
   Begin osenxpsuite2005.OsenXPHookMenu OsenXPHookMenu1 
      Height          =   375
      Left            =   3960
      Top             =   810
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
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
      MCountMenu      =   3
      XMenuA1         =   "System "
      XMenuACS1       =   ""
      XMenuC1         =   "mnu_System"
      XMenuE1         =   -1  'True
      XMenuH1         =   0   'False
      XMenuA2         =   "Right Event "
      XMenuACS2       =   ""
      XMenuC2         =   "Mnu_Right"
      XMenuE2         =   -1  'True
      XMenuH2         =   0   'False
      XMenuA3         =   "Help "
      XMenuACS3       =   ""
      XMenuC3         =   "mnu_Help"
      XMenuE3         =   -1  'True
      XMenuH3         =   0   'False
   End
   Begin osenxpsuite2005.MyImageList SmallIcons 
      Left            =   1020
      Top             =   3600
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin osenxpsuite2005.MyImageList LargeIcons 
      Left            =   1860
      Top             =   3660
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      Iconsize        =   2
   End
   Begin osenxpsuite2005.OsenXPToolBar OsenXPToolBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   1485
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowEndPanel    =   -1  'True
      XPBlend         =   0   'False
      Begin osenxpsuite2005.OsenXPComboBox cboScheme 
         Height          =   255
         Left            =   7560
         TabIndex        =   9
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   "Default"
         ComboStyle      =   1
         LBN             =   16777215
         LBS             =   10841658
         LBG1            =   16777215
         LBG2            =   14854529
         LAR             =   -1  'True
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
   Begin VB.PictureBox pResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   780
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1725
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   3600
      Width           =   30
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   450
      ScaleHeight     =   1815
      ScaleWidth      =   30
      TabIndex        =   3
      Top             =   3540
      Visible         =   0   'False
      Width           =   30
   End
   Begin osenxpsuite2005.OsenXPTreeView tvwMenus 
      Height          =   4665
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   8229
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
      LostFocusSelectedBackColor=   33023
      MouseIcon       =   "frm_main.frx":058A
      MousePointer    =   99
      ShowNumber      =   -1  'True
      BorderStyle     =   0
      HeaderCaption   =   "Northwind Traders"
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowHeader      =   -1  'True
      Gradient        =   -1  'True
      GradientColor2  =   14854529
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   2
      Top             =   450
      Width           =   10035
      Begin VB.Image Image1 
         Height          =   720
         Left            =   9000
         Picture         =   "frm_main.frx":06EC
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone: 1-206-555-1417   Fax: 1-206-555-5938"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   3
         Left            =   270
         TabIndex        =   7
         Top             =   720
         Width           =   2850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One Portals Way, Twin Points WA  98156"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   510
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Northwind Traders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   2610
      End
   End
   Begin osenxpsuite2005.OsenXPListBox vList 
      Height          =   4845
      Left            =   3060
      TabIndex        =   1
      Top             =   1920
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8546
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
      MousePointer    =   99
      MouseIcon       =   "frm_main.frx":15B6
      ShowHeader      =   -1  'True
      ShowGridLines   =   -1  'True
      AlternateRowColors=   -1  'True
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ASURC           =   -1  'True
      IMGLIST         =   "vImg"
   End
   Begin osenxpsuite2005.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
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
      BorderStyle     =   1
      AutoBackColor   =   0   'False
   End
   Begin osenxpsuite2005.OsenXPStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      Top             =   6750
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   767
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
      NumberOfPanels  =   8
      PWidth1         =   300
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   ""
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_main.frx":1718
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   150
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   ""
      pTextAlignment2 =   0
      PanelPicture2   =   "frm_main.frx":1734
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   240
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Powered By http://osenxpsuite.net"
      pTextAlignment3 =   0
      pTextBold3      =   -1  'True
      PanelPicture3   =   "frm_main.frx":1750
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
      PWidth4         =   55
      PMinWidth4      =   0
      pTTText4        =   ""
      pType4          =   5
      pText4          =   "CAPS"
      pTextAlignment4 =   0
      PanelPicture4   =   "frm_main.frx":1AA2
      PanelPicAlignment4=   0
      pBckgColor4     =   0
      pGradient4      =   0
      pEdgeSpacing4   =   0
      pEdgeInner4     =   0
      pEdgeOuter4     =   0
      PWidth5         =   50
      PMinWidth5      =   0
      pTTText5        =   ""
      pType5          =   6
      pText5          =   "NUM"
      pTextAlignment5 =   0
      PanelPicture5   =   "frm_main.frx":1ABE
      PanelPicAlignment5=   0
      pBckgColor5     =   0
      pGradient5      =   0
      pEdgeSpacing5   =   0
      pEdgeInner5     =   0
      pEdgeOuter5     =   0
      PWidth6         =   60
      PMinWidth6      =   0
      pTTText6        =   ""
      pType6          =   7
      pText6          =   "SCROLL"
      pTextAlignment6 =   0
      PanelPicture6   =   "frm_main.frx":1ADA
      PanelPicAlignment6=   0
      pBckgColor6     =   0
      pGradient6      =   0
      pEdgeSpacing6   =   0
      pEdgeInner6     =   0
      pEdgeOuter6     =   0
      PWidth7         =   75
      PMinWidth7      =   0
      pTTText7        =   ""
      pType7          =   3
      pText7          =   "2005-06-27"
      pTextAlignment7 =   0
      PanelPicture7   =   "frm_main.frx":1AF6
      PanelPicAlignment7=   0
      pBckgColor7     =   0
      pGradient7      =   0
      pEdgeSpacing7   =   0
      pEdgeInner7     =   0
      pEdgeOuter7     =   0
      PWidth8         =   65
      PMinWidth8      =   0
      pTTText8        =   ""
      pType8          =   2
      pText8          =   "01:21:58"
      pTextAlignment8 =   0
      PanelPicture8   =   "frm_main.frx":1B12
      PanelPicAlignment8=   0
      pBckgColor8     =   0
      pGradient8      =   0
      pEdgeSpacing8   =   0
      pEdgeInner8     =   0
      pEdgeOuter8     =   0
      DrawMode        =   1
      HaveXPForm      =   -1  'True
      Begin osenxpsuite2005.OsenXPProgressBar pBar 
         Height          =   225
         Left            =   4590
         TabIndex        =   10
         Top             =   90
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   397
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
   End
   Begin VB.Menu mnu_System 
      Caption         =   "System"
      Visible         =   0   'False
      Begin VB.Menu MnuSysChild 
         Caption         =   "Available"
         Index           =   0
      End
   End
   Begin VB.Menu Mnu_Right 
      Caption         =   "Right Event"
      Visible         =   0   'False
      Begin VB.Menu MnuAction 
         Caption         =   "User Action"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Hlp_Child 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCrLeft             As Long     ' Resize
Private xL                  As Long     ' Resize
Private mWidth              As Long     ' Resize
Private bMove               As Boolean  ' Resize
Private lngTime             As Long
Private m_back              As Collection
Private m_Forward           As Collection
Private b_Allow_Back        As Boolean
Private b_Allow_Forward     As Boolean
Private strPrivileges       As String


' Purpose : Insert Image Collection to Current MyImageList
Private Sub PreparedImage()

    
    ' Add image collection into SmallIcons (MyImageList)
    SmallIcons.OpenImagesCollection App.Path & "\resources\smallicons.img"
    
    ' Add image collection into LargeIcons (MyImageList)
    LargeIcons.OpenImagesCollection App.Path & "\resources\largeicons.img"
    
    ' Set Image handle
    tvwMenus.ImageList = SmallIcons.hIml
    vList.LargeIcons = LargeIcons.hIml
    
    ' Create Toolbar Button
    InitMenusandToolbar
    
End Sub

' Purpose : Create Dynamic toolbar button and menus
Private Sub InitMenusandToolbar()
On Error Resume Next

    Dim I As Long
    Dim StrA
    
    Me.OsenXPToolBar1.SmallIcon = SmallIcons.hIml
    For I = 1 To 20
        StrA = Split(LoadResString(98 + I), "|")
        OsenXPToolBar1.AddNewButton CInt(StrA(1)), StrA(0), CLng(StrA(2)), , , , InStr(1, StrA(0), "&"), CLng(StrA(3))
    Next
    
    ' Now Redraw Toolbar button
    OsenXPToolBar1.RefreshButton
    
    ' Set Up ImageList for Hook Menu
    Set Me.OsenXPHookMenu1.ImageList = SmallIcons
    
    ' Now trying to create dynamic menus for User Activity
    For I = 1 To 10
        Load MnuAction(I)
        StrA = Split(LoadResString(104 + I), "|")
        MnuAction(I).Caption = StrA(0)
        OsenXPHookMenu1.SetIconIndex MnuAction(I), CLng(StrA(2)) + 1
    Next
    
    MnuAction(0).Visible = False
    
    
    ' Create Child of System Menu
    For I = 1 To 4
    
        Load MnuSysChild(I)
        StrA = Split(LoadResString(118 + I), "|")
        MnuSysChild(I).Caption = StrA(0)
        OsenXPHookMenu1.SetIconIndex MnuSysChild(I), CLng(StrA(1)) + 1

    Next
    
    MnuSysChild(0).Visible = False
    
    ' Create Child of Help Menu
    For I = 1 To 4
        Load Mnu_Hlp_Child(I)
        Mnu_Hlp_Child(I).Caption = LoadResString(122 + I)
    Next
    
    Mnu_Hlp_Child(0).Visible = False
    
End Sub



' Purpose : Reposition Controls
Private Sub ResizePosition()
    On Error Resume Next
    
    ' Reposition TreeView and ListBox(ListView:))
    tvwMenus.Move 4, 126, mWidth + mCrLeft, Me.ScaleHeight - 126 - OsenXPStatusBar1.Height - 2
    vList.Move tvwMenus.Left + tvwMenus.Width + 2, 126, ScaleWidth - tvwMenus.Width - 12, tvwMenus.Height
    
    ' Repos PicHandle for Spliter
    pResize.Move tvwMenus.Left + tvwMenus.Width, 126, 2, tvwMenus.Height
    Picture2.Move tvwMenus.Left + tvwMenus.Width, 126, 2, tvwMenus.Height
    
End Sub

'Purpose: Handle User to Change Color Scheme
Private Sub cboScheme_Click()

    On Error Resume Next
    
    ' Change Color Scheme of FrmMain and their Controls
    OsenXPForm1.ColorScheme = cboScheme.ListIndex
    
    
End Sub

' Purpose : Exit App?
Private Sub Form_Unload(Cancel As Integer)

    ' exit application
    Cancel = CloseProgram
    
End Sub

Private Sub Mnu_Hlp_Child_Click(Index As Integer)
    If Index = 4 Then
        MsgBoxGT "Northwind Database Enterprise System v 0.9", vbInformation, "Northwind Traders", 2
        
    Else
        MsgBoxGT "Sorry the help system does not exist.", vbExclamation, "Help"
    End If
End Sub

'Purpose: Do it as sama as toolbar do
Private Sub MnuAction_Click(Index As Integer)
    ' Calll Procedure Toolbar_Button_Click
    OsenXPToolBar1_ButtonClick Index + 6, ""
End Sub

Private Sub MnuSysChild_Click(Index As Integer)
    Select Case Index
        Case 1
            frm_login.Show 1
        Case 2
        Case 4
            Unload Me
    End Select
End Sub

Private Sub OsenXPForm1_ColorSchemeChange(ByVal NewColorScheme As osenxpsuite2005.XPTheme)
    
    ' Change the default theme (Default Color Scheme for this Project)
    DefaultXPTheme = NewColorScheme
    
    ' Change gradient color for picture1
    DrawGradient4Pic Picture1

End Sub

'purpose: User want to change the records or other ....
Private Sub OsenXPToolBar1_ButtonClick(Index As Integer, sText As String)
On Error GoTo Err_MSG
    Dim vt As CLS_xpNode
    
    Select Case Index
    
        Case 3
           tvwMenus.Back
        Case 4
            tvwMenus.Forward
        Case 5
           tvwMenus.UpOneLevel
        Case 7, 8
            mStrSQL = "select form_name from nodes where nodekey='" & tvwMenus.CurrentID & "' "
            FrmName = ADO_SQL_RESULT(mStrSQL)
            KeyValue = vList.GetItemKey
            If FrmName <> "" Then
                IsNew = (Index = 7)
                DisplayForm
            End If
        Case 9
            ' user allow to delete record if viewmode=detail not LargeIcon
            If vList.ViewMode = lvwDetail Then
            
                ' Check listview have data or not
                If vList.ListCount Then
                
                    
                    If vList.ListIndex > -1 Then
                    
                        ' Confirmation before delete
                        If MsgBoxGT("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                            
                            vList.Rs_Delete
                            
                        End If
                    
                    End If
                    
                Else
                    
                    MsgBoxGT "There are no record(s) to delete.", vbExclamation, "Delete", 3
                
                End If
            
            End If
        Case 11
            If vList.ViewMode = lvwDetail Then
                vList.List_Search
            End If
        Case 12
            If vList.ViewMode = lvwDetail Then
                vList.List_Filter
            End If
        Case 13
            If vList.ViewMode = lvwDetail Then
                vList.List_Refresh
            End If
        Case 15
            mStrSQL = "select report_name from nodes where nodekey='" & tvwMenus.CurrentID & "' "
            RptName = ADO_SQL_RESULT(mStrSQL)
            If RptName <> "" Then DisplayReport
        Case 16
            If vList.ViewMode = lvwDetail Then
                vList.ExportToExcel
            End If
        
    End Select
    Exit Sub
Err_MSG:
    MsgBox Err.Description
End Sub

' Purpose: Show Popup Menu
Private Sub OsenXPToolBar1_PopUpMainMenu(Index As Integer, sText As String, X As Long, Y As Long)
    On Error Resume Next
    Select Case Index
        Case 1 ' Connection >> Login,ChangePassword,Exit
            PopupMenu mnu_System, , X, Y
        Case 20 ' Help,About
            PopupMenu mnu_Help, , X, Y
    End Select
    
End Sub

' Purpose: Draw Gradient Color and Reposition Image1
Private Sub Picture1_Resize()
    On Error Resume Next
    
    DrawGradient4Pic Picture1
    Image1.Move Picture1.ScaleWidth - 64
    
End Sub

' Purpose : Resize event handle on runtime
Private Sub pResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then bMove = True

End Sub

' Purpose : Resize event handle on runtime

Private Sub pResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If bMove Then
        If Button = 1 Then
            xL = CLng((X / 15))
            Picture2.Left = tvwMenus.Width + xL
            Picture2.Visible = True
        End If
    End If

End Sub

' Purpose : Resize event handle on runtime
Private Sub pResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Picture2.Visible = False
    mCrLeft = xL
    mWidth = tvwMenus.Width
    ResizePosition

End Sub

' Purpose: Initialize system ...
Private Sub Form_Load()
On Error Resume Next
    

    ' Make sure the osenxpform controls work fine ....
    Me.OsenXPForm1.Init Me, True

    Me.OsenXPForm1.MaximizeWindows
    
    ' Prepared icons collection
    PreparedImage
    
    ' Repaint and resize
    Picture1_Resize
    
    ' Initialize CboScheme Items
    With cboScheme
        .Clear
        .AddItem "Blue"
        .AddItem "Olive Green"
        .AddItem "Silver"
        .ListIndex = 0
    End With
    
    ' Reposition
    mWidth = 200
    ResizePosition
    
    ' Prepared item of treeView (Nodes)
    CreateNode
    AlreadyExist = True
    
    vList.InsertViewFromNode tvwMenus.CurrentNode
    
    
End Sub

' Purpose: Resize ....
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState <> 1 Then ResizePosition
    
End Sub

 
' Purpose : Insert Node item by recordset >> From table "Nodes" <<
Public Sub CreateNode()
On Error Resume Next
    
    
    
    ' Prepared query to get User Privileges from users table
    mStrSQL = "select privileges from users where userid='" & StrUserID & "' "
    
    ' Get .....
    strPrivileges = ADO_SQL_RESULT(mStrSQL)
    
    
    vList.Clear
    
    With tvwMenus
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
    
    ' Now prepare Back and Forward collection
    ' Clean Up and create new collection
    Set m_back = Nothing
    Set m_back = New Collection
    Set m_Forward = Nothing
    Set m_Forward = New Collection
    
    b_Allow_Back = True
    b_Allow_Forward = True
    
End Sub

' Prepared History of user navigation
Private Sub tvwMenus_UpdateHistory(EnableBack As Boolean, EnableForward As Boolean, EnableUpOneLevel As Boolean)
    On Error Resume Next
    
    OsenXPToolBar1.EnabledButton(3) = EnableBack
    OsenXPToolBar1.EnabledButton(4) = EnableForward
    OsenXPToolBar1.EnabledButton(5) = EnableUpOneLevel

End Sub

' Purpose: Display Current Row and Column
Private Sub vList_CellClick(lrow As Long, iCol As Integer, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, Value As String)

    ' Display current cell position [Row,Col]
    OsenXPStatusBar1.PanelCaption(2) = "Ln " & lrow + 1 & " , Col " & iCol + 1

End Sub

'Purpose: there are function as same as edit command
Private Sub vList_DblClick()
    
    ' This function just available if type of Viewmode=lvwDetail , not LargeIcons
    If vList.ViewMode = lvwDetail And tvwMenus.CurrentNode.CurrentPrivileges(2) Then
        OsenXPToolBar1_ButtonClick 8, ""
    End If
    
End Sub

'Purpose: Hide the progressbar
Private Sub vList_EndProgress()
    
    pBar.Visible = False
    
    ' Display total of record(s) and taken time
    OsenXPStatusBar1.PanelCaption(1) = vList.ListCount & " row(s) retrieved [ " & lngTime & " ms taken ]"
    
End Sub

' Purpose: send message into treeview when icon selected
Private Sub vList_IconClick(Index As Long)
    On Error Resume Next
        
    tvwMenus.SetNodeClick tvwMenus.CurrentNode.Child(Index)
        
End Sub

Private Sub vList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 And vList.ViewMode = lvwDetail Then
    
        ' Enable this menu by user privileges setting
        MnuAction(1).Enabled = Me.OsenXPToolBar1.EnabledButton(7)
        MnuAction(2).Enabled = Me.OsenXPToolBar1.EnabledButton(8)
        MnuAction(3).Enabled = Me.OsenXPToolBar1.EnabledButton(9)
        MnuAction(9).Enabled = Me.OsenXPToolBar1.EnabledButton(15)
        MnuAction(10).Enabled = Me.OsenXPToolBar1.EnabledButton(16)
        
        ' Show popup menu
        PopupMenu Mnu_Right
        
    End If
End Sub

' Purpose: Show progress ...
Private Sub vList_ProgressStatus(ByVal lngProgress As Long)
    pBar.Value = lngProgress
End Sub

' Purpose: raise when start insert item into list view
Private Sub vList_StartProgress()
    
    pBar.Value = 0
    pBar.Visible = True
    
End Sub

' Purpose : User Privileges validation
Private Sub tvwMenus_AllowAddNode(NewNodeKey As String, Allow As Boolean, Data As String, iStart As Long)

    Dim X As Long
    X = InStr(1, strPrivileges, Valid_Node_Key(NewNodeKey))
    Data = Mid(strPrivileges, X + iStart, Total_Access)
    Allow = X

End Sub

' Purpose: Create Child Nodes from progress, when insert node by recordset
Private Sub tvwMenus_NodeAdded(NewParentKey As String)
On Error Resume Next

    ' Prepared SQL
     mStrSQL = "select nodekey,caption,normalicon,selectedicon from nodes where parent='" & NewParentKey & "' order by nodekey"

    ' Create node by recordset ...
    ' GetADORecordset is a resultset from Query on the Above
    tvwMenus.AddNodeByRecordset GetADORecordset, 0, 1, 2, 3, NewParentKey, , 1, 0
        
End Sub

' Purpose: Create dynamic nodes (child dynamic) if availables, when first time selected
Private Sub tvwMenus_NodeChange(Node As osenxpsuite2005.CLS_xpNode)
On Error Resume Next
    
    ' Make sure this node have not dynamic child and first time to select
    If Node.FirstSelected = False Then
    
        ' Check from database, is nodekey or dynamic nodekey available to create dynamic child or not
        ' Prepare SQL to do it
        mStrSQL = "select * from nodes where nodekey='" & Node.DynamicParentKey & "' and havechild>=" & Node.DynamicLevel
        
        ' Check from the resultset
        If GetADORecordset.RecordCount Then
        
            ' This node available to create dynamic child on runtime
            ' Now prepare query from View of database base on node.key
            mStrSQL = "vw" & Node.DynamicLevel & "_" & Node.sp_SQL
            
            tvwMenus.AddNodeByRecordset GetADORecordset, 0, 1, , , Node.Key, 1, , , 2, 3, True
            
        End If
        
        ' Done
        Node.FirstSelected = True
        
    End If
    
End Sub

' Purpose: Display resultset into listview if available or display large icon of child into listview
Private Sub tvwMenus_NodeClick(Node As osenxpsuite2005.CLS_xpNode)
On Error Resume Next

    ' Display FullPath of Node
    Me.OsenXPForm1.Caption = "Northwind Traders - " & Node.FullPath
    
    ' Check total child from this node
    If Node.ChildCount Then
        
        vList.HeaderAlignment = enAlignLeft
        vList.InsertViewFromNode Node
        
    Else
    
        vList.Clear
    
        ' Now prepare query from View of database base on node.key
        mStrSQL = "vw" & Node.DynamicLevel & "_" & Node.sp_SQL
    
        ' Check the resultset
        If GetADORecordset.RecordCount Then
        
            Dim StrA As String
            StrA = ADO_SQL_RESULT("select flags from nodes where nodekey='" & Node.DynamicParentKey & "'")
            If StrA <> "" Then
                ' Conditional formating function here ....
                Dim StrB
                StrB = Split(StrA, "|")
                vList.InsertItemBySQL AdoCN, mStrSQL, True, , True, lngTime, CInt(StrB(0)), , CInt(StrB(1)), CInt(StrB(2))
            Else
                vList.InsertItemBySQL AdoCN, mStrSQL, True, , True, lngTime
            End If
        End If
        
        ' If record(s) not found, show message on the header of listview
        If vList.ListCount = 0 Then

            vList.HeaderAlignment = enAlignCenter
            vList.HeaderCaption = "There are no items to show in this view."

        End If
        
        ' Get Privileges setting on current node [Menu] '00000'
        Me.OsenXPToolBar1.EnabledButton(7) = Node.CurrentPrivileges(1) ' Check Access to AddNew Records
        Me.OsenXPToolBar1.EnabledButton(8) = Node.CurrentPrivileges(2) ' Check Access to Edit Records
        Me.OsenXPToolBar1.EnabledButton(9) = Node.CurrentPrivileges(3) ' Check Access to Delete Records
        Me.OsenXPToolBar1.EnabledButton(15) = Node.CurrentPrivileges(4) ' Check Access to Print Preview Records
        Me.OsenXPToolBar1.EnabledButton(16) = Node.CurrentPrivileges(5) ' Check Access to Export to Excel
        
    End If
    
    
End Sub

' Purpose: Call Add/Edit form [Like Property Page on Windows System]
Private Sub DisplayForm()
On Error Resume Next

    Select Case FrmName
        Case "user"
            frm_user_mgmt.Show 0, Me
        Case "employee"
            frm_employees.Show 0, Me
        Case "category"
            frm_Category.Show 0, Me
        Case "customer"
            frm_customers.Show 0, Me
        Case "product"
            frm_products.Show 0, Me
        Case "supplier"
            frm_suppliers.Show 0, Me
        Case Else
            DoEvents
    End Select
    
End Sub

'Purpose: Display Report
Private Sub DisplayReport()
On Error Resume Next
    ' Check resultset view
    If vList.ViewMode = lvwDetail Then
    
        ' Check have record(s) or not
        If vList.ListCount Then
        
            ' Now select the report designer by node.key ....
            Select Case RptName
                Case "customer"
                    vList.ActiveRst.MoveFirst
                    Set rpt_customers.DataSource = vList.ActiveRst
                    rpt_customers.Show 0, Me
                Case "supplier"
                    vList.ActiveRst.MoveFirst
                    Set rpt_suppliers.DataSource = vList.ActiveRst
                    rpt_suppliers.Show 0, Me
                Case "product"
                    vList.ActiveRst.MoveFirst
                    Set rpt_products.DataSource = vList.ActiveRst
                    rpt_products.Show 0, Me
                Case Else
            End Select
        
        End If
    
    End If
End Sub









 



