Attribute VB_Name = "Mod_Nwind"
Option Explicit
Public IsNew As Boolean
Public KeyValue As String
Public FrmName  As String
Public RptName  As String
Public StrUserID As String
Public StrUserName As String
Public AlreadyExist As Boolean

'Count of Type User Privileges
Public Const Total_Access          As Long = 5 ' there are [AddNew, Edit, Delete, Preview, Export]

' Purpose : Exit Application
Public Function CloseProgram() As Boolean
    
    If MsgBoxGT("Do you really want to exit?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
    
        ' unload all forms
        Unload frmMain
    Else
        CloseProgram = True
    End If
    
End Function

' Purpose: Open Connection ....
Public Function OpenConnection() As Boolean

    ' Database Filename : nwind.mdb
    ' Database Password: xpsuite
    
    OpenConnection = Ado_Open(MsAccessConnString(App.Path & "\database\nwind.mdb", "xpsuite"))
    
End Function


'Purpose: Start Application
Sub Main()
    
    frmSplash.Show 1
    
    If OpenConnection Then
        
        frm_login.Show
        
    Else
        MsgBoxGT "Could not connect to database.", vbCritical, "Connection Failed", 5
    End If
        
End Sub









