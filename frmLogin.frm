VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Form"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "frmLogin"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2220
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton exitBtn 
      Caption         =   "E&xit"
      Height          =   300
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   " Exit "
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   10
      TabIndex        =   6
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton btnClear 
         Caption         =   "&Clear"
         Height          =   310
         Left            =   3600
         TabIndex        =   4
         ToolTipText     =   " Clear Entries "
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox AccLevel 
         Height          =   345
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox PWORD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "01234567890123456789"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "&Login"
         Height          =   310
         Left            =   3600
         TabIndex        =   3
         Tag             =   " Login "
         ToolTipText     =   " Login "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox LOGNAME 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "01234567890123456789"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Access Level"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   200
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   200
         TabIndex        =   8
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   200
         TabIndex        =   7
         Top             =   285
         Width           =   960
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   1905
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5698
            MinWidth        =   5645
            Text            =   "Caption"
            TextSave        =   "Caption"
            Key             =   "capt"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Caption"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "1/14/00"
            Key             =   "dt"
            Object.Tag             =   ""
            Object.ToolTipText     =   " Current Date "
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "9:08 PM"
            Key             =   "time"
            Object.Tag             =   ""
            Object.ToolTipText     =   " Current Time "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Login_Attempts As Integer

Private Sub AccLevel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call btnLogin_Click
  End If
End Sub

Private Sub btnClear_Click()
  LOGNAME.Text = ""
  PWORD.Text = ""
  AccLevel.Clear
  AccLevel.AddItem "Administrator"
  AccLevel.AddItem "User"
  AccLevel.ListIndex = 0
  LOGNAME.SetFocus
End Sub

Private Sub btnClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Clear Entries"
End Sub

Private Sub btnLogin_Click()
  Dim Record_Found As Boolean
  Dim LoginDatabase As Database
  Dim LoginRecordset As Recordset
  On Error GoTo LoginErr
  
  Record_Found = False
  Set LoginDatabase = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
  Login_Attempts = Login_Attempts + 1
  Set LoginRecordset = LoginDatabase.OpenRecordset("Users")
    
  Do While Not LoginRecordset.EOF
     If LoginRecordset.Fields("LoginName") = EncryptText(LOGNAME.Text, Database_Password) And _
        LoginRecordset.Fields("Password") = EncryptText(PWORD.Text, Database_Password) And _
        LoginRecordset.Fields("Accesslevel") = EncryptText(AccLevel.Text, Database_Password) Then
        Record_Found = True
        Exit Do
       Else
        LoginRecordset.MoveNext
      End If
  Loop
  
  If Record_Found = True Then
     If LoginRecordset.Fields("LoggedIn") = True Then
        MsgBox LOGNAME.Text & " is currently listed as Logged In. If you are sure that " & LOGNAME.Text & " is currently not Logged In, contact an Administrator or view readme.txt", vbInformation + vbOKOnly
        Unload Me
     End If
     
     Current_LoginName = DecryptText(LoginRecordset.Fields("LoginName"), Database_Password)
     Current_AccessLevel = DecryptText(LoginRecordset.Fields("AccessLevel"), Database_Password)
     
     WriteIniFile App.Path & "\Family2.ini", Current_LoginName, "Last-Logged-In", Format(Now, "Long Date")
     WriteIniFile App.Path & "\Family2.ini", "LOG", "Last-Used-By", Current_LoginName
     
     LoginRecordset.Edit
     LoginRecordset.Fields("LoggedIn") = True
     LoginRecordset.Update
     LoginRecordset.Close
     LoginDatabase.Close
     
     If Table_Ok(Database_Path & "\" & Database_Name, Current_LoginName) = True Then
        MsgBox "Welcome " & Current_LoginName & "...", , "Logged In Successfully"
        Load frmMain
        frmMain.Show
       Else
        MsgBox "Unable to load " & Current_LoginName & "'s database", vbCritical + vbOKOnly
        Unload Me
     End If
    Else ' Record_Found =False
     If Login_Attempts < 4 Then
        MsgBox "The entries that you have made are invalid. Note: Values are Case Sensitive.", vbInformation + vbOKOnly
       Else
        MsgBox "Contact an Administrator for a valid login name and password.", vbInformation
        Unload Me
     End If
     Call btnClear_Click
  End If
   
LoginErr:
  If Err.Number <> 0 Then
     MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Sub

Private Sub btnLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Login"
End Sub


Private Sub exitBtn_Click()
   Unload Me
   End
End Sub

Private Sub exitBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Exit"
End Sub

Private Sub Form_Load()
 'Set The Database Path
  Database_Path = App.Path & "\Dbase"
 'Set Database Name
  Database_Name = "Family.FM2"
 'Set The Database Password
  Database_Password = EncryptText("SmileyOmar", "Jesus")

  Login_Attempts = 0
  LOGNAME.Text = ""
  PWORD.Text = ""
  AccLevel.Clear
  AccLevel.AddItem "Administrator"
  AccLevel.AddItem "User"
  AccLevel.ListIndex = 0
   
  WriteIniFile App.Path & "\Family2.ini", "DATABASE", "PATH", Database_Path
  WriteIniFile App.Path & "\Family2.ini", "DATABASE", "FileName", Database_Name
     
  If InitLogin <> True Then
     MsgBox "Unable to load database. View the file Readme.txt for mor Information", vbCritical + vbOKOnly
     End
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set frmAbout = Nothing
  Set frmEdit = Nothing
  Set frmHelp = Nothing
  Set frmLinks = Nothing
  Set frmMain = Nothing
  Set frmProfile = Nothing
  Set frmSearch = Nothing
  Set frmSplash = Nothing
  Set frmLogin = Nothing
  End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = ""
End Sub


Private Sub LOGNAME_KeyPress(KeyAscii As Integer)
 'Prevent the user from entering Char(39) [']
  If KeyAscii = 39 Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub

Private Sub LOGNAME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "LoginName :" & LOGNAME.Text
End Sub

Private Sub PWORD_KeyPress(KeyAscii As Integer)
 'Prevent the user from entering Char(39) [']
  If KeyAscii = 39 Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
  
  If KeyAscii = 13 Then
     Call btnLogin_Click
  End If
End Sub

Private Sub PWORD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Password"
End Sub

Private Function InitLogin() As Boolean
  Dim Msg1 As VbMsgBoxResult
  
 'Check if the database exist
  If Database_Found(Database_Path) = True Then
    'Check the database to see if good
     If Database_Ok(Database_Path & "\" & Database_Name) = True Then
        InitLogin = True
        Exit Function
       Else
       InitLogin = False
       Exit Function
     End If
     
    Else ' Database_Found(Database_Path) <> True
     
     Msg1 = MsgBox("The database file was not found in it's destinated path [" & Database_Path & "]. Do you want to recreate the database file?", vbInformation + vbYesNo)
     If Msg1 = vbYes Then
       'Check if the database directory exist
        If DirectoryExist(Database_Path) = True Then
          'Recreate the database
           If Recreate_DB = True Then
             'Check if the database is ok
              If Database_Ok(Database_Path & "\" & Database_Name) = True Then
                 WriteIniFile App.Path & "\Family2.ini", "DATABASE", "Created", Format(Now, "Long Date")
                 InitLogin = True
                 Exit Function
                Else 'Database_Ok <> True
                 InitLogin = False
                 Exit Function
              End If
             Else 'Recreate_DB <> True
              InitLogin = False
              Exit Function
           End If ' Recreate
          Else 'DirectoryExist <> True
          'Create the database directory
           Call CreateNewDirectory(Database_Path & "\")
          'Recreate the database
           If Recreate_DB = True Then
             'Check if the database if ok
              If Database_Ok(Database_Path & "\" & Database_Name) = True Then
                 WriteIniFile App.Path & "\Family2.ini", "DATABASE", "Created", Format(Now, "Long Date")
                 InitLogin = True
                 Exit Function
                Else 'Database_Ok <> True
                 InitLogin = False
                 Exit Function
              End If
             Else 'Recreate_DB <> True
              InitLogin = False
              Exit Function
           End If 'Recreate
        End If 'Directory
     End If 'Msg1
  End If 'Database_Found(Database_Path) = True
End Function
