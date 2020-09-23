VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin Family.TrayArea TrayArea1 
      Left            =   1920
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Data Profile_Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Omar\Omar's Projects\NewProj\Family.FM1"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Users"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton profClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   3240
         TabIndex        =   15
         Top             =   3200
         Width           =   2055
         Begin VB.CommandButton pCancel 
            Caption         =   "&Cancel"
            Height          =   300
            Left            =   1080
            TabIndex        =   17
            Top             =   200
            Width           =   855
         End
         Begin VB.CommandButton pSave 
            Caption         =   "&Save"
            Height          =   300
            Left            =   120
            TabIndex        =   16
            Top             =   200
            Width           =   855
         End
      End
      Begin VB.CommandButton pAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   5520
         TabIndex        =   14
         Top             =   3180
         Width           =   975
      End
      Begin VB.ComboBox usrCStatus 
         Height          =   345
         Left            =   4320
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton pDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   9
         Top             =   3510
         Width           =   975
      End
      Begin VB.CommandButton pEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   8
         Top             =   2830
         Width           =   975
      End
      Begin VB.ComboBox usrAccLvl 
         Height          =   345
         Left            =   4320
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox usrPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4320
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   " njtyghuj"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox usrName1 
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "Hellerytyr5 ty"
         Top             =   600
         Width           =   2175
      End
      Begin ComctlLib.TreeView TVUsers 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   6376
         _Version        =   327682
         Indentation     =   706
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6000
         Picture         =   "frmProfile.frx":0BC2
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Logged In"
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
         Left            =   3250
         TabIndex        =   10
         Top             =   1920
         Width           =   765
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
         Left            =   3250
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   3250
         TabIndex        =   4
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
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
         Left            =   3250
         TabIndex        =   3
         Top             =   660
         Width           =   870
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   4425
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8441
            MinWidth        =   5292
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
            TextSave        =   "9:01 PM"
            Key             =   "time"
            Object.Tag             =   ""
            Object.ToolTipText     =   " Current Time "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":1784
            Key             =   "Crowd"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":1A9E
            Key             =   "Admins"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":1DB8
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":20D2
            Key             =   "ShowFolders"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":23EC
            Key             =   "OpenFolder"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuRestorer 
      Caption         =   "Restorer"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu tfy 
         Caption         =   "-"
      End
      Begin VB.Menu prAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu prHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu lij 
         Caption         =   "-"
      End
      Begin VB.Menu resExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public TmpUserName As String
  Public TmpAccessLevel As String
  Public TmpPassword As String
  Public Last_AccessLevel As String
  Public Last_NameAccessed As String
  Public pCurrently_Editting As Boolean
  Public pCurrently_Adding As Boolean
  Public LVState As Long


Private Sub Form_Load()
  pCurrently_Editting = False
  pCurrently_Adding = False
  Last_AccessLevel = Current_AccessLevel
  Last_NameAccessed = Current_LoginName
  frmMain.Hide
  Call Init_Profile_DB
  Call Clear_Profile_Fields
 'Call pNo_Changes
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (pCurrently_Editting = True) Or (pCurrently_Adding = True) Then
     Cancel = True
     Exit Sub
  End If

  Load frmMain
  frmMain.Enabled = True
  frmMain.Init_Main
  frmMain.Show
End Sub

Private Sub Form_Resize()
  'Minimized
  If Me.WindowState = 1 Then
     If Minimize_To_Tray Then
        Set TrayArea1.Icon = Me.Icon
        TrayArea1.ToolTip = " Double-Click To Restore " & Me.Caption & " "
        TrayArea1.Visible = True
        Me.Hide
     End If
  End If
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = ""
End Sub

Private Sub mnuRestore_Click()
  On Error Resume Next
  TrayArea1.Visible = False
  frmProfile.WindowState = 0
  frmProfile.Show
End Sub

Private Sub pAdd_Click()
  usrName1.SetFocus
  pCurrently_Adding = True
  pCurrently_Editting = False
  Clear_Profile_Fields
  Call pMaking_Changes
End Sub

Private Sub pAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Add"
End Sub

Private Sub pCancel_Click()
 'Profile_Data1.Recordset.CancelUpdate
 'Call Init_Profile_DB
  Call pNo_Changes
  Call Clear_Profile_Fields
End Sub

Private Sub pCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Cancel"
End Sub

Private Sub pDelete_Click()
  Dim Qs As VbMsgBoxResult
  
 'Check if user trying to remove his/her own record
  If usrName1.Text = Current_LoginName Then
     MsgBox Current_LoginName & " you can't remove your own record.", vbInformation + vbOKOnly
     Exit Sub
  End If
  
 'Check if user logged in
  If UserLoggedIn(usrName1.Text) = True Then
     MsgBox "The Database states that the User [" & usrName1.Text & "] is Currently Logged in." & vbNewLine & _
            "If you are sure that [" & usrName1.Text & "] is currently not logged in" & vbNewLine & _
            "you can correct this problem by using the Menu Option [Logout User] " & vbNewLine & vbNewLine & _
            "This error may have occured because " & usrName1.Text & " did not log-out properly" & vbNewLine & _
            "For more information please view the [Help]", vbCritical + vbOKOnly
     Exit Sub
  End If
  
  Qs = MsgBox("Are you sure that you want to remove " & Last_AccessLevel & " " & Last_NameAccessed, vbQuestion + vbYesNo)
  If Qs = vbYes Then
     If Remove_User(Last_NameAccessed) = True Then
        MsgBox Last_NameAccessed & " has been removed successfully.", vbInformation + vbOKOnly
        Call Init_Profile_DB
        Call pNo_Changes
        Call Clear_Profile_Fields
        Exit Sub
       Else
        MsgBox Last_NameAccessed & " was not removed successfully.", vbInformation + vbOKOnly
        Call Init_Profile_DB
        Call pNo_Changes
        Exit Sub
     End If
    Else
     Call pNo_Changes
     Exit Sub
  End If
End Sub

Private Sub pDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Delete"
End Sub

Private Sub pEdit_Click()
 'Temporarily Store the Name and Access Level
 'of the user to be edited
  usrName1.SetFocus
  TmpUserName = usrName1.Text
  TmpAccessLevel = usrAccLvl.Text
  TmpPassword = usrPassword
  
  usrAccLvl.Locked = True
  
  If UserLoggedIn(usrName1.Text) = True Then
     If Current_LoginName <> usrName1.Text Then
        MsgBox "The user [" & TmpUserName & "] is currently Logged-In." & vbNewLine & vbNewLine & _
               "If you are sure that [" & TmpUserName & "] is currently not Logged-In, use the Menu Option " & vbNewLine & _
               "[Fix - Logout User] to correct this problem." & vbNewLine & vbNewLine & _
               "Please view the [Help] for more information.", vbExclamation + vbOKOnly
        Exit Sub
     End If
  End If
  
 'Set pCurrently_Editing = True
  pCurrently_Editting = True
 'Set pCurrently_Adding = False
  pCurrently_Adding = False
   
 'Administrator
  If Current_AccessLevel = "Administrator" Then
    'Allow the Administrator to change another
    'User's Accesslevel
     If (Current_LoginName <> TmpUserName) Then
         Call pMaking_Changes
         usrAccLvl.Locked = False
         Exit Sub
     End If
     
     
     If (AdminCount > 1) Then
         Call pMaking_Changes
         usrAccLvl.Locked = False
         Exit Sub
        Else 'AdminCount < 1
         Call pMaking_Changes
         usrAccLvl.Locked = True
         Exit Sub
     End If
     
    Else 'User
     Call pMaking_Changes
     usrAccLvl.Locked = True
     Exit Sub
  End If
End Sub


Private Sub pEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Edit"
End Sub

Private Sub prAbout_Click()
  Load frmAbout
  frmAbout.Show
End Sub

Private Sub prHelp_Click()
  Load frmHelp
  frmHelp.Show
End Sub

Private Sub profClose_Click()
   Unload Me
End Sub

Private Sub profClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Close"
End Sub

Private Sub Profile_Data1_Reposition()
  On Error Resume Next
  usrName1.Text = DecryptText(Profile_Data1.Recordset.Fields("LoginName"), Database_Password)
  usrPassword.Text = DecryptText(Profile_Data1.Recordset.Fields("Password"), Database_Password)
  usrAccLvl.Text = DecryptText(Profile_Data1.Recordset.Fields("AccessLevel"), Database_Password)
  usrCStatus.Text = Profile_Data1.Recordset.Fields("LoggedIn")
End Sub


Private Sub pSave_Click()
  usrName1.Text = Trim(usrName1.Text)
  usrPassword.Text = Trim(usrPassword.Text)
  
  If (Len(usrName1.Text) < 4) Or (Len(usrPassword.Text) < 4) Then
     MsgBox "The login name and the password field, should both have at least 4 characters.", vbInformation + vbOKOnly
     Exit Sub
  End If
  
'Check For Apostrophes
 If (InStr(usrName1.Text, "'") > 0) Then
    MsgBox " Please Remove the Apostrophe(/s) ['] from the User Name field", vbInformation + vbOKOnly
    Exit Sub
   Else
    If (InStr(usrPassword.Text, "'") > 0) Then
        MsgBox " Please Remove the Apostrophe(/s) ['] from the Password field", vbInformation + vbOKOnly
        Exit Sub
    End If
 End If

'Don't Allow the users to use the following
 If (usrName1.Text = "Users") Or (usrName1.Text = "User") Or _
    (usrName1.Text = "Administrator") Or (usrName1.Text = "Administrators") Then
    MsgBox "You are not allowed to use " & usrName1.Text & " as a Login Name.", vbInformation + vbOKOnly
    Exit Sub
 End If

'pCurrently_Editing = True
 If pCurrently_Editting = True Then
   'Checks if user name was changed if
   'and if so check if the user already exist
    If (TmpUserName <> usrName1.Text) And (User_Exist(usrName1.Text) = True) Then
       MsgBox "The User Name (" & usrName1.Text & ") already Exist.  Try using a different loginname.", vbInformation + vbOKOnly, usrName1.Text & " already exist."
       Exit Sub
      Else
       Profile_Data1.Recordset.Edit
       Profile_Data1.Recordset.Fields("LoginName") = EncryptText(usrName1.Text, Database_Password)
       Profile_Data1.Recordset.Fields("Password") = EncryptText(usrPassword.Text, Database_Password)
       Profile_Data1.Recordset.Fields("AccessLevel") = EncryptText(usrAccLvl.Text, Database_Password)
       Last_AccessLevel = usrAccLvl.Text
      
      
       If TmpUserName = usrName1.Text Then
              
         'Check if is edditing his or her own record
          If TmpUserName = Current_LoginName Then
            'Update the global variables
             Current_LoginName = usrName1.Text
             Current_Password = usrPassword.Text
             Current_AccessLevel = usrAccLvl.Text
             Profile_Data1.Recordset.Fields("LoggedIn") = True
            Else
             Profile_Data1.Recordset.Fields("LoggedIn") = False
          End If 'If TmpUserName = Current_LoginName Then
                  
          Profile_Data1.Recordset.Update
          Profile_Data1.Refresh
        
          Call Init_Profile_DB
          Call pNo_Changes
          Call Clear_Profile_Fields
          Exit Sub
          
         Else
         
          If Rename_Database_Table(TmpUserName, usrName1.Text) = True Then
             MsgBox usrName1.Text & " has been updated successfully", vbInformation + vbOKOnly
            
            'Check if is edditing his or her own record
             If TmpUserName = Current_LoginName Then
               'Update the global variables
                Current_LoginName = usrName1.Text
                Current_Password = usrPassword.Text
                Current_AccessLevel = usrAccLvl.Text
                Profile_Data1.Recordset.Fields("LoggedIn") = True
               Else
                Profile_Data1.Recordset.Fields("LoggedIn") = False
             End If 'If TmpUserName = Current_LoginName Then

             Profile_Data1.Recordset.Update
             Profile_Data1.Refresh
             Call Init_Profile_DB
             Call pNo_Changes
             Call Clear_Profile_Fields
             Exit Sub
            Else 'Rename db = false
             MsgBox usrName1.Text & " has not been updated successfully", vbCritical + vbOKOnly
             Call Init_Profile_DB
             Call pNo_Changes
             Call Clear_Profile_Fields
             Exit Sub
          End If 'Rename_Database_Table(TmpUserName, usrName1.Text) = True
       End If 'TmpUserName = usrName1.Text
    End If '(TmpUserName <> usrName1.Text) And (User_Exist(usrName1.Text) = True)
    Exit Sub
 End If 'pCurrently_Editing = True
 
 
'pCurrently_Adding
 If pCurrently_Adding = True Then
    If (Table_Exist(usrName1.Text) = False) And (User_Exist(usrName1.Text) = False) Then
       If Create_User(usrName1.Text, usrPassword.Text, usrAccLvl.Text) = True Then
          MsgBox usrName1.Text & " has successfully been added to the database.", vbInformation + vbOKOnly
          Call Init_Profile_DB
          Call pNo_Changes
          Call Clear_Profile_Fields
          Exit Sub
         Else
          MsgBox "Unable To Add " & usrName1.Text & " to the Database.", vbCritical + vbOKOnly
          Call Init_Profile_DB
          Call pNo_Changes
          Call Clear_Profile_Fields
          Exit Sub
       End If
      Else
       MsgBox usrName1.Text & " already exist. Use a different Login Name.", vbInformation + vbOKOnly
       Call Init_Profile_DB
       Call pNo_Changes
       Call Clear_Profile_Fields
       Exit Sub
    End If
    Exit Sub
 End If
End Sub


Private Sub pSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Save"
End Sub

Private Sub resExit_Click()
  Call mnuRestore_Click
  Call profClose_Click
End Sub

Private Sub TrayArea1_DblClick()
  Call mnuRestore_Click
End Sub


Private Sub TrayArea1_MouseDown(Button As Integer)
  PopupMenu mnuRestorer
End Sub

Private Sub TVUsers_Collapse(ByVal Node As ComctlLib.Node)
  If (pCurrently_Editting = True) Or (pCurrently_Adding = True) Then
     Exit Sub
  End If
    
  Select Case Node.Text
   Case "Users"
        Last_AccessLevel = "Administrator"
        Call Clear_Profile_Fields
  
    Case "Administrator", "User"
         Last_AccessLevel = Node.Text
         Call Clear_Profile_Fields
  End Select
End Sub

Private Sub TVUsers_Expand(ByVal Node As ComctlLib.Node)
  If (pCurrently_Editting = True) Or (pCurrently_Adding = True) Then
     Exit Sub
  End If
  
  If (Node.Text = "Administrator") And (Current_AccessLevel = "User") Then
      Node.Expanded = False
      MsgBox Current_LoginName & ", your current access is " & Current_AccessLevel & ". You are not allowed to expand " & Node.Text & ".", vbInformation + vbOKOnly
      Exit Sub
  End If
  
  Select Case Node.Text
   Case "Users"
        Last_AccessLevel = "Administrator"
        Call Clear_Profile_Fields
  
    Case "Administrator", "User"
         Last_AccessLevel = Node.Text
         Call Clear_Profile_Fields
  End Select
End Sub

Private Sub TVUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim t As Node
  Set t = TVUsers.HitTest(X, Y)
  If t Is Nothing Then
     StatusBar1.Panels(1).Text = ""
     Exit Sub
    Else
     StatusBar1.Panels(1).Text = t.Text
  End If
End Sub

Private Sub TVUsers_NodeClick(ByVal Node As ComctlLib.Node)
  If pCurrently_Editting = True Then
     Exit Sub
  End If
  
  If pCurrently_Adding = True Then
     Exit Sub
  End If
 
 'Root
  If (Node.Text = "Users") Then
     Last_AccessLevel = "Administrator"
     Call Clear_Profile_Fields
  End If

  
  If (Node.Text <> "Administrator") And (Node.Text <> "User") And (Node.Text <> "Users") Then
     Last_AccessLevel = Node.Parent
     If (Current_AccessLevel = "User") And (Node.Text = Current_LoginName) Then
        Profile_Data1.RecordSource = "SELECT * FROM Users WHERE LoginName = '" & EncryptText(Node.Text, Database_Password) & "'"
        Profile_Data1.Refresh
        pEdit.Enabled = True
        pDelete.Enabled = True
        Last_NameAccessed = Node.Text
        Last_AccessLevel = "User"
     End If
  
     If (Current_AccessLevel = "User") And (Node.Text <> Current_LoginName) Then
        Last_AccessLevel = "User"
        Call Clear_Profile_Fields
     End If
     
     If (Current_AccessLevel = "Administrator") Then
        Profile_Data1.RecordSource = "SELECT * FROM Users WHERE LoginName = '" & EncryptText(Node.Text, Database_Password) & "'"
        Profile_Data1.Refresh
        pEdit.Enabled = True
        pDelete.Enabled = True
        Last_NameAccessed = Node.Text
        Last_AccessLevel = Node.Parent.Text
     End If
  End If
  
  If (Node.Text = "Administrator") Or (Node.Text = "User") Then
     Last_AccessLevel = Node.Text
     Clear_Profile_Fields
  End If
End Sub


Private Sub usrName1_KeyPress(KeyAscii As Integer)
 'Prevent the user from entering Char(39) [']
  If KeyAscii = 39 Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub


Private Sub Init_Profile_DB()
  TmpUserName = ""
  TmpAccessLevel = ""
  TmpPassword = ""
  
  usrAccLvl.Clear
  usrAccLvl.AddItem "Administrator"
  usrAccLvl.AddItem "User"
  usrAccLvl.ListIndex = 0
  usrCStatus.Clear
  usrCStatus.AddItem False
  usrCStatus.AddItem True
  usrCStatus.ListIndex = 0
  
  If Current_AccessLevel = "Administrator" Then
     pAdd.Enabled = True
    Else
     pAdd.Enabled = False
  End If

  Profile_Data1.DatabaseName = Database_Path & "\" & Database_Name
  Profile_Data1.Connect = ";pwd=" & Database_Password
  Profile_Data1.RecordSource = "SELECT * FROM Users WHERE LoginName = '" & EncryptText(Current_LoginName, Database_Password) & "'"
  Profile_Data1.Refresh
 'Loads The Users Database Into The Tree View
  Load_User_DB_TO_Treeview TVUsers, ImageList1
  
  TVUsers.Nodes(Last_AccessLevel).Expanded = True
  TVUsers.Nodes(Last_AccessLevel).Selected = True
  Me.Caption = " User Profile(s) [" & Current_LoginName & "-" & Current_AccessLevel & "] "
End Sub


Private Sub Clear_Profile_Fields()
  usrName1.Text = ""
  usrName1.Locked = True
  usrPassword.Text = ""
  usrPassword.Locked = True
  usrAccLvl.Text = Last_AccessLevel
  usrAccLvl.Locked = True
  usrCStatus.ListIndex = 0
  usrCStatus.Locked = True
  pEdit.Enabled = False
  pDelete.Enabled = False
  pSave.Enabled = False
  pCancel.Enabled = False
End Sub


Private Sub pMaking_Changes()
'Unlock Fields
 usrName1.Locked = False
 usrPassword.Locked = False

'Enable and Disable Butttons
 pSave.Enabled = True
 pCancel.Enabled = True
 profClose.Enabled = False
 pAdd.Enabled = False
 pEdit.Enabled = False
 pDelete.Enabled = False
 
 If Current_AccessLevel <> "Administrator" Then
    usrAccLvl.Locked = True
 End If
 
 If Current_AccessLevel = "Administrator" Then
    If pCurrently_Editting Then
       If AdminCount > 1 Then
          usrAccLvl.Locked = False
         Else
          usrAccLvl.Locked = True
       End If
    End If
    
    If pCurrently_Adding = True Then
       usrAccLvl.Locked = False
    End If
 End If
  
End Sub


Private Sub pNo_Changes()
'Lock Fields
 usrName1.Locked = True
 usrPassword.Locked = True
 usrAccLvl.Locked = True
'Enable and Disable Butttons
 pSave.Enabled = False
 pCancel.Enabled = False
 profClose.Enabled = True
 pEdit.Enabled = True
 pDelete.Enabled = True
 If (Current_AccessLevel = "Administrator") Then
    pAdd.Enabled = True
   Else
    pAdd.Enabled = False
 End If
 pCurrently_Adding = False
 pCurrently_Editting = False
End Sub

Private Sub usrPassword_KeyPress(KeyAscii As Integer)
 'Prevent the user from entering Char(39) [']
  If KeyAscii = 39 Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub
