VERSION 5.00
Begin VB.Form fixLogin2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loggout Database User"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "fixLogin2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton LogExitBtn1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   " Exit "
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   10
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton LoginClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   3
         ToolTipText     =   " Clear Entries "
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton LoginBtn1 
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   2
         ToolTipText     =   " Login "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox LogName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "01234567890123456789"
         ToolTipText     =   " Enter Your Login Name Here "
         Top             =   360
         Width           =   2000
      End
      Begin VB.TextBox Pword 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "01234567890123456789"
         ToolTipText     =   " Enter Your Password Here "
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "fixLogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim CmdSQL As String
  Dim ATTEMPTS As Integer
  Dim Max_Attempts As Integer


Private Sub Form_Load()
  DatabaseName = "Family.FM2"
  
 'Set The Database Path
  Database_Path = App.Path & "\Dbase"
  Database_Password = EncryptText("SmileyOmar", "Jesus")
  
  If Len(Database_Path) = 0 Then
     MsgBox "Unable to load the database path from Family2.ini." & vbNewLine & _
            "Make sure that the file exist and the database path is correct.", vbCritical + vbOKOnly
     End
  End If
     
  ATTEMPTS = 0
  Max_Attempts = 4
  
  If InitCMDLogin <> True Then
      Unload Me
      End
  End If
   
End Sub


Private Sub LogExitBtn1_Click()
  Unload Me
  End
End Sub

Private Sub LoginBtn1_Click()
  Dim CmdDB As Database
  Dim CmdRec As Recordset
  Dim Record_Found As Boolean
  On Error GoTo LogErr
  
  Record_Found = False
  ATTEMPTS = ATTEMPTS + 1
  If (Len(LogName.Text) > 0) And (Len(Pword.Text) > 0) Then
     Set CmdDB = OpenDatabase(Database_Path & "\" & DatabaseName, False, True, ";pwd=" & Database_Password)
     Set CmdRec = CmdDB.OpenRecordset("Users")
     Do While Not CmdRec.EOF
        If CmdRec.Fields("LoginName") = EncryptText(LogName.Text, Database_Password) And _
           CmdRec.Fields("Password") = EncryptText(Pword.Text, Database_Password) And _
           CmdRec.Fields("Accesslevel") = EncryptText("Administrator", Database_Password) Then
           Record_Found = True
           Exit Do
          Else
           CmdRec.MoveNext
        End If
     Loop
      
     If Record_Found = True Then
        Unload Me
        Load frmFix
        frmFix.Show
       Else
        Call LoginClear_Click
        MsgBox "The entries that you have made are invalid" _
             , vbExclamation + vbOKOnly, App.ProductName & " [" & Str(ATTEMPTS) & "/" & Str(Max_Attempts) & "]"
     End If
    Else
     Call LoginClear_Click
     MsgBox "Please make sure that you enter a valid Administrator Login Name and Password"
  End If
      
  If ATTEMPTS = Max_Attempts Then
     MsgBox "Contact your local Administrator for a Login Name and Password.", vbExclamation + vbOKOnly
     Call LogExitBtn1_Click
  End If
 
LogErr:
 If Err.Number <> 0 Then
   MsgBox "Error : " & Err.Description & " " & Err.Number, vbCritical + vbOKOnly
   Err.Clear
 End If
End Sub

Private Sub LoginClear_Click()
  LogName.Text = ""
  Pword.Text = ""
  LogName.SetFocus
End Sub

Private Function InitCMDLogin() As Boolean
  Dim tmpDB As Database
  Dim tmpRec As Recordset
  On Error GoTo initErr
  
  LogName.Text = ""
  Pword.Text = ""
  
  Set tmpDB = OpenDatabase(Database_Path & "\" & DatabaseName, False, True, ";pwd=" & Database_Password)
  Set tmpRec = tmpDB.OpenRecordset("Users")
  tmpRec.Fields.Refresh
  tmpRec.Close
  tmpDB.Close
  InitCMDLogin = True
    
initErr:
  If Err.Number <> 0 Then
    InitCMDLogin = False
    Set tmpDB = Nothing
    Set tmpRec = Nothing
    MsgBox " Unable to open " & DatabaseName & vbNewLine & "Error : " & Err.Description & " " & Err.Number, vbCritical + vbOKOnly
  End If
End Function

Private Sub Pword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call LoginBtn1_Click
  End If
End Sub
