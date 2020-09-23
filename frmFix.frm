VERSION 5.00
Begin VB.Form frmFix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmFix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnrefresh 
      Caption         =   "&Refresh"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton exitFix 
      Caption         =   "E&xit"
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
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton LogOut 
         Caption         =   "&Log-Out User"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   " Log-Out This User "
         Top             =   2760
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnrefresh_Click()
   Call InitFixDB
End Sub

Private Sub LogOut_Click()
  Dim LUser1 As VbMsgBoxResult
  If Len(List1.Text) > 0 Then
     LUser1 = MsgBox("Are you sure that you want to Log-Out " & List1.Text, vbQuestion + vbYesNo)
     If LUser1 = vbYes Then
         If LogOutUser(List1.Text) Then
            MsgBox List1.Text & " has been Logged-Out Successfully", vbInformation + vbOKOnly
            Call InitFixDB
           Else
            MsgBox List1.Text & " has not been Logged-Out Successfully", vbCritical + vbOKOnly
         End If
     End If
   End If
End Sub

Private Sub exitFix_Click()
  Call closefix
  Unload Me
  End
End Sub

Private Sub Form_Load()
  Call InitFixDB
End Sub

Public Function InitFixDB()
  Dim fix_DB As Database
  Dim fixRec As Recordset
  On Error GoTo InitDbErr
  
  List1.Clear

  Set fix_DB = OpenDatabase(Database_Path & "\" & DatabaseName, False, True, ";pwd=" & Database_Password)
  Set fixRec = fix_DB.OpenRecordset("SELECT * FROM Users", dbOpenForwardOnly)
  
  If fixRec.RecordCount > 0 Then
     Do While Not fixRec.EOF
       If fixRec.Fields("LoggedIN") = True Then
          List1.AddItem DecryptText(fixRec.Fields("LoginName"), Database_Password)
       End If
       fixRec.MoveNext
     Loop
  End If

InitDbErr:
  If Err.Number <> 0 Then
    MsgBox "An error has occured while trying to open " & DatabaseName & vbNewLine & Err.Description & " " & Err.Number, vbCritical + vbOKOnly
  End If
End Function

Public Function closefix()
  DatabaseName = ""
  Database_Password = ""
  
End Function

Public Function LogOutUser(theUserName As String) As Boolean
  Dim rDB As Database
  Dim rRec As Recordset
  Dim rSql As String
  On Error GoTo LoggOutErr
   
  Set rDB = OpenDatabase(Database_Path & "\" & DatabaseName, False, False, ";pwd=" & Database_Password)
  rSql = "SELECT * FROM Users WHERE LoginName = '" & EncryptText(theUserName, Database_Password) & "'"
  Set rRec = rDB.OpenRecordset(rSql)
  
  If rRec.RecordCount > 0 Then
     rRec.Edit
     rRec.Fields("LoggedIn") = False
     rRec.Update
     rRec.Close
     rDB.Close
     LogOutUser = True
    Else
     MsgBox "The User " & theUserName & " was not found", vbCritical + vbOKOnly
     rDB.Close
     rRec.Close
     LogOutUser = False
  End If

LoggOutErr:
  If Err.Number <> 0 Then
    MsgBox "Unable To Logg-Out " & theUserName & vbNewLine & Err.Description & " " & Err.Number, vbCritical + vbOKOnly
    LogOutUser = False
  End If
End Function

