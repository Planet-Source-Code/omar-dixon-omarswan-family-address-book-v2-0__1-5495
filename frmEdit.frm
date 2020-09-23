VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   20
      TabIndex        =   13
      Top             =   0
      Width           =   5520
      Begin VB.ComboBox Sex 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox First_Name 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox Last_Name 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox Telephone 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox Address 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox City_State 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox ZipCode 
         Height          =   315
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   6
         Text            =   "111111111111111"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Email 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2760
         Width           =   4095
      End
      Begin VB.ComboBox Relation 
         Height          =   345
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   21
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   20
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   19
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   18
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "City-State"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Zip-Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   16
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Em@il"
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
         Left            =   100
         TabIndex        =   15
         Top             =   2850
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Relation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   100
         TabIndex        =   14
         Top             =   3240
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   3480
      TabIndex        =   24
      Top             =   3600
      Width           =   2055
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   20
      TabIndex        =   23
      Top             =   3600
      Width           =   2055
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public Edit_SQL As String
  Dim Edit_Database As Database
  Dim Edit_Recordset As Recordset
 'Used for temporary storage
  Dim TmpFName As String
  Dim TmpLName As String
  Dim TmpRelation As String

Private Sub btnCancel_Click()
  Unload Me
End Sub

Private Sub btnDelete_Click()
 'Delete
  On Error GoTo DelErr
  Dim DelMSG As VbMsgBoxResult
  DelMSG = MsgBox(" Are You Sure That You Want To Delete This Record", vbQuestion + vbYesNo)
  If DelMSG = vbYes Then
     Edit_Recordset.Delete
  End If
  Unload Me
DelErr:
  If Err.Number <> 0 Then
     MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
     Unload Me
  End If
End Sub

Private Sub btnEdit_Click()
  'Edit
   Call Editing
   First_Name.SetFocus
End Sub

Private Sub btnSave_Click()
  Dim SaveMSG As VbMsgBoxResult
  On Error GoTo SaveErr
  SaveMSG = MsgBox("Do You Want To Save The Changes", vbQuestion + vbYesNo)
  
  If SaveMSG = vbYes Then
    'Check For Invalid Characters ( ' & _ )
     If InStr(First_Name.Text, "'") > 0 Then
        MsgBox " Please Remove the Apostrophe(/s) ['] from the First Name field", vbInformation + vbOKOnly
        Exit Sub
       Else
       If InStr(First_Name.Text, "_") > 0 Then
          MsgBox " Please Remove the Underscore(/s) [_] from the First Name field", vbInformation + vbOKOnly
          Exit Sub
       End If
     End If
      
     If InStr(Last_Name.Text, "'") > 0 Then
        MsgBox " Please Remove the Apostrophe(/s) ['] from the Last Name field", vbInformation + vbOKOnly
        Exit Sub
       Else
        If InStr(Last_Name.Text, "_") > 0 Then
           MsgBox " Please Remove the Underscore(/s) [_] from the Last Name field", vbInformation + vbOKOnly
           Exit Sub
        End If
     End If
      
     If Len(Trim(Last_Name.Text)) < 1 Then
        Last_Name.Text = "Unknown"
     End If

     First_Name.Text = Trim(First_Name.Text)
     If Len(Trim(First_Name.Text)) < 3 Then
        MsgBox "Note: The First Name Field should contain at least 3 characters", vbInformation + vbOKOnly
        Exit Sub
     End If
        
    'check if the record already exists
     If (TmpFName = First_Name.Text) And (TmpLName = Last_Name.Text) And (TmpRelation = Relation.Text) Then
        Edit_Recordset.Edit
        Edit_Recordset.Fields("FirstName") = First_Name.Text
        Edit_Recordset.Fields("LastName") = Last_Name.Text
        Edit_Recordset.Fields("Sex") = Sex.Text
        Edit_Recordset.Fields("Relation") = Relation.Text
        Edit_Recordset.Fields("Telephone") = Telephone.Text
        Edit_Recordset.Fields("Address") = Address.Text
        Edit_Recordset.Fields("City_State") = City_State.Text
        Edit_Recordset.Fields("ZipCode") = ZipCode.Text
        Edit_Recordset.Fields("EmailAddress") = Email.Text
        Edit_Recordset.Update
        Edit_Recordset.Close
       Else
        If Record_Exist(Current_LoginName, First_Name.Text, Last_Name.Text, Relation.Text) = True Then
           MsgBox "Sorry " & Current_LoginName & ", but this record already exists.", vbInformation + vbOKOnly
           Exit Sub
          Else
           Edit_Recordset.Edit
           Edit_Recordset.Fields("FirstName") = First_Name.Text
           Edit_Recordset.Fields("LastName") = Last_Name.Text
           Edit_Recordset.Fields("Sex") = Sex.Text
           Edit_Recordset.Fields("Relation") = Relation.Text
           Edit_Recordset.Fields("Telephone") = Telephone.Text
           Edit_Recordset.Fields("Address") = Address.Text
           Edit_Recordset.Fields("City_State") = City_State.Text
           Edit_Recordset.Fields("ZipCode") = ZipCode.Text
           Edit_Recordset.Fields("EmailAddress") = Email.Text
           Edit_Recordset.Update
           Edit_Recordset.Close
        End If
     End If
  End If
  
  Unload Me
  
SaveErr:
 If Err.Number <> 0 Then
    MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
    Err.Clear
 End If
End Sub


Private Sub First_Name_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 95) Or (KeyAscii = 39) Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  frmEdit_Editting = True
  frmSearch.Enabled = False
  If Init_Edit <> True Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Set The Form NotOntop
 'Call Set_Form_NotOnTop(frmEdit)
  
  Set Edit_Database = Nothing
  Set Edit_Recordset = Nothing
  Edit_SQL = ""

  frmSearch.WindowState = vbNormal
  frmSearch.Show
  frmSearch.Enabled = True
  frmEdit_Editting = False
  frmSearch.SrchBtn_Click
End Sub


Private Function Init_Edit() As Boolean
  On Error GoTo InitErr
  
  btnSave.Enabled = False
  btnCancel.Enabled = False
  btnEdit.Enabled = True
  btnDelete.Enabled = True
 
 'Load the Combos
  Sex.Clear
  Sex.AddItem "Male"
  Sex.AddItem "Female"
  Sex.ListIndex = 0
   
  Relation.Clear
  Relation.AddItem "Family"
  Relation.AddItem "Spouse"
  Relation.AddItem "Friend"
  Relation.AddItem "Co_Worker"
  Relation.AddItem "Acquaintance"
  Relation.ListIndex = 0
 '-------------------------------------------
 '-------------------------------------------
 'Lock Fields
  First_Name.Locked = True
  Last_Name.Locked = True
  Sex.Locked = True
  Relation.Locked = True
  Telephone.Locked = True
  Address.Locked = True
  City_State.Locked = True
  ZipCode.Locked = True
  Email.Locked = True
  
 'Open Database
  Set Edit_Database = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
  Set Edit_Recordset = Edit_Database.OpenRecordset("SELECT * FROM " & Current_LoginName & Edit_SQL)
  Edit_Recordset.Fields.Refresh
  Edit_Recordset.MoveFirst
  
  If Edit_Recordset.RecordCount > 0 Then
     Init_Edit = True
    Else
     Init_Edit = False
  End If
  
  First_Name.Text = Edit_Recordset.Fields("FirstName")
  Last_Name.Text = Edit_Recordset.Fields("LastName")
  Sex.Text = Edit_Recordset.Fields("Sex")
  Relation.Text = Edit_Recordset.Fields("Relation")
  Telephone.Text = Edit_Recordset.Fields("Telephone")
  Address.Text = Edit_Recordset.Fields("Address")
  City_State.Text = Edit_Recordset.Fields("City_State")
  ZipCode.Text = Edit_Recordset.Fields("ZipCode")
  Email.Text = Edit_Recordset.Fields("EmailAddress")
  
  TmpFName = First_Name.Text
  TmpLName = Last_Name.Text
  TmpRelation = Relation.Text

  Me.Caption = Current_LoginName & "[Edditing-" & First_Name.Text & "]"
InitErr:
 If Err.Number <> 0 Then
    MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
    Init_Edit = False
    Err.Clear
 End If
End Function


Private Sub Editing()
  btnSave.Enabled = True
  btnCancel.Enabled = True
  btnEdit.Enabled = False
  btnDelete.Enabled = False

 'Unlock Fields
  First_Name.Locked = False
  Last_Name.Locked = False
  Sex.Locked = False
  Relation.Locked = False
  Telephone.Locked = False
  Address.Locked = False
  City_State.Locked = False
  ZipCode.Locked = False
  Email.Locked = False
End Sub

Private Sub Last_Name_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 95) Or (KeyAscii = 39) Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub
