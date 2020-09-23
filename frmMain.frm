VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data MainData1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   2890
      TabIndex        =   21
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame4 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         TabIndex        =   31
         Top             =   3840
         Width           =   5055
         Begin VB.CommandButton btnAdd 
            Caption         =   "&Add"
            Height          =   310
            Left            =   3840
            TabIndex        =   12
            ToolTipText     =   " Add A New Record "
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton btnDelete 
            Caption         =   "&Delete"
            Height          =   310
            Left            =   3840
            TabIndex        =   13
            ToolTipText     =   " Delete The Current Record "
            Top             =   720
            Width           =   1095
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3495
            Begin VB.CommandButton btnCancel 
               Caption         =   "&Cancel"
               Height          =   310
               Left            =   1200
               TabIndex        =   10
               ToolTipText     =   " Cancel Changes Made "
               Top             =   300
               Width           =   975
            End
            Begin VB.CommandButton btnSave 
               Caption         =   "&Save"
               Height          =   310
               Left            =   120
               TabIndex        =   9
               ToolTipText     =   " Save Changes Made "
               Top             =   300
               Width           =   975
            End
            Begin VB.CommandButton btnEdit 
               Caption         =   "&Edit"
               Height          =   310
               Left            =   2280
               TabIndex        =   11
               ToolTipText     =   " Edit The Current Record "
               Top             =   300
               Width           =   1095
            End
         End
      End
      Begin VB.ComboBox Relation 
         Height          =   345
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Email 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3000
         Width           =   4095
      End
      Begin VB.TextBox ZipCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   6
         Text            =   "111111111111111"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox City_State 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox Address 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox Telephone 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox Last_Name 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox First_Name 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
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
         Left            =   120
         TabIndex        =   30
         Top             =   3480
         Width           =   675
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
         Left            =   120
         TabIndex        =   29
         Top             =   3090
         Width           =   540
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
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   720
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
         Left            =   100
         TabIndex        =   27
         Top             =   2260
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
         TabIndex        =   26
         Top             =   1890
         Width           =   615
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
         TabIndex        =   25
         Top             =   1520
         Width           =   810
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
         TabIndex        =   24
         Top             =   1200
         Width           =   285
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
         TabIndex        =   23
         Top             =   800
         Width           =   825
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
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   10
      TabIndex        =   19
      Top             =   0
      Width           =   2865
      Begin ComctlLib.TreeView MTView 
         Height          =   5055
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   8916
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
   End
   Begin VB.CommandButton btnLinks 
      Caption         =   "Update &Links..."
      Height          =   330
      Left            =   5400
      TabIndex        =   16
      ToolTipText     =   " Internet Links "
      Top             =   5520
      Width           =   1400
   End
   Begin VB.CommandButton btnProfile 
      Caption         =   "&User Profile..."
      Height          =   330
      Left            =   6900
      TabIndex        =   14
      ToolTipText     =   " User Profile... "
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "&Print / Search..."
      Height          =   330
      Left            =   3840
      TabIndex        =   17
      ToolTipText     =   " Print / Search or Delete Record(s) ... "
      Top             =   5520
      Width           =   1400
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   5895
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11192
            MinWidth        =   10583
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
            TextSave        =   "1/15/00"
            Key             =   "dt"
            Object.Tag             =   ""
            Object.ToolTipText     =   " Current Date "
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "1:10 PM"
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
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   " Close/Loggout "
      Top             =   5520
      Width           =   1095
   End
   Begin Family.TrayArea TrayArea1 
      Left            =   2760
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":030A
            Key             =   "People"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0624
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":093E
            Key             =   "Person2"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Options"
      Begin VB.Menu yjut 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search ..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuUserProfile 
         Caption         =   "&User Profile ..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu kljh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Enable"
         Begin VB.Menu rtyrt 
            Caption         =   "-"
         End
         Begin VB.Menu tytty 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTray 
            Caption         =   "&Minimize to tray"
         End
         Begin VB.Menu rd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAutoSend 
            Caption         =   "&Auto Send Email"
         End
         Begin VB.Menu serw 
            Caption         =   "-"
         End
         Begin VB.Menu dergedr 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuFix 
         Caption         =   "Fix"
         Begin VB.Menu mnuLogout_User 
            Caption         =   "&Logout User ..."
         End
         Begin VB.Menu hd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCheckDBErr 
            Caption         =   "&Check Database For Errors"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuLinks 
         Caption         =   "Links"
         Begin VB.Menu er 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUpDateLinks 
            Caption         =   "&Update Links ..."
         End
         Begin VB.Menu hguyj 
            Caption         =   "-"
         End
         Begin VB.Menu fh 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link1"
            Index           =   1
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link2"
            Index           =   2
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link3"
            Index           =   3
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link4"
            Index           =   4
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link5"
            Index           =   5
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link6"
            Index           =   6
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link7"
            Index           =   7
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link8"
            Index           =   8
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link9"
            Index           =   9
         End
         Begin VB.Menu mnuLink 
            Caption         =   "Link10"
            Index           =   10
         End
         Begin VB.Menu guyj 
            Caption         =   "-"
         End
         Begin VB.Menu ty 
            Caption         =   "-"
         End
      End
      Begin VB.Menu erg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu eer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
      Begin VB.Menu rtyt 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuRestorer 
      Caption         =   "mnuRestorer"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu hlbj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout2 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelp2 
         Caption         =   "&Help"
      End
      Begin VB.Menu gi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Edit_Mode As Boolean
  Dim Currently_Editting As Boolean
  Dim Currently_Adding As Boolean
  Dim TmpRelation As String
 'Stores the name of the last parent node clicked
  Dim Last_Parent As String


Private Sub btnAdd_Click()
 ' On Error GoTo AddErr
  btnAdd.Enabled = False
  Currently_Adding = True
  Currently_Editting = False
  Call Empty_Main_Fields
  Make_Changes (True)
  First_Name.SetFocus
  btnAdd.Enabled = False
  MainData1.Recordset.AddNew
  Exit Sub
AddErr:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Sub

Private Sub btnAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusBar1.Panels(1).Text = "Add A New Record"
End Sub

Private Sub btnCancel_Click()
  On Error GoTo CancelErr
 'Cancel Changes
  MainData1.Recordset.CancelUpdate
 'Refresh
  MainData1.Recordset.Fields.Refresh
  Make_Changes (False)
  Currently_Editting = False
  Currently_Adding = False
  Empty_Main_Fields
 'Used to select and expand the last Parent Node Used
  MTView.Nodes(Last_Parent).Selected = True
  MTView.Nodes(Last_Parent).Expanded = True
  Exit Sub
CancelErr:
  If Err.Number <> 0 Then
     MsgBox "Error Add " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Make_Changes (False)
     Empty_Main_Fields
     Err.Clear
  End If
End Sub


Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Cancel Changes Made"
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Close/Loggout"
End Sub

Private Sub btnDelete_Click()
  Dim DelYN As VbMsgBoxResult
  On Error GoTo DelErr
  
  DelYN = MsgBox("Do you want to delete this record ?", vbQuestion + vbYesNo, "Delete current record")
  If DelYN = vbYes Then
     MainData1.Recordset.Delete
     Load_DB_TO_Treeview Current_LoginName, MTView, ImageList1
     
    'Used to select and expand the last Parent Node Used
     MTView.Nodes(Last_Parent).Selected = True
     MTView.Nodes(Last_Parent).Expanded = True
     
     Make_Changes (False)
     Empty_Main_Fields
  End If
  Exit Sub
DelErr:
  If Err.Number <> 0 Then
     MsgBox "Error Add " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Make_Changes (False)
     Empty_Main_Fields
     Err.Clear
  End If
End Sub

Private Sub btnDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Delete The Current Record"
End Sub

Private Sub btnEdit_Click()
  On Error GoTo EditErr
  TmpRelation = Relation.Text
  Currently_Editting = True
  Call Make_Changes(True)
  First_Name.SetFocus
  MainData1.Recordset.Edit
  Exit Sub
EditErr:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Currently_Editting = False
     Call Make_Changes(False)
     Err.Clear
  End If
End Sub

Private Sub btnEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Edit The Current Record"
End Sub

Private Sub btnLinks_Click()
  TrayArea1.Visible = False
  MainData1.Recordset.Close
  MainData1.Database.Close
  Load frmLinks
  frmLinks.Show
End Sub

Private Sub btnLinks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     PopupMenu mnuLinks
  End If
End Sub

Private Sub btnLinks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Internet Links"
End Sub


Private Sub btnProfile_Click()
  TrayArea1.Visible = False
  MainData1.Recordset.Close
  MainData1.Database.Close
  Load frmProfile
  frmProfile.Show
End Sub

Private Sub btnProfile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "User Profile(s)"
End Sub

Private Sub btnSave_Click()
  Dim MsgRes(1 To 2) As VbMsgBoxResult
  On Error GoTo SaveErr
  
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
 
  MsgRes(1) = MsgBox("Do you want to save the changes made", vbQuestion + vbYesNo)
    
  If MsgRes(1) = vbYes Then
     If (Len(Trim(First_Name.Text)) < 3) Then
         MsgBox "Note: The First Name Field should contain at least 3 characters", vbInformation + vbOKOnly
         Exit Sub
     End If
     
     If Currently_Editting = True Then
        Last_Parent = Relation.Text
        If (TmpRelation <> Relation.Text) Then
           If ChildExist(MTView, Relation.Text, First_Name.Text & "_" & Last_Name.Text) = True Then
              MsgBox First_Name.Text & "_" & Last_Name.Text & " already exists in " & Relation.Text, vbInformation + vbOKOnly
              Exit Sub
             Else
              GoTo Label1
           End If
          Else 'tmpRelation = Relation.Text
           GoTo Label1
        End If
     End If
     
     If Currently_Adding = True Then
        Last_Parent = Relation.Text
       'Search if record already exist
        If ChildExist(MTView, Relation.Text, First_Name.Text & "_" & Last_Name.Text) = True Then
           MsgBox First_Name.Text & "_" & Last_Name.Text & " already exist in " & Relation.Text, vbInformation + vbOKOnly
           Exit Sub
          Else
           Currently_Adding = False
           GoTo Label1
        End If
     End If
          
Label1:
     MainData1.Recordset.Fields("FirstName") = ProperString(Trim(First_Name.Text))
     MainData1.Recordset.Fields("LastName") = ProperString(Trim(Last_Name.Text))
     MainData1.Recordset.Fields("Sex") = Sex.Text
     MainData1.Recordset.Fields("Telephone") = Telephone.Text
     MainData1.Recordset.Fields("Address") = Address.Text
     MainData1.Recordset.Fields("City_State") = City_State.Text
     MainData1.Recordset.Fields("ZipCode") = ZipCode.Text
     MainData1.Recordset.Fields("EmailAddress") = Email.Text
     MainData1.Recordset.Fields("Relation") = Relation.Text
     MainData1.Recordset.Update
     MainData1.Recordset.Fields.Refresh
     
     Make_Changes (False)
     Call Empty_Main_Fields
     Load_DB_TO_Treeview Current_LoginName, MTView, ImageList1
     Call Empty_Main_Fields
     Currently_Editting = False
     
    'Used to select and expand the last Parent Node Used
     MTView.Nodes(Last_Parent).Selected = True
     MTView.Nodes(Last_Parent).Expanded = True
  End If
  Exit Sub
  
SaveErr:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Make_Changes (False)
     Call Empty_Main_Fields
     Err.Clear
  End If
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Save Changes Made"
End Sub

Private Sub btnSearch_Click()
  TrayArea1.Visible = False
  MainData1.Recordset.Close
  MainData1.Database.Close
  Load frmSearch
  frmSearch.Show
End Sub

Private Sub btnSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Print / Search or Delete Record(s)..."
End Sub

Private Sub btnClose_Click()
  Unload Me
End Sub


Private Sub Email_DblClick()
  If (Auto_Send_Email = On_) And Valid_Email_Address(Email.Text) Then
      Send_Email_To (Email.Text)
  End If
End Sub

Private Sub Email_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     Email.Enabled = False
     PopupMenu mnuEnable
     Email.Enabled = True
  End If
End Sub

Private Sub Email_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Len(First_Name.Text) > 0 Then
     If Auto_Send_Email = On_ Then
        If Valid_Email_Address(Email.Text) = True Then
           StatusBar1.Panels(1).Text = "Double Mouse-Click to AutoSend Email To " & First_Name.Text
          Else
           StatusBar1.Panels(1).Text = "Enter A Valid Email Address to AutoSend Email To " & First_Name.Text
        End If
          
       Else
        StatusBar1.Panels(1).Text = "Right Mouse-Click to Enable AutoSend Email To " & First_Name.Text
     End If
    Else
     If Auto_Send_Email = On_ Then
        StatusBar1.Panels(1).Text = "Right Mouse-Click to Disable AutoSend Email"
       Else
        StatusBar1.Panels(1).Text = "Right Mouse-Click to Enable AutoSend Email"
     End If
  End If
End Sub

Private Sub First_Name_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 95) Or (KeyAscii = 39) Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  frmLogin.Hide
  Set frmLogin = Nothing
  Call Load_Links
  Call Init_Main
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     PopupMenu mnuFile
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "My Family Address Book v2.0 by SmileyOmar inc."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim CloseDbase As Database
  Dim CloseRecordset As Recordset
  Dim loggedOut As Boolean
 'On Error Resume Next
 
  loggedOut = False
  
  If (Currently_Editting = True) Or (Currently_Adding = True) Then
     Cancel = True
     Exit Sub
  End If
   
  If MsgBox(Current_LoginName & " are you sure that you want to quit ?", vbQuestion + vbYesNo + vbDefaultButton2, "Do you want to quit " & Current_LoginName & "?") = vbNo Then
      Cancel = True
     Else
     'Close Main Database
      MainData1.Recordset.Close
      MainData1.Database.Close
      
      Set CloseDbase = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
      Set CloseRecordset = CloseDbase.OpenRecordset("SELECT * FROM Users WHERE LoginName = '" & EncryptText(Current_LoginName, Database_Password) & "'")
      CloseRecordset.Fields.Refresh
      loggedOut = False
         
      If CloseRecordset.RecordCount > 0 Then
         CloseRecordset.Edit
         CloseRecordset.Fields("LoggedIn") = False
         CloseRecordset.Update
         CloseRecordset.Close
         CloseDbase.Close
         MsgBox Current_LoginName & " logged out successfully.", vbInformation + vbOKOnly
         loggedOut = True
         WriteIniFile App.Path & "\Family2.ini", Current_LoginName, "Last-Logged-Out", Format(Now, "Long Date")
      End If
       
      If loggedOut = False Then
         If Current_AccessLevel = "Administrator" Then
            MsgBox Current_LoginName & " not logged out successfully.", vbCritical + vbOKOnly
           Else
            MsgBox Current_LoginName & " not logged out successfully. Contact an Administrator.", vbCritical + vbOKOnly
         End If
      End If
          
      Current_LoginName = ""
      Current_AccessLevel = ""
      
      Set frmAbout = Nothing
      Set frmEdit = Nothing
      Set frmHelp = Nothing
      Set frmLinks = Nothing
      Set frmProfile = Nothing
      Set frmSearch = Nothing
      Set frmSplash = Nothing
      Set frmLogin = Nothing
      Set frmMain = Nothing
      End
   End If
End Sub

Private Sub Form_Resize()
  'Minimized
  If Me.WindowState = 1 Then
     If Minimize_To_Tray Then
        Set TrayArea1.Icon = Me.Icon
        TrayArea1.ToolTip = " Double-Click To Restore " & frmMain.Caption & " "
        TrayArea1.Visible = True
        frmMain.Hide
     End If
  End If
End Sub



Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     PopupMenu mnuFile
  End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "My Family Address Book v2.0 by SmileyOmar inc."
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     PopupMenu mnuFile
  End If
End Sub

Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
     PopupMenu mnuFile
  End If
End Sub


Private Sub Last_Name_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 95) Or (KeyAscii = 39) Then
     MsgBox "Sorry, but the character ( " & Chr(KeyAscii) & " ) that is an invalid character", vbInformation + vbOKOnly
     KeyAscii = 0
  End If
End Sub

'Used to update the textboxes and combo boxes
Private Sub UpdateFileds()
  On Error Resume Next
  First_Name.Text = MainData1.Recordset.Fields("FirstName") & ""
  Last_Name.Text = MainData1.Recordset.Fields("LastName") & ""
  If IsNull(MainData1.Recordset.Fields("Sex")) Then
     Sex.ListIndex = 0
    Else
      Sex.Text = MainData1.Recordset.Fields("Sex")
  End If
  Telephone.Text = MainData1.Recordset.Fields("Telephone") & ""
  Address.Text = MainData1.Recordset.Fields("Address") & ""
  City_State.Text = MainData1.Recordset.Fields("City_State") & ""
  ZipCode.Text = MainData1.Recordset.Fields("ZipCode") & ""
  Email.Text = MainData1.Recordset.Fields("EmailAddress") & ""
  If IsNull(MainData1.Recordset.Fields("Relation")) Then
     Relation.ListIndex = 0
    Else
     Relation.Text = MainData1.Recordset.Fields("Relation")
  End If
End Sub

Private Sub mnuAbout_Click()
  Load frmAbout
  frmAbout.Show
End Sub

Private Sub mnuAbout2_Click()
  Load frmAbout
  frmAbout.Show
End Sub

Private Sub mnuAutoSend_Click()
  If mnuAutoSend.Checked = True Then
     Set_Auto_Send_Email (Off_)
     mnuAutoSend.Checked = False
    Else
     Set_Auto_Send_Email (On_)
     mnuAutoSend.Checked = True
  End If
End Sub

Private Sub mnuCheckDBErr_Click()
  '
  
End Sub

Private Sub mnuClose_Click()
  Call btnClose_Click
End Sub

Private Sub mnuEnable_Click()
  If Auto_Send_Email = On_ Then
     mnuAutoSend.Checked = True
    Else
     mnuAutoSend.Checked = False
  End If
  
  If Minimize_To_Tray Then
     mnuTray.Checked = True
    Else
     mnuTray.Checked = False
   End If
End Sub

Private Sub mnuExit_Click()
  Call mnuRestore_Click
  Call btnClose_Click
End Sub

Private Sub mnuFile_Click()
  If Auto_Send_Email = On_ Then
     mnuAutoSend.Checked = True
    Else
     mnuAutoSend.Checked = False
  End If
  
  If Minimize_To_Tray Then
     mnuTray.Checked = True
    Else
     mnuTray.Checked = False
   End If
End Sub

Private Sub mnuHelp_Click()
  Load frmHelp
  frmHelp.Show
End Sub

Private Sub mnuHelp2_Click()
  Load frmHelp
  frmHelp.Show
End Sub

Private Sub mnuLink_Click(Index As Integer)
 'Opens the default internet browser
  OpenURL (mnuLink(Index).Caption)
End Sub

Private Sub mnuLogout_User_Click()
  On Error GoTo FixErr
  Dim OpenFixDB As Long
  'Open Fam-Fix2.exe
   OpenFixDB = Shell(App.Path & "\Fam-Fix2.exe", vbNormalFocus)
 
FixErr:
  If Err <> 0 Then
    MsgBox "Error " & Err.Description, vbCritical + vbOKOnly
    Err.Clear
  End If
End Sub

Private Sub mnuRestore_Click()
  On Error Resume Next
  TrayArea1.Visible = False
  frmMain.WindowState = 0
  frmMain.Show
End Sub

Private Sub mnuSearch_Click()
  Call btnSearch_Click
End Sub

Private Sub mnuTray_Click()
  If mnuTray.Checked = True Then
     mnuTray.Checked = False
     Set_Minimize_To_Tray (False)
    Else
     mnuTray.Checked = True
     Set_Minimize_To_Tray (True)
  End If
End Sub

Private Sub mnuUpDateLinks_Click()
  Call btnLinks_Click
End Sub

Private Sub mnuUserProfile_Click()
  Call btnProfile_Click
End Sub

Private Sub MTView_Collapse(ByVal Node As ComctlLib.Node)
  
  If (Currently_Editting = True) Or (Currently_Adding = True) Then
     Exit Sub
  End If
  
  Select Case Node.Text
   Case "People"
        Call Empty_Main_Fields
        Relation.Text = "Family"
        Last_Parent = "Family"
        Node.Expanded = True
         
    Case "Family", "Spouse", "Friend", "Co_Worker", "Acquaintance"
         Call Empty_Main_Fields
         Relation.Text = Node.Text
         Last_Parent = Node.Text
  End Select
  
End Sub

Private Sub MTView_Expand(ByVal Node As ComctlLib.Node)
  If Currently_Editting = True Then
     Exit Sub
  End If
    
  If Currently_Adding = True Then
     Exit Sub
  End If
  
  Select Case Node.Text
   Case "People"
       Call Empty_Main_Fields
       Relation.Text = "Family"
       Last_Parent = "Family"
         
    Case "Family", "Spouse", "Friend", "Co_Worker", "Acquaintance"
       Call Empty_Main_Fields
       Relation.Text = Node.Text
       Last_Parent = Node.Text
  End Select
End Sub


Private Sub MTView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
End Sub

Private Sub MTView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim t As Node
  Set t = MTView.HitTest(X, Y)
  
  If t Is Nothing Then
     StatusBar1.Panels(1).Text = ""
     Exit Sub
    Else
     StatusBar1.Panels(1).Text = t.Text
  End If
End Sub

Private Sub MTView_NodeClick(ByVal Node As ComctlLib.Node)
  Dim Pos1 As Integer
  Dim FN1 As String 'First Name
  Dim LN1 As String 'Last Name
  Dim RL As String 'Relation
  Dim tmpSQL As String
    
  If Currently_Editting = True Then
     Exit Sub
  End If
    
  If Currently_Adding = True Then
     Exit Sub
  End If
  
  Select Case Node.Text
    Case "People"
         Call Empty_Main_Fields
         Relation.Text = "Family"
         Last_Parent = "Family"
         
    Case "Family", "Spouse", "Friend", "Co_Worker", "Acquaintance"
         Call Empty_Main_Fields
         Relation.Text = Node.Text
         Last_Parent = Node.Text
    
    Case Else
         Pos1 = InStr(1, Node.Text, "_")
         If Pos1 > 0 Then
            Changable (True)
            FN1 = Apostrophe(Mid$(Node.Text, 1, Pos1 - 1))
            LN1 = Apostrophe(Mid$(Node.Text, Pos1 + 1))
            RL = Node.Parent.Text
            tmpSQL = ""
            tmpSQL = "SELECT * FROM " & Current_LoginName
            tmpSQL = tmpSQL & " WHERE FirstName = '" & FN1 & "'"
            tmpSQL = tmpSQL & " and LastName = '" & LN1 & "'"
            tmpSQL = tmpSQL & " and Relation = '" & RL & "'"
            MainData1.RecordSource = tmpSQL
            MainData1.Refresh
            Call UpdateFileds
            Changable (True)
            Last_Parent = Node.Parent.Text
          End If
  End Select
End Sub

Public Function Changable(Emode As Boolean)
  btnEdit.Enabled = Emode
  btnDelete.Enabled = Emode
End Function

Public Function Make_Changes(CMODE As Boolean)
  btnSave.Enabled = CMODE
  btnCancel.Enabled = CMODE
  
 If CMODE = True Then
   'UnLock Fields
    Sex.Locked = False
    Relation.Locked = False
    First_Name.Locked = False
    Last_Name.Locked = False
    Telephone.Locked = False
    Address.Locked = False
    City_State.Locked = False
    ZipCode.Locked = False
    Email.Locked = False
     
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    btnAdd.Enabled = False
    btnClose.Enabled = False
  End If
  
  If CMODE = False Then
    'Lock Fields
     Sex.Locked = True
     Relation.Locked = True
     First_Name.Locked = True
     Last_Name.Locked = True
     Telephone.Locked = True
     Address.Locked = True
     City_State.Locked = True
     ZipCode.Locked = True
     Email.Locked = True
          
     btnAdd.Enabled = True
     btnClose.Enabled = True
  End If
  
  
End Function

Private Sub Empty_Main_Fields()
  Sex.ListIndex = 0
  Relation.Text = Last_Parent
  First_Name.Text = ""
  Last_Name.Text = ""
  Telephone.Text = ""
  Address.Text = ""
  City_State.Text = ""
  ZipCode.Text = ""
  Email.Text = ""
  Changable (False)
End Sub

Private Sub TrayArea1_DblClick()
  Call mnuRestore_Click
End Sub

Private Sub TrayArea1_MouseDown(Button As Integer)
  If Button = 2 Then
     PopupMenu mnuRestorer
  End If
End Sub


Public Sub Init_Main()
  btnSave.Enabled = False
  btnCancel.Enabled = False
  On Error GoTo InitError
 '-------------------------------------------
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
 '-------------------------------------------
 '-------------------------------------------
  If (Table_Ok(Database_Path & "\" & Database_Name, Current_LoginName) = False) Then
     MsgBox "Error, unable to load " & Current_LoginName & "'s database", vbCritical + vbOKOnly
     Exit Sub
  End If
  
  
  Load_DB_TO_Treeview Current_LoginName, MTView, ImageList1
  Last_Parent = "Family"
  MTView.Nodes("Root").Selected = True
  MTView.Nodes("Root").Expanded = True
 '-------------------------------------------
 '-------------------------------------------
  MainData1.DatabaseName = Database_Path & "\" & Database_Name
  MainData1.Connect = ";pwd=" & Database_Password
  MainData1.RecordSource = "SELECT * FROM " & Current_LoginName
  MainData1.Refresh
  Call Empty_Main_Fields
  frmMain.Caption = "Family Address Book v2 - [" & Current_AccessLevel & " - " & Current_LoginName & "]"
  
InitError:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Sub
