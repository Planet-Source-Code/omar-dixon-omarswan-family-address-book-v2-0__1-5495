VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record(s) Print / Search  etc.."
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin Family.TrayArea SrchTrayArea1 
      Left            =   7320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton srchClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   5850
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin ComctlLib.ListView srchListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CheckBox ListAllCheck1 
         Caption         =   "List All Records"
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
         Left            =   4800
         TabIndex        =   13
         Top             =   1510
         Width           =   1575
      End
      Begin VB.CheckBox SexCheck1 
         Caption         =   "Sex"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   1490
         Width           =   615
      End
      Begin VB.CheckBox RelationCheck1 
         Caption         =   "Relation"
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
         Left            =   2880
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox SexCombo1 
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1370
         Width           =   2655
      End
      Begin VB.ComboBox RelCombo1 
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox LNameCheck1 
         Caption         =   "Last Name"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   660
         Width           =   1215
      End
      Begin VB.CheckBox FNameCheck1 
         Caption         =   "First Name"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox LName 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox FName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton SrchBtn 
         Caption         =   "&Search"
         Height          =   300
         Left            =   6960
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7440
         Picture         =   "frmSearch.frx":0442
         Top             =   840
         Width           =   480
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   6210
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11183
            MinWidth        =   9701
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
            TextSave        =   "8:17 AM"
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
      Left            =   6960
      Top             =   5640
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
            Picture         =   "frmSearch.frx":1284
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSearch.frx":159E
            Key             =   "person1"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSearch.frx":18B8
            Key             =   "person2"
         EndProperty
      EndProperty
   End
   Begin VB.Menu m 
      Caption         =   "f"
      Visible         =   0   'False
      Begin VB.Menu srchRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu fty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu fth 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuChanges 
      Caption         =   "Changes"
      Visible         =   0   'False
      Begin VB.Menu jg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Selected Record"
      End
      Begin VB.Menu jhfjv 
         Caption         =   "-"
      End
      Begin VB.Menu kjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Selected Record"
      End
      Begin VB.Menu lih 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteall 
         Caption         =   "Delete &All Records Listed"
      End
      Begin VB.Menu gtjyg 
         Caption         =   "-"
      End
      Begin VB.Menu kgb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSelcted 
         Caption         =   "&Print Selected Record"
      End
      Begin VB.Menu fghgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintAll 
         Caption         =   "Prin&t All Records Listed"
      End
      Begin VB.Menu fgtyhf 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public Search_Database As Database
  Public Search_Recordset As Recordset
  Public Search_Sql As String
  Public Public_Sql As String
  Public lvHeader As ColumnHeader
  Public lvListItems As ListItem
  Dim Record_To_Delete As String


Private Sub FNameCheck1_Click()
  If FNameCheck1.Value = 1 Then
     ListAllCheck1.Value = 0
  End If
End Sub

Private Sub Form_Load()
 frmEdit_Editting = False
 frmMain.Enabled = False
 frmMain.Hide
 Call Load_LV_Header
 
 SexCombo1.Clear
 SexCombo1.AddItem "Male"
 SexCombo1.AddItem "Female"
 SexCombo1.ListIndex = 0
 RelCombo1.Clear
 RelCombo1.AddItem "Family"
 RelCombo1.AddItem "Spouse"
 RelCombo1.AddItem "Friend"
 RelCombo1.AddItem "Co_Worker"
 RelCombo1.AddItem "Acquaintance"
 RelCombo1.ListIndex = 0
 StatusBar1.Panels(1).Text = "My Family Address Book v2.0 :[Record Search]"
 
 FName.Text = ""
 LName.Text = ""
 FNameCheck1.Value = 0
 LNameCheck1.Value = 0
 RelationCheck1.Value = 0
 SexCheck1.Value = 0
End Sub


Private Sub Form_Resize()
  'Minimized
  If Me.WindowState = 1 Then
     If Minimize_To_Tray Then
        Set SrchTrayArea1.Icon = Me.Icon
        SrchTrayArea1.ToolTip = " Double-Click To Restore " & frmSearch.Caption & " "
        SrchTrayArea1.Visible = True
        frmSearch.Hide
     End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Show
  frmMain.Enabled = True
  frmMain.Init_Main
End Sub


Private Sub ListAllCheck1_Click()
 If ListAllCheck1.Value = 1 Then
    FNameCheck1.Value = 0
    LNameCheck1.Value = 0
    RelationCheck1.Value = 0
    SexCheck1.Value = 0
 End If
End Sub


Private Sub LNameCheck1_Click()
  If LNameCheck1.Value = 1 Then
     ListAllCheck1.Value = 0
  End If
End Sub

Private Sub mnuAbout_Click()
  Load frmAbout
  frmAbout.Show
End Sub

Private Sub mnuDelete_Click()
  Dim DelAns As VbMsgBoxResult
  Dim DelTokens() As String
  Dim DNumberOfTokens As Integer
  Dim DelDB As Database
  On Error GoTo DEL_ALL_ERR
  
  DNumberOfTokens = ParseDelimitedString(Record_To_Delete, DelTokens, "_")
  DelAns = MsgBox("Are you sure that you want to delete " & DelTokens(0) & " " & DelTokens(1) & " ?", vbQuestion + vbYesNo)
  
  If DelAns = vbYes Then
     If DNumberOfTokens > 0 Then
        Set DelDB = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
        DelDB.Execute "DELETE FROM " & Current_LoginName & " WHERE FirstName = '" & DelTokens(0) _
                     & "' AND LastName = '" & DelTokens(1) & "' AND Relation = '" & DelTokens(2) & "'"
        DelDB.Close
        Call SrchBtn_Click
     End If
  End If
  Exit Sub
  
DEL_ALL_ERR:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Sub


Private Sub mnuDeleteall_Click()
  Dim DelallAns As VbMsgBoxResult
 ' On Error GoTo DelAllErr
  DelallAns = MsgBox("Are you sure that you want to delete the " & Str(srchListView1.ListItems.Count) & " record(/s) listed?", vbQuestion + vbYesNo)
  
  If DelallAns = vbYes Then
     Dim DelDB As Database
     Dim DelRec As Recordset
    'Open Database
     Set DelDB = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
    'Set DelRec =
     DelDB.Execute "DELETE FROM " & Public_Sql
    'DelRec.Close
     DelDB.Close
     Call SrchBtn_Click
  End If
  Exit Sub
DelAllErr:
  If Err.Number <> 0 Then
     Err.Clear
  End If
End Sub

Private Sub mnuEdit_Click()
  Dim tmpStr As String
  Dim Token() As String
  Dim NumberOfTokens As Integer
  
  NumberOfTokens = ParseDelimitedString(Record_To_Delete, Token, "_")
 
  tmpStr = " WHERE FirstName = '" & Token(0) & "'"
  tmpStr = tmpStr & "AND LastName = '" & Token(1) & "'"
  tmpStr = tmpStr & "AND Relation = '" & Token(2) & "'"
  frmEdit.Edit_SQL = tmpStr
  Me.WindowState = vbMinimized
  Load frmEdit
  frmEdit.Show
End Sub


Private Sub mnuExit_Click()
 Call SrchTrayArea1_DblClick
 Call srchClose_Click
End Sub

Private Sub mnuHelp_Click()
  Load frmHelp
  frmHelp.Show
End Sub

Private Sub mnuPrintAll_Click()
 'Print All Records Listed
  Dim PR_ALL_DB As Database
  Dim PR_ALL_REC As Recordset
  On Error Resume Next
  
  Set PR_ALL_DB = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set PR_ALL_REC = PR_ALL_DB.OpenRecordset("SELECT * FROM " & Public_Sql)
  
  PR_ALL_REC.Fields.Refresh
  PR_ALL_REC.MoveFirst
  
  Printer.Font = "Times New Roman"
  Printer.FontBold = False
  Printer.FontUnderline = True
  Printer.FontSize = 10
  Printer.Print vbNewLine
  PrintCenter (Current_LoginName & "'s " & App.ProductName)
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.Print vbNewLine
  Do While Not PR_ALL_REC.EOF
     Printer.Print Space(6) & "Name : " & ProperString(PR_ALL_REC.Fields("FirstName")) & " " & ProperString(PR_ALL_REC.Fields("LastName"))
     Printer.Print Space(6) & "Sex : " & PR_ALL_REC.Fields("Sex")
     Printer.Print Space(6) & "Telephone : " & PR_ALL_REC.Fields("Telephone") & " "
     Printer.Print Space(6) & "Address : " & PR_ALL_REC.Fields("Address") & ""
     Printer.Print Space(6) & "City-State-ZipCode : " & PR_ALL_REC.Fields("City_State"); "-"; PR_ALL_REC.Fields("ZipCode")
     Printer.Print vbNewLine
     PR_ALL_REC.MoveNext
  Loop
  Printer.EndDoc
End Sub

Sub PrintCenter(PrintString$)
  'print the string in the center of the page
   Printer.CurrentX = (Printer.ScaleWidth / 2) - ((Printer.FontSize * _
                      (TextWidth(PrintString$) / 8.28)) / 2)
   'where the 8.28 is the PC
   'default font size   (where the width of the letters comnes from)
    Printer.Print PrintString$
End Sub


Private Sub mnuPrintSelcted_Click()
 'Print Selected Record
  Dim PStr As String
  Dim Token() As String
  Dim NumberOfPTokens As Integer
  Dim PDB As Database
  Dim PRec As Recordset
  On Error Resume Next
  
  NumberOfPTokens = ParseDelimitedString(Record_To_Delete, Token, "_")
  PStr = "SELECT * FROM " & Current_LoginName & " WHERE FirstName = '" & Token(0) & "'"
  PStr = PStr & "AND LastName = '" & Token(1) & "'"
  PStr = PStr & "AND Relation = '" & Token(2) & "'"
  
  Set PDB = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set PRec = PDB.OpenRecordset(PStr)
  PRec.Fields.Refresh
  PRec.MoveFirst
  
  Printer.Font = "Times New Roman"
  Printer.FontBold = False
  Printer.FontUnderline = True
  Printer.FontSize = 10
  Printer.Print vbNewLine
  PrintCenter (Current_LoginName & "'s " & App.ProductName)
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.Print vbNewLine
  Do While Not PRec.EOF
     Printer.Print vbNewLine
     Printer.Print Space(6) & "Name : " & ProperString(PRec.Fields("FirstName")) & " " & ProperString(PRec.Fields("LastName"))
     Printer.Print Space(6) & "Sex : " & PRec.Fields("Sex")
     Printer.Print Space(6) & "Telephone : " & PRec.Fields("Telephone") & " "
     Printer.Print Space(6) & "Address : " & PRec.Fields("Address") & ""
     Printer.Print Space(6) & "City-State-ZipCode : " & PRec.Fields("City_State"); "-"; PRec.Fields("ZipCode")
     Printer.Print vbNewLine
     PRec.MoveNext
  Loop
  Printer.EndDoc
End Sub

Private Sub RelationCheck1_Click()
  If RelationCheck1.Value = 1 Then
     ListAllCheck1.Value = 0
  End If
End Sub

Private Sub SexCheck1_Click()
  If SexCheck1.Value = 1 Then
     ListAllCheck1.Value = 0
  End If
End Sub

Public Sub SrchBtn_Click()
  Dim TMP_KEY As String
  Dim TmpFN As String
  Dim TmpLN As String
  On Error Resume Next

  Search_Sql = ""
  Public_Sql = ""
  TmpFN = ""
  TmpLN = ""
  
  If ListAllCheck1.Value = 1 Then
     Public_Sql = Current_LoginName
     Search_Sql = "SELECT * FROM " & Public_Sql
  End If
     
 'Store First Name and Last Name in Temporary Variables
  TmpFN = Trim(FName.Text)
  TmpLN = Trim(LName.Text)
     
 'Fast Name
  If FNameCheck1.Value = 1 Then
     If Len(TmpFN) < 1 Then
        StatusBar1.Panels(1).Text = "You need to enter a value for [First Name]"
        Exit Sub
       Else
       'Check for apostrophies
        TmpFN = Apostrophe(TmpFN)
        
        Public_Sql = Current_LoginName & " WHERE FirstName LIKE '*" & TmpFN & "*'"
        Search_Sql = "SELECT * FROM " & Public_Sql
     End If
  End If
  
 'Last Name
  If LNameCheck1.Value = 1 Then
     LName.Text = Trim(TmpLN)
     If Len(TmpLN) < 1 Then
        StatusBar1.Panels(1).Text = "You need to enter a value for [Last Name]"
        Exit Sub
       Else
      'Check for apostrophies
        TmpLN = Apostrophe(TmpLN)
        If Len(Search_Sql) > 0 Then
          ' Search_Sql = Search_Sql & " AND LastName LIKE '*" & TmpLN & "*'"
           Public_Sql = Public_Sql & " AND LastName LIKE '*" & TmpLN & "*'"
           Search_Sql = "SELECT * FROM " & Public_Sql
          Else
           Public_Sql = Current_LoginName & " WHERE LastName LIKE '*" & TmpLN & "*'"
           Search_Sql = "SELECT * FROM " & Public_Sql
        End If
     End If
  End If
    
 'Relation
  If RelationCheck1.Value = 1 Then
     If Len(Search_Sql) > 0 Then
        Public_Sql = Public_Sql & " AND Relation = '" & RelCombo1.Text & "'"
        Search_Sql = "SELECT * FROM " & Public_Sql
       Else
        Public_Sql = Current_LoginName & " WHERE Relation = '" & RelCombo1.Text & "'"
        Search_Sql = "SELECT * FROM " & Public_Sql
     End If
  End If
  
 'Sex
  If SexCheck1.Value = 1 Then
     If Len(Search_Sql) > 0 Then
        Public_Sql = Public_Sql & " AND Sex = '" & SexCombo1.Text & "'"
        Search_Sql = "SELECT * FROM " & Public_Sql
       Else
        Public_Sql = Current_LoginName & " WHERE Sex = '" & SexCombo1.Text & "'"
        Search_Sql = "SELECT * FROM " & Public_Sql
     End If
  End If
    
 'Check if the search string is empty
  If Len(Search_Sql) < 1 Then
     StatusBar1.Panels(1).Text = "You need to select one or more of the search options above"
     Exit Sub
  End If
  
  Set Search_Database = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set Search_Recordset = Search_Database.OpenRecordset(Search_Sql)
 
 'ListSubItems 1 = Last Name
 'ListSubItems 2 = Sex
 'ListSubItems 3 = Telephone
 'ListSubItems 4 = Address
 'ListSubItems 5 = City-State
 'ListSubItems 6 = Zip Code
 'ListSubItems 7 = Email Address
 
 'clear the listview
  srchListView1.ListItems.Clear
  
  If Search_Recordset.RecordCount > 0 Then
     Search_Recordset.Fields.Refresh
     Do While Not Search_Recordset.EOF
        TMP_KEY = ProperString(Search_Recordset.Fields("FirstName")) & "_" & _
                    ProperString(Search_Recordset.Fields("LastName")) & "_" & _
                    ProperString(Search_Recordset.Fields("Relation")) & "_" & _
                    Search_Recordset.Fields("Sex")
        
        If Search_Recordset.Fields("Sex") = "Male" Then
           Set lvListItems = srchListView1.ListItems.Add(, TMP_KEY, ProperString(Search_Recordset.Fields("FirstName")), "person1", "person1")
          Else
           Set lvListItems = srchListView1.ListItems.Add(, TMP_KEY, ProperString(Search_Recordset.Fields("FirstName")), "person2", "person2")
        End If
        lvListItems.SubItems(1) = ProperString(Search_Recordset.Fields("LastName"))
        lvListItems.SubItems(2) = Search_Recordset.Fields("Sex")
        lvListItems.SubItems(3) = Search_Recordset.Fields("Telephone")
        lvListItems.SubItems(4) = Search_Recordset.Fields("Address")
        lvListItems.SubItems(5) = Search_Recordset.Fields("City_State")
        lvListItems.SubItems(6) = Search_Recordset.Fields("ZipCode")
        lvListItems.SubItems(7) = Search_Recordset.Fields("EmailAddress")
        lvListItems.SubItems(8) = Search_Recordset.Fields("Relation")
        Search_Recordset.MoveNext
     Loop
     StatusBar1.Panels(1).Text = Str(Search_Recordset.RecordCount) & " record(s) found in the last search"
    'Close the recordset and the database
     Search_Recordset.Close
     Search_Database.Close
    Else
     StatusBar1.Panels(1).Text = "No Match Found"
  End If
End Sub

Private Sub SrchBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Search"
End Sub

Private Sub srchClose_Click()
   Unload Me
End Sub


Private Sub srchClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Close"
End Sub


Private Sub srchListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  With srchListView1
       If .SortKey <> ColumnHeader.Index - 1 Then
          .SortKey = ColumnHeader.Index - 1
          .SortOrder = lvwAscending
         Else
          If .SortOrder = lvwAscending Then
             .SortOrder = lvwDescending
            Else
             .SortOrder = lvwAscending
          End If
       End If
       .Sorted = True
  End With
 End Sub


Private Sub Load_LV_Header()
  srchListView1.ListItems.Clear
  Set lvHeader = Nothing
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C1", "First Name", 2500, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C2", "Last Name", 2500, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C3", "Sex", 1000, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C4", "Telephone #", 1300, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C5", "Address", 3000, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C6", "City-State", 3000, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C7", "Zip Code", 1000, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C8", "Email Address", 2000, lvwColumnLeft)
  Set lvHeader = srchListView1.ColumnHeaders.Add(, "C9", "Relationship", 1000, lvwColumnLeft)
End Sub


Private Sub srchListView1_ItemClick(ByVal item As ComctlLib.ListItem)
 Dim Select_State As Long
 Select_State = 1
'set full row select
 Call SendMessage(srchListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, Select_State)
 
End Sub

Private Sub srchListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim t As ListItem
  Dim Select_State As Long
  If Button = 2 Then
     Set t = srchListView1.HitTest(X, Y)
     If t Is Nothing Then
        Exit Sub
       Else
        srchListView1.ListItems(t.Index).Selected = True
        Record_To_Delete = srchListView1.ListItems(t.Index).Key
        Select_State = 1
       'set full row select
        Call SendMessage(srchListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, Select_State)
        PopupMenu mnuChanges
     End If
  End If
End Sub


Private Sub srchListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim t As ListItem
  Set t = srchListView1.HitTest(X, Y)
  If t Is Nothing Then
     StatusBar1.Panels(1).Text = ""
     Exit Sub
    Else
     StatusBar1.Panels(1).Text = t.Text
  End If
End Sub


Private Sub srchRestore_Click()
  Call SrchTrayArea1_DblClick
End Sub


Private Sub SrchTrayArea1_DblClick()
  On Error Resume Next
  If frmEdit_Editting = True Then
     Exit Sub
  End If
  SrchTrayArea1.Visible = False
  frmSearch.WindowState = 0
  frmSearch.Show
End Sub

Private Sub SrchTrayArea1_MouseDown(Button As Integer)
  If Button = 2 Then
     PopupMenu m
  End If
End Sub

