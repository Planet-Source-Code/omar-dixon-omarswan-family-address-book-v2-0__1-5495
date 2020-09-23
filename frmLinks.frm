VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLinks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Links"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLinks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   4440
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9780
            MinWidth        =   9701
            Text            =   "Caption"
            TextSave        =   "Caption"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Current Date"
            TextSave        =   "1/15/00"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   ""
            TextSave        =   "8:19 AM"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton LinksClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   5160
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
         Begin VB.CommandButton lnkCancel 
            Caption         =   "&Cancel"
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   860
         End
         Begin VB.CommandButton lnkSave 
            Caption         =   "&Save"
            Height          =   300
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   860
         End
      End
      Begin Family.TrayArea TrayArea1 
         Left            =   3960
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.CommandButton lnkEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   6360
         TabIndex        =   5
         Top             =   1320
         Width           =   825
      End
      Begin VB.TextBox UrlText 
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   720
         Width           =   3375
      End
      Begin ComctlLib.TreeView URL_TreeView 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6376
         _Version        =   327682
         Indentation     =   706
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6720
         Picture         =   "frmLinks.frx":014A
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "URL"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   480
         Width           =   345
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLinks.frx":0A14
            Key             =   "folders"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLinks.frx":0D2E
            Key             =   "wpage"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuF 
      Caption         =   "t"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu fth 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu drt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public Lnk_Edit As Boolean
 'Stores the name of the last link clicked
  Public Last_Link As String
  


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Close"
End Sub

Private Sub Form_Load()
  StatusBar1.Panels(1).Text = ""
  frmMain.Hide
   
  lnkEdit.Enabled = False
  lnkSave.Enabled = False
  lnkCancel.Enabled = False
  Lnk_Edit = False
  UrlText.Locked = True
   
  frmLinks.Caption = Current_LoginName & "'s Internet Links."
  
  Call Load_TV_Links
End Sub

Private Sub Load_TV_Links()
  Dim URL_NodeX As Node
  Dim Cnt As Integer
  Dim Tmp_Link(1 To 10) As String
  On Error Resume Next
  
  URL_TreeView.Nodes.Clear
  
 'Set The TreeView Image List
  Set URL_TreeView.ImageList = ImageList1
 
 'Add Root Node
  Set URL_NodeX = URL_TreeView.Nodes.Add(, , "Root", "Links", "folders")
 
 'Set Expanded Image for the Root node
  URL_NodeX.ExpandedImage = "folders"
 'Expand root node so we can see what's under it
  URL_NodeX.Expanded = True
     
  For Cnt = 1 To 10
      Tmp_Link(Cnt) = ReadIniFile(App.Path & "\Flinks.ini", Current_LoginName, "Link" & Str(Cnt), "")
      Tmp_Link(Cnt) = Trim(Tmp_Link(Cnt))
      If Len(Tmp_Link(Cnt)) > 0 Then
        'Create a child node under the root node
         Set URL_NodeX = URL_TreeView.Nodes.Add("Root", tvwChild, "Link" & Str(Cnt), Tmp_Link(Cnt), "wpage")
        Else
         Tmp_Link(Cnt) = "Empty Slot"
        'Create a child node under the root node
         Set URL_NodeX = URL_TreeView.Nodes.Add("Root", tvwChild, "Link" & Str(Cnt), Tmp_Link(Cnt), "wpage")
         WriteIniFile App.Path & "\FLinks.ini", Current_LoginName, "Link" & Str(Cnt), Tmp_Link(Cnt)
      End If
  Next Cnt
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = ""
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


Private Sub Form_Unload(Cancel As Integer)
  Load_Links
  frmMain.Show
  frmMain.Enabled = True
  frmMain.Init_Main
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = ""
End Sub

Private Sub LinksClose_Click()
  Unload Me
End Sub

Private Sub LinksClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Close"
End Sub

Private Sub lnkCancel_Click()
  UrlText.Locked = True
  Lnk_Edit = False
  UrlText.Text = ""
  lnkEdit.Enabled = False
  lnkSave.Enabled = False
  lnkCancel.Enabled = False
End Sub

Private Sub lnkCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Cancel Changes"
End Sub

Private Sub lnkEdit_Click()
  lnkEdit.Enabled = False
  UrlText.Locked = False
  Lnk_Edit = True
  
  UrlText.SetFocus
  lnkSave.Enabled = True
  lnkCancel.Enabled = True
End Sub

Private Sub lnkEdit_GotFocus()
  '
End Sub

Private Sub lnkEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Edit [" & UrlText.Text & "]"
End Sub

Private Sub lnkSave_Click()
  WriteIniFile App.Path & "\FLinks.ini", Current_LoginName, Last_Link, UrlText.Text
  UrlText.Locked = True
  Lnk_Edit = False
  UrlText.Text = ""
  lnkEdit.Enabled = False
  lnkSave.Enabled = False
  lnkCancel.Enabled = False
  Call Load_TV_Links
End Sub

Private Sub lnkSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = "Save Changes"
End Sub

Private Sub mnuAbout_Click()
  Load frmAbout
  frmAbout.Show
End Sub

Private Sub mnuExit_Click()
  Call mnuRestore_Click
  Call LinksClose_Click
End Sub

Private Sub mnuHelp_Click()
  Load frmHelp
  frmHelp.Show
End Sub

Private Sub mnuRestore_Click()
  On Error Resume Next
  TrayArea1.Visible = False
  frmLinks.WindowState = 0
  frmLinks.Show
End Sub

Private Sub TrayArea1_DblClick()
  Call mnuRestore_Click
End Sub

Private Sub TrayArea1_MouseDown(Button As Integer)
  If Button = 2 Then
     PopupMenu mnuF
  End If
End Sub

Private Sub URL_TreeView_Collapse(ByVal Node As ComctlLib.Node)
  If Lnk_Edit = True Then
     Node.Expanded = True
     Exit Sub
  End If
End Sub

Private Sub URL_TreeView_Expand(ByVal Node As ComctlLib.Node)
  If Lnk_Edit = True Then
      Exit Sub
  End If
End Sub

Private Sub URL_TreeView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim t As Node
  Set t = URL_TreeView.HitTest(X, Y)
  If t Is Nothing Then
     StatusBar1.Panels(1).Text = ""
     Exit Sub
    Else
     StatusBar1.Panels(1).Text = t.Text
  End If
End Sub

Private Sub URL_TreeView_NodeClick(ByVal Node As ComctlLib.Node)
  If Lnk_Edit = True Then
     lnkEdit.Enabled = False
     Exit Sub
  End If
  
  If Node.Key = "Root" Then
     lnkEdit.Enabled = False
    Else
     lnkEdit.Enabled = True
     Last_Link = Node.Key
     UrlText.Text = Node.Text
  End If
End Sub


Private Sub UrlText_GotFocus()
  UrlText.SelStart = 0
  UrlText.SelLength = Len(UrlText.Text)
End Sub

Private Sub UrlText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.Panels(1).Text = UrlText.Text
End Sub

