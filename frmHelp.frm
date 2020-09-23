VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Family Address Book"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton hlpClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox hlpText 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":27A2
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Function is used to load a text file into a text box
Function GetTextFromFile(txtFile, txtopen As TextBox)
    Dim sfile As String
    Dim nfile As Integer
    On Error Resume Next
    
    nfile = FreeFile
    sfile = txtFile
    Open sfile For Input As nfile
    txtopen = Input(LOF(nfile), nfile)
    Close nfile
End Function

Private Sub Form_Load()
  hlpText.Locked = False
 'Load Readme.txt into the text box
  Call GetTextFromFile(App.Path & "\README.txt", hlpText)
  hlpText.Locked = True
  hlpText.Enabled = True
End Sub

Private Sub hlpClose_Click()
  Unload Me
End Sub

