VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6405
   FillColor       =   &H80000012&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton AboutCloseBtn1 
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
      Height          =   300
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   " ... Close ... "
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmAbout.frx":0442
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Image Image3 
         Height          =   1605
         Left            =   120
         Picture         =   "frmAbout.frx":0448
         Top             =   240
         Width           =   1470
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2040
         Picture         =   "frmAbout.frx":12D7
         Top             =   720
         Width           =   3585
      End
   End
   Begin VB.Label AuthorURL 
      AutoSize        =   -1  'True
      Caption         =   "http://www.omarswan.cjb.net"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2760
      MouseIcon       =   "frmAbout.frx":279E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   " Visit The Autho'r Web-Site http://www.omarswan.cjb.net "
      Top             =   5280
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "frmAbout.frx":2AA8
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":2DB2
      ToolTipText     =   " Send Email To The Author "
      Top             =   5160
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Private Sub AboutCloseBtn1_Click()
   Unload Me
End Sub

Private Sub AuthorURL_Click()
  OpenURL ("http://www.omarswan.cjb.net")
End Sub

Private Sub Form_Load()
  Text1.Locked = True
  Text1.Text = ""
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & "First I Must Give Thanks & Savoir - Jesus Christ." & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & "SmileyOrange - For Being There For Me For Me." & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & "Shiva And The Staff @ FUZION - http://www.fuzion.com [Thanks Alot]" & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & "Everybody That Contributes" & vbNewLine
  Text1.Text = Text1.Text & "http://www.planetsourcecode.com" & vbNewLine
  Text1.Text = Text1.Text & "http://www.vb-world.net" & vbNewLine
  Text1.Text = Text1.Text & "http://www.vbAccelerator.com" & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & "To God Be The Glory" & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & vbNewLine
  Text1.Text = Text1.Text & "Copyright (c) 2000 OmarSwan Software inc." & vbNewLine
End Sub

Private Sub Image2_Click()
   Send_Email_To ("omarswan@yahoo.com")
End Sub
