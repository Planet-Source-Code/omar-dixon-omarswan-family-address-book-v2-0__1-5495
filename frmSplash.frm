VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4560
      Top             =   0
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2325
      Left            =   120
      Picture         =   "frmSplash.frx":0000
      Top             =   480
      Width           =   7305
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  MakeTransparent frmSplash
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Load frmLogin
  frmLogin.Show
  Screen.MousePointer = 1
  Unload Me
End Sub

Private Sub Image1_Click()
  Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Screen.MousePointer = 11
End Sub

Private Sub Timer1_Timer()
 If Timer1.Interval = 3000 Then
    Unload Me
  End If
End Sub
