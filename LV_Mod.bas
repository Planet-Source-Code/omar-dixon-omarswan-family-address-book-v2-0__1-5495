Attribute VB_Name = "LV_Mod"
Option Explicit
'This Module is used for enhance the listview control


 Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long


 Public Const LVM_FIRST = &H1000
 Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
 Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
 
'USED FOR ENTIRE ROW SELECT
 Public Const LVS_EX_FULLROWSELECT = &H20
'--end block--'


