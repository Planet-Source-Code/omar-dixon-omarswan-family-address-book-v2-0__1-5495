Attribute VB_Name = "Module1"
Option Explicit
'============================================================================================================
 Public Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
 End Type

Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

Public Enum MySwich
       On_
       Off_
End Enum

'Task: Create a multi-level directory structure using CreateDirectory API call
'Declarations
Private Type SECURITY_ATTRIBUTES

nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long




'============================================================================================================
' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
'============================================================================================================
Public Function BrowseForDirectory(hWndOwner As Long, sPrompt As String) As String
  Dim browse_info As BrowseInfo
  Dim item As Long
  Dim dir_name As String
   
  browse_info.hWndOwner = hWndOwner
  browse_info.pidlRoot = 0
  browse_info.sDisplayName = Space$(260)
  browse_info.sTitle = sPrompt
  browse_info.ulFlags = 1 ' Return directory name.
  browse_info.lpfn = 0
  browse_info.lParam = 0
  browse_info.iImage = 0
   
  item = SHBrowseForFolder(browse_info)
  If item Then
     dir_name = Space$(260)
     If SHGetPathFromIDList(item, dir_name) Then
        BrowseForDirectory = Left$(dir_name, InStr(dir_name, Chr$(0)) - 1)
      Else
        BrowseForDirectory = ""
     End If
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Credits to :
'============================================================================================================
Public Function ProperString(strText As String) As String
  Dim FirstChar As String
  Dim TheRest As String
  'Check If String is empty
    If Len(Trim(strText)) <> 0 Then
      'Store The First Character in a string variable
       FirstChar = UCase(Left$(Trim(strText), 1))
      'place the rest of the string in another variable
       TheRest = LCase(Mid$(Trim(strText), 2))
      'combine the two string variables
       ProperString = FirstChar & TheRest
      Else
      'if empty, do essentially nothing
       ProperString = strText
    End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Original Author : By: Gaetan Savoie
'Used To Format a SQL string incase it has an Apostrophe [']
'============================================================================================================
Public Function Apostrophe(sFieldString As String) As String
  If InStr(sFieldString, "'") Then
     Dim iLen As Integer
     Dim ii As Integer
     Dim apostr As Integer
     iLen = Len(sFieldString)
     ii = 1
     
     Do While ii <= iLen
        If Mid$(sFieldString, ii, 1) = "'" Then
           apostr = ii
           sFieldString = Left$(sFieldString, apostr) & "'" & _
           Right$(sFieldString, iLen - apostr)
           iLen = Len(sFieldString)
           ii = ii + 1
         End If
         ii = ii + 1
     Loop
  End If
  Apostrophe = sFieldString
End Function
'============================================================================================================
'============================================================================================================

'============================================================================================================
'Check The Minimize Option
'============================================================================================================
Public Function Minimize_To_Tray() As Boolean
  Dim TmpTray As String
  TmpTray = ReadIniFile(App.Path & "\Family2.ini", "OPTIONS", "Minimize_To_Tray", "")
  If TmpTray = "True" Then
     Minimize_To_Tray = True
    Else
     WriteIniFile App.Path & "\Family2.ini", "OPTIONS", "Minimize_To_Tray", "False"
     Minimize_To_Tray = False
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Used To Set Minimize State
'============================================================================================================
Public Sub Set_Minimize_To_Tray(Set_State As Boolean)
  If Set_State = True Then
     WriteIniFile App.Path & "\Family2.ini", "OPTIONS", "Minimize_To_Tray", "True"
    Else
     WriteIniFile App.Path & "\Family2.ini", "OPTIONS", "Minimize_To_Tray", "False"
  End If
End Sub
'============================================================================================================
'============================================================================================================


'============================================================================================================
'This Function is used to load the internet links from
'flinks.ini to the menu on frmMAIN
'============================================================================================================
Public Sub Load_Links()
  Dim TmpLink(1 To 10) As String
  Dim Cnt As Integer
  
  For Cnt = 1 To 10
      TmpLink(Cnt) = ReadIniFile(App.Path & "\Flinks.ini", Current_LoginName, "Link" & Str(Cnt), "")
      TmpLink(Cnt) = Trim(TmpLink(Cnt))
      If Len(TmpLink(Cnt)) > 0 Then
         frmMain.mnuLink(Cnt).Caption = TmpLink(Cnt)
        Else
         frmMain.mnuLink(Cnt).Caption = "Empty Slot"
         WriteIniFile App.Path & "\FLinks.ini", Current_LoginName, "Link" & Str(Cnt), "Empty Slot"
      End If
  Next Cnt
End Sub
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Checks if the em@il is valid
'============================================================================================================
Public Function Valid_Email_Address(EmailAddress As String) As Boolean
    Valid_Email_Address = EmailAddress Like "*@[A-Z,a-z,0-9]*.*"
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Open INTERNET Link using the default internet browser
'Example : Internet Explorer, Netscape
'============================================================================================================
Public Sub OpenURL(link As String)
  On Error Resume Next
  Dim jmp
 'Check For http://
  If (Left$(link, 7) <> "http://") And (Len(link) > 3) And (Left$(link, 4) <> "http://") Then
     link = "http://" + link
     jmp = Shell("start.exe " & link, vbHide)
    Else
     jmp = Shell("start.exe " & link, vbHide)
  End If
End Sub
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Opens The Default Email Client to send an email
'Example : Outlook, etc..)
'============================================================================================================
Public Sub Send_Email_To(EmailAddress As String)
  On Error Resume Next
  Dim jmpEmail
 'check if the email address is valid
  If Valid_Email_Address(EmailAddress) Then
     jmpEmail = Shell("start.exe mailto:" & EmailAddress, vbHide)
    Else
     MsgBox "Sorry, but [" & EmailAddress & "] is not a valid email address" _
            , vbExclamation + vbOKOnly, "Invalid Email Address"
  End If
End Sub
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Check Family2.ini to see if Auto-Send-Email = On
'============================================================================================================
Public Function Auto_Send_Email() As MySwich
  Dim ASE As String
  
  ASE = ReadIniFile(App.Path & "\Family2.ini", "OPTIONS", "Auto-Send-Email", "")
  If ASE = "On" Then
     Auto_Send_Email = On_
    Else
     WriteIniFile App.Path & "\Family2.ini", "OPTIONS", "Auto-Send-Email", "Off"
     Auto_Send_Email = Off_
  End If
End Function
'============================================================================================================
'============================================================================================================

'============================================================================================================
'Sets the Auto-Send-Email option in Family2.ini
'============================================================================================================
Public Sub Set_Auto_Send_Email(SetIt As MySwich)
  If SetIt = On_ Then
     WriteIniFile App.Path & "\Family2.ini", "OPTIONS", "Auto-Send-Email", "On"
    Else
     WriteIniFile App.Path & "\Family2.ini", "OPTIONS", "Auto-Send-Email", "Off"
  End If
End Sub
'============================================================================================================
'============================================================================================================

'============================================================================================================
'Sub Used to create new directory
'============================================================================================================
Public Sub CreateNewDirectory(NewDirectory As String)
    Dim sDirTest As String
    Dim SecAttrib As SECURITY_ATTRIBUTES
    Dim bSuccess As Boolean
    Dim sPath As String
    Dim iCounter As Integer
    Dim sTempDir As String

    sPath = NewDirectory
    
    If Right(sPath, Len(sPath)) <> "\" Then
        sPath = sPath & "\"
    End If
    
    iCounter = 1
    
    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
        sTempDir = Left(sPath, iCounter)
        sDirTest = Dir(sTempDir)
        iCounter = iCounter + 1
        'create directory
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
    Loop

End Sub
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Checks if directory exist
'============================================================================================================
Public Function DirectoryExist(DirPath As String) As Boolean
    DirectoryExist = Dir(DirPath, vbDirectory) <> ""
End Function
'============================================================================================================
'============================================================================================================
