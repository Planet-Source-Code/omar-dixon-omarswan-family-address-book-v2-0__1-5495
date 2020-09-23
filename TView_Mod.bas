Attribute VB_Name = "TView_Mod"
Option Explicit

'============================================================================================================
'Used To Load The Database Into a Treview
'============================================================================================================
Public Function Load_DB_TO_Treeview(UserTable As String, TView As TreeView, ImgLst As ImageList) As Boolean
  Dim TrvDbase As Database
  Dim TrvRecSet As Recordset
 'Treeview Node
  Dim NodeX As Node
 'Local Variables
  Dim tmpName As String
  Dim TmpRelation As String
  Dim tCounter As Long
  On Error GoTo LoadTVErr:
   
  TView.Nodes.Clear
  
 'Set The TreeView Image List
  Set TView.ImageList = ImgLst
 
 'Add Root Node
  Set NodeX = TView.Nodes.Add(, , "Root", "People", "People")
 
 'Set Expanded Image for the Root node
  NodeX.ExpandedImage = "People"
 'Expand root node so we can see what's under it
  NodeX.Expanded = True
     
 'Create a child node under the root node called Family
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "Family", "Family", "Person")
 'Create a child node under the root node called Spouse
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "Spouse", "Spouse", "Person")
 'Create a child node under the root node called Friend
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "Friend", "Friend", "Person")
 'Create a child node under the root node called Co_Worker
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "Co_Worker", "Co_Worker", "Person")
 'Create a child node under the root node called Acquaintance
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "Acquaintance", "Acquaintance", "Person")
  
  tCounter = 0
  Set TrvDbase = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set TrvRecSet = TrvDbase.OpenRecordset("SELECT * FROM " & Current_LoginName)
  TrvRecSet.Fields.Refresh
'  TrvRecSet.MoveFirst
  
  If TrvRecSet.RecordCount > 0 Then
   Do
    tCounter = tCounter + 1
    TmpRelation = TrvRecSet.Fields("Relation")
    tmpName = ProperString(TrvRecSet.Fields("FirstName")) & "_" & _
    ProperString(TrvRecSet.Fields("LastName"))
         
    Select Case TmpRelation
      Case "Family"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & Str(tCounter), tmpName, "Person2")
      Case "Spouse"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & Str(tCounter), tmpName, "Person2")
      Case "Friend"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & Str(tCounter), tmpName, "Person2")
      Case "Co_Worker"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & Str(tCounter), tmpName, "Person2")
      Case "Acquaintance"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & Str(tCounter), tmpName, "Person2")
      Case Else 'Add to Acquaintance if No match found
            Set NodeX = TView.Nodes.Add("Acquaintance", tvwChild, TmpRelation & Str(tCounter), tmpName, "Person2")
    End Select
   'Move To the Next Record
    TrvRecSet.MoveNext
   Loop Until TrvRecSet.EOF
  End If
  
  TrvRecSet.Close
  TrvDbase.Close
  Exit Function
  
LoadTVErr:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
'Used To Load The Users Database Into a Treview
'============================================================================================================
Public Function Load_User_DB_TO_Treeview(TView As TreeView, ImgLst As ImageList) As Boolean
  Dim TrvDbase As Database
  Dim TrvRecSet As Recordset
 'Treeview Node
  Dim NodeX As Node
 'Local Variables
  Dim tmpName As String
  Dim TmpRelation As String
  'On Error GoTo LoadTVErr:
   
  TView.Nodes.Clear
  
 'Set The TreeView Image List
  Set TView.ImageList = ImgLst
 
 'Add Root Node
  Set NodeX = TView.Nodes.Add(, , "Root", "Users", "ShowFolders")
 'Set Expanded Image for the Root node
  NodeX.ExpandedImage = "OpenFolder"
 'Set Closed Image for the Root node
'  NodeX.Image = "ClosedFolder"
 'Expand root node so we can see what's under it
  NodeX.Expanded = True
     
 'Create a child node under the root node called Administrator
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "Administrator", "Administrator", "Crowd")
 'Create a child node under the root node called User
  Set NodeX = TView.Nodes.Add("Root", tvwChild, "User", "User", "Crowd")
  
  Set TrvDbase = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
  Set TrvRecSet = TrvDbase.OpenRecordset("SELECT * FROM Users")
  
 ' tCounter = 0
  TrvRecSet.Fields.Refresh
  If TrvRecSet.RecordCount > 0 Then
   Do
   ' tCounter = tCounter + 1
    TmpRelation = DecryptText(TrvRecSet.Fields("AccessLevel"), Database_Password)
    tmpName = DecryptText(TrvRecSet.Fields("LoginName"), Database_Password)
         
    Select Case TmpRelation
      Case "Administrator"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & tmpName, tmpName, "Admins")
      Case "User"
            Set NodeX = TView.Nodes.Add(TmpRelation, tvwChild, TmpRelation & tmpName, tmpName, "Users")
    End Select
   'Move To the Next Record
    TrvRecSet.MoveNext
   Loop Until TrvRecSet.EOF
  End If
  
  TrvRecSet.Close
  TrvDbase.Close
  
LoadTVErr:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
'Used To Check If A Specific Child Exist
'============================================================================================================
Public Function ChildExist(ParentTreeView As TreeView, pNode As String, cNode As String) As Boolean
  Dim nodeChild As Node
  Dim t As TreeView
  Dim i As Integer
  
  ChildExist = False
 
  Set nodeChild = ParentTreeView.Nodes(pNode).Child
  
  Do While Not (nodeChild Is Nothing)
    If UCase(nodeChild.Text) = UCase(cNode) Then
       ChildExist = True
       'MsgBox nodeChild.Text
       Exit Function
    End If
    Set nodeChild = nodeChild.Next
  Loop
End Function
'============================================================================================================
'============================================================================================================

