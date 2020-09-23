Attribute VB_Name = "DataBase1"
Option Explicit
 'Declaring Global Variables
  Public NewUser_LoginName As String
  Public NewUser_Password As String
  Public NewUser_AccessLevel As String
  
 'Stores The Login Name Of The User Currently Logged In
  Public Current_LoginName As String
 'Stores The Password Of The User Currently Logged In
  Public Current_Password As String
 'Stores The Access Level Of The User Currently Logged In
  Public Current_AccessLevel As String
 'Stores The Name Of The Database
  Public Database_Name As String
 'Stores The Database Password
  Public Database_Password As String
 'Stores The Database Path
  Public Database_Path As String
 'Tells If the A New database was found
  Public New_Database_Found As Boolean
 'Stores The New Database Location
  Public New_Database_Location As String
 'Tells If A Record is been edited in frmEdit
  Public frmEdit_Editting As Boolean
 

'============================================================================================================
'Call Recreate Database an Add_To_User_Dbase
'============================================================================================================
Public Function Recreate_DB() As Boolean
  If DirectoryExist(Database_Path) <> True Then
     Recreate_DB
     Exit Function
  End If
 'Recreate The Database
  If Recreate_Database_File(Database_Path & "\" & Database_Name) = True Then
    'Creates A New Database Called Admin
     If Create_New_User_Dbase("Admin") = True Then
       'Add Defafault Values To the "Admin" Table
        If Add_To_User_Dbase("Admin", "Admin", "Administrator") = True Then
          'Everything seems OK So Set Recreate_DB = True
           Recreate_DB = True
           Exit Function
          Else 'Unable To Add "Admin" to the Database
           Recreate_DB = False
           Exit Function
        End If
      Else 'Unable To Create_New_User_Dbase("Admin")
        Recreate_DB = False
        Exit Function
     End If
    Else 'Unable To Recreate_Database_File(Database_Path & "\" & Database_Name)
     Recreate_DB = False
     Exit Function
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
'Used To Recreate The Database File
'============================================================================================================
Private Function Recreate_Database_File(dbName As String) As Boolean
  Dim MsgAns As VbMsgBoxResult
  Dim tdfNewTable As TableDef
  Dim newDb As Database
  On Error GoTo CreateDB_Err
  
 'Check If The Database File Exist
  If Dir(dbName) <> "" Then
     MsgAns = MsgBox("Database - " & dbName & " already exist." & _
                     vbNewLine & "Are you sure that you want to recreate it?", vbCritical + vbYesNo, "Create Database")
                     
          
     If MsgAns = vbYes Then
       'Delete File
        Kill (dbName)
       Else
        Recreate_Database_File = False
        Exit Function
     End If
  End If
    
 'Create A New Database "PasswordProtected"
  Set newDb = CreateDatabase(dbName, dbLangGeneral & ";pwd=" & Database_Password)
     
 'Create a new tabe called "Users"
 'Used to store informations about the Users
  Set tdfNewTable = newDb.CreateTableDef("Users")
       
 'Add Fields to the "Users" Table
  With tdfNewTable
    .Fields.Append .CreateField("LoginName", dbText, 20)
    .Fields.Append .CreateField("Password", dbText, 20)
    .Fields.Append .CreateField("AccessLevel", dbText, 13)
    .Fields.Append .CreateField("LoggedIn", dbBoolean)
  End With
      
 'Add the Users table to the database
  newDb.TableDefs.Append tdfNewTable
  newDb.TableDefs.Refresh
    
 'Close The Database
  newDb.Close
  
  Recreate_Database_File = True
  Exit Function
  
CreateDB_Err:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str$(Err.Number) & " Creating Database." & Err.Description & vbNewLine & _
            "Make sure that the database is not open by another user or application", vbCritical + vbOKOnly
     Recreate_Database_File = False
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
' Adds The New User Records To The "Users" Table
'===========================================================================================================
Public Function Add_To_User_Dbase(New_UName As String, New_UPass As String, New_UAccess As String) As Boolean
  Dim TmpDb As Database
  Dim TmpRecSet As Recordset
  On Error GoTo AddToUserDBErr
  
  Add_To_User_Dbase = False
  
 'Open Database & Table (Shared)
  Set TmpDb = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
  Set TmpRecSet = TmpDb.OpenRecordset("Users")
                 
 'Add New Fields With Encryption
  TmpRecSet.AddNew
  TmpRecSet.Fields("LoginName") = EncryptText(New_UName, Database_Password)
  TmpRecSet.Fields("Password") = EncryptText(New_UPass, Database_Password)
  TmpRecSet.Fields("AccessLevel") = EncryptText(New_UAccess, Database_Password)
  TmpRecSet.Fields("LoggedIn") = False
  TmpRecSet.Update
  
 'Closing
  TmpDb.Close
  Set TmpRecSet = Nothing
  Add_To_User_Dbase = True
  
AddToUserDBErr:
  If Err.Number <> 0 Then
     MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly
     Add_To_User_Dbase = False
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
' Creates a new database for the new user
'============================================================================================================
Public Function Create_New_User_Dbase(UName As String) As Boolean
  Dim NewUserDb As Database
  Dim NewUserTable As TableDef
  Dim NewField(1 To 9) As Field
  Dim i As Integer
  On Error GoTo CreateUserErr
   
 'Set Create_New_User_Dbase = False
  Create_New_User_Dbase = False
 'Open Database
  Set NewUserDb = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
            
 'Create a new tabe
  Set NewUserTable = NewUserDb.CreateTableDef(UName)
  
 'Add Fields to the New Table
  Set NewField(1) = NewUserTable.CreateField("FirstName", dbText, 50)
      NewField(1).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(1)
    
  Set NewField(2) = NewUserTable.CreateField("LastName", dbText, 50)
      NewField(2).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(2)
   
  Set NewField(3) = NewUserTable.CreateField("Sex", dbText, 6)
      NewField(3).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(3)
                      
  Set NewField(4) = NewUserTable.CreateField("Telephone", dbText, 20)
      NewField(4).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(4)
    
  Set NewField(5) = NewUserTable.CreateField("Address", dbText, 50)
      NewField(5).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(5)
                      
  Set NewField(6) = NewUserTable.CreateField("City_State", dbText, 50)
      NewField(6).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(6)
    
  Set NewField(7) = NewUserTable.CreateField("ZipCode", dbText, 11)
      NewField(7).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(7)
        
  Set NewField(8) = NewUserTable.CreateField("EmailAddress", dbText, 50)
      NewField(8).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(8)
  
  Set NewField(9) = NewUserTable.CreateField("Relation", dbText, 40)
      NewField(9).AllowZeroLength = True
      NewUserTable.Fields.Append NewField(9)
  
 'Add The New Table to the database
  NewUserDb.TableDefs.Append NewUserTable

 'Closing
  NewUser_LoginName = ""
  NewUser_Password = ""
  NewUserDb.Close
  For i = 1 To 9
   Set NewField(i) = Nothing
  Next i
  Set NewUserTable = Nothing
  Create_New_User_Dbase = True
   
CreateUserErr:
  If Err.Number <> 0 Then
     MsgBox "Error " & Str$(Err.Number) & " Creating Database." & vbCrLf & _
     Err.Description, vbCritical + vbOKOnly
     Create_New_User_Dbase = False
     Err.Clear
     Exit Function
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Creates New User
'============================================================================================================
Public Function Create_User(User_Name As String, Pwd As String, Access_Lvl As String) As Boolean
  If (Table_Exist(User_Name) = False) And (User_Exist(User_Name) = False) Then
      If Add_To_User_Dbase(User_Name, Pwd, Access_Lvl) = True Then
         If Create_New_User_Dbase(User_Name) = True Then
            Create_User = True
            Exit Function
           Else
            Create_User = False
            Exit Function
         End If
        Else
         Create_User = False
         Exit Function
      End If
     Else
      Create_User = False
      Exit Function
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
'Used To Check If Database Is OK
'============================================================================================================
Public Function Database_Ok(dbName As String) As Boolean
  Dim tstDB As Database
  Dim tstRecSet As Recordset
  On Error GoTo DatabaseErr
  
 'Open The Database
  Set tstDB = OpenDatabase(dbName, False, True, ";pwd=" & Database_Password)
 'Open A Table that should always be in the database
 'This Table Stores Information on All Users
  Set tstRecSet = tstDB.OpenRecordset("Users")
 'Refresh Table
  tstRecSet.Fields.Refresh
  tstRecSet.MoveFirst
  tstRecSet.Fields.Refresh
 'Close
  tstRecSet.Close
  tstDB.Close
 'Set Database_Ok = true
  Database_Ok = True
  Exit Function
  
DatabaseErr:
  If Err.Number <> 0 Then
    Database_Ok = False
    MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical
    Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Checks If The Database Exist
'============================================================================================================
Public Function Database_Found(Dbase_Path As String) As Boolean
  If Dir(Dbase_Path & "\" & Database_Name) <> "" Then
     Database_Found = True
    Else
     Database_Found = False
  End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
'Used To Check If A Database Table Is OK
'============================================================================================================
Public Function Table_Ok(dbName As String, Table As String) As Boolean
  Dim tDB As Database
  Dim tRecSet As Recordset
  On Error GoTo TableErr
  
  If Database_Ok(dbName) = True Then
    'Open Database (Shared-Read Only)
     Set tDB = OpenDatabase(dbName, False, True, ";pwd=" & Database_Password)
    'Open Table
     Set tRecSet = tDB.OpenRecordset(Table)
    'Refresh Table
     tRecSet.Fields.Refresh
    'Closing Database and Recordset
     tRecSet.Close
     tDB.Close
     Table_Ok = True
     Exit Function
    Else
     Table_Ok = False
     Exit Function
  End If

TableErr:
 If Err.Number <> 0 Then
    Table_Ok = False
    MsgBox "Error : " & Str(Err.Number) & " " & Err.Description, vbCritical
    Err.Clear
 End If
End Function
'============================================================================================================
'============================================================================================================



'============================================================================================================
'Counts the Records within a Table
'============================================================================================================
Public Function RecCount(db_Name As String, rTable As String) As Long
  Dim rcDB As Database
  Dim rcRecSet As Recordset
  On Error GoTo RecCountErr
  
  If Table_Ok(db_Name, rTable) = True Then
     'Opening Database and Recordset
      Set rcDB = OpenDatabase(db_Name, False, True, ";pwd=" & Database_Password)
      Set rcRecSet = rcDB.OpenRecordset(rTable)
      rcRecSet.Fields.Refresh
     'Closing Database and Recordset
      rcRecSet.Close
      rcDB.Close
      RecCount = rcRecSet.RecordCount
     Else
      RecCount = -1
  End If
   
RecCountErr:
  If Err.Number <> 0 Then
     RecCount = -1
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Counts The Amount of Administrators in User's Table
'============================================================================================================
Public Function AdminCount() As Long
  Dim aDb As Database
  Dim aRecSet As Recordset
  On Error GoTo AdminCountErr
 
 'Open Database Shared-Readonly
  Set aDb = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set aRecSet = aDb.OpenRecordset("Users")
  
  AdminCount = 0
  
  aRecSet.Fields.Refresh
  
  Do While Not aRecSet.EOF
     If aRecSet.Fields("AccessLevel") = EncryptText("Administrator", Database_Password) Then
        AdminCount = AdminCount + 1
     End If
     aRecSet.MoveNext
  Loop
  
  aRecSet.Close
  aDb.Close
  
AdminCountErr:
  If Err.Number <> 0 Then
     AdminCount = -1
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Used to check if a record exist
'============================================================================================================
Public Function Record_Exist(Table As String, F_Name As String, L_Name As String, Relation_ As String) As Boolean
  Dim rDB As Database
  Dim rRecSet As Recordset
  
  
 'Check if the table is OK
  If (Table_Ok(Database_Path & "\" & Database_Name, Table) = True) Then
     
    'Check The Amount of records in the Table
     If RecCount(Database_Path & "\" & Database_Name, Table) = 1 Then
       'Since there is no records
        Record_Exist = False
        Exit Function
     End If
      
     Set rDB = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
     Set rRecSet = rDB.OpenRecordset(Table)
    'Refresh
     rRecSet.Fields.Refresh
     
     Do While Not rRecSet.EOF
        If (rRecSet.Fields("FirstName") = F_Name) And _
           (rRecSet.Fields("LastName") = L_Name) And _
           (rRecSet.Fields("Relation") = Relation_) Then
           Record_Exist = True
           Exit Function
        End If
        rRecSet.MoveNext
     Loop
     
    'Since it reaches here
     Record_Exist = False
     Exit Function
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Checks if a recordset is empty
'============================================================================================================
Public Function EmptyRS(RS As Recordset) As Boolean
  EmptyRS = ((RS.BOF = True) And (RS.EOF = True))
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Checks If A Table Exists
'============================================================================================================
Public Function Table_Exist(TableName As String) As Boolean
  Dim i As Integer
  Dim db As Database
  On Error GoTo Table_Err
  Table_Exist = False
  
 'Open the password protected database
  Set db = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  For i = 0 To db.TableDefs.Count - 1
   If UCase(db.TableDefs(i).Name) = UCase(TableName) Then
      Table_Exist = True
      db.Close
      Exit Function
   End If
  Next i
Table_Err:
  If Err.Number <> 0 Then
     Table_Exist = False
     Err.Clear
     Exit Function
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'This Function is used to search the "Users" table to see if
'specific user exist
'============================================================================================================
Public Function User_Exist(U_Name As String) As Boolean
  Dim usrDB As Database
  Dim usrRec As Recordset
  Dim tmpStr As String
  On Error GoTo UserExistErr
  
  User_Exist = False
  
 'Open the password protected database
  Set usrDB = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set usrRec = usrDB.OpenRecordset("Users")
  
  usrRec.Fields.Refresh
  usrRec.MoveFirst
  
  If usrRec.RecordCount < 1 Then
     User_Exist = False
     usrRec.Close
     usrDB.Close
     Exit Function
    Else
     Do While Not usrRec.EOF
        tmpStr = DecryptText(usrRec.Fields("LoginName"), Database_Password)
        If (LCase(tmpStr)) = (LCase(U_Name)) Then
           User_Exist = True
           usrRec.Close
           usrDB.Close
           Exit Function
        End If
        usrRec.MoveNext
     Loop
   End If 'usrRec.RecordCount > 0
   
UserExistErr:
  If Err.Number <> 0 Then
     MsgBox "Error : " & Err.Description & " " & Err.Number
     User_Exist = False
     Set usrRec = Nothing
     Set usrDB = Nothing
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'This Function Is Used To Remove A Table From The Database
'============================================================================================================
Public Function Remove_Table(TableName As String) As Boolean
  Dim dropDB As Database
  Dim dropTableDef
  Dim dropDB_Open As Boolean
  On Error GoTo DropError
  
  dropDB_Open = False
  
  Set dropDB = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
  dropDB_Open = True
  dropDB.Execute "DROP TABLE " & TableName
  dropDB.Close
  Remove_Table = True
  Exit Function

DropError:
  If Err.Number <> 0 Then
     Remove_Table = False
     MsgBox "Error " & Format$(Err.Number) & " dropping table." & _
            Err.Description, vbCritical + vbOKOnly
     Set dropDB = Nothing
     Err.Clear
   End If
End Function
'============================================================================================================
'============================================================================================================

'============================================================================================================
'Used To Remove A Specific user's Record from the "Users"
'Table
'============================================================================================================
Function Delete_User_Record(UName As String) As Boolean
  Dim RS As Recordset
  Dim db As Database
  On Error GoTo Delete_Err
   
 'Open Database The Password Protected Database
  Set db = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
 'Open the User table
  Set RS = db.OpenRecordset("SELECT * FROM Users WHERE LoginName = '" & EncryptText(UName, Database_Password) & "'")
  RS.Fields.Refresh
 'Check If The User Is Logged In
  If RS.Fields("LoggedIN") = False Then
     With RS
        .Delete  'Delete it
        .Close   'Close it
     End With
     db.Close
     Delete_User_Record = True
    Else
     MsgBox UName & " is currently logged in."
  End If
  
Delete_Err:
  If Err.Number <> 0 Then
     Delete_User_Record = False
     Err.Clear
     Exit Function
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'Used To Completely Remove A User
'Remove the user's from record from the "User's" Table
'Remove The Table That matches the UserName
'============================================================================================================
Public Function Remove_User(User_Name As String) As Boolean
 'First Check if the The User's Record and
 'check if The User Table is ok
  Remove_User = False
 
 If (User_Exist(User_Name) = True) And (Table_Ok(Database_Path & "\" & Database_Name, User_Name) = True) Then
    If (Delete_User_Record(User_Name) = True) And (Remove_Table(User_Name) = True) Then
       Remove_User = True
       Exit Function
      Else
       Remove_User = False
       Exit Function
    End If
   Else
    Remove_User = False
    Exit Function
 End If

End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'This Function is used to Rename a Database Table 'Used
'============================================================================================================
Public Function Rename_Database_Table(Old_Table As String, New_Table As String) As Boolean
  Dim DBase As Database
  Dim TDef As TableDef
  Dim Table_Found As Boolean
  On Error GoTo Rename_Table_Error
   
  Rename_Database_Table = False
  Table_Found = False
  
 'Open The Database
  Set DBase = OpenDatabase(Database_Path & "\" & Database_Name, False, False, ";pwd=" & Database_Password)
   
 'Search For The Matching Table
  For Each TDef In DBase.TableDefs
      If TDef.Name = Old_Table Then
         Table_Found = True
         Exit For
      End If
  Next

  If Table_Found = True Then
    'the varable is still holding the
    'object reference here!
     TDef.Name = New_Table
     DBase.TableDefs.Refresh
  End If

  Set TDef = Nothing
  DBase.Close
  Set DBase = Nothing
  Rename_Database_Table = True
  Exit Function

Rename_Table_Error:
  If Err.Number <> 0 Then
     Rename_Database_Table = False
     MsgBox "Error " & Str(Err.Number) & ". Unable to rename the table." & vbCrLf & Err.Description, vbCritical + vbOKOnly
     Set TDef = Nothing
     Set DBase = Nothing
     Err.Clear
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'This Fuction Is Used To check If A USER is Logged in
'============================================================================================================
Public Function UserLoggedIn(theUserName As String) As Boolean
  Dim tmpUserDb As Database
  Dim tmpUserRec As Recordset
 'On Error Resume Next
  
 'Note : this method is used because The UserName is Unique
  Set tmpUserDb = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set tmpUserRec = tmpUserDb.OpenRecordset("SELECT * FROM Users WHERE LoginName = '" & EncryptText(theUserName, Database_Password) & "'")
  tmpUserRec.Fields.Refresh
      
 'Check If Found
  If tmpUserRec.RecordCount > 0 Then
     UserLoggedIn = tmpUserRec.Fields("LoggedIn")
    'Closing
     tmpUserRec.Close
     tmpUserDb.Close
  End If
End Function
'============================================================================================================
'============================================================================================================


'============================================================================================================
'This Function is used to check if someone is logged in
'============================================================================================================
Public Function Users_Logged_In() As Long
  Dim sDbase As Database
  Dim sRecordset As Recordset
  On Error Resume Next
  
  Set sDbase = OpenDatabase(Database_Path & "\" & Database_Name, False, True, ";pwd=" & Database_Password)
  Set sRecordset = sDbase.OpenRecordset("SELECT * FROM Users")
  sRecordset.Fields.Refresh
  sRecordset.MoveFirst
  Users_Logged_In = 0
  
  Do While Not sRecordset.EOF
     If sRecordset.Fields("LoggedIn") = True Then
         Users_Logged_In = Users_Logged_In + 1
    End If
    sRecordset.MoveNext
  Loop
  sRecordset.Close
  sDbase.Close

End Function
'============================================================================================================
'============================================================================================================


