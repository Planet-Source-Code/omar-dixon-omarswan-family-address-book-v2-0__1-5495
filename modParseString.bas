Attribute VB_Name = "modParseString"
Option Explicit
Option Compare Text


'Retrieve and array of tokens delimited by a list of tokens
'This function takes the following arguments
'1 - The string to be searched (this argument is passed byval
'    to preserve it in the calling function)
'2 - A REFERENCE to an array fo strings in which the tokens are stored
'3 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns the number of tokens found

Public Function ParseDelimitedString(ByVal SearchString As String, ByRef Tokens As Variant, _
    Optional TokenList As String = " " & vbTab) As Integer

Dim StringLength As Integer
Dim TokensFound As Integer
ReDim TempToken(50) As String
Dim StartPos As Integer

'Initialize local variables
ParseDelimitedString = 0
TokensFound = 0

StartPos = skipDelimitor(SearchString, TokenList)
If StartPos = 0 Then
  'Empty string or just demilitors.
  Tokens = Array()
  Exit Function
ElseIf StartPos > 1 Then
  'String starts with delimitors. Skip initial delimitors
  SearchString = Right$(SearchString, Len(SearchString) - StartPos + 1)
End If
StringLength = Len(SearchString)

If StringLength = 0 Then
  'Empty string
  Tokens = Array()
  Exit Function
End If

'make Tokens somewhat large to avoid redim-ming to often
ReDim Tokens(50)
Do While getNextToken(SearchString, Tokens(TokensFound), TokenList)
  TokensFound = TokensFound + 1
  If TokensFound > UBound(Tokens) Then
    'Running out of space.
    ReDim Preserve Tokens(TokensFound * 2)
  End If
Loop
ReDim Preserve Tokens(TokensFound)
ParseDelimitedString = TokensFound + 1
End Function

'Get the next token from a string and remove token+delimitors from the string
'This function takes the following arguments:
'1 - The string to be searched. THIS STRING WILL BE CHANGED BY THIS FUNCTION
'2 - The string to contain the token
'3 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns True if another token can befound in the string and
'return False if this is the last token in the string

Function getNextToken(ByRef SearchString As String, ByRef Token As Variant, _
    Optional TokenList As String = " " & vbTab) As Boolean

Dim StartPos As Integer
Dim DelimitorPos As Integer

DelimitorPos = FindDelimitor(SearchString, TokenList)
If DelimitorPos = 0 Then
  Token = SearchString
  getNextToken = False
  Exit Function
Else
  'Found a delimitor.
  'Store string in Tokens
  Token = Left$(SearchString, DelimitorPos - 1)
  SearchString = Right$(SearchString, Len(SearchString) - DelimitorPos + 1)
  StartPos = skipDelimitor(SearchString, TokenList)
  If StartPos > 1 Then
    'Skip delimitor characters
    SearchString = Right$(SearchString, Len(SearchString) - StartPos + 1)
    getNextToken = True
    Exit Function
  Else
    If StartPos = 0 Then
      'only delimitors left
    Else
      'No non-delimitor characters left in the string
      Token = SearchString
    End If
    getNextToken = False
    Exit Function
  End If
End If
End Function


'Find the position of the first token delimitor character in a string.
'This function takes the following arguments:
'1 - The string to be searched
'2 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns the position of the found character or
'zero if the character was not found

Public Function FindDelimitor(SearchString As String, _
    Optional TokenList As String = " " & vbTab) As Integer

Dim StringLength As Integer
Dim Counter As Integer

StringLength = Len(SearchString)
For Counter = 1 To StringLength
  If InStr(TokenList, Mid$(SearchString, Counter, 1)) > 0 Then
    FindDelimitor = Counter
    Exit Function
  End If
Next
FindDelimitor = 0
End Function


'Find the position of the first character, which is NOT a token delimitor character, in a string.
'This function takes the following arguments:
'1 - The string to be searched
'2 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns the position of the found character or
'zero if the character was not found

Public Function skipDelimitor(SearchString As String, _
    Optional TokenList As String = " " & vbTab) As Integer

Dim StringLength As Integer
Dim Counter As Integer

StringLength = Len(SearchString)
For Counter = 1 To StringLength
  If InStr(TokenList, Mid$(SearchString, Counter, 1)) = 0 Then
    skipDelimitor = Counter
    Exit Function
  End If
Next
skipDelimitor = 0

End Function




