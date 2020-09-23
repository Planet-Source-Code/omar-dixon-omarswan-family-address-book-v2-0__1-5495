Attribute VB_Name = "Encrypt"
Option Explicit

'Set to True to make the password case-sensitive
#Const CASE_SENSITIVE_PASSWORD = False
 
'Encrypt text
Public Function EncryptText(strText As String, ByVal strPwd As String) As String
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function



