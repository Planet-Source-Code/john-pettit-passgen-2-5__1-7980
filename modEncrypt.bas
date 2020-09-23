Attribute VB_Name = "modEncrypt"

' THIS WILL NEED TO BE MORE SECURE EVENTUALLY!
'**********************************************************************************
'  Encrypt by Igguk
'
'  Allows you to encrypt/decrypt strings containing any ASCII values from 0 to 255
'  limited to a string length of 39 chars.
'
'  v1.0 created and completed by Igguk : 31 Jan 2000
'**********************************************************************************
Option Explicit

Public Function Encrypt(sText As String) As String
' Encryption of a string
' Parameters :
'           sText : string to encrypt
' Return value :
'           The encrypted string

Dim i As Integer
Dim sChar As String

    Encrypt = ""
    For i = 1 To Len(sText)
        sChar = Mid(sText, i, 1)
        sChar = Format(Asc(sChar) * i, "0000")
        sChar = 9 - Mid(sChar, 4, 1) & 9 - Mid(sChar, 3, 1) & 9 - Mid(sChar, 2, 1) & 9 - Mid(sChar, 1, 1)
        Encrypt = Encrypt & Chr(Mid(sChar, 3, 2)) & Chr(Mid(sChar, 1, 2))
    Next
    
End Function

Public Function Decrypt(sText As String) As String
' Decryption of a string
' Parameters :
'           sText : string to decrypt
' Return value :
'           The decrypted string

Dim i As Integer
Dim sChar As String

    Decrypt = ""
    For i = 1 To Len(sText) Step 2
        sChar = Mid(sText, i, 2)
        sChar = Format(Asc(Mid(sChar, 2, 1)), "00") & Format(Asc(Mid(sChar, 1, 1)), "00")
        sChar = 9 - Mid(sChar, 4, 1) & 9 - Mid(sChar, 3, 1) & 9 - Mid(sChar, 2, 1) & 9 - Mid(sChar, 1, 1)
        If sChar / (1 + Int(i / 2)) < 256 Then
            sChar = Chr(sChar / (1 + Int(i / 2)))
        Else
            sChar = ""
        End If
        Decrypt = Decrypt & sChar
    Next
    
End Function

