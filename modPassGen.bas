Attribute VB_Name = "PassGen"
Option Explicit
'**********************************************************************
'   PassGen by Master Yoda
'
'  Allows you to generate passwords based on random characters to
' prevent a dictionary attack, etc... from succeeding.
'
' v1.0 created and completed by Master Yoda: 29 Dec 1999
' v2.5 completed by Master Yoda: 5 May 2000
'
'  You may edit and modify this program only if you keep my name on
' it as the original creator.  I would also appreciate any additional
' functions and design improvements.
'
'   VERSION 2.5 CREDITS
'
'   Master Yoda         - Programming
'                       - Design
'                       - Lead Testing
'                       - Stupid Idiot who didn't do this project in BCB!
'
'   Igguk               - Registry Entry
'                       - Code Optimization
'                       - Testing
'
'   Jonathan Walkyier   - Phrase Dependant Generation Concept
'                       - Testing
'
'   Renegade            - Additional Testing
'
'   ?ue                 - Still more testing
'
'
'**********************************************************************
' API Code for registry editing
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedBuffer As Any, ByVal nSize As Long, ByVal lpFileName As String) As Long

' GLOBAL VARIABLES AND CONTANTS
Public sININame As String
Public Key As String
Public Number As Long
Public OldOnline As Boolean
Public file_name As String
Attribute file_name.VB_VarProcData = "StandardPicture;Data"
Public msgTitle As String
Public msgSubject As String
Public msgDetail As String
Public msgStandard As Boolean
Public Section As String
Public Password As String


Public Function CheckCharacter(newChar As Integer) As Boolean
Const StartSymbol1 As Integer = 33
Const EndSymbol1 As Integer = 47

Const StartNumeric As Integer = 48
Const EndNumeric As Integer = 57

Const StartSymbol2 As Integer = 58
Const EndSymbol2 As Integer = 64

Const StartUpper As Integer = 65
Const EndUpper As Integer = 90

Const StartLower As Integer = 97
Const EndLower As Integer = 122

Const StartExtended As Integer = 128
Const EndExtended As Integer = 255

' NOTE: When all characters can be used for generating passwords the speed
' of creating passwords is significantly faster because of the way this SUB
' is designed.  I'm sure this can be tweaked to gain better performance.
' Athough, I'd rather use this way than using InStr...this is a bit more elegant.
' Please let me know if you come up with a better solution.
    
    CheckCharacter = False
    
    ' This Select Case that will branch of into different groups
    ' depending on what type of character the newly created random
    ' number represents. Most common groups of characters put first
    '
    Select Case newChar
            ' Case Statement for UpperCase Characters
            Case StartUpper To EndUpper
                CheckCharacter = frmPassGen.chkUpper.Value
                   
            ' Case Statement for LowerCase Characters
            Case StartLower To EndLower
                CheckCharacter = frmPassGen.chkLower.Value
                   
            ' Case Statement for Numerics
            Case StartNumeric To EndNumeric
                CheckCharacter = frmPassGen.chkNumeric.Value
                    
            ' Case Statement for symbols
            Case StartSymbol1 To EndSymbol1, StartSymbol2 To EndSymbol2
                CheckCharacter = frmPassGen.chkSymbols.Value
            
            Case StartExtended To EndExtended
                CheckCharacter = frmPassGen.chkExtended.Value
    End Select

End Function

Public Sub CheckSecurity()

    'Set Security Settings for 'Lowest'
    If frmPassGen.cmbSecurity.ListIndex = 0 Then
        frmPassGen.txtLength.Text = 4
        frmPassGen.chkUpper.Value = 0
        frmPassGen.chkLower.Value = 1
        frmPassGen.chkNumeric.Value = 0
        frmPassGen.chkSymbols.Value = 0
        frmPassGen.chkExtended.Value = 0
        frmPassGen.chkAllChars.Value = 0
        frmPassGen.Status.Caption = "Security Level Set To LOWEST"
    End If

    'Set Security Settings for 'Low'
    If frmPassGen.cmbSecurity.ListIndex = 1 Then
        frmPassGen.txtLength.Text = 6
        frmPassGen.chkUpper.Value = 0
        frmPassGen.chkLower.Value = 1
        frmPassGen.chkNumeric.Value = 0
        frmPassGen.chkSymbols.Value = 0
        frmPassGen.chkExtended.Value = 0
        frmPassGen.chkAllChars.Value = 0
        frmPassGen.Status.Caption = "Security Level Set To LOW"
    End If

    'Set Security Settings for 'Medium'
    If frmPassGen.cmbSecurity.ListIndex = 2 Then
        frmPassGen.txtLength.Text = 8
        frmPassGen.chkUpper.Value = 1
        frmPassGen.chkLower.Value = 1
        frmPassGen.chkNumeric.Value = 0
        frmPassGen.chkSymbols.Value = 0
        frmPassGen.chkExtended.Value = 0
        frmPassGen.chkAllChars.Value = 0
        frmPassGen.Status.Caption = "Security Level Set To MEDIUM"
    End If

    'Set Security Settings for 'High'
    If frmPassGen.cmbSecurity.ListIndex = 3 Then
        frmPassGen.txtLength.Text = 10
        frmPassGen.chkUpper.Value = 1
        frmPassGen.chkLower.Value = 1
        frmPassGen.chkNumeric.Value = 0
        frmPassGen.chkSymbols.Value = 0
        frmPassGen.chkExtended.Value = 0
        frmPassGen.chkAllChars.Value = 0
        frmPassGen.Status.Caption = "Security Level Set To HIGH"
    End If

    'Set Security Settings for 'Very High'
    If frmPassGen.cmbSecurity.ListIndex = 4 Then
        frmPassGen.txtLength.Text = 16
        frmPassGen.chkUpper.Value = 1
        frmPassGen.chkLower.Value = 1
        frmPassGen.chkNumeric.Value = 0
        frmPassGen.chkSymbols.Value = 0
        frmPassGen.chkExtended.Value = 0
        frmPassGen.chkAllChars.Value = 0
        frmPassGen.Status.Caption = "Security Level Set To VERY HIGH"
    End If

    'Set Security Settings for 'Extremely High'
    If frmPassGen.cmbSecurity.ListIndex = 5 Then
        frmPassGen.txtLength.Text = 32
        frmPassGen.chkUpper.Value = 1
        frmPassGen.chkLower.Value = 1
        frmPassGen.chkNumeric.Value = 1
        frmPassGen.chkSymbols.Value = 0
        frmPassGen.chkExtended.Value = 0
        frmPassGen.chkAllChars.Value = 0
        frmPassGen.Status.Caption = "Security Level Set To EXTREME"
    End If


End Sub


Public Sub StandardKey(Key As String, Number As Long)
Dim byte1 As Integer
Dim byte2 As Integer
Dim byte3 As Integer
Dim Looped_Key As String
Dim NumLoops As Integer

Looped_Key = vbNullString
Password = vbNullString

For NumLoops = 1 To frmPassGen.lstPasswords.ListCount
    frmPassGen.lstPasswords.RemoveItem 0
Next NumLoops

Do
    Looped_Key = Looped_Key & frmPassGen.txtKey.Text
Loop Until Len(Looped_Key) > Val(frmPassGen.txtLength.Text)

Rnd -1
Randomize (Val(frmPassGen.txtKeyNum.Text) Xor Val(frmPassGen.txtLength.Text))
For NumLoops = 1 To Val(frmPassGen.txtLength.Text)
redo:
    byte1 = Int(Rnd * 512)
    byte2 = Asc(Mid$(Looped_Key, NumLoops, 1))
    byte3 = (byte1 + byte2) Mod 256
    If CheckCharacter(byte3) = True Then
        Password = Password & Chr$(byte3)
    Else
        GoTo redo
    End If
Next NumLoops
frmPassGen.lstPasswords.AddItem Password, 0
End Sub

Public Sub msg(msgDetail As String, msgSubject As String, msgTitle As String, msgStandard As Boolean)
    
    With frmMsg
        .title = msgTitle
        .lblSubject.Caption = msgSubject
        .txtDetail.Text = msgDetail
      
        If msgStandard = True Then
            .txtDetail.Top = 600
            .txtDetail.Height = 1095
        Else
            .txtDetail.Top = 360
            .txtDetail.Height = 1335
        End If
        .Show 1
    End With
End Sub


Public Sub ShowHelp(Filename As String)
file_name = Filename
frmHelp.Show

End Sub

Public Sub SaveSettings(Section As String)
SaveSetting App.ProductName, Section, "WTop", frmPassGen.Top
SaveSetting App.ProductName, Section, "WLeft", frmPassGen.Left
SaveSetting App.ProductName, Section, "Security", frmPassGen.cmbSecurity.ListIndex
SaveSetting App.ProductName, Section, "AutoGen", frmPassGen.chkAutoGen.Value
SaveSetting App.ProductName, Section, "Uppercase", frmPassGen.chkUpper.Value
SaveSetting App.ProductName, Section, "Lowercase", frmPassGen.chkLower.Value
SaveSetting App.ProductName, Section, "Numeric", frmPassGen.chkNumeric.Value
SaveSetting App.ProductName, Section, "Symbols", frmPassGen.chkSymbols.Value
SaveSetting App.ProductName, Section, "Extended", frmPassGen.chkExtended.Value
SaveSetting App.ProductName, Section, "All", frmPassGen.chkAllChars.Value
SaveSetting App.ProductName, Section, "PasswordLength", Val(frmPassGen.txtLength.Text)
SaveSetting App.ProductName, Section, "TotalPasswords", Val(frmPassGen.txtTotalPasswords.Text)
If frmPassGen.optTimer.Value Then
    SaveSetting App.ProductName, Section, "optTimer", "-1"
Else
    SaveSetting App.ProductName, Section, "optTimer", "0"
End If
If frmPassGen.optDepend.Value Then
    SaveSetting App.ProductName, Section, "optDepend", "-1"
Else
    SaveSetting App.ProductName, Section, "optDepend", "0"
End If
If frmPassGen.optSeed.Value Then
    SaveSetting App.ProductName, Section, "optSeed", "-1"
Else
    SaveSetting App.ProductName, Section, "optSeed", "0"
End If
SaveSetting App.ProductName, Section, "SeedNum", Val(frmPassGen.txtSeed.Text)
SaveSetting App.ProductName, Section, "UseNumber", Val(frmPassGen.txtKeyNum.Text)
SaveSetting App.ProductName, Section, "PassPhrase", frmPassGen.txtKey.Text
End Sub

Public Sub GetSettings(Section As String)
On Error Resume Next
frmPassGen.optTimer.Value = GetSetting(App.ProductName, Section, "optTimer")
frmPassGen.optDepend.Value = GetSetting(App.ProductName, Section, "optDepend")
frmPassGen.optSeed.Value = GetSetting(App.ProductName, Section, "optSeed")
If Err <> 0 Then
    frmPassGen.Status.Caption = "Registry Settings Will Complete When PassGen is Terminated"
    frmPassGen.Top = (Screen.Height - frmPassGen.Height) / 2
    frmPassGen.Left = (Screen.Width - frmPassGen.Width) / 2
    frmPassGen.chkAutoGen.Value = 1
    frmPassGen.cmbSecurity.ListIndex = 2
    frmPassGen.txtTotalPasswords.Text = 5
Else
    frmPassGen.Top = Val(GetSetting(App.ProductName, Section, "WTop"))
    frmPassGen.Left = Val(GetSetting(App.ProductName, Section, "WLeft"))
    frmPassGen.cmbSecurity.ListIndex = Val(GetSetting(App.ProductName, Section, "Security"))
    frmPassGen.chkAutoGen.Value = Val(GetSetting(App.ProductName, Section, "AutoGen"))
    frmPassGen.chkUpper.Value = Val(GetSetting(App.ProductName, Section, "Uppercase"))
    frmPassGen.chkLower.Value = Val(GetSetting(App.ProductName, Section, "Lowercase"))
    frmPassGen.chkNumeric = Val(GetSetting(App.ProductName, Section, "Numeric"))
    frmPassGen.chkSymbols.Value = Val(GetSetting(App.ProductName, Section, "Symbols"))
    frmPassGen.chkExtended.Value = Val(GetSetting(App.ProductName, Section, "Extended"))
    frmPassGen.chkAllChars.Value = Val(GetSetting(App.ProductName, Section, "All"))
    frmPassGen.txtLength.Text = Val(GetSetting(App.ProductName, Section, "PasswordLength"))
    frmPassGen.txtTotalPasswords.Text = Val(GetSetting(App.ProductName, Section, "TotalPasswords"))
    frmPassGen.txtSeed.Text = Val(GetSetting(App.ProductName, Section, "SeedNum"))
    frmPassGen.txtKeyNum.Text = Val(GetSetting(App.ProductName, Section, "UseNumber"))
    frmPassGen.txtKey.Text = GetSetting(App.ProductName, Section, "PassPhrase")
End If
End Sub

Public Sub EndProgram(StoreSettings As Boolean)
If StoreSettings Then
    ' Save Current Settings
    Call SaveSettings("Settings")
End If
End
End Sub
