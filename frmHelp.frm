VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5310
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   Icon            =   "frmHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   4575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   7335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   5300
      Y2              =   5300
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   7560
      X2              =   7560
      Y1              =   5280
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   7560
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Label cmdX 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   190
      Left            =   7320
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   225
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PassGen Help"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7590
   End
   Begin VB.Menu mnutest 
      Caption         =   "test"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Move form without a border declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112




 


Private Sub Form_Load()

On Error Resume Next

Dim file_byte As String
Dim file_pos As Integer
Dim file_content As String

file_pos = 0
Open file_name For Binary As 1

    If LOF(1) = 0 Then
        txtHelp.Text = "File could not be loaded!" & vbNewLine & "Error " & Err.Number & ": " & Err.Description
        
        Exit Sub
    End If
    Do While EOF(1) <> True
        file_byte = " "
        file_pos = file_pos + 1
        Get 1, file_pos, file_byte
        Select Case Asc(file_byte)
        Case Else:
            file_content = file_content & file_byte
        End Select
    Loop
    txtHelp.Text = file_content

Close 1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdX.BackColor = &HFF8080

End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'This code makes the form move when the mouse
    'is down on the label
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, &HF012, 0
    End If

End Sub
Private Sub cmdX_Click()

    frmHelp.Hide
    
End Sub

Private Sub cmdX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdX.BackColor = &HC00000

End Sub

Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdX.BackColor = &HFF8080

End Sub

