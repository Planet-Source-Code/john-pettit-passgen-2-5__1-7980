VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2115
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDetail 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMsg.frx":0E42
      Top             =   600
      Width           =   3440
   End
   Begin VB.Label lblSubject 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3435
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   0
      X2              =   3675
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   3675
      X2              =   3675
      Y1              =   2100
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   3675
      X2              =   0
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2100
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1260
      TabIndex        =   1
      Top             =   1785
      Width           =   1170
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MessageTitle"
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
      Width           =   3705
   End
   Begin VB.Menu mnuTest 
      Caption         =   "test"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Move form without a border declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub cmdClose_Click()

    Unload frmMsg
    
End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdClose.BackColor = &HFFC0C0

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdClose.BackColor = &HFF8080

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 112:
End Select


End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdClose.BackColor = &H800000

End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'This code makes the form move when the mouse
    'is down on the label
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, &HF012, 0
    End If


End Sub
