VERSION 5.00
Begin VB.Form frmPassGen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   75
   ClientTop       =   -210
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
   Icon            =   "frmPassGen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Timer INETUpdate 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   7080
      Top             =   4800
   End
   Begin VB.Frame fraDependance 
      BackColor       =   &H00000000&
      Caption         =   "Dependance"
      Enabled         =   0   'False
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
      Height          =   1800
      Left            =   80
      TabIndex        =   29
      Top             =   3120
      Width           =   4560
      Begin VB.TextBox txtKeyNum 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   3120
         MaxLength       =   9
         TabIndex        =   32
         ToolTipText     =   "Input a number to be used to alter the resulting output"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtKey 
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
         Height          =   1200
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Enter your phrase here. This will be used to generate passwords"
         Top             =   525
         Width           =   4345
      End
      Begin VB.Label lblPhrase 
         BackColor       =   &H00000000&
         Caption         =   "Key Phrase:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblKeyNum 
         BackColor       =   &H00000000&
         Caption         =   "Use Number:"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraSecurity 
      BackColor       =   &H00000000&
      Caption         =   "Security"
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
      Height          =   915
      Left            =   4725
      TabIndex        =   17
      Top             =   240
      Width           =   2775
      Begin VB.CheckBox chkAutoGen 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Auto Generate"
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
         Height          =   225
         Left            =   105
         TabIndex        =   20
         ToolTipText     =   "Generate passwords automatically when security level is changed"
         Top             =   600
         Width           =   1620
      End
      Begin VB.ComboBox cmbSecurity 
         BackColor       =   &H00800000&
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
         Height          =   315
         ItemData        =   "frmPassGen.frx":0E42
         Left            =   105
         List            =   "frmPassGen.frx":0E58
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Select Security Level"
         Top             =   210
         Width           =   2535
      End
   End
   Begin VB.Frame fraUse 
      BackColor       =   &H00000000&
      Caption         =   "Characters to Use:"
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
      Height          =   1095
      Left            =   4725
      TabIndex        =   10
      Top             =   1200
      Width           =   2775
      Begin VB.CheckBox chkAllChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Use All"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         ToolTipText     =   "Use All Characters"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkNumeric 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Numeric"
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Use Numbers"
         Top             =   735
         Width           =   1215
      End
      Begin VB.CheckBox chkExtended 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Extended"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         ToolTipText     =   "Use Extended Characters"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkSymbols 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Symbols"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         ToolTipText     =   "Use Symbols"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkLower 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Lowercase"
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
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Use Lowercase Letters"
         Top             =   480
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkUpper 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Uppercase"
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Use Uppercase Letters"
         Top             =   210
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00000000&
      Caption         =   "Options"
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
      Height          =   1425
      Left            =   4725
      TabIndex        =   4
      Top             =   2330
      Width           =   2775
      Begin VB.TextBox txtTotalPasswords 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
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
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   36
         Text            =   "10"
         ToolTipText     =   "Length (in bytes) of passwords"
         Top             =   450
         Width           =   615
      End
      Begin VB.OptionButton optDepend 
         BackColor       =   &H00000000&
         Caption         =   "Dependance"
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
         Height          =   255
         Left            =   960
         TabIndex        =   31
         ToolTipText     =   "Input a phrase and it will provide a fixed password."
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtSeed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   300
         TabIndex        =   9
         Top             =   1065
         Width           =   1635
      End
      Begin VB.OptionButton optSeed 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Seed #"
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Generate random numbers based on a number you select"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optTimer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Timer"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Use timer to generate randomized numbers"
         Top             =   780
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.TextBox txtLength 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
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
         Height          =   285
         Left            =   720
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "6"
         ToolTipText     =   "Length (in bytes) of passwords"
         Top             =   210
         Width           =   615
      End
      Begin VB.Label lblNumPW1 
         BackColor       =   &H00000000&
         Caption         =   "Make:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblNumPW2 
         BackColor       =   &H00000000&
         Caption         =   "passwords"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   2310
         X2              =   2310
         Y1              =   210
         Y2              =   735
      End
      Begin VB.Label lblBits 
         BackColor       =   &H00000000&
         Caption         =   "chars [48]"
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
         Height          =   225
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblHowManyChars 
         BackColor       =   &H00000000&
         Caption         =   "Length:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraButtons 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1160
      Left            =   4725
      TabIndex        =   3
      Top             =   3770
      Width           =   2775
      Begin VB.Label cmdAbout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "About"
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
         Left            =   105
         TabIndex        =   28
         Top             =   840
         Width           =   1170
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
         Left            =   1470
         TabIndex        =   24
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label cmdSite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Site"
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
         Left            =   105
         TabIndex        =   23
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label cmdHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Help"
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
         Left            =   1470
         TabIndex        =   22
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label cmdGenerate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Generate"
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
         Left            =   105
         TabIndex        =   21
         Top             =   210
         Width           =   2535
      End
   End
   Begin VB.Frame fraPasswords 
      BackColor       =   &H00000000&
      Caption         =   "Passwords"
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
      Height          =   2820
      Left            =   80
      TabIndex        =   0
      Top             =   240
      Width           =   4560
      Begin VB.TextBox txtSelected 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   4315
      End
      Begin VB.ListBox lstPasswords 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
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
         Height          =   2550
         Left            =   105
         TabIndex        =   1
         ToolTipText     =   "Generated  Passwords"
         Top             =   240
         Width           =   4355
      End
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   120
      TabIndex        =   38
      Top             =   4960
      Width           =   7335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   5230
      Y2              =   5230
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
   Begin VB.Label cmdMin 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "?"
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
      Left            =   25
      TabIndex        =   27
      ToolTipText     =   "Help"
      Top             =   25
      Width           =   225
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
      TabIndex        =   26
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   225
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PassGen v2.5"
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
      TabIndex        =   25
      Top             =   0
      Width           =   7590
   End
   Begin VB.Menu mnuPopups 
      Caption         =   "popups"
      Enabled         =   0   'False
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuClearPU 
         Caption         =   "Clear"
         Begin VB.Menu mnuClear 
            Caption         =   "Clear"
         End
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
         Begin VB.Menu mnuStatus_DelReg 
            Caption         =   "Delete Registry Settings"
         End
      End
   End
End
Attribute VB_Name = "frmPassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Move form without a border declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Sub chkAllChars_Click()
    
    ' If no character groups are selected disable generate button
    If chkAllChars.Value = 0 And chkUpper.Value = 0 And chkLower.Value = 0 And chkNumeric.Value = 0 And chkSymbols.Value = 0 And chkExtended.Value = 0 Then
        cmdGenerate.Enabled = False
    Else
        cmdGenerate.Enabled = True
    End If
    
    ' Toggle all other character groups if chkAllChars is selected/deselected
    ' This keeps pre-existing values
    If chkAllChars.Value = 1 Then
        chkUpper.Enabled = False
        chkLower.Enabled = False
        chkNumeric.Enabled = False
        chkSymbols.Enabled = False
        chkExtended.Enabled = False
    Else
        chkUpper.Enabled = True
        chkLower.Enabled = True
        chkNumeric.Enabled = True
        chkSymbols.Enabled = True
        chkExtended.Enabled = True
    End If

End Sub

Private Sub chkAllChars_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Use ALL Characters"

End Sub

Private Sub chkAutoGen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Generate Passwords on Security Level Change"
End Sub

Private Sub chkExtended_Click()

    ' If no character groups are selected disable generate button
    If chkAllChars.Value = 0 And chkUpper.Value = 0 And chkLower.Value = 0 And chkNumeric.Value = 0 And chkSymbols.Value = 0 And chkExtended.Value = 0 Then
        cmdGenerate.Enabled = False
    Else
        cmdGenerate.Enabled = True
    End If

End Sub

Private Sub chkExtended_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Use Extended Characters"

End Sub

Private Sub chkLower_Click()

    ' If no character groups are selected disable generate button
    If chkAllChars.Value = 0 And chkUpper.Value = 0 And chkLower.Value = 0 And chkNumeric.Value = 0 And chkSymbols.Value = 0 And chkExtended.Value = 0 Then
        cmdGenerate.Enabled = False
    Else
        cmdGenerate.Enabled = True
    End If

End Sub

Private Sub chkLower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Use Lowercase Characters"

End Sub

Private Sub chkNumeric_Click()

    ' If no character groups are selected disable generate button
    If chkAllChars.Value = 0 And chkUpper.Value = 0 And chkLower.Value = 0 And chkNumeric.Value = 0 And chkSymbols.Value = 0 And chkExtended.Value = 0 Then
        cmdGenerate.Enabled = False
    Else
        cmdGenerate.Enabled = True
    End If

End Sub

Private Sub chkNumeric_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Use Numeric Characters"

End Sub

Private Sub chkSymbols_Click()

    ' If no character groups are selected disable generate button
    If chkAllChars.Value = 0 And chkUpper.Value = 0 And chkLower.Value = 0 And chkNumeric.Value = 0 And chkSymbols.Value = 0 And chkExtended.Value = 0 Then
        cmdGenerate.Enabled = False
    Else
        cmdGenerate.Enabled = True
    End If

End Sub

Private Sub chkSymbols_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Use Symbolic Characters"

End Sub

Private Sub chkUpper_Click()

    ' If no character groups are selected disable generate button
    If chkAllChars.Value = 0 And chkUpper.Value = 0 And chkLower.Value = 0 And chkNumeric.Value = 0 And chkSymbols.Value = 0 And chkExtended.Value = 0 Then
        cmdGenerate.Enabled = False
    Else
        cmdGenerate.Enabled = True
    End If

End Sub

Private Sub chkUpper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Use Uppercase Characters"
End Sub

Private Sub cmbSecurity_Click()

    Call PassGen.CheckSecurity
    If chkAutoGen.Value = 1 Then
        Call cmdGenerate_Click
    End If

End Sub

Private Sub cmdAbout_Click()
    
    frmAbout.Show 0
    
End Sub

Private Sub cmdAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdAbout.BackColor = &HFFC0C0

End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdAbout.BackColor = &HFF8080
    Status.Caption = "About PassGen"

End Sub

Private Sub cmdClose_Click()

    ' Close button pressed...terminate PassGen
    Call EndProgram(True)

End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdClose.BackColor = &HFFC0C0

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &HFF8080
    cmdAbout.BackColor = &H800000
    Status.Caption = "Close PassGen"


End Sub

Private Sub cmdGenerate_Click()

If Val(txtLength.Text) = 0 Then
    Status.Caption = "Password Length = 0"
    Beep
    Exit Sub
End If
If Val(txtTotalPasswords.Text) = 0 Then
    Status.Caption = "Generate No Passwords? :)"
    Beep
    Exit Sub
End If


If optDepend.Value = False Then
'  THIS GENERATES THE PASSWORDS WITH THE HELP OF THE CheckCharacter SUB
Dim GeneratePasswords As Integer
Dim PasswordCreate As Integer
Dim IsCharGood As Boolean
Dim StartTime As Single          ' Holds timer value when generation of passwords start
Dim EndTime As Single            ' Holds timer value when generation is complete
Dim newPassword As String        ' Holds the newPassword currently being created
Dim newChar As Integer           ' newly generated character for use in the password currently being generated

    txtSeed.Text = Val(txtSeed.Text)
    
    StartTime = Timer   ' Get start time reference
    
    ' Remove all passwords in the list box
    If lstPasswords.ListCount > 0 Then
        lstPasswords.Clear
    End If
    
    'Randomize Numbers based on the timer
    If optTimer.Value = True Then
        Randomize Timer
    End If
    
    If optSeed.Value = True Then
        'Randomize numbers with manual seed #
        If Len(txtSeed.Text) > 0 Then
            Randomize txtSeed.Text
        Else
            Randomize
        End If
    End If
    
    'Begin Generating passwords...
    For GeneratePasswords = 1 To Val(txtTotalPasswords.Text)
        newPassword = ""    ' Start a new password
        ' This FOR-NEXT loop creates a single password
        For PasswordCreate = 1 To txtLength
            newChar = Int((Rnd * 255))  ' set value of new character
            If chkAllChars.Value = 0 Then
                While Not CheckCharacter(newChar)
                    newChar = Int((Rnd * 255))  ' set value of new character
                Wend
            End If
            newPassword = newPassword & Chr(newChar)
        Next
        ' Places the new password at the bootom of the list
        lstPasswords.AddItem newPassword, lstPasswords.ListCount
    Next
    
    EndTime = Timer     ' get the time that password generation stopped
    
    ' Calculate and display time taken
    If EndTime - StartTime > 0.05 Then
        txtSelected = "Password generation took: " & Format(EndTime - StartTime, "###.##") & " secs... "
    Else
        txtSelected = "Password generation took: < 0.05 secs... "
    End If
Else
    If Len(txtKey.Text) > 0 Then
        Call PassGen.StandardKey(txtKey.Text, Val(txtKeyNum.Text))
    Else
        Call PassGen.msg("No Phrase Entered", "Dependant Generation", "Generate Password", False)
    End If
End If
Status.Caption = txtTotalPasswords.Text & " random passwords generated."

End Sub

Private Sub cmdGenerate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &HFFC0C0

End Sub

Private Sub cmdGenerate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &HFF8080
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdAbout.BackColor = &H800000
    Status.Caption = "Generate Passwords Using Selected Options"

End Sub

Private Sub cmdHelp_Click()
        Call ShowHelp(App.Path & "\main.pgh")

End Sub

Private Sub cmdMin_Click()
Call ShowHelp(App.Path & "\main.pgh")
End Sub

Private Sub cmdMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdMin.BackColor = &HC00000
    Status.Caption = "Get Help"
End Sub

Private Sub cmdSite_Click()
Dim xshell As Integer
 Shell "explorer http://webone.com.au/~jpettit"
End Sub

Private Sub cmdSite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSite.BackColor = &HFFC0C0
End Sub

Private Sub cmdSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &HFF8080
    cmdClose.BackColor = &H800000
    cmdAbout.BackColor = &H800000
    Status.Caption = "Visit Blade++ Software Web Site"

End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdHelp.BackColor = &HFFC0C0
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &HFF8080
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdAbout.BackColor = &H800000
    Status.Caption = "Get Help"

End Sub

Private Sub cmdX_Click()

    Call EndProgram(True)
End Sub

Private Sub cmdX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdX.BackColor = &HC00000
    Status.Caption = "Close PassGen"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 112:
        ' invoke context sensitive help
        Call ShowHelp(App.Path & "\main.pgh")
    Case 113:
End Select

End Sub

Private Sub Form_Load()
Dim Version As String

    'PC 280100 Remove MHINI32.OCX
    sININame = App.Path & "\" & App.EXEName & ".INI"
    Version = App.Major & "." & App.Minor
    title.Caption = "PassGen " & Version
    If App.PrevInstance = True Then
        Call msg("PassGen is already running", "", "PassGen 2.5", False)
        End
    End If
    Call GetSettings("Settings")
   
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdX.BackColor = &HFF8080
    cmdMin.BackColor = &HFF8080
    cmdAbout.BackColor = &H800000
    Status.Caption = ""

End Sub

Private Sub optPsuedoRandom_Click()

    ' Disable the manual seed function
    txtSeed.Enabled = False

End Sub

Private Sub Form_Terminate()
    Call EndProgram(True)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EndProgram(True)

End Sub

Private Sub fraButtons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdX.BackColor = &HFF8080
    cmdMin.BackColor = &HFF8080
    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdAbout.BackColor = &H800000
    Status.Caption = ""

End Sub

Private Sub fraOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdX.BackColor = &HFF8080
    cmdMin.BackColor = &HFF8080
    cmdAbout.BackColor = &H800000
    Status.Caption = ""

End Sub

Private Sub fraSecurity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdX.BackColor = &HFF8080
    cmdMin.BackColor = &HFF8080
    cmdAbout.BackColor = &H800000
    Status.Caption = "Set Security Level"

End Sub

Private Sub fraUse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdX.BackColor = &HFF8080
    cmdMin.BackColor = &HFF8080
    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    cmdAbout.BackColor = &H800000

End Sub


Private Sub lstPasswords_DblClick()

    ' When a password is double-clicked in the list box it gets set to frmSave to get stored
    If lstPasswords.ListIndex <> -1 Then
        Password = lstPasswords.List(lstPasswords.ListIndex)
    End If

End Sub

Private Sub lstPasswords_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuClearPU
End Sub

Private Sub lstPasswords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdGenerate.BackColor = &H800000
    cmdHelp.BackColor = &H800000
    cmdSite.BackColor = &H800000
    cmdClose.BackColor = &H800000
    Status.Caption = "List Of Passwords"

End Sub

Private Sub mnuClear_Click()
Dim NumLoops As Integer
For NumLoops = 1 To lstPasswords.ListCount
    lstPasswords.RemoveItem 0
Next NumLoops
End Sub


Private Sub mnuStatus_DelReg_Click()
On Error Resume Next
DeleteSetting App.ProductName, "settings"
End Sub

Private Sub optDepend_Click()
    ' Enable Key Dependant Generation Mode
    txtSeed.Enabled = False
    fraDependance.Enabled = True
    fraSecurity.Enabled = False
    Status.Caption = "Enter Phrase and Randomize Number"
'    Call Msg("This function is NOT complete!", "Incomplete Function", "PassGen 2.5", False)
End Sub

Private Sub optDepend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Create Phrase Dependant Passwords"

End Sub

Private Sub optSeed_Click()

    ' Enable th manual Seed function
    txtSeed.Enabled = True
    fraDependance.Enabled = False
    fraSecurity.Enabled = True
End Sub

Private Sub optSeed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Create Seeded Passwords"
End Sub

Private Sub optTimer_Click()

    ' Disable the manual seed function
    txtSeed.Enabled = False
    fraDependance.Enabled = False
    fraSecurity.Enabled = True
End Sub

Private Sub tmrUpdate_Timer()

    txtSelected.Text = cmbSecurity.ListIndex

End Sub



Private Sub optTimer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Create Timer Generated Passwords"

End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuStatus
End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'This code makes the form move when the mouse
    'is down on the label
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, &HF012, 0
    End If

End Sub

Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdX.BackColor = &HFF8080
    cmdMin.BackColor = &HFF8080

End Sub


Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
    Case 121
        Call cmdGenerate_Click
        
        
End Select
End Sub

Private Sub txtKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Enter Phrase"

End Sub

Private Sub txtKeyNum_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        Call cmdGenerate_Click

End Select

End Sub

Private Sub txtKeyNum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Number to Randomize Result Password"
End Sub

Private Sub txtLength_Change()

    If Len(txtLength.Text) > 0 Then
        txtLength.Text = Val(txtLength.Text)
        lblBits.Caption = "chars [" & txtLength.Text * 8 & "]"
    Else
    txtLength.Text = 0
    txtLength.SelStart = 0
    txtLength.SelLength = 1
    End If
    
End Sub


Private Sub txtLength_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Set Password Length"

End Sub

Private Sub txtTotalPasswords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Caption = "Set Total Number of Passwords to Generate"

End Sub
