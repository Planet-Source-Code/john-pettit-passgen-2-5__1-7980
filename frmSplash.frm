VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    DeskHdc = GetDC(0)
    ret = BitBlt(picDC.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
    ret = ReleaseDC(0&, DeskHdc)
    Blend Me, picDC, 50, 0, 0, Me.ScaleWidth, Me.ScaleHeight

    Me.Show
    Me.Refresh

End Sub
