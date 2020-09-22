VERSION 5.00
Begin VB.Form Hello 
   BorderStyle     =   0  'None
   Caption         =   "Hello"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1590
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Hello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hrgn

Dim LastX As Single, LastY As Single

Private Sub Form_DblClick()
    End
End Sub

Private Sub Form_Load()
    If hrgn Then DeleteObject hrgn
    hrgn = GetBitmapRegion(Picture, vbWhite)
    SetWindowRgn Me.hwnd, hrgn, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        LastX = X
        LastY = Y
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim newleft&, newtop&

    If Button = 1 Then
        newleft = Left + (X - LastX)
        newtop = Top + (Y - LastY)
        Move newleft, newtop
    End If
    
End Sub
