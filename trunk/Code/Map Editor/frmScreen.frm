VERSION 5.00
Begin VB.Form frmScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Game Screen"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Call the MouseMove event
    Form_MouseMove Button, Shift, x, y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tX As Integer
Dim tY As Integer

    'Convert the click position to tile position
    Engine_ConvertCPtoTP 0, 0, x, y, tX, tY
    HovertX = tX
    HovertY = tY

    'Update caption
    HoverX = x + ParticleOffsetX - 288
    HoverY = y + ParticleOffsetY - 288
    frmMain.MouseLbl.Caption = "(" & HoverX & "," & HoverY & ")"
    frmMain.TileLbl.Caption = "(" & HovertX & "," & HovertY & ")"
    If tX < 0 Then Exit Sub
    If tY < 0 Then Exit Sub
    If tX > MapInfo.Width Then Exit Sub
    If tY > MapInfo.Height Then Exit Sub
    
    'Click the tile
    SetTile tX, tY, Button, Shift
             
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    UpdateEffectList

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1

End Sub
