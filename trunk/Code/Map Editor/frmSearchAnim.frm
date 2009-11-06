VERSION 5.00
Begin VB.Form frmSearchAnim 
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Animations Using Texture: 0"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSearchAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'If IsUnloading = 0 Then Cancel = 1

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    SetInfo vbNullString
    For i = 1 To ShownTextureAnims.NumGrhs
        With ShownTextureAnims.Grh(i)
            If Engine_RectCollision(X, Y, 1, 1, .X, .Y, .Width, .Height) Then
                If Button = vbLeftButton Then
                    frmSetTile.GrhTxt.Text = .GrhIndex
                Else
                    SetInfo "Click to select Grh " & .GrhIndex
                End If
                Exit For
            End If
        End With
    Next i

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ShownTextureAnims.NumGrhs = 0

End Sub
