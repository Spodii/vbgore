VERSION 5.00
Begin VB.Form frmSearchTexture 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Texture: 0"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSearchTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Form_MouseMove Button, Shift, x, y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
    
    SetInfo vbNullString
    For i = 1 To ShownTextureGrhs.NumGrhs
        With ShownTextureGrhs.Grh(i)
            If Engine_RectCollision(x, y, 1, 1, .x, .y, .Width, .Height) Then
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Form_MouseMove Button, Shift, x, y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'If IsUnloading = 0 Then Cancel = 1

End Sub

Private Sub Form_Resize()

    If SearchTextureFileNum > 0 Then Engine_Render_FullTexture frmSearchTexture.hWnd, SearchTextureFileNum

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ShownTextureGrhs.NumGrhs = 0

End Sub
