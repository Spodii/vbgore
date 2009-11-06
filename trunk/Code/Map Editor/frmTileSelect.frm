VERSION 5.00
Begin VB.Form frmTileSelect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tile Selection"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmTileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim RetVal As Long

    'Select a tile and place it in the clipboard
    If Button = vbLeftButton Then
        RetVal = PreviewGrhList((Int(X / tsTileWidth) * tsHeight) + Int(Y / tsTileHeight)).GrhIndex
        Select Case stBoxID
            Case 0
                frmTile.GrhTxt.Text = RetVal
            Case Else
                frmSetTile.GrhTxt.Text = RetVal
        End Select
        HideFrmTileSelect
        
    'Show menu
    ElseIf Button = vbRightButton Then
        Me.Enabled = False
        ShowFrmTSOpt
    End If
    
End Sub

Private Sub LeftPic_Click()

    tsStart = tsStart - (tsWidth * tsHeight)
    Engine_SetTileSelectionArray

End Sub

Private Sub RightPic_Click()

    tsStart = tsStart + (tsWidth * tsHeight)
    Engine_SetTileSelectionArray

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    HideFrmTileSelect

End Sub
