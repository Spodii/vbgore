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
   Begin VB.PictureBox RightPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   360
      Picture         =   "frmTileSelect.frx":0000
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox LeftPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "frmTileSelect.frx":0222
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   0
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "frmTileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LeftPic_Click()

    tsStart = tsStart - (tsWidth * tsHeight)
    Engine_SetTileSelectionArray

End Sub

Private Sub RightPic_Click()

    tsStart = tsStart + (tsWidth * tsHeight)
    Engine_SetTileSelectionArray

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim RetVal As Long

    'Select a tile and place it in the clipboard
    If Button = vbLeftButton Then
        RetVal = PreviewGrhList((Int(x / tsTileWidth) * tsHeight) + Int(y / tsTileHeight)).GrhIndex
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    HideFrmTileSelect

End Sub
