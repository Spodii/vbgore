VERSION 5.00
Begin VB.Form frmFloods 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Floods"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ShowInTaskbar   =   0   'False
   Begin VB.Label ScreenLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Flood all the tiles currently in the screen"
      Top             =   480
      Width           =   900
   End
   Begin VB.Label AllLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   375
      TabIndex        =   0
      ToolTipText     =   "Flood every tile on the map"
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmFloods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AllLbl_Click()
Dim X As Byte
Dim Y As Byte

    'Flood the map
    If MsgBox("Are you sure you wish to flood the whole map with the selected content?" & _
        vbCrLf & "Set NPCs: " & CBool(frmNPCs.Visible And frmNPCs.SetOpt.Value = True) & _
        vbCrLf & "Erase NPCs: " & CBool(frmNPCs.Visible And frmNPCs.EraseOpt.Value = True) & _
        vbCrLf & "Set Tiles: " & CBool(frmSetTile.Visible), vbYesNo) = vbYes Then
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
                SetTile X, Y, vbLeftButton, 0, True
            Next Y
        Next X
        Engine_BuildMiniMap
        Engine_CreateTileLayers
    End If

End Sub

Private Sub AllLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Flood every tile on the map."

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub ScreenLbl_Click()
Dim X As Integer
Dim Y As Integer

    'Flood the border
    If MsgBox("Are you sure you wish to flood the screen with the selected content?" & _
        vbCrLf & "Set NPCs: " & CBool(frmNPCs.Visible And frmNPCs.SetOpt.Value = True) & _
        vbCrLf & "Erase NPCs: " & CBool(frmNPCs.Visible And frmNPCs.EraseOpt.Value = True) & _
        vbCrLf & "Set Tiles: " & CBool(frmSetTile.Visible), vbYesNo) = vbYes Then
        For X = (UserPos.X - AddtoUserPos.X) - WindowTileWidth \ 2 To (UserPos.X - AddtoUserPos.X) + WindowTileWidth \ 2
            For Y = (UserPos.Y - AddtoUserPos.Y) - WindowTileHeight \ 2 To (UserPos.Y - AddtoUserPos.Y) + WindowTileHeight \ 2
                If X > 0 Then
                    If Y > 0 Then
                        If X <= MapInfo.Width Then
                            If Y <= MapInfo.Height Then
                                SetTile X, Y, vbLeftButton, 0, True
                            End If
                        End If
                    End If
                End If
            Next Y
        Next X
        Engine_BuildMiniMap
        Engine_CreateTileLayers
    End If
    
End Sub

Private Sub ScreenLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Flood all the tiles currently in the screen."

End Sub
