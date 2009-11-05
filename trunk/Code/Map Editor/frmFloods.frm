VERSION 5.00
Begin VB.Form frmFloods 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Floods"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFloods.frx":0000
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   93
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Flood all the tiles shown on the screen only"
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label AllLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Map"
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
      Height          =   195
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Flood the whole map, border and non-border"
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label InnerLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inner Map"
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
      Height          =   195
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Flood the inner map - the whole map except for the border"
      Top             =   840
      Width           =   900
   End
   Begin VB.Label BorderLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Border"
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
      Height          =   195
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Flood the border of the map only - this area is always blocked off"
      Top             =   600
      Width           =   900
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
        vbCrLf & "Set OBJs: " & CBool(frmObj.Visible And frmObj.SetOpt.Value = True) & _
        vbCrLf & "Erase OBJs: " & CBool(frmObj.Visible And frmObj.EraseOpt.Value = True) & _
        vbCrLf & "Set Tiles: " & CBool(frmSetTile.Visible), vbYesNo) = vbYes Then
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                SetTile X, Y, vbLeftButton, 0
            Next Y
        Next X
    End If

End Sub

Private Sub BorderLbl_Click()
Dim X As Byte
Dim Y As Byte

    'Flood the border
    If MsgBox("Are you sure you wish to flood the map border with the selected content?" & _
        vbCrLf & "Set NPCs: " & CBool(frmNPCs.Visible And frmNPCs.SetOpt.Value = True) & _
        vbCrLf & "Erase NPCs: " & CBool(frmNPCs.Visible And frmNPCs.EraseOpt.Value = True) & _
        vbCrLf & "Set OBJs: " & CBool(frmObj.Visible And frmObj.SetOpt.Value = True) & _
        vbCrLf & "Erase OBJs: " & CBool(frmObj.Visible And frmObj.EraseOpt.Value = True) & _
        vbCrLf & "Set Tiles: " & CBool(frmSetTile.Visible), vbYesNo) = vbYes Then
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                    SetTile X, Y, vbLeftButton, 0
                End If
            Next Y
        Next X
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Data2Path & "MapEditor.ini", "FLOODS", "X", Me.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "FLOODS", "Y", Me.Top
    HideFrmFloods

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

    'Close form
    If Button = vbLeftButton Then
        If X >= Me.ScaleWidth - 23 Then
            If X <= Me.ScaleWidth - 10 Then
                If Y <= 26 Then
                    If Y >= 11 Then
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub InnerLbl_Click()
Dim X As Byte
Dim Y As Byte

    'Flood the inner map
    If MsgBox("Are you sure you wish to flood the inner map (all but but the border) with the selected content?" & _
        vbCrLf & "Set NPCs: " & CBool(frmNPCs.Visible And frmNPCs.SetOpt.Value = True) & _
        vbCrLf & "Erase NPCs: " & CBool(frmNPCs.Visible And frmNPCs.EraseOpt.Value = True) & _
        vbCrLf & "Set OBJs: " & CBool(frmObj.Visible And frmObj.SetOpt.Value = True) & _
        vbCrLf & "Erase OBJs: " & CBool(frmObj.Visible And frmObj.EraseOpt.Value = True) & _
        vbCrLf & "Set Tiles: " & CBool(frmSetTile.Visible), vbYesNo) = vbYes Then
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                If Not (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then SetTile X, Y, vbLeftButton, 0
            Next Y
        Next X
    End If

End Sub

Private Sub ScreenLbl_Click()
Dim X As Byte
Dim Y As Byte

    'Flood the border
    If MsgBox("Are you sure you wish to flood the screen with the selected content?" & _
        vbCrLf & "Set NPCs: " & CBool(frmNPCs.Visible And frmNPCs.SetOpt.Value = True) & _
        vbCrLf & "Erase NPCs: " & CBool(frmNPCs.Visible And frmNPCs.EraseOpt.Value = True) & _
        vbCrLf & "Set OBJs: " & CBool(frmObj.Visible And frmObj.SetOpt.Value = True) & _
        vbCrLf & "Erase OBJs: " & CBool(frmObj.Visible And frmObj.EraseOpt.Value = True) & _
        vbCrLf & "Set Tiles: " & CBool(frmSetTile.Visible), vbYesNo) = vbYes Then
        For X = (UserPos.X - AddtoUserPos.X) - WindowTileWidth \ 2 To (UserPos.X - AddtoUserPos.X) + WindowTileWidth \ 2
            For Y = (UserPos.Y - AddtoUserPos.Y) - WindowTileHeight \ 2 To (UserPos.Y - AddtoUserPos.Y) + WindowTileHeight \ 2
                SetTile X, Y, vbLeftButton, 0
            Next Y
        Next X
    End If
    
End Sub
