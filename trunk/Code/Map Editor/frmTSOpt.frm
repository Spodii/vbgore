VERSION 5.00
Begin VB.Form frmTSOpt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tile Select Options"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTSOpt.frx":0000
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox HeightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Height of the previewed tile"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox WidthTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Width of the previewed tile"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox StartTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The starting number to view the tiles from"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Height:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Width:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Number:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label CancelLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label SaveLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save Changes"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1245
   End
End
Attribute VB_Name = "frmTSOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    HideFrmTSOpt
    ShowFrmTileSelect stBoxID

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

Private Sub HeightTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub SaveLbl_Click()

    tsTileHeight = Val(HeightTxt.Text)
    tsTileWidth = Val(WidthTxt.Text)
    tsStart = Val(StartTxt.Text)

    Engine_Var_Write Data2Path & "MapEditor.ini", "TSOPT", "W", WidthTxt.Text
    Engine_Var_Write Data2Path & "MapEditor.ini", "TSOPT", "H", HeightTxt.Text
    Engine_Var_Write Data2Path & "MapEditor.ini", "TSOPT", "S", StartTxt.Text
    
    tsWidth = CLng(frmTileSelect.ScaleWidth / tsTileWidth)  'Use clng to make sure we round down
    tsHeight = CLng(frmTileSelect.ScaleHeight / tsTileHeight)
    ReDim PreviewGrhList(tsWidth * tsHeight)    'Resize our array accordingly to fit all our Grhs
    Engine_SetTileSelectionArray
    
    HideFrmTSOpt
    ShowFrmTileSelect stBoxID

End Sub

Private Sub StartTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub WidthTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
