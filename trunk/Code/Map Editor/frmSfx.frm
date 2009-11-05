VERSION 5.00
Begin VB.Form frmSfx 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Add Sfx"
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSfx.frx":0000
   ScaleHeight     =   61
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SfxTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The number of the .wav file that will be looped on the tile for stuff like waterfalls, birds, etc - set to 0 for nothing"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index:"
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
      TabIndex        =   1
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "frmSfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Ini2Path & "MapEditor.ini", "SFX", "X", Me.Left
    Engine_Var_Write Ini2Path & "MapEditor.ini", "SFX", "Y", Me.Top
    HideFrmSfx

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
Private Sub SfxTxt_Change()
Dim i As Integer

    On Error GoTo ErrOut
    
    i = Val(SfxTxt.Text)

    Exit Sub
    
ErrOut:

    SfxTxt.Text = 0

End Sub
