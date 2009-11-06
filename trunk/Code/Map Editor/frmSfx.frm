VERSION 5.00
Begin VB.Form frmSfx 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add Sfx"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   120
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SfxTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The number of the .wav file that will be looped on the tile for stuff like waterfalls, birds, etc - set to 0 for nothing"
      Top             =   120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   135
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

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub SfxTxt_Change()
Dim i As Integer

    On Error GoTo ErrOut
    
    i = Val(SfxTxt.Text)

    Exit Sub
    
ErrOut:

End Sub

Private Sub SfxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub SfxTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The number of the .wav file that will be looped on the tile for sounds of waterfalls, birds, etc. Set to 0 for nothing."

End Sub
