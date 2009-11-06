VERSION 5.00
Begin VB.Form frmSfx 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Add Sfx"
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   124
   ShowInTaskbar   =   0   'False
   Begin MapEditor.cForm cForm 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "Add Sfx"
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.TextBox SfxTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
    Var_Write Data2Path & "MapEditor.ini", "SFX", "X", Me.Left
    Var_Write Data2Path & "MapEditor.ini", "SFX", "Y", Me.Top
    HideFrmSfx

End Sub

Private Sub Form_Load()

    cForm.LoadSkin Me
    Skin_Set Me

End Sub

Private Sub SfxTxt_Change()
Dim i As Integer

    On Error GoTo ErrOut
    
    i = Val(SfxTxt.Text)

    Exit Sub
    
ErrOut:

    SfxTxt.Text = 0

End Sub

Private Sub SfxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
