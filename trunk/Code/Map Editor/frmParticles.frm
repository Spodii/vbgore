VERSION 5.00
Begin VB.Form frmParticles 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Particle Effects"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   Begin MapEditor.cButton RefreshBtn 
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      ToolTipText     =   "Refresh the list of effects in play"
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Refresh"
   End
   Begin MapEditor.cForm cForm 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "Particle Effects"
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.TextBox DirTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "Direction the effect is animating "
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox YTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Y co-ordinate of the effect"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox XTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "X co-ordinate of the effect"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox ParticlesTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "Number of particles the effect has"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox GfxTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "The particle graphic index, based off of the p#.bmp number"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox IndexTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "The ID of the effect, based off the EffectNum_ value in Particles module"
      Top             =   2280
      Width           =   615
   End
   Begin VB.ListBox ParticlesList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "The list of particle effect slots, whether used or unused"
      Top             =   360
      Width           =   2895
   End
   Begin MapEditor.cButton CreateBtn 
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      ToolTipText     =   "Create an effect with the values entered to the left"
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Create"
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(?)"
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
      Index           =   8
      Left            =   1200
      TabIndex        =   15
      ToolTipText     =   "Help on what the *s mean"
      Top             =   3360
      Width           =   240
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dir:"
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
      Index           =   7
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
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
      Index           =   6
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
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
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particles:"
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
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gfx:"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create Effect:"
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
      TabIndex        =   7
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Effects:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "frmParticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CreateBtn_Click()
Dim Gfx As Byte
Dim Particles As Integer
Dim EffectIndex As Byte
Dim Dir As Single
Dim X As Single
Dim Y As Single

On Error GoTo ErrOut

    'Set the values to the variables first, just to be sure an invalid range is caught
    Gfx = Val(GfxTxt.Text)
    Particles = Val(ParticlesTxt.Text)
    EffectIndex = Val(IndexTxt.Text)
    Dir = Val(DirTxt.Text)
    X = Val(XTxt.Text)
    Y = Val(YTxt.Text)

    'Create the particle effect
    Effect_Begin EffectIndex, X - (ParticleOffsetX - 288), Y - (ParticleOffsetY - 288), Gfx, Particles, Dir

    'Update list
    UpdateEffectList

Exit Sub

ErrOut:

    MsgBox "Error creating the particle effect! Aborting...", vbOKOnly

End Sub

Private Sub DirTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()

    cForm.LoadSkin Me
    Skin_Set Me
    Me.Refresh
    
    'Update list
    UpdateEffectList

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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Var_Write Data2Path & "MapEditor.ini", "PART", "X", Me.Left
    Var_Write Data2Path & "MapEditor.ini", "PART", "Y", Me.Top
    HideFrmParticles

End Sub

Private Sub GfxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub IndexTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub MiscLbl_Click(Index As Integer)

    'Display help on the *
    If Index = 8 Then
        MsgBox "Items marked with a * are optional. Whether or not they are applicable to the effect" & vbCrLf _
         & " you are using can be found by checking if that variable is set on the effect in sub Effect_Begin.", vbOKOnly
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Control
    
    For Each c In Me
        If TypeName(c) = "cButton" Then
            c.Refresh
            c.DrawState = 0
        End If
    Next c
    Set c = Nothing
    
End Sub

Private Sub ParticlesList_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrOut

    'Delete effect
    If KeyCode = vbKeyDelete Then
        If Effect(ParticlesList.ListIndex + 1).Used = True Then
            Effect_Kill ParticlesList.ListIndex + 1
            UpdateEffectList
        End If
    End If
    
Exit Sub

ErrOut:
    
    MsgBox "Error deleting effect " & ParticlesList.ListIndex + 1 & "!", vbOKOnly
    
End Sub

Private Sub ParticlesTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub RefreshBtn_Click()

    UpdateEffectList

End Sub

Private Sub XTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub YTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
