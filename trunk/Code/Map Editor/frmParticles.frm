VERSION 5.00
Begin VB.Form frmParticles 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Particle Effects"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmParticles.frx":0000
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox DirTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "Direction the effect is animating "
      Top             =   3480
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
      Top             =   3000
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
      Top             =   3000
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
      Top             =   3240
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
      Top             =   2760
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
      Top             =   2760
      Width           =   615
   End
   Begin VB.ListBox ParticlesList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "The list of particle effect slots, whether used or unused"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label RefreshLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
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
      Left            =   2040
      TabIndex        =   17
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* = ?"
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
      Left            =   2280
      TabIndex        =   16
      ToolTipText     =   "Help on what the *s mean"
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dir*:"
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
      TabIndex        =   15
      Top             =   3480
      Width           =   390
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
      TabIndex        =   14
      Top             =   3000
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
      TabIndex        =   12
      Top             =   3000
      Width           =   195
   End
   Begin VB.Label CreateLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   570
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
      Top             =   3240
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
      Top             =   2760
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
      Top             =   2760
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
      Top             =   2520
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
      Top             =   600
      Width           =   1350
   End
End
Attribute VB_Name = "frmParticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CreateLbl_Click()
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

Private Sub Form_Load()

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
    Engine_Var_Write Data2Path & "MapEditor.ini", "PART", "X", Me.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "PART", "Y", Me.Top
    HideFrmParticles

End Sub

Private Sub MiscLbl_Click(Index As Integer)

    'Display help on the *
    If Index = 8 Then
        MsgBox "Items marked with a * are optional. Whether or not they are applicable to the effect" & vbCrLf _
         & " you are using can be found by checking if that variable is set on the effect in sub Effect_Begin.", vbOKOnly
    End If
    
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

Private Sub RefreshLbl_Click()

    UpdateEffectList

End Sub
