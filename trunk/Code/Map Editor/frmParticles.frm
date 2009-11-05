VERSION 5.00
Begin VB.Form frmParticles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Particle Effects"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin VB.CheckBox EditChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Edit Mode"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox DirTxt 
      Appearance      =   0  'Flat
      Height          =   285
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
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Y co-ordinate of the effect (in pixels)"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox XTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "X co-ordinate of the effect (in pixels)"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox ParticlesTxt 
      Appearance      =   0  'Flat
      Height          =   285
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
      Height          =   285
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
      Height          =   285
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
      Height          =   1590
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "The list of particle effect slots, whether used or unused"
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label CreateBtnl 
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
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   3360
      Width           =   570
   End
   Begin VB.Label RefreshBtn 
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
      Height          =   195
      Left            =   2160
      TabIndex        =   15
      Top             =   3000
      Width           =   675
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   14
      Top             =   3390
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   1080
      TabIndex        =   13
      Top             =   2670
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2670
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3030
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1290
      TabIndex        =   9
      Top             =   2295
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2295
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
      ForeColor       =   &H80000008&
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
      ForeColor       =   &H80000008&
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

Private Sub CreateBtnl_Click()
Dim Gfx As Byte
Dim Particles As Integer
Dim EffectIndex As Integer
Dim Dir As Single
Dim X As Single
Dim Y As Single

On Error GoTo ErrOut

    'Set the values to the variables first, just to be sure an invalid range is caught
    Gfx = Val(GfxTxt.Text)
    Particles = Val(ParticlesTxt.Text)
    EffectIndex = Val(IndexTxt.Text)
    If EffectIndex < 1 Then GoTo ErrOut
    Dir = Val(DirTxt.Text)
    X = Val(XTxt.Text)
    Y = Val(YTxt.Text)

    'Create the particle effect
    Effect_Begin EffectIndex, X - (ParticleOffsetX - 288), Y - (ParticleOffsetY - 288), Gfx, Particles, Dir, True

    'Update list
    UpdateEffectList

Exit Sub

ErrOut:

    MsgBox "Error creating the particle effect! Aborting...", vbOKOnly

End Sub

Private Sub CreateBtnl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Create a particle effect with the defined values."

End Sub

Private Sub DirTxt_Change()
Dim Parts As Integer

    If EditChk.Value = 0 Then Exit Sub
    On Error GoTo ErrOut
    
    Parts = Val(ParticlesTxt.Text)
    If Effect(ParticlesList.ListIndex + 1).Used Then Effect(ParticlesList.ListIndex + 1).Direction = Parts
    
ErrOut:

End Sub

Private Sub DirTxt_KeyPress(KeyAscii As Integer)
    If ParticlesList.ListIndex + 1 = WeatherEffectIndex Or EditChk.Value = False Then
        KeyAscii = 0
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub DirTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Direction the effect is animating. Only applies to certain effects."

End Sub

Private Sub EditChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "When enabled, modifying effect values will update the current effect. When off, changes have no affect."

End Sub

Private Sub Form_Load()

    'Update list
    UpdateEffectList

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub GfxTxt_Change()
Dim Gfx As Byte

    If EditChk.Value = 0 Then Exit Sub
    On Error GoTo ErrOut
    
    Gfx = Val(GfxTxt.Text)
    If Gfx < 1 Then GoTo ErrOut
    If Gfx > UBound(ParticleTexture) Then GoTo ErrOut
    If Effect(ParticlesList.ListIndex + 1).Used Then Effect(ParticlesList.ListIndex + 1).Gfx = Gfx
    
ErrOut:

End Sub

Private Sub GfxTxt_KeyPress(KeyAscii As Integer)
    If ParticlesList.ListIndex + 1 = WeatherEffectIndex Or EditChk.Value = False Then
        KeyAscii = 0
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub GfxTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The particle graphic index, based off of the p#.bmp number."

End Sub

Private Sub IndexTxt_Change()
Dim i As Byte

    If EditChk.Value = 0 Then Exit Sub
    On Error GoTo ErrOut
    
    i = Val(IndexTxt.Text)
    If i < 1 Then GoTo ErrOut
    If Effect(ParticlesList.ListIndex + 1).Used Then Effect(ParticlesList.ListIndex + 1).EffectNum = i
    
ErrOut:

End Sub

Private Sub IndexTxt_KeyPress(KeyAscii As Integer)
    If ParticlesList.ListIndex + 1 = WeatherEffectIndex Or EditChk.Value = False Then
        KeyAscii = 0
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub IndexTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The ID of the effect, based off the EffectNum_ value in Particles module."

End Sub

Private Sub ParticlesList_Click()

    On Error GoTo ErrOut

    If Effect(ParticlesList.ListIndex + 1).Used = False Then GoTo ErrOut
    If WeatherEffectIndex = ParticlesList.ListIndex + 1 Then GoTo ErrOut
    
    With Effect(ParticlesList.ListIndex + 1)
        IndexTxt.Text = .EffectNum
        ParticlesTxt.Text = .ParticleCount
        GfxTxt.Text = .Gfx
        XTxt.Text = .X + ParticleOffsetX - 288
        YTxt.Text = .Y + ParticleOffsetY - 288
        DirTxt.Text = .Direction
    End With
    
    Exit Sub
    
ErrOut:

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

Private Sub ParticlesList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "List of currently active particle effects on this map. Shift + RightClick on the screen to move the selected effect."

End Sub

Private Sub ParticlesTxt_Change()
Dim Parts As Long
Dim j As Long

    If EditChk.Value = 0 Then Exit Sub
    On Error GoTo ErrOut
    
    Parts = Val(ParticlesTxt.Text)
    If Parts < 1 Then GoTo ErrOut
    If Effect(ParticlesList.ListIndex + 1).Used Then
        ReDim Effect(ParticlesList.ListIndex + 1).Particles(0 To Parts)
        ReDim Effect(ParticlesList.ListIndex + 1).PartVertex(0 To Parts)
        For j = 0 To Parts
            Set Effect(ParticlesList.ListIndex + 1).Particles(j) = New Particle
            Effect(ParticlesList.ListIndex + 1).Particles(j).Used = True
            Effect(ParticlesList.ListIndex + 1).PartVertex(j).Rhw = 1
        Next j
        Effect(ParticlesList.ListIndex + 1).ParticleCount = Parts
    End If
    
ErrOut:

End Sub

Private Sub ParticlesTxt_KeyPress(KeyAscii As Integer)
    If ParticlesList.ListIndex + 1 = WeatherEffectIndex Or EditChk.Value = False Then
        KeyAscii = 0
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub ParticlesTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Number of particles the effect has."

End Sub

Private Sub RefreshBtn_Click()

    UpdateEffectList

End Sub

Private Sub RefreshBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Refresh the active particle effects list."

End Sub

Private Sub Timer1_Timer()

    If frmParticles.Visible Then
        UpdateEffectList
    End If

End Sub

Private Sub XTxt_Change()
Dim X As Single

    On Error GoTo ErrOut
    
    If EditChk.Value = 0 Then Exit Sub
    X = Val(XTxt.Text) - ParticleOffsetX
    If Effect(ParticlesList.ListIndex + 1).Used Then Effect(ParticlesList.ListIndex + 1).X = X + 288
    
ErrOut:

End Sub

Private Sub XTxt_KeyPress(KeyAscii As Integer)
    If ParticlesList.ListIndex + 1 = WeatherEffectIndex Or EditChk.Value = False Then
        KeyAscii = 0
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub XTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "X co-ordinate of the effect (in pixels)."

End Sub

Private Sub YTxt_Change()
Dim Y As Single

    On Error GoTo ErrOut
    
    If EditChk.Value = 0 Then Exit Sub
    Y = Val(YTxt.Text) - ParticleOffsetY
    If Effect(ParticlesList.ListIndex + 1).Used Then Effect(ParticlesList.ListIndex + 1).Y = Y + 288
    
ErrOut:

End Sub

Private Sub YTxt_KeyPress(KeyAscii As Integer)
    If ParticlesList.ListIndex + 1 = WeatherEffectIndex Or EditChk.Value = False Then
        KeyAscii = 0
        Exit Sub
    End If
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub YTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Y co-ordinate of the effect (in pixels)."

End Sub
