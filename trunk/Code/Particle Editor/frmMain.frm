VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Particle Editor"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox LoopChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Force Loop"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DispEffect As Byte  'Index of the displayed effect

Private ResetX As Single
Private ResetY As Single

Private Sub Form_Load()

'Init particle engine

    Me.Show
    Engine_Init_TileEngine Me.hWnd, 32, 32, 1, 1, 1, 0.011

    'Set initial reset position (center screen)
    ResetX = frmMain.ScaleWidth * 0.5
    ResetY = frmMain.ScaleHeight * 0.5

    'Create initial effect
    ResetEffect

    'Main loop
    EngineRun = True

    Do While EngineRun

        'Reset if effect stopped and forceloop is on
        If Effect(DispEffect).Used = False Then
            If LoopChk.Value Then ResetEffect
        End If

        'Draw
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        D3DDevice.BeginScene
        Effect_UpdateAll
        D3DDevice.EndScene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

        'FPS
        ElapsedTime = Engine_ElapsedTime()
        If FPS_Last_Check + 1000 < timeGetTime Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            FPS_Last_Check = timeGetTime
            frmMain.Caption = "Particle Editor: FPS " & FPS
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If

        DoEvents

    Loop

    'Unload engine
    Engine_Init_UnloadTileEngine
    Engine_UnloadAllForms

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Recreate the effect

    If Button = vbRightButton Then
        Effect(DispEffect).Used = False
        ResetEffect
        Effect(DispEffect).x = x
        Effect(DispEffect).Y = Y
        ResetX = x
        ResetY = Y
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Reposition the effect

    If Button = vbLeftButton Then
        If Effect(DispEffect).Used = False Then ResetEffect
        Effect(DispEffect).x = x
        Effect(DispEffect).Y = Y
        ResetX = x
        ResetY = Y
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Stop the engine

    EngineRun = False

End Sub

Private Sub ResetEffect()
    
'Resets the effect - use this sub to change the effect displayed
    
    'DispEffect = Effect_Fire_Begin(ResetX, ResetY, 1, 150)
    'DispEffect = Effect_Heal_Begin(ResetX, ResetY, 9, 200, 1)
    
    'DispEffect = Effect_Bless_Begin(ResetX, ResetY, 3, 100, 40, 15)
    'DispEffect = Effect_Protection_Begin(ResetX, ResetY, 11, 100, 40, 15)
    'DispEffect = Effect_Strengthen_Begin(ResetX, ResetY, 12, 100, 40, 15)
    
    DispEffect = Effect_EquationTemplate_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 1, 1000)
    
    'DispEffect = Effect_Waterfall_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 2, 75)

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 18:13)  Decl: 62  Code: 106  Total: 168 Lines
':) CommentOnly: 71 (42.3%)  Commented: 1 (0.6%)  Empty: 30 (17.9%)  Max Logic Depth: 3
