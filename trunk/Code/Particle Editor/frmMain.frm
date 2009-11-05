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
'**       ____        _________   ______   ______  ______   _______           **
'**       \   \      /   /     \ /  ____\ /      \|      \ |   ____|          **
'**        \   \    /   /|      |  /     |        |       ||  |____           **
'***        \   \  /   / |     /| |  ___ |        |      / |   ____|         ***
'****        \   \/   /  |     \| |  \  \|        |   _  \ |  |____         ****
'******       \      /   |      |  \__|  |        |  | \  \|       |      ******
'********      \____/    |_____/ \______/ \______/|__|  \__\_______|    ********
'*******************************************************************************
'*******************************************************************************
'************ vbGORE - Visual Basic 6.0 Graphical Online RPG Engine ************
'************            Official Release: Version 0.1.1            ************
'************                 http://www.vbgore.com                 ************
'*******************************************************************************
'*******************************************************************************
'***** Source Distribution Information: ****************************************
'*******************************************************************************
'** If you wish to distribute this source code, you must distribute as-is     **
'** from the vbGORE website unless permission is given to do otherwise. This  **
'** comment block must remain in-tact in the distribution. If you wish to     **
'** distribute modified versions of vbGORE, please contact Spodi (info below) **
'** before distributing the source code. You may never label the source code  **
'** as the "Official Release" or similar unless the code and content remains  **
'** unmodified from the version downloaded from the official website.         **
'** You may also never sale the source code without permission first. If you  **
'** want to sell the code, please contact Spodi (below). This is to prevent   **
'** people from ripping off other people by selling an insignificantly        **
'** modified version of open-source code just to make a few quick bucks.      **
'*******************************************************************************
'***** Creating Engines With vbGORE: *******************************************
'*******************************************************************************
'** If you plan to create an engine with vbGORE that, please contact Spodi    **
'** before doing so. You may not sell the engine unless told elsewise (the    **
'** engine must has substantial modifications), and you may not claim it as   **
'** all your own work - credit must be given to vbGORE, along with a link to  **
'** the vbGORE homepage. Failure to gain approval from Spodi directly to      **
'** make a new engine with vbGORE will result in first a friendly reminder,   **
'** followed by much more drastic measures.                                   **
'*******************************************************************************
'***** Helping Out vbGORE: *****************************************************
'*******************************************************************************
'** If you want to help out with vbGORE's progress, theres a few things you   **
'** can do:                                                                   **
'**  *Donate - Great way to keep a free project going. :) Info and benifits   **
'**        for donating can be found at:                                      **
'**        http://www.vbgore.com/modules.php?name=Content&pa=showpage&pid=11  **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        create tutorials for the Knowledge Base. :)                        **
'**  *Ads - Advertisements have been placed on the site for those who can     **
'**        not or do not want to donate. Not donating is understandable - not **
'**        everyone has access to credit cards / paypal or spair money laying **
'**        around. These ads allow for a free way for you to help out the     **
'**        site. Those who do donate have the option to hide/remove the ads.  **
'*******************************************************************************
'***** Conact Information: *****************************************************
'*******************************************************************************
'** Please contact the creator of vbGORE (Spodi) directly with any questions: **
'** AIM: Spodii                          Yahoo: Spodii                        **
'** MSN: Spodii@hotmail.com              Email: spodi@vbgore.com              **
'** 2nd Email: spodii@hotmail.com        Website: http://www.vbgore.com       **
'*******************************************************************************
'***** Credits: ****************************************************************
'*******************************************************************************
'** Below are credits to those who have helped with the project or who have   **
'** distributed source code which has help this project's creation. The below **
'** is listed in no particular order of significance:                         **
'**                                                                           **
'** ORE (Aaron Perkins): Used as base engine and for learning experience      **
'**   http://www.baronsoft.com/                                               **
'** SOX (Trevor Herselman): Used for all the networking                       **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=35239&lngWId=1      **
'** Compression Methods (Marco v/d Berg): Provided compression algorithms     **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1      **
'** All Files In Folder (Jorge Colaccini): Algorithm implimented into engine  **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51435&lngWId=1      **
'** Game Programming Wiki (All community): Help on many different subjects    **
'**   http://wwww.gpwiki.org/                                                 **
'** ORE Maraxus's Edition (Maraxus): Used the map editor from this project    **
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'** Big thanks goes to Van, Nex666 and ChAsE01!                               **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

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
    
    'DispEffect = Effect_EquationTemplate_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 1, 1000)
    'DispEffect = Effect_EquationTemplate_Begin(0, Me.ScaleHeight * 0.5, 1, 1000)
    
    DispEffect = Effect_Waterfall_Begin(Me.ScaleWidth * 0.5, Me.ScaleHeight * 0.5, 2, 75)

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 18:13)  Decl: 62  Code: 106  Total: 168 Lines
':) CommentOnly: 71 (42.3%)  Commented: 1 (0.6%)  Empty: 30 (17.9%)  Max Logic Depth: 3
