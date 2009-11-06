Attribute VB_Name = "General"
Option Explicit

Public Enum LogType
    General = 0
    CodeTracker = 1
    PacketIn = 2
    PacketOut = 3
    CriticalError = 4
    InvalidPacketData = 5
End Enum

Public Type NPCTradeItems
    name As String
    Price As Long
    GrhIndex As Long
End Type

Public NumBytesForSkills As Long

Public NPCTradeItems() As NPCTradeItems
Public NPCTradeItemArraySize As Byte
Private SkillPos As Long

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Log(ByVal DummyT As String, ByVal DummyB As LogType)

'***************************************************
'Dummy routine for logs from the server since some files are shared between multiple projects
'***************************************************

End Sub

Public Function Engine_BuildSkinsList() As String

'***************************************************
'Returns the list of all the skins
'***************************************************
Dim TempSplit() As String
Dim Files() As String
Dim i As Long

    'Get the list of files
    Files() = AllFilesInFolders(DataPath & "Skins\", False)
    
    'Show the header message
    Engine_AddToChatTextBuffer "The following skins are available:", FontColor_Info
    
    'Look for files ending with ".ini" only
    For i = LBound(Files) To UBound(Files)
        If Right$(Files(i), 4) = ".ini" Then
            
            'Crop out the skin name and add it to the function
            TempSplit() = Split(Files(i), "\")
            If Engine_BuildSkinsList <> "" Then Engine_BuildSkinsList = Engine_BuildSkinsList & vbCrLf
            Engine_BuildSkinsList = Engine_BuildSkinsList & " * |" & Left$(TempSplit(UBound(TempSplit)), Len(TempSplit(UBound(TempSplit))) - 4) & "|"

        End If
    Next i
    
End Function

Private Sub Draw_Stat(ByVal SkillName As String, ByVal Base As Long, ByVal Modi As Long)

'***************************************************
'Renders the skills to the skill box
'***************************************************

Dim RaiseCost As Long

'Calculate the cost to raise the skill

    RaiseCost = Game_GetEXPCost(Base)

    'Draw the skill's information
    If BaseStats(SID.Points) >= RaiseCost Then
        Engine_Render_Text "+", 0, SkillPos, -16777216    'ARGB(255,0,0,0)
    End If
    Engine_Render_Text SkillName, 8, SkillPos, -16777216
    Engine_Render_Text Base & "(" & Modi & ")", 90, SkillPos, -16777216  'ARGB(255,0,0,0)
    Engine_Render_Text Str$(RaiseCost), 150, SkillPos, -16777216 'ARGB(255,0,0,0)

    'Raise the skill pos
    SkillPos = SkillPos + 12

End Sub

Function Game_CheckUserData() As Boolean

'*****************************************************************
'Checks all user data for mistakes and reports them.
'*****************************************************************
Dim LoopC As Integer

    'Password
    If Len(UserPassword) < 3 Then
        MsgBox ("Password box is empty.")
        Exit Function
    End If
    If Len(UserPassword) > 10 Then
        MsgBox ("Password must be 10 characters or less.")
        Exit Function
    End If
    If Game_LegalString(UserPassword) = False Then
        MsgBox ("Invalid Password.")
        Exit Function
    End If
    
    'Name
    If Len(UserName) < 3 Then
        MsgBox ("Name box is empty.")
        Exit Function
    End If
    If Len(UserName) > 10 Then
        MsgBox ("Name must be 10 characters or less.")
        Exit Function
    End If
    If Game_LegalString(UserName) = False Then
        MsgBox ("Invalid Name.")
        Exit Function
    End If
    
    'If all good send true
    Game_CheckUserData = True

End Function

Function Game_ClickItem(ByVal ItemIndex As Byte, Optional ByVal InventoryType As Long = 1) As Long

'***************************************************
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    
    'Make sure item index is within limits
    If ItemIndex <= 0 Then Exit Function
    If ItemIndex > MAX_INVENTORY_SLOTS Then Exit Function
    
    'Check by the appropriate window
    Select Case InventoryType
        
        'User inventory
        Case 1
            If UserInventory(ItemIndex).GrhIndex > 0 Then Game_ClickItem = 1
            
        'Shop inventory
        Case 2
            If NPCTradeItems(ItemIndex).GrhIndex > 0 Then Game_ClickItem = 1
        
        'Bank depot
        Case 3
            If UserBank(ItemIndex).GrhIndex > 0 Then Game_ClickItem = 1
            
    End Select

End Function

Function Game_GetEXPCost(BaseSkill As Long) As Long

'*****************************************************************
'Calculate the exp required to raise a skill up to the next point
'*****************************************************************

    Game_GetEXPCost = Int(0.17376 * (BaseSkill ^ 3) + 0.44 * (BaseSkill ^ 2) - 0.48 * BaseSkill + 1.035) + 1

End Function

Function Game_ValidCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is used for chat/etc)
'Make sure you update the server's Server_ValidCharacter, too!
'*****************************************************************

    Log "Call Game_ValidCharacter(" & KeyAscii & ")", CodeTracker '//\\LOGLINE//\\

    If KeyAscii >= 32 Then Game_ValidCharacter = True

End Function

Function Game_LegalCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is for username/pass)
'Make sure you update the server's Server_ValidCharacter, too!
'*****************************************************************

    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
    'Allow characters A to Z
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
    'Allow characters a to z
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
    'Allow foreign characters
    If KeyAscii >= 128 And KeyAscii <= 168 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
Exit Function

ErrOut:

    'Something bad happened, so the character must be invalid
    Game_LegalCharacter = False
    
End Function

Function Game_ValidString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for illegal characters in the string (wrapper for Game_ValidCharacter)
'*****************************************************************
Dim i As Long

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then Exit Function
    If LenB(CheckString) < 1 Then Exit Function

    'Loop through the string
    For i = 1 To Len(CheckString)
        
        'Check the values
        If Game_ValidCharacter(AscB(Mid$(CheckString, i, 1))) = False Then Exit Function
        
    Next i
    
    'If we have made it this far, then all is good
    Game_ValidString = True

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Game_ValidString = False

End Function

Function Game_LegalString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for illegal characters in the string (wrapper for Server_LegalCharacter)
'*****************************************************************
Dim i As Long

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then Exit Function
    If LenB(CheckString) < 1 Then Exit Function

    'Loop through the string
    For i = 1 To Len(CheckString)
        
        'Check the values
        If Game_LegalCharacter(AscB(Mid$(CheckString, i, 1))) = False Then Exit Function
        
    Next i
    
    'If we have made it this far, then all is good
    Game_LegalString = True

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Game_LegalString = False

End Function

Public Sub Game_Config_Load()

'***************************************************
'Load the user configuration
'***************************************************

Dim i As Byte

    'Quickbar
    For i = 1 To 12
        QuickBarID(i).ID = Val(Engine_Var_Get(DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "ID"))
        QuickBarID(i).Type = Val(Engine_Var_Get(DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "Type"))
    Next i
    
    'Skin
    CurrentSkin = Engine_Var_Get(DataPath & "Game.ini", "INIT", "CurrentSkin")

End Sub

Sub Game_Map_Switch(Map As Integer)

'*****************************************************************
'Loads and switches to a new map
'*****************************************************************
Dim GetParticleCount As Integer
Dim GetEffectNum As Byte
Dim GetDirection As Integer
Dim GetGfx As Byte
Dim GetX As Integer
Dim GetY As Integer
Dim ByFlags As Long
Dim MapNum As Byte
Dim i As Integer
Dim Y As Byte
Dim X As Byte

    'Clear the offset values for the particle engine
    ParticleOffsetX = 0
    ParticleOffsetY = 0
    LastOffsetX = 0
    LastOffsetY = 0

    'Erase characters
    For i = 1 To LastChar
        If CharList(i).Active Then Engine_Char_Erase i
    Next i

    'Erase objects
    For i = 1 To LastObj
        OBJList(i).Grh.GrhIndex = 0
    Next i
    
    'Erase map-bound particle effects
    For i = 1 To NumEffects
        If Effect(i).Used Then
            If Effect(i).BoundToMap Then Effect_Kill i
        End If
    Next i

    'Open map file
    MapNum = FreeFile
    Open MapPath & Map & ".map" For Binary As #MapNum
    Seek #MapNum, 1

    'Map Header
    Get #MapNum, , MapInfo.MapVersion

    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            'Clear the graphic layers
            For i = 1 To 6
                MapData(X, Y).Graphic(i).GrhIndex = 0
            Next i

            'Get flag's byte
            Get #MapNum, , ByFlags

            'Blocked
            If ByFlags And 1 Then Get #MapNum, , MapData(X, Y).Blocked Else MapData(X, Y).Blocked = 0

            'Graphic layers
            If ByFlags And 2 Then
                Get #MapNum, , MapData(X, Y).Graphic(1).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            End If
            If ByFlags And 4 Then
                Get #MapNum, , MapData(X, Y).Graphic(2).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            End If
            If ByFlags And 8 Then
                Get #MapNum, , MapData(X, Y).Graphic(3).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            End If
            If ByFlags And 16 Then
                Get #MapNum, , MapData(X, Y).Graphic(4).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            End If
            If ByFlags And 32 Then
                Get #MapNum, , MapData(X, Y).Graphic(5).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(5), MapData(X, Y).Graphic(5).GrhIndex
            End If
            If ByFlags And 64 Then
                Get #MapNum, , MapData(X, Y).Graphic(6).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(6), MapData(X, Y).Graphic(6).GrhIndex
            End If
            
            'Set light to default (-1) - it will be set again if it is not -1 from the code below
            For i = 1 To 24
                MapData(X, Y).Light(i) = -1
            Next i
            
            'Get lighting values
            If ByFlags And 128 Then
                For i = 1 To 4
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If

            'Store the lighting in the SaveLightBuffer
            For i = 1 To 24
                SaveLightBuffer(X, Y).Light(i) = MapData(X, Y).Light(i)
            Next i

            'Mailbox
            If ByFlags And 8192 Then MapData(X, Y).Mailbox = 1 Else MapData(X, Y).Mailbox = 0

            'Shadows
            If ByFlags And 16384 Then MapData(X, Y).Shadow(1) = 1 Else MapData(X, Y).Shadow(1) = 0
            If ByFlags And 32768 Then MapData(X, Y).Shadow(2) = 1 Else MapData(X, Y).Shadow(2) = 0
            If ByFlags And 65536 Then MapData(X, Y).Shadow(3) = 1 Else MapData(X, Y).Shadow(3) = 0
            If ByFlags And 131072 Then MapData(X, Y).Shadow(4) = 1 Else MapData(X, Y).Shadow(4) = 0
            If ByFlags And 262144 Then MapData(X, Y).Shadow(5) = 1 Else MapData(X, Y).Shadow(5) = 0
            If ByFlags And 524288 Then MapData(X, Y).Shadow(6) = 1 Else MapData(X, Y).Shadow(6) = 0
            
            'Clear any old sfx
            If Not MapData(X, Y).Sfx Is Nothing Then
                MapData(X, Y).Sfx.Stop
                Set MapData(X, Y).Sfx = Nothing
            End If
            
            'Set the sfx
            If ByFlags And 1048576 Then
                Get #MapNum, , i
                Engine_Sound_SetToMap i, X, Y
            End If
            
            'Blocked attack
            If ByFlags And 2097152 Then MapData(X, Y).BlockedAttack = 1 Else MapData(X, Y).BlockedAttack = 0
            
            'Sign
            If ByFlags And 4194304 Then Get #MapNum, , MapData(X, Y).Sign Else MapData(X, Y).Sign = 0
            
            'If there is a warp
            If ByFlags And 8388608 Then MapData(X, Y).Warp = 1 Else MapData(X, Y).Warp = 0

        Next X
    Next Y
    
    'Get the number of effects
    Get #MapNum, , Y

    'Store the individual particle effect types
    If Y > 0 Then
        For X = 1 To Y
            Get #MapNum, , GetEffectNum
            Get #MapNum, , GetX
            Get #MapNum, , GetY
            Get #MapNum, , GetParticleCount
            Get #MapNum, , GetGfx
            Get #MapNum, , GetDirection
            Effect_Begin GetEffectNum, GetX, GetY, GetGfx, GetParticleCount, GetDirection
        Next X
    End If
    
    Close #MapNum
    
    'Create the minimap
    Engine_BuildMiniMap

    'Clear out old mapinfo variables
    MapInfo.name = vbNullString

    'Set current map
    CurMap = Map

End Sub

Public Sub Game_Config_Save()

'***************************************************
'Load the user configuration
'***************************************************
Dim t As String
Dim i As Byte

    'Quickbar
    For i = 1 To 12
        Engine_Var_Write DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "ID", Str$(QuickBarID(i).ID)
        Engine_Var_Write DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "Type", Str$(QuickBarID(i).Type)
    Next i
    
    'Skin
    Engine_Var_Write DataPath & "Game.ini", "INIT", "CurrentSkin", CurrentSkin
    
    'Skin positions
    t = DataPath & "Skins\" & CurrentSkin & ".dat"   'Set the custom positions file for the skin
    With GameWindow
        Engine_Var_Write t, "QUICKBAR", "ScreenX", Str(.QuickBar.Screen.X)
        Engine_Var_Write t, "QUICKBAR", "ScreenY", Str(.QuickBar.Screen.Y)
        Engine_Var_Write t, "CHATWINDOW", "ScreenX", Str(.ChatWindow.Screen.X)
        Engine_Var_Write t, "CHATWINDOW", "ScreenY", Str(.ChatWindow.Screen.Y)
        Engine_Var_Write t, "INVENTORY", "ScreenX", Str(.Inventory.Screen.X)
        Engine_Var_Write t, "INVENTORY", "ScreenY", Str(.Inventory.Screen.Y)
        Engine_Var_Write t, "SHOP", "ScreenX", Str(.Shop.Screen.X)
        Engine_Var_Write t, "SHOP", "ScreenY", Str(.Shop.Screen.Y)
        Engine_Var_Write t, "MAILBOX", "ScreenX", Str(.Mailbox.Screen.X)
        Engine_Var_Write t, "MAILBOX", "ScreenY", Str(.Mailbox.Screen.Y)
        Engine_Var_Write t, "VIEWMESSAGE", "ScreenX", Str(.ViewMessage.Screen.X)
        Engine_Var_Write t, "VIEWMESSAGE", "ScreenY", Str(.ViewMessage.Screen.Y)
        Engine_Var_Write t, "WRITEMESSAGE", "ScreenX", Str(.WriteMessage.Screen.X)
        Engine_Var_Write t, "WRITEMESSAGE", "ScreenY", Str(.WriteMessage.Screen.Y)
        Engine_Var_Write t, "AMOUNT", "ScreenX", Str(.Amount.Screen.X)
        Engine_Var_Write t, "AMOUNT", "ScreenY", Str(.Amount.Screen.Y)
        Engine_Var_Write t, "MENU", "ScreenX", Str(.Menu.Screen.X)
        Engine_Var_Write t, "MENU", "ScreenY", Str(.Menu.Screen.Y)
        Engine_Var_Write t, "BANK", "ScreenX", Str(.Bank.Screen.X)
        Engine_Var_Write t, "BANK", "ScreenY", Str(.Bank.Screen.Y)
    End With

End Sub

Sub UpdateShownTextBuffer()

'*****************************************************************
'Updates the ShownTextBuffer
'*****************************************************************
Dim X As Long
Dim Y As Long
Dim j As Long
    
    'Check if the width is larger then the screen
    If EnterTextBufferWidth > GameWindow.ChatWindow.Text.Width - 24 Then
        
        'Loop through the characters backwards
        For X = Len(EnterTextBuffer) To 1 Step -1
            
            'Add up the size
            j = j + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(EnterTextBuffer, X, 1)))
            
            'Check if the size has become too large
            If j > GameWindow.ChatWindow.Text.Width - 24 Then
            
                'If the size has become too large, the character before (since we are looping backwards, it is + 1) is the limit
                ShownText = Right$(EnterTextBuffer, Len(EnterTextBuffer) - X + 1)
                Exit For
                
            End If
        Next X
    Else
    
        'Set the shown text buffer to the full buffer
        ShownText = EnterTextBuffer
    
    End If

End Sub

Sub Main()

'*****************************************************************
'Main
'*****************************************************************
Dim StartTime As Long
Dim FileNum As Byte
Dim i As Integer

    'Init file paths
    InitFilePaths
    
    'Load frmMain
    Load frmMain
    frmMain.Hide
    DoEvents

    'Check if we need to run the updater
    If ForceUpdateCheck = True Then
    
        'Check for the right parameter
        If Command$ <> "-sdf@041jkdf0)21`~" Then

            'Force the creation of frmConnect, thus forcing the creation of its hWnd
            Load frmConnect
            frmConnect.Show
            frmConnect.Hide
            
            'Load the updater
            ShellExecute frmConnect.hWnd, vbNullString, App.Path & "\UpdateClient.exe", vbNullString, vbNullString, 1   'The 1 means "show normal"
    
            'Unload the client
            Engine_UnloadAllForms
            End
        
        End If
    End If
    
    'Generate the packet keys
    GenerateEncryptionKeys
    
    'Number of bytes required to fill the skills
    NumBytesForSkills = Int((NumSkills - 1) / 8) + 1
    
    'Load the font information
    Engine_Init_FontSettings
    
    'Load the messages
    Engine_Init_Messages LCase$(Engine_Var_Get(DataPath & "Game.ini", "INIT", "Language"))

    'Fill startup variables for the tile engine
    TilePixelWidth = 32
    TilePixelHeight = 32
    WindowTileHeight = 18
    WindowTileWidth = 25
    TileBufferSize = 10
    EnterTextBufferWidth = 1
    EngineBaseSpeed = 0.011
    ReDim SkillListIDs(1 To NumSkills)
    LineBreakChr = Chr$(10)

    'Setup borders
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = 4
    ScrollPixelsPerFrameY = 4
    ShowGameWindow(QuickBarWindow) = 1
    ShowGameWindow(ChatWindow) = 1

    'Set the array sizes by the number of graphic files
    NumGrhFiles = CInt(Engine_Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    ReDim SurfaceTimer(1 To NumGrhFiles)
    
    'Load graphic data into memory
    Engine_Init_GrhData
    Engine_Init_BodyData
    Engine_Init_WeaponData
    Engine_Init_WingData
    Engine_Init_HeadData
    Engine_Init_HairData
    Engine_Init_MapData
    Engine_Init_Signs
    
    'Load the config
    Game_Config_Load
    Engine_Init_GUI

    'Create the buffer
    Set sndBuf = New DataBuffer
    sndBuf.Clear

    'Set the form starting positions
    DoEvents

    'Load the data commands
    InitDataCommands

    'Display connect window
    frmConnect.Visible = True

    'Main Loop
    Do
    
        'Calculate the starttime - this is the absolute time it takes from start to finish, disincluding DoEvents
        ' The idea is that it works just like the ElapsedTime, but in slightly different placing
        StartTime = timeGetTime
    
        'Check if unloading
        If IsUnloading = 1 Then
            GOREsock_UnHook
            Exit Do
        End If

        'Don't draw frame is window is minimized or there is no map loaded
        If frmMain.WindowState <> 1 Then
            If CurMap > 0 Then

                'Show the next frame
                Engine_ShowNextFrame

                'Check for key inputs
                Engine_Input_CheckKeys
                
                'Keep the music looping
                If MapInfo.Music > 0 Then Engine_Music_Loop 1

                'Check to unload surfaces
                For i = 1 To NumGrhFiles

                    'Only update surfaces in use
                    If SurfaceTimer(i) > 0 Then

                        'Lower the counter
                        SurfaceTimer(i) = SurfaceTimer(i) - ElapsedTime

                        'Unload the surface
                        If SurfaceTimer(i) <= 0 Then
                            Set SurfaceDB(i) = Nothing
                            SurfaceTimer(i) = 0
                        End If

                    End If

                Next i
                
                'Check to unload sound buffers
                For i = 1 To NumSfx
                
                    'Only update sound buffers in use
                    If SoundBufferTimer(i) > 0 Then
                        
                        'Lower the counter
                        SoundBufferTimer(i) = SoundBufferTimer(i) - ElapsedTime
                        
                        'Unload the sound buffer
                        If SoundBufferTimer(i) <= 0 Then
                            Set DSBuffer(i) = Nothing
                            SoundBufferTimer(i) = 0
                        End If
                        
                    End If
                    
                Next i

            End If
        End If

        If SocketOpen Then
        
            'Send the data buffer
            Data_Send

        End If
        
        'Too many failed pings, we disconnect
        If FailedPings > 2 Then IsUnloading = 1
        
        'Do other events
        DoEvents
        
        'Do sleep event - force FPS at ~60 (62.5) average (prevents extensive processing)
        If (timeGetTime - StartTime) < 16 Then  'If Elapsed Time < Time required for 60 FPS
            Sleep 16 - (timeGetTime - StartTime)
        End If

    Loop

    'Save the config
    Game_Config_Save
    
    'Close down
    frmMain.ShutdownTimer.Enabled = True

End Sub
