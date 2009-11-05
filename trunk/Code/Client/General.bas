Attribute VB_Name = "General"
Option Explicit

Public Type NPCTradeItems
    Name As String
    Price As Long
    GrhIndex As Integer
End Type

Public NPCTradeItems() As NPCTradeItems
Public NPCTradeItemArraySize As Byte
Private SkillPos As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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
    If LenB(UserPassword) = 0 Then
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
    If LenB(UserName) = 0 Then
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

Function Game_ClickItem(ItemIndex As Byte, Optional ByVal InventoryType As Long = 1) As Long

'***************************************************
'Selects the item clicked if it's valid and return's it's index
'***************************************************
'Make sure item index is within limits

    If ItemIndex <= 0 Then Exit Function
    If ItemIndex > MAX_INVENTORY_SLOTS Then Exit Function
    'Make sure it's within limits
    Select Case InventoryType
    Case 1
        If UserInventory(ItemIndex).GrhIndex Then Game_ClickItem = 1
    End Select

End Function

Function Game_GetEXPCost(BaseSkill As Long) As Long

'*****************************************************************
'Calculate the exp required to raise a skill up to the next point
'*****************************************************************

    Game_GetEXPCost = Int(0.17376 * (BaseSkill ^ 3) + 0.44 * (BaseSkill ^ 2) - 0.48 * BaseSkill + 1.035) + 1

End Function

Function Game_LegalCharacter(KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters
'*****************************************************************

    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 Or KeyAscii <= 57 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
    'Allow letters A to Z
    If KeyAscii >= 65 Or KeyAscii <= 90 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
    'Allow letters a to z
    If KeyAscii >= 97 Or KeyAscii <= 122 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    
Exit Function

ErrOut:

    'Something bad happened, so the character must be invalid
    Game_LegalCharacter = False
    
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
Dim x As Byte

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
        For x = XMinMapSize To XMaxMapSize
        
            'Clear the graphic layers
            For i = 1 To 6
                MapData(x, Y).Graphic(i).GrhIndex = 0
            Next i

            'Get flag's byte
            Get #MapNum, , ByFlags

            'Blocked
            If ByFlags And 1 Then Get #MapNum, , MapData(x, Y).Blocked Else MapData(x, Y).Blocked = 0

            'Graphic layers
            If ByFlags And 2 Then
                Get #MapNum, , MapData(x, Y).Graphic(1).GrhIndex
                Engine_Init_Grh MapData(x, Y).Graphic(1), MapData(x, Y).Graphic(1).GrhIndex
            End If
            If ByFlags And 4 Then
                Get #MapNum, , MapData(x, Y).Graphic(2).GrhIndex
                Engine_Init_Grh MapData(x, Y).Graphic(2), MapData(x, Y).Graphic(2).GrhIndex
            End If
            If ByFlags And 8 Then
                Get #MapNum, , MapData(x, Y).Graphic(3).GrhIndex
                Engine_Init_Grh MapData(x, Y).Graphic(3), MapData(x, Y).Graphic(3).GrhIndex
            End If
            If ByFlags And 16 Then
                Get #MapNum, , MapData(x, Y).Graphic(4).GrhIndex
                Engine_Init_Grh MapData(x, Y).Graphic(4), MapData(x, Y).Graphic(4).GrhIndex
            End If
            If ByFlags And 32 Then
                Get #MapNum, , MapData(x, Y).Graphic(5).GrhIndex
                Engine_Init_Grh MapData(x, Y).Graphic(5), MapData(x, Y).Graphic(5).GrhIndex
            End If
            If ByFlags And 64 Then
                Get #MapNum, , MapData(x, Y).Graphic(6).GrhIndex
                Engine_Init_Grh MapData(x, Y).Graphic(6), MapData(x, Y).Graphic(6).GrhIndex
            End If
            
            'Set light to default (-1) - it will be set again if it is not -1 from the code below
            For i = 1 To 24
                MapData(x, Y).Light(i) = -1
            Next i
            
            'Get lighting values
            If ByFlags And 128 Then
                For i = 1 To 4
                    Get #MapNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    Get #MapNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    Get #MapNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    Get #MapNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    Get #MapNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    Get #MapNum, , MapData(x, Y).Light(i)
                Next i
            End If

            'Store the lighting in the SaveLightBuffer
            For i = 1 To 24
                SaveLightBuffer(x, Y).Light(i) = MapData(x, Y).Light(i)
            Next i

            'Mailbox
            If ByFlags And 8192 Then
                MapData(x, Y).Mailbox = 1
            Else
                MapData(x, Y).Mailbox = 0
            End If

            'Shadows
            If ByFlags And 16384 Then MapData(x, Y).Shadow(1) = 1 Else MapData(x, Y).Shadow(1) = 0
            If ByFlags And 32768 Then MapData(x, Y).Shadow(2) = 1 Else MapData(x, Y).Shadow(2) = 0
            If ByFlags And 65536 Then MapData(x, Y).Shadow(3) = 1 Else MapData(x, Y).Shadow(3) = 0
            If ByFlags And 131072 Then MapData(x, Y).Shadow(4) = 1 Else MapData(x, Y).Shadow(4) = 0
            If ByFlags And 262144 Then MapData(x, Y).Shadow(5) = 1 Else MapData(x, Y).Shadow(5) = 0
            If ByFlags And 524288 Then MapData(x, Y).Shadow(6) = 1 Else MapData(x, Y).Shadow(6) = 0
            
            'Clear any old sfx
            If Not MapData(x, Y).Sfx Is Nothing Then
                MapData(x, Y).Sfx.Stop
                Set MapData(x, Y).Sfx = Nothing
            End If
            
            'Set the sfx
            If ByFlags And 1048576 Then
                Get #MapNum, , i
                Engine_Sound_SetToMap i, x, Y
            End If

        Next x
    Next Y
    
    'Get the number of effects
    Get #MapNum, , Y

    'Store the individual particle effect types
    If Y > 0 Then
        For x = 1 To Y
            Get #MapNum, , GetEffectNum
            Get #MapNum, , GetX
            Get #MapNum, , GetY
            Get #MapNum, , GetParticleCount
            Get #MapNum, , GetGfx
            Get #MapNum, , GetDirection
            Effect_Begin GetEffectNum, GetX + 288, GetY + 288, GetGfx, GetParticleCount, GetDirection
        Next x
    End If
    
    Close #MapNum

    'Clear out old mapinfo variables
    MapInfo.Name = vbNullString

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
        Engine_Var_Write t, "QUICKBAR", "ScreenX", Str(.QuickBar.Screen.x)
        Engine_Var_Write t, "QUICKBAR", "ScreenY", Str(.QuickBar.Screen.Y)
        Engine_Var_Write t, "CHATWINDOW", "ScreenX", Str(.ChatWindow.Screen.x)
        Engine_Var_Write t, "CHATWINDOW", "ScreenY", Str(.ChatWindow.Screen.Y)
        Engine_Var_Write t, "INVENTORY", "ScreenX", Str(.Inventory.Screen.x)
        Engine_Var_Write t, "INVENTORY", "ScreenY", Str(.Inventory.Screen.Y)
        Engine_Var_Write t, "SHOP", "ScreenX", Str(.Shop.Screen.x)
        Engine_Var_Write t, "SHOP", "ScreenY", Str(.Shop.Screen.Y)
        Engine_Var_Write t, "MAILBOX", "ScreenX", Str(.Mailbox.Screen.x)
        Engine_Var_Write t, "MAILBOX", "ScreenY", Str(.Mailbox.Screen.Y)
        Engine_Var_Write t, "VIEWMESSAGE", "ScreenX", Str(.ViewMessage.Screen.x)
        Engine_Var_Write t, "VIEWMESSAGE", "ScreenY", Str(.ViewMessage.Screen.Y)
        Engine_Var_Write t, "WRITEMESSAGE", "ScreenX", Str(.WriteMessage.Screen.x)
        Engine_Var_Write t, "WRITEMESSAGE", "ScreenY", Str(.WriteMessage.Screen.Y)
        Engine_Var_Write t, "AMOUNT", "ScreenX", Str(.Amount.Screen.x)
        Engine_Var_Write t, "AMOUNT", "ScreenY", Str(.Amount.Screen.Y)
        Engine_Var_Write t, "MENU", "ScreenX", Str(.Menu.Screen.x)
        Engine_Var_Write t, "MENU", "ScreenY", Str(.Menu.Screen.Y)
    End With

End Sub

Sub Game_SaveMapData(SaveAs As Integer)

'*****************************************************************
'Saves map data to file
'*****************************************************************

Dim ByFlags As Long
Dim FileNum As Byte
Dim i As Integer
Dim Y As Long
Dim x As Long

    'Remove old file if it exists
    If Engine_FileExist(MapPath & SaveAs & ".map", vbNormal) = True Then Kill MapPath & SaveAs & ".map"

    'Write header info on Map.dat
    Call Engine_Var_Write(DataPath & "Map.dat", "INIT", "NumMaps", Str$(NumMaps))

    'Open .map file
    FileNum = FreeFile
    Open MapPath & SaveAs & ".map" For Binary As #FileNum
    Seek #FileNum, 1

    'Map header
    Put #FileNum, , MapInfo.MapVersion

    'Loop through each tile
    For Y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            'Reset flags
            ByFlags = 0

            'Blocked
            If MapData(x, Y).Blocked > 0 Then ByFlags = ByFlags Or 1

            'Graphic layers
            If MapData(x, Y).Graphic(1).GrhIndex Then ByFlags = ByFlags Or 2
            If MapData(x, Y).Graphic(2).GrhIndex Then ByFlags = ByFlags Or 4
            If MapData(x, Y).Graphic(3).GrhIndex Then ByFlags = ByFlags Or 8
            If MapData(x, Y).Graphic(4).GrhIndex Then ByFlags = ByFlags Or 16
            If MapData(x, Y).Graphic(5).GrhIndex Then ByFlags = ByFlags Or 32
            If MapData(x, Y).Graphic(6).GrhIndex Then ByFlags = ByFlags Or 64

            'Light 1-4 used
            For i = 1 To 4
                If MapData(x, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 128
            Next i
            'Light 5-8 used
            For i = 5 To 8
                If MapData(x, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 256
            Next i
            'Light 9-12 used
            For i = 9 To 12
                If MapData(x, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 512
            Next i
            'Light 13-16 used
            For i = 13 To 16
                If MapData(x, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 1024
            Next i
            'Light 17-20 used
            For i = 17 To 20
                If MapData(x, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 2048
            Next i
            'Light 21-24 used
            For i = 21 To 24
                If MapData(x, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 4096
            Next i
            
            'Mailbox
            If MapData(x, Y).Mailbox = 1 Then ByFlags = ByFlags Or 8192
            
            'Shadows
            If MapData(x, Y).Shadow(1) = 1 Then ByFlags = ByFlags Or 16384
            If MapData(x, Y).Shadow(2) = 1 Then ByFlags = ByFlags Or 32768
            If MapData(x, Y).Shadow(3) = 1 Then ByFlags = ByFlags Or 65536
            If MapData(x, Y).Shadow(4) = 1 Then ByFlags = ByFlags Or 131072
            If MapData(x, Y).Shadow(5) = 1 Then ByFlags = ByFlags Or 262144
            If MapData(x, Y).Shadow(6) = 1 Then ByFlags = ByFlags Or 524288

            'Store layers
            Put #FileNum, , ByFlags
            
            'Save blocked value
            If MapData(x, Y).Blocked > 0 Then Put #FileNum, , MapData(x, Y).Blocked
            
            'Save needed grh indexes
            For i = 1 To 6
                If MapData(x, Y).Graphic(i).GrhIndex > 0 Then
                    Put #FileNum, , MapData(x, Y).Graphic(i).GrhIndex
                End If
            Next i

            'Save needed lights
            If ByFlags And 128 Then
                For i = 1 To 4
                    Put #FileNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    Put #FileNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    Put #FileNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    Put #FileNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    Put #FileNum, , MapData(x, Y).Light(i)
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    Put #FileNum, , MapData(x, Y).Light(i)
                Next i
            End If

        Next x
    Next Y

    'Close .map file
    Close #FileNum

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
    
    'Check if we need to run the updater
    If ForceUpdateCheck = True Then
    
        'Check for the right parameter
        If Command$ <> "-sdf@041jkdf0)21`~" Then

            'Force the creation of frmConnect, thus forcing the creation of its hWnd
            Load frmConnect
            frmConnect.Show
            frmConnect.Hide
            
            'Load the updater
            ShellExecute frmConnect.hwnd, vbNullString, App.Path & "\UpdateClient.exe", vbNullString, vbNullString, 1   'The 1 means "show normal"
    
            'Unload the client
            Engine_UnloadAllForms
            End
        
        End If
    End If
    
    'Kill the temp key
    If Engine_FileExist(DataPath & "tempkey.dat", vbNormal) Then Kill DataPath & "tempkey.dat"
    
    'Fill startup variables for the tile engine
    TilePixelWidth = 32
    TilePixelHeight = 32
    WindowTileHeight = 18
    WindowTileWidth = 25
    TileBufferSize = 10
    EnterTextBufferWidth = 1
    EngineBaseSpeed = 0.011
    ReDim SkillListIDs(1 To NumSkills)

    'Setup borders
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.Y = MinYBorder

    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
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

    'Create the buffer
    Set sndBuf = New DataBuffer
    
    'Initialize our encryption
    Encryption_Misc_Init
    
    'Load the config
    Game_Config_Load

    'Set the form starting positions
    DoEvents

    'Load the data commands
    InitDataCommands

    'Display connect window
    frmConnect.Visible = True

    'Main Loop
    prgRun = True
    Do While prgRun

        'Don't draw frame is window is minimized or there is no map loaded
        If frmMain.WindowState <> 1 Then
            If CurMap > 0 Then
            
                'Calculate the starttime - this is the absolute time it takes from start to finish, disincluding DoEvents
                ' The idea is that it works just like the ElapsedTime, but in slightly different placing
                StartTime = timeGetTime

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
                
                'Do sleep event - force FPS at ~60 (62.5) average (prevents extensive processing)
                If (timeGetTime - StartTime) < 16 Then  'If Elapsed Time < Time Required For 60 FPS
                    Sleep 16 - (timeGetTime - StartTime)
                End If

            End If
        End If
        
        'Do other events
       DoEvents

        'Check if unloading
        If IsUnloading = 1 Then
            If frmMain.Sox.ShutDown <> soxERROR Then
                frmMain.Sox.UnHook
                prgRun = False
                Exit Do
            End If
        End If

        'Send the data buffer
        If SocketOpen = 1 Then Data_Send

    Loop

    'Save the config
    Game_Config_Save
    
    'Close down
    frmMain.ShutdownTimer.Enabled = True
    
    'Errors dont matter right now
    On Error Resume Next
    
    'Make sure the socket is closed
    frmMain.Sox.ShutDown
    frmMain.Sox.UnHook

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 11  Code: 602  Total: 613 Lines
':) CommentOnly: 103 (16.8%)  Commented: 4 (0.7%)  Empty: 108 (17.6%)  Max Logic Depth: 7
