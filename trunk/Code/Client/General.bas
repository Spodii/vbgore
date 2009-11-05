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
Dim CharAscii As Integer
'Password

    If LenB(UserPassword) = 0 Then
        MsgBox ("Password box is empty.")
        Exit Function
    End If
    If Len(UserPassword) > 10 Then
        MsgBox ("Password must be 10 characters or less.")
        Exit Function
    End If
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(Mid$(UserPassword, LoopC, 1))
        If Game_LegalCharacter(CharAscii) = False Then
            MsgBox ("Invalid Password.")
            Exit Function
        End If
    Next LoopC
    'Name
    If LenB(UserName) = 0 Then
        MsgBox ("Name box is empty.")
        Exit Function
    End If
    If Len(UserName) > 30 Then
        MsgBox ("Name must be 30 characters or less.")
        Exit Function
    End If
    For LoopC = 1 To Len(UserName)
        CharAscii = Asc(Mid$(UserName, LoopC, 1))
        If Game_LegalCharacter(CharAscii) = False Then
            MsgBox ("Invalid Name.")
            Exit Function
        End If
    Next LoopC
    'If all good send true
    Game_CheckUserData = True

End Function

Public Sub Game_ClearMapTileChanged()

Dim x As Byte
Dim Y As Byte

    For x = 1 To 100
        For Y = 1 To 100
            MapTileChanged(x, Y) = 0
        Next Y
    Next x

End Sub

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

Private Sub Game_InitDataCommands()

'Load the values for the data commands

    With EmoID
        .Dots = 1
        .Exclimation = 2
        .Question = 3
        .Surprised = 4
        .Heart = 5
        .Hearts = 6
        .HeartBroken = 7
        .Utensils = 8
        .Meat = 9
        .ExcliQuestion = 10
    End With

    With SkID
        .Bless = 1
        .Curse = 2
        .Heal = 3
        .IronSkin = 4
        .Protection = 5
        .Strengthen = 6
        .Warcry = 7
        .SpikeField = 8
    End With

    With SID
        .Agil = 1
        .Clairovoyance = 2
        .Dagger = 3
        .DEF = 4
        .DefensiveMag = 5
        .ELU = 6
        .ELV = 7
        .EXP = 8
        .Fist = 9
        .Gold = 10
        .Immunity = 11
        .Mag = 12
        .MaxHIT = 13
        .MaxHP = 14
        .MaxMAN = 15
        .MaxSTA = 16
        .Meditate = 17
        .MinHIT = 18
        .MinHP = 19
        .MinMAN = 20
        .MinSTA = 21
        .OffensiveMag = 22
        .Parry = 23
        .Points = 24
        .Regen = 25
        .Rest = 26
        .Staff = 27
        .Str = 28
        .SummoningMag = 29
        .Sword = 30
        .WeaponSkill = 31
    End With

    With DataCode
        .Comm_UMsgbox = 2
        .Server_IconSpellExhaustion = 3
        .Comm_Shout = 4
        .Server_UserCharIndex = 5
        .Comm_Emote = 6
        .Server_SetUserPosition = 7
        .Map_LoadMap = 8
        .Map_DoneLoadingMap = 9
        .Map_RequestUpdate = 10
        .Map_StartTransfer = 11
        .Server_CharHP = 12
        .Map_EndTransfer = 13
        .Map_DoneSwitching = 14
        .Map_SendName = 15
        .User_Attack = 16
        .Server_MakeChar = 17
        .Server_EraseChar = 18
        .Server_MoveChar = 19
        .Server_ChangeChar = 20
        .Server_MakeObject = 21
        .Server_EraseObject = 22
        .User_KnownSkills = 23
        .User_SetInventorySlot = 24
        .User_StartQuest = 25
        '26
        .Server_PlaySound = 27
        .User_Login = 28
        .User_NewLogin = 29
        .Comm_Whisper = 30
        .Server_Who = 31
        .User_Move = 32
        .User_Rotate = 33
        .User_LeftClick = 34
        .User_RightClick = 35
        .Map_RequestPositionUpdate = 36
        .User_Get = 37
        .User_Drop = 38
        .User_Use = 39
        '40
        .Comm_Talk = 41
        .Server_SetCharDamage = 42
        .User_ChangeInvSlot = 43
        .User_Emote = 44
        .Server_CharMP = 45
        .Server_Disconnect = 46
        'All numbers between the above and below are free...
        .User_BaseStat = 90
        .User_ModStat = 91
        .Comm_FontType_Fight = 92
        .Comm_FontType_Info = 93
        .Comm_FontType_Quest = 94
        .Comm_FontType_Talk = 95
        '. = 96
        '. = 97
        '. = 98
        '. = 99
        '. = 100
        '. = 101
        .User_CastSkill = 102
        .Server_IconCursed = 103
        .Server_IconWarCursed = 104
        .Server_IconBlessed = 105
        .Server_IconStrengthened = 106
        .Server_IconProtected = 107
        .Server_IconIronSkin = 108
        .Server_MailBox = 109
        .Server_MailMessage = 110
        .Server_MailItemInfo = 111
        .Server_MailItemTake = 112
        .Server_MailItemRemove = 113
        .Server_MailDelete = 114
        .Server_MailCompose = 115
        '. = 116
        '. = 117
        '. = 118
        .User_LookLeft = 119
        .User_LookRight = 120
        .User_Blink = 121
        .User_AggressiveFace = 122
        .User_Trade_BuyFromNPC = 123
        .User_Trade_SellToNPC = 124
        .User_Trade_StartNPCTrade = 125
        .Dev_SetBlocked = 126
        .Dev_SetExit = 127
        .Dev_SetLight = 128
        .Dev_SetMailbox = 129
        .Dev_SetMapInfo = 130
        .Dev_SetNPC = 131
        .Dev_SetObject = 132
        .User_Target = 133
        .Dev_SetSurface = 134
        ' = 135
        .Map_UpdateTile = 136
        .Dev_UpdateTile = 137
        .Dev_SaveMap = 138
        .Server_Ping = 139
        '140
        .User_Desc = 141
        .Server_Help = 142
        .GM_Approach = 143
        .GM_Summon = 144
        .GM_Kick = 145
        .GM_Raise = 146
        .Dev_SetMode = 147
        .Dev_SetTile = 148
    End With

End Sub

Function Game_LegalCharacter(KeyAscii As Integer) As Boolean

'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
'if backspace allow

    If KeyAscii = 8 Then
        Game_LegalCharacter = True
        Exit Function
    End If
    'Only allow space,numbers,letters and special characters
    If KeyAscii < 32 Then
        Game_LegalCharacter = False
        Exit Function
    End If
    If KeyAscii > 126 Then
        Game_LegalCharacter = False
        Exit Function
    End If
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Game_LegalCharacter = False
        Exit Function
    End If
    'else everything is cool
    Game_LegalCharacter = True

End Function

Public Sub Game_Config_Load()

'***************************************************
'Load the user configuration
'***************************************************

Dim i As Byte

    'Quickbar
    For i = 1 To 12
        QuickBarID(i).ID = Val(Engine_Var_Get(IniPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "ID"))
        QuickBarID(i).Type = Val(Engine_Var_Get(IniPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "Type"))
    Next i
    
    'Skin
    CurrentSkin = Engine_Var_Get(IniPath & "Game.ini", "INIT", "CurrentSkin")

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
        Engine_Var_Write IniPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "ID", Str$(QuickBarID(i).ID)
        Engine_Var_Write IniPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "Type", Str$(QuickBarID(i).Type)
    Next i
    
    'Skin
    Engine_Var_Write IniPath & "Game.ini", "INIT", "CurrentSkin", CurrentSkin
    
    'Skin positions
    t = IniPath & "Skins\" & CurrentSkin & ".dat"   'Set the custom positions file for the skin
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
        Engine_Var_Write t, "DEV", "ScreenX", Str(.Dev.Screen.x)
        Engine_Var_Write t, "DEV", "ScreenY", Str(.Dev.Screen.Y)
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
    Call Engine_Var_Write(IniPath & "Map.dat", "INIT", "NumMaps", Str$(NumMaps))

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

Dim i As Integer

    'Create the buffer
    Set sndBuf = New DataBuffer

    'Init file paths
    MusicPath = App.Path & "\Music\"
    IniPath = App.Path & "\Data\"
    MapPath = App.Path & "\Maps\"
    SfxPath = App.Path & "\Sfx\"
    GrhPath = App.Path & "\Grh\"
    
    'Initialize our encryption
    Encryption_Misc_Init
    
    'Load the config
    Game_Config_Load

    'Set the form starting positions
    DoEvents

    'Load the data commands
    Game_InitDataCommands

    'Display connect window
    frmConnect.Visible = True

    'Main Loop
    prgRun = True
    Do While prgRun

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

        'Do other events
        DoEvents

        'Check if unloading
        If IsUnloading Then
            If frmMain.Sox.ShutDown <> soxERROR Then
                frmMain.Sox.UnHook
                prgRun = False
                Exit Do
            End If
        End If

        'Send the data buffer
        Data_Send

    Loop

    'Close Down
    Game_Config_Save
    Engine_Init_UnloadTileEngine
    Engine_UnloadAllForms
    End

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 11  Code: 602  Total: 613 Lines
':) CommentOnly: 103 (16.8%)  Commented: 4 (0.7%)  Empty: 108 (17.6%)  Max Logic Depth: 7
