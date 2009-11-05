Attribute VB_Name = "General"
Option Explicit

Function Server_LegalString(ByVal CheckString As String) As Boolean

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
        If Server_LegalCharacter(AscB(Mid$(CheckString, i, 1))) = False Then Exit Function
        
    Next i
    
    'If we have made it this far, then all is good
    Server_LegalString = True

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Server_LegalString = False

End Function

Function Server_LegalCharacter(KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters
'*****************************************************************

    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 Or KeyAscii <= 57 Then
        Server_LegalCharacter = True
        Exit Function
    End If
    
    'Allow letters A to Z
    If KeyAscii >= 65 Or KeyAscii <= 90 Then
        Server_LegalCharacter = True
        Exit Function
    End If
    
    'Allow letters a to z
    If KeyAscii >= 97 Or KeyAscii <= 122 Then
        Server_LegalCharacter = True
        Exit Function
    End If
    
Exit Function

ErrOut:

    'Something bad happened, so the character must be invalid
    Server_LegalCharacter = False
    
End Function

Function Server_CalcEXPCost(BaseSkill As Long) As Long

'*****************************************************************
'Calculate the exp required to raise a skill up to the next point
'*****************************************************************
On Error Resume Next

    Server_CalcEXPCost = Int(0.17376 * (BaseSkill ^ 3) + 0.44 * (BaseSkill ^ 2) - 0.48 * BaseSkill + 1.035) + 1

End Function

Function Server_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single

'*****************************************************************
'Finds the distance between two points
'*****************************************************************

    Server_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))

End Function

Function Server_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
On Error GoTo ErrOut

    If Dir$(File, FileType) <> "" Then Server_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Server_FileExist = False

End Function

Public Sub Server_InitDataCommands()

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
        .Server_Connect = 26
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

Function Server_RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Integer

'*****************************************************************
'Find a Random number between a range
'*****************************************************************

    Server_RandomNumber = Fix((UpperBound - LowerBound + 1) * Rnd) + LowerBound

End Function

Sub Server_RefreshUserListBox()

'*****************************************************************
'Refreshes the User list box
'*****************************************************************

Dim LoopC As Long

    If LastUser < 0 Then
        frmMain.Userslst.Clear
        Exit Sub
    End If

    frmMain.Userslst.Clear
    CurrConnections = 0
    For LoopC = 1 To LastUser
        If UserList(LoopC).Name <> "" Then
            frmMain.Userslst.AddItem UserList(LoopC).Name
            CurrConnections = CurrConnections + 1
        End If
    Next LoopC
    TrayModify ToolTip, "Game Server: " & CurrConnections & " connections"

End Sub

Public Sub Server_UpdateMapTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte)

'*****************************************************************
'Takes the map tile info and compiles it into the conversion buffer
'Does not make the send so you can send multiple tiles to Data_Send at once
'*****************************************************************

Dim ChunkData As Long
Dim LoopC As Byte

    With MapData(Map, x, Y)

        'Set the initial values
        ConBuf.Put_Byte DataCode.Map_UpdateTile
        ConBuf.Put_Byte x
        ConBuf.Put_Byte Y

        'Build the chunk data
        ChunkData = 0
        If .Blocked Or BlockedNorth Then ChunkData = ChunkData Or 1
        If .Blocked Or BlockedEast Then ChunkData = ChunkData Or 2
        If .Blocked Or BlockedSouth Then ChunkData = ChunkData Or 4
        If .Blocked Or BlockedWest Then ChunkData = ChunkData Or 8
        If .Mailbox Then ChunkData = ChunkData Or 16
        If .Graphic(1) > 0 Then ChunkData = ChunkData Or 32
        If .Graphic(2) > 0 Then ChunkData = ChunkData Or 64
        If .Graphic(3) > 0 Then ChunkData = ChunkData Or 128
        If .Graphic(4) > 0 Then ChunkData = ChunkData Or 256
        If .Graphic(5) > 0 Then ChunkData = ChunkData Or 512
        If .Graphic(6) > 0 Then ChunkData = ChunkData Or 1024
        If .Light(1) <> -1 Then ChunkData = ChunkData Or 2048
        If .Light(2) <> -1 Then ChunkData = ChunkData Or 2048
        If .Light(3) <> -1 Then ChunkData = ChunkData Or 2048
        If .Light(4) <> -1 Then ChunkData = ChunkData Or 2048
        If .Light(5) <> -1 Then ChunkData = ChunkData Or 4096
        If .Light(6) <> -1 Then ChunkData = ChunkData Or 4096
        If .Light(7) <> -1 Then ChunkData = ChunkData Or 4096
        If .Light(8) <> -1 Then ChunkData = ChunkData Or 4096
        If .Light(9) <> -1 Then ChunkData = ChunkData Or 8192
        If .Light(10) <> -1 Then ChunkData = ChunkData Or 8192
        If .Light(11) <> -1 Then ChunkData = ChunkData Or 8192
        If .Light(12) <> -1 Then ChunkData = ChunkData Or 8192
        If .Light(13) <> -1 Then ChunkData = ChunkData Or 16384
        If .Light(14) <> -1 Then ChunkData = ChunkData Or 16384
        If .Light(15) <> -1 Then ChunkData = ChunkData Or 16384
        If .Light(16) <> -1 Then ChunkData = ChunkData Or 16384
        If .Light(17) <> -1 Then ChunkData = ChunkData Or 32768
        If .Light(18) <> -1 Then ChunkData = ChunkData Or 32768
        If .Light(19) <> -1 Then ChunkData = ChunkData Or 32768
        If .Light(20) <> -1 Then ChunkData = ChunkData Or 32768
        If .Light(21) <> -1 Then ChunkData = ChunkData Or 65536
        If .Light(22) <> -1 Then ChunkData = ChunkData Or 65536
        If .Light(23) <> -1 Then ChunkData = ChunkData Or 65536
        If .Light(24) <> -1 Then ChunkData = ChunkData Or 65536
        If .Shadow(1) > 0 Then ChunkData = ChunkData Or 131072
        If .Shadow(2) > 0 Then ChunkData = ChunkData Or 262144
        If .Shadow(3) > 0 Then ChunkData = ChunkData Or 524288
        If .Shadow(4) > 0 Then ChunkData = ChunkData Or 1048576
        If .Shadow(5) > 0 Then ChunkData = ChunkData Or 2097152
        If .Shadow(6) > 0 Then ChunkData = ChunkData Or 4194304
        If .Sfx > 0 Then ChunkData = ChunkData Or 8388608

        'Send the chunk data
        ConBuf.Put_Long ChunkData

        'Send the graphics
        For LoopC = 1 To 6
            If .Graphic(LoopC) > 0 Then ConBuf.Put_Integer .Graphic(LoopC)
        Next LoopC

        'Send the lights
        If ChunkData And 2048 Then
            For LoopC = 1 To 4
                ConBuf.Put_Long .Light(LoopC)
            Next LoopC
        End If
        If ChunkData And 4096 Then
            For LoopC = 5 To 8
                ConBuf.Put_Long .Light(LoopC)
            Next LoopC
        End If
        If ChunkData And 8192 Then
            For LoopC = 9 To 12
                ConBuf.Put_Long .Light(LoopC)
            Next LoopC
        End If
        If ChunkData And 16384 Then
            For LoopC = 13 To 16
                ConBuf.Put_Long .Light(LoopC)
            Next LoopC
        End If
        If ChunkData And 32768 Then
            For LoopC = 17 To 20
                ConBuf.Put_Long .Light(LoopC)
            Next LoopC
        End If
        If ChunkData And 65536 Then
            For LoopC = 21 To 24
                ConBuf.Put_Long .Light(LoopC)
            Next LoopC
        End If
        
        'Send the Sfx
        If ChunkData And 8388608 Then
            ConBuf.Put_Integer .Sfx
        End If

    End With

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:47)  Decl: 1  Code: 368  Total: 369 Lines
':) CommentOnly: 42 (11.4%)  Commented: 0 (0%)  Empty: 46 (12.5%)  Max Logic Depth: 4
