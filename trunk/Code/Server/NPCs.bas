Attribute VB_Name = "NPCs"
Option Explicit

Public Sub NPC_UpdateModStats(ByVal NPCIndex As Integer)

Dim Temp As Integer

'Set the HP

    Temp = NPCList(NPCIndex).ModStat(SID.MinHP)

    'Copy over the base stats to the mod stats
    CopyMemory NPCList(NPCIndex).ModStat(1), NPCList(NPCIndex).BaseStat(1), 4 * NumStats

    'Put back the HP
    NPCList(NPCIndex).ModStat(SID.MinHP) = Temp

End Sub

Sub NPC_AI(ByVal NPCIndex As Integer)

'*****************************************************************
'Moves NPC based on it's .movement value
'*****************************************************************

Dim nPos As WorldPos
Dim HeadingLoop As Long
Dim tHeading As Byte
Dim t1 As Byte
Dim t2 As Byte
Dim Y As Long
Dim X As Long

    'Leave if map is in devmode (dont do NPC AI)
    If MapInfo(NPCList(NPCIndex).Pos.Map).DevMode Then Exit Sub

    'Update the action delay counter
    If NPCList(NPCIndex).Flags.ActionDelay > 0 Then
        NPCList(NPCIndex).Flags.ActionDelay = NPCList(NPCIndex).Flags.ActionDelay - Elapsed
        Exit Sub
        
    Else
    
        'Look for someone to attack if hostile
        If NPCList(NPCIndex).Hostile Then
    
            'Check in all directions
            For HeadingLoop = NORTH To NORTHWEST
                nPos = NPCList(NPCIndex).Pos
                Server_HeadToPos HeadingLoop, nPos
    
                'If a legal pos and a user is found attack
                If MapData(nPos.Map, nPos.X, nPos.Y).UserIndex > 0 Then
    
                    'Face NPC to target
                    NPC_ChangeChar ToMap, NPCIndex, NPCIndex, NPCList(NPCIndex).Char.Body, NPCList(NPCIndex).Char.Head, CByte(HeadingLoop), NPCList(NPCIndex).Char.Weapon, NPCList(NPCIndex).Char.Hair
    
                    'Tell everyone in the PC area to show the attack animation
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.User_Attack
                    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
    
                    'Attack
                    NPC_AttackUser NPCIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
    
                    'Don't move if fighting
                    Exit Sub
    
                End If
                
            Next HeadingLoop
        End If
        
    End If

    'Movement
    Select Case NPCList(NPCIndex).Movement

    Case 2  '*** Random movement ***
        NPC_MoveChar NPCIndex, Int(Rnd * 8) + 1

    Case 3  '*** Go towards nearby players - simple/fast AI ***

        'Look for a user
        X = NPC_AI_ClosestPC(NPCIndex, 10, 8)
        If X > 0 Then

            'Find the direction to move
            tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, UserList(X).Pos)
            
            'Move towards the retrieved position
            If NPC_MoveChar(NPCIndex, tHeading) = 0 Then
            
                'Move towards alternate positions (the two directions that surround the selected direction)
                Select Case tHeading
                    Case 1
                        t1 = 5
                        t2 = 8
                    Case 2
                        t1 = 5
                        t2 = 6
                    Case 3
                        t1 = 7
                        t2 = 6
                    Case 4
                        t1 = 7
                        t2 = 8
                    Case 5
                        t1 = 1
                        t2 = 2
                    Case 6
                        t1 = 2
                        t2 = 3
                    Case 7
                        t1 = 3
                        t2 = 4
                    Case 8
                        t1 = 4
                        t2 = 1
                End Select
                
                'Do the alternate movement
                If NPC_MoveChar(NPCIndex, t1) = 0 Then
                    NPC_MoveChar NPCIndex, t2   'If this doesn't happen, then we're out of stuff to do
                End If
            
            End If
                
            Exit Sub

        End If

    End Select

End Sub

Public Function NPC_AI_ClosestPC(ByVal NPCIndex As Integer, ByVal SearchX As Byte, ByVal SearchY As Byte) As Integer

'*****************************************************************
'Return the index of the closest player character (PC)
'*****************************************************************
Dim tX As Integer
Dim tY As Integer
Dim X As Integer
Dim Y As Integer
    
    'Expand the search range
    For tX = 1 To SearchX
        For tY = 1 To SearchY
            'Loop through the search area (only look on the outside of the search rectangle to prevent checking the same thing multiple times)
            For X = NPCList(NPCIndex).Pos.X - tX To NPCList(NPCIndex).Pos.X + tX Step tX
                For Y = NPCList(NPCIndex).Pos.Y - tY To NPCList(NPCIndex).Pos.Y + tY Step tY
                    'Make sure tile is legal
                    If X > MinXBorder Then
                        If X < MaxXBorder Then
                            If Y > MinYBorder Then
                                If Y < MaxYBorder Then
                                    'Look for a user
                                    If MapData(NPCList(NPCIndex).Pos.Map, X, Y).UserIndex > 0 Then
                                        NPC_AI_ClosestPC = MapData(NPCList(NPCIndex).Pos.Map, X, Y).UserIndex
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next Y
            Next X
        Next tY
    Next tX

End Function

Sub NPC_AttackUser(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)

'*****************************************************************
'Have a NPC attack a User
'*****************************************************************

Dim Hit As Integer

    'Check for an action delay
    If NPCList(NPCIndex).Flags.ActionDelay > 0 Then Exit Sub

    'Don't allow if switchingmaps maps
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub

    'Set the action delay
    NPCList(NPCIndex).Flags.ActionDelay = NPCDelayFight

    'Check if the user has a 100% chance to miss
    If NPCList(NPCIndex).ModStat(SID.WeaponSkill) + 50 < UserList(UserIndex).Stats.ModStat(SID.Parry) Then Exit Sub

    'If user weapon skill is at least 50 points greater, 100% chance to hit
    If NPCList(NPCIndex).ModStat(SID.WeaponSkill) - 50 <= UserList(UserIndex).Stats.ModStat(SID.Parry) Then

        'Since the user doesn't have 100% chance to hit, calculate if they hit
        If Server_RandomNumber(1, 100) >= ((NPCList(NPCIndex).ModStat(SID.WeaponSkill) + 50) - UserList(UserIndex).Stats.ModStat(SID.Parry)) Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_SetCharDamage
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            ConBuf.Put_Integer -1
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
            Exit Sub
        End If

    End If

    'Calculate hit
    Hit = Server_RandomNumber(NPCList(NPCIndex).ModStat(SID.MinHIT), NPCList(NPCIndex).ModStat(SID.MaxHIT))
    Hit = Hit - (UserList(UserIndex).Stats.ModStat(SID.DEF) \ 2)
    If Hit < 1 Then Hit = 1

    'Hit user
    UserList(UserIndex).Stats.ModStat(SID.MinHP) = UserList(UserIndex).Stats.ModStat(SID.MinHP) - Hit

    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_SetCharDamage
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Integer Hit
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

    'User Die
    If UserList(UserIndex).Stats.ModStat(SID.MinHP) <= 0 Then
        'Kill user
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "The " & NPCList(NPCIndex).Name & " kills you!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Fight
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        User_Kill UserIndex
    End If

End Sub

Sub NPC_ChangeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, NPCIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, Weapon As Integer, Hair As Integer)

'*****************************************************************
'Changes a NPC char's head,body and heading
'*****************************************************************

    NPCList(NPCIndex).Char.Body = Body
    NPCList(NPCIndex).Char.Head = Head
    NPCList(NPCIndex).Char.Heading = Heading
    NPCList(NPCIndex).Char.HeadHeading = Heading

    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_ChangeChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Integer Body
    ConBuf.Put_Integer Head
    ConBuf.Put_Byte Heading
    ConBuf.Put_Integer Weapon
    ConBuf.Put_Integer Hair
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

End Sub

Sub NPC_Close(ByVal NPCIndex As Integer)

'*****************************************************************
'Closes a NPC
'*****************************************************************

    NPCList(NPCIndex).Flags.NPCActive = 0

    'Update LastNPC
    If NPCIndex = LastNPC Then
        Do Until NPCList(LastNPC).Flags.NPCActive = 1
            LastNPC = LastNPC - 1
            If LastNPC = 0 Then Exit Do
        Loop
        If NPCIndex <> LastNPC Then
            If NPCIndex <> 0 Then
                ReDim Preserve NPCList(1 To LastNPC)
            Else
                ReDim Preserve NPCList(1)
            End If
        End If
    End If

    'Update number of NPCs
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If

End Sub

Public Sub NPC_Damage(NPCIndex As Integer, UserIndex As Integer, Damage As Integer)

'*****************************************************************
'Do damage to a NPC - ONLY USE THIS SUB TO HURT NPCS
'*****************************************************************

Dim HPA As Byte         'HP percentage before reducing hp
Dim HPB As Byte         'HP percentage after reducing hp
Dim i As Integer

'Check if the NPC can be attacked

    If NPCList(NPCIndex).Attackable = 0 Then Exit Sub

    'Get the pre-damage percentage
    HPA = CByte((NPCList(NPCIndex).ModStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)

    'Lower the NPC's life
    NPCList(NPCIndex).ModStat(SID.MinHP) = NPCList(NPCIndex).ModStat(SID.MinHP) - Damage

    'Check to update health percentage client-side
    If NPCList(NPCIndex).ModStat(SID.MinHP) > 0 Then
        HPB = CByte((NPCList(NPCIndex).ModStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
        If HPA <> HPB Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_CharHP
            ConBuf.Put_Byte HPB
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
        End If
    End If

    'Turn the NPC aggressive-faced
    If NPCList(NPCIndex).Counters.AggressiveCounter <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_AggressiveFace
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Byte 1
        Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
    End If
    NPCList(NPCIndex).Counters.AggressiveCounter = AGGRESSIVEFACETIME

    'Display the damage on the client screen
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_SetCharDamage
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Integer Damage
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer

    'Check if the NPC died
    If NPCList(NPCIndex).ModStat(SID.MinHP) <= 0 Then

        If UserIndex > 0 Then

            'Check on quests
            For i = 1 To MaxQuests
                If UserList(UserIndex).Quest(i) > 0 Then
                    If QuestData(UserList(UserIndex).Quest(i)).FinishReqNPC = NPCList(NPCIndex).NPCNumber Then

                        'User must kill at least one more of the NPC
                        If UserList(UserIndex).QuestStatus(i).NPCKills <= QuestData(UserList(UserIndex).Quest(i)).FinishReqNPCAmount Then
                            UserList(UserIndex).QuestStatus(i).NPCKills = UserList(UserIndex).QuestStatus(i).NPCKills + 1
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Comm_Talk
                            ConBuf.Put_String "You have killed " & UserList(UserIndex).QuestStatus(i).NPCKills & " of " & QuestData(UserList(UserIndex).Quest(i)).FinishReqNPCAmount & " " & NPCList(QuestData(UserList(UserIndex).Quest(i)).FinishReqNPC).Name & "s!"
                            ConBuf.Put_Byte DataCode.Comm_FontType_Quest
                        End If

                    End If
                End If
            Next i

            'Give EXP and gold
            User_RaiseExp UserIndex, NPCList(NPCIndex).GiveEXP
            UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + NPCList(NPCIndex).GiveGLD

            'Display kill message to the user
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You kill " & NPCList(NPCIndex).Name & "!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Fight
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        End If

        'Kill off the NPC
        NPC_Kill NPCIndex

    End If

End Sub

Sub NPC_EraseChar(ByVal NPCIndex As Integer)

'*****************************************************************
'Erase a character
'*****************************************************************

    'Remove from list
    CharList(NPCList(NPCIndex).Char.CharIndex).Index = 0
    CharList(NPCList(NPCIndex).Char.CharIndex).CharType = 0
    
    'Remove pathfinding values
    NPCList(NPCIndex).Flags.HasPath = 0

    'Remove from map
    MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0

    'Send erase command to clients
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_EraseChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    Data_Send ToMap, 0, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

    'Clear the variables
    NPCList(NPCIndex).Char.CharIndex = 0
    NPCList(NPCIndex).Flags.NPCAlive = 0

    'Set at the respawn spot
    NPCList(NPCIndex).Pos.Map = NPCList(NPCIndex).StartPos.Map
    NPCList(NPCIndex).Pos.X = NPCList(NPCIndex).StartPos.X
    NPCList(NPCIndex).Pos.Y = NPCList(NPCIndex).StartPos.Y

End Sub

Sub NPC_Kill(ByVal NPCIndex As Integer)

'*****************************************************************
'Kill a NPC
'*****************************************************************
    
    'Set health back to 100%
    NPCList(NPCIndex).ModStat(SID.MinHP) = NPCList(NPCIndex).ModStat(SID.MaxHP)

    'Erase it from map
    NPC_EraseChar NPCIndex

    'Set death time for respawn wait
    NPCList(NPCIndex).Counters.RespawnCounter = timeGetTime

End Sub

Sub NPC_MakeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal NPCIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Makes and places a NPC character
'*****************************************************************

Dim EmptySkills As Declares.Skills  'VB thinks we are trying to declare the module if we dont put the Declares in there (sigh)
Dim SndHP As Byte
Dim SndMP As Byte

'Place character on map

    MapData(Map, X, Y).NPCIndex = NPCIndex

    'Set alive flag
    NPCList(NPCIndex).Flags.NPCAlive = 1

    'Set the hp/mp to send
    If NPCList(NPCIndex).ModStat(SID.MaxHP) > 0 Then SndHP = CByte((NPCList(NPCIndex).ModStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
    If NPCList(NPCIndex).ModStat(SID.MaxMAN) > 0 Then SndMP = CByte((NPCList(NPCIndex).ModStat(SID.MinMAN) / NPCList(NPCIndex).ModStat(SID.MaxMAN)) * 100)

    'Send make character command to clients
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_MakeChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Body
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Head
    ConBuf.Put_Byte NPCList(NPCIndex).Char.Heading
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Byte X
    ConBuf.Put_Byte Y
    ConBuf.Put_String NPCList(NPCIndex).Name
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Weapon
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Hair
    ConBuf.Put_Byte SndHP
    ConBuf.Put_Byte SndMP

    'NPCs wont be created with active spells
    NPCList(NPCIndex).Skills = EmptySkills

    'Send the NPC
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, Map

End Sub

Function NPC_MoveChar(ByVal NPCIndex As Integer, ByVal nHeading As Byte) As Byte

'*****************************************************************
'Moves a NPC from one tile to another
'*****************************************************************

Dim nPos As WorldPos

'Move

    nPos = NPCList(NPCIndex).Pos
    Server_HeadToPos nHeading, nPos

    'Move if legal pos
    If Server_LegalPos(NPCList(NPCIndex).Pos.Map, nPos.X, nPos.Y, nHeading) = True Then

        'Set the move delay
        NPCList(NPCIndex).Flags.ActionDelay = NPCDelayWalk

        'Send the movement update packet
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MoveChar
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Byte nPos.X
        ConBuf.Put_Byte nPos.Y
        Data_Send ToMap, 0, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

        'Update map and user pos
        MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0
        NPCList(NPCIndex).Pos = nPos
        NPCList(NPCIndex).Char.Heading = nHeading
        NPCList(NPCIndex).Char.HeadHeading = nHeading
        MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = NPCIndex

        'NPC moved, return the flag
        NPC_MoveChar = 1

    End If

End Function

Function NPC_NextOpen() As Integer

'*****************************************************************
'Finds the next open NPC Index in NPCList
'*****************************************************************

Dim LoopC As Long

    Do
        LoopC = LoopC + 1
        If LoopC > LastNPC Then
            LoopC = LastNPC + 1
            Exit Do
        End If
    Loop While NPCList(LoopC).Flags.NPCActive = 1

    NPC_NextOpen = LoopC

End Function

Sub NPC_Spawn(ByVal NPCIndex As Integer)

'*****************************************************************
'Places a NPC that has been Opened
'*****************************************************************

Dim TempPos As WorldPos
Dim CharIndex As Integer

'Give it a char index if needed

    If NPCList(NPCIndex).Char.CharIndex = 0 Then
        CharIndex = Server_NextOpenCharIndex
        NPCList(NPCIndex).Char.CharIndex = CharIndex
        CharList(CharIndex).Index = NPCIndex
        CharList(CharIndex).CharType = CharType_NPC
    End If

    'Find a place to put npc
    Server_ClosestLegalPos NPCList(NPCIndex).StartPos, TempPos
    If Not Server_LegalPos(TempPos.Map, TempPos.X, TempPos.Y, 0) Then Exit Sub

    'Set vars
    NPCList(NPCIndex).Pos = TempPos

    'Make NPC Char
    If UBound(ConnectionGroups(TempPos.Map).UserIndex) > 0 Then NPC_MakeChar ToMap, ConnectionGroups(TempPos.Map).UserIndex(1), NPCIndex, TempPos.Map, TempPos.X, TempPos.Y

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:46)  Decl: 13  Code: 684  Total: 697 Lines
':) CommentOnly: 126 (18.1%)  Commented: 25 (3.6%)  Empty: 166 (23.8%)  Max Logic Depth: 13
