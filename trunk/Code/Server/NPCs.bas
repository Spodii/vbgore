Attribute VB_Name = "NPCs"
Option Explicit

Public Sub NPC_UpdateModStats(ByRef NPCIndex As Integer)

    Log "Call NPC_UpdateModStats(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    With NPCList(NPCIndex)
    
        'Copy the base stats to the mod stats (we can use copymemory since we dont have to give item bonuses)
        CopyMemory .ModStat(FirstModStat), .BaseStat(FirstModStat), ((NumStats - FirstModStat) + 1) * 4 '* 4 since we are using longs (4 bytes)
         
        'War curse
        If .Skills.WarCurse > 0 Then
            Log "NPC_UpdateModStats: Updating modstats with effects from WarCurse", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.Agi) = .ModStat(SID.Agi) - (.Skills.WarCurse \ 4)  'Remember, AGI for NPCs is the hit rate!
            .ModStat(SID.DEF) = .ModStat(SID.DEF) - (.Skills.WarCurse \ 4)
            .ModStat(SID.Mag) = .ModStat(SID.Mag) - (.Skills.WarCurse \ 4)
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - (.Skills.WarCurse \ 4)
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - (.Skills.WarCurse \ 4)
        End If
            
        'Strengthen
        If .Skills.Strengthen > 0 Then
            Log "NPC_UpdateModStats: Updating modstats with effects from Strengthen", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + .Skills.Strengthen
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + .Skills.Strengthen
        End If
        
        'Protection
        If .Skills.Protect > 0 Then
            Log "NPC_UpdateModStats: Updating modstats with effects from Protect", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + .Skills.Protect
        End If
        
        'Bless
        If .Skills.Bless > 0 Then
            Log "NPC_UpdateModStats: Updating modstats with effects from Bless", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.Agi) = .ModStat(SID.Agi) + .Skills.Bless \ 2 'Remember, AGI for NPCs is the hit rate!
            .ModStat(SID.Mag) = .ModStat(SID.Mag) + .Skills.Bless \ 2
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + .Skills.Bless \ 4
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + .Skills.Bless \ 4
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + .Skills.Bless \ 4
        End If
        
        'Iron skin
        If .Skills.IronSkin > 0 Then
            Log "NPC_UpdateModStats: Updating modstats with effects from IronSkin", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + .Skills.IronSkin * 2
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - .Skills.IronSkin * 1.5
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - .Skills.IronSkin * 1.5
        End If

    End With

End Sub

Private Function NPC_AI_Attack(ByVal NPCIndex As Integer) As Byte

'*****************************************************************
'Calls the NPC attack AI - only call by the NPC_AI routine!
'*****************************************************************
Dim HeadingLoop As Long
Dim nPos As WorldPos
Dim Damage As Long
Dim X As Long

    Log "Call NPC_AI_Attack(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    '*** Melee attacking ***
    If NPCList(NPCIndex).AttackRange <= 1 Then
        
        'Check in all directions
        For HeadingLoop = NORTH To NORTHWEST
            nPos = NPCList(NPCIndex).Pos
            Server_HeadToPos HeadingLoop, nPos

            'If a legal pos and a user is found attack
            If MapInfo(nPos.Map).Data(nPos.X, nPos.Y).UserIndex > 0 Then
                X = MapInfo(nPos.Map).Data(nPos.X, nPos.Y).UserIndex
            
                'Get the damage
                Damage = NPC_AttackUser(NPCIndex, X)

                'Send the attack packet
                ConBuf.PreAllocate 12
                ConBuf.Put_Byte DataCode.Combo_SlashSoundRotateDamage
                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                ConBuf.Put_Integer UserList(X).Char.CharIndex
                ConBuf.Put_Long NPCList(NPCIndex).AttackGrh
                ConBuf.Put_Byte NPCList(NPCIndex).AttackSfx
                If Damage > 32000 Then ConBuf.Put_Integer 32000 Else ConBuf.Put_Integer Damage
                Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

                'Apply damage
                NPC_AttackUser_ApplyDamage NPCIndex, X, Damage
                NPC_AI_Attack = 1
                Exit Function

            End If
            
        Next HeadingLoop
        
    '*** Ranged attacking ***
    Else
        
        'Check for the closest user
        X = NPC_AI_ClosestPC(NPCIndex, NPCList(NPCIndex).AttackRange \ 2, NPCList(NPCIndex).AttackRange \ 2)
        If X > 0 Then
            
            'Check for a valid range
            If Server_Distance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(X).Pos.X, UserList(X).Pos.Y) <= NPCList(NPCIndex).AttackRange Then

                'Check for a valid path
                If Engine_ClearPath(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(X).Pos.X, UserList(X).Pos.Y) Then
                    
                    'Get the damage
                    Damage = NPC_AttackUser(NPCIndex, X)

                    'Face the NPC to the target
                    ConBuf.PreAllocate 13
                    ConBuf.Put_Byte DataCode.Combo_ProjectileSoundRotateDamage
                    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                    ConBuf.Put_Integer UserList(X).Char.CharIndex
                    ConBuf.Put_Long NPCList(NPCIndex).AttackGrh
                    ConBuf.Put_Byte NPCList(NPCIndex).ProjectileRotateSpeed
                    ConBuf.Put_Byte NPCList(NPCIndex).AttackSfx
                    If Damage > 32000 Then ConBuf.Put_Integer 32000 Else ConBuf.Put_Integer Damage
                    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

                    'Apply damage
                    NPC_AttackUser_ApplyDamage NPCIndex, X, Damage
                    NPC_AI_Attack = 1
                    Exit Function
                
                End If
            End If
        End If
    End If

End Function

Private Function NPC_AI_AttackNPC(ByVal NPCIndex As Integer, Optional ByVal NotSlaveOfUserIndex As Integer = 0) As Byte

'*****************************************************************
'Calls the NPC attack AI to attack another NPC - only call by the NPC_AI routine!
'*****************************************************************
Dim Damage As Long
Dim HeadingLoop As Long
Dim nPos As WorldPos
Dim X As Long

    Log "Call NPC_AI_AttackNPC(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    '*** Melee attacking ***
    If NPCList(NPCIndex).AttackRange <= 1 Then
        
        'Check in all directions
        For HeadingLoop = NORTH To NORTHWEST
            nPos = NPCList(NPCIndex).Pos
            Server_HeadToPos HeadingLoop, nPos

            'If a legal pos and a NPC is found attack
            If MapInfo(nPos.Map).Data(nPos.X, nPos.Y).NPCIndex > 0 Then
                X = MapInfo(nPos.Map).Data(nPos.X, nPos.Y).NPCIndex
                If NPCList(X).Attackable Then
                    If NPCList(X).Hostile Then
                        If NPCList(X).OwnerIndex <> NotSlaveOfUserIndex Then
                        
                            'Get the damage
                            Damage = NPC_AttackNPC(NPCIndex, X)
            
                            'Send the attack packet
                            ConBuf.PreAllocate 12
                            ConBuf.Put_Byte DataCode.Combo_SlashSoundRotateDamage
                            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                            ConBuf.Put_Integer NPCList(X).Char.CharIndex
                            ConBuf.Put_Long NPCList(NPCIndex).AttackGrh
                            ConBuf.Put_Byte NPCList(NPCIndex).AttackSfx
                            If Damage > 32000 Then ConBuf.Put_Integer 32000 Else ConBuf.Put_Integer Damage
                            Data_Send ToNPCArea, X, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
            
                            'Apply damage
                            NPC_Damage X, 0, Damage, NPCList(NPCIndex).Char.CharIndex
                            NPC_AI_AttackNPC = 1
                            Exit Function
                        
                        End If
                    End If
                End If
            End If
            
        Next HeadingLoop
        
    '*** Ranged attacking ***
    Else
        
        'Check for the closest npc to attack
        X = NPC_AI_ClosestNPC(NPCIndex, NPCList(NPCIndex).AttackRange \ 2, NPCList(NPCIndex).AttackRange \ 2, NotSlaveOfUserIndex)
        If X > 0 Then
            
            'Check for a valid range
            If Server_Distance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, NPCList(X).Pos.X, NPCList(X).Pos.Y) <= NPCList(NPCIndex).AttackRange Then

                'Check for a valid path
                If Engine_ClearPath(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, NPCList(X).Pos.X, NPCList(X).Pos.Y) Then
    
                    'Get the damage
                    Damage = NPC_AttackNPC(NPCIndex, X)

                    'Face the NPC to the target
                    ConBuf.PreAllocate 13
                    ConBuf.Put_Byte DataCode.Combo_ProjectileSoundRotateDamage
                    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                    ConBuf.Put_Integer NPCList(X).Char.CharIndex
                    ConBuf.Put_Long NPCList(NPCIndex).AttackGrh
                    ConBuf.Put_Byte NPCList(NPCIndex).ProjectileRotateSpeed
                    ConBuf.Put_Byte NPCList(NPCIndex).AttackSfx
                    If Damage > 32000 Then ConBuf.Put_Integer 32000 Else ConBuf.Put_Integer Damage
                    Data_Send ToNPCArea, X, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

                    'Apply damage
                    NPC_Damage X, 0, Damage, NPCList(NPCIndex).Char.CharIndex
                    NPC_AI_AttackNPC = 1
                    Exit Function
                
                End If
            End If
        End If
        
    End If

End Function

Public Sub NPC_AI(ByVal NPCIndex As Integer)

'*****************************************************************
'Moves NPC based on it's .movement value
'*****************************************************************
Dim tHeading As Byte
Dim t1 As Byte
Dim t2 As Byte
Dim t3 As Byte
Dim t4 As Byte
Dim Y As Long
Dim X As Long
Dim b As Byte
Dim tX As Long
Dim tY As Long
Dim i As Integer
Dim tPos As WorldPos

    Log "Call NPC_AI(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Movement
    Select Case NPCList(NPCIndex).AI
    
        '*** Random movement ***
        Case 2
        
            'Attack
            If NPCList(NPCIndex).Hostile Then b = NPC_AI_Attack(NPCIndex)
            If b = 1 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
                Exit Sub
            End If
            
            'Move
            If Int(Rnd * 3) = 0 Then
                NPC_MoveChar NPCIndex, Int(Rnd * 8) + 1
            End If
            NPCList(NPCIndex).Counters.ActionDelay = NPCList(NPCIndex).Counters.ActionDelay + 500

        '*** Go towards nearby players - simple/fast AI ***
        Case 3
    
            'Attack
            If NPCList(NPCIndex).Hostile Then b = NPC_AI_Attack(NPCIndex)
            If b = 0 Then
                If NPCList(NPCIndex).Hostile Then b = NPC_AI_AttackNPC(NPCIndex)
            End If
            If b = 1 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
                Exit Sub
            End If
            
            'Look for a user
            X = NPC_AI_ClosestPC(NPCIndex, (ScreenWidth \ 64) - 1, (ScreenHeight \ 64) - 1)
            
            'If no users are nearby, then put a moderate delay to lighten the CPU load
            If X = 0 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000
                Exit Sub
            
            Else
            
                'Find the direction to move
                tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, UserList(X).Pos)
                
                'Move towards the retrieved position
                If NPC_MoveChar(NPCIndex, tHeading) = 0 Then
                
                    'Move towards alternate positions (the two directions that surround the selected direction)
                    Select Case tHeading
                        Case NORTH
                            t1 = NORTHEAST
                            t2 = NORTHWEST
                        Case EAST
                            t1 = NORTHEAST
                            t2 = SOUTHEAST
                        Case SOUTH
                            t1 = SOUTHWEST
                            t2 = SOUTHEAST
                        Case WEST
                            t1 = SOUTHWEST
                            t2 = NORTHWEST
                        Case NORTHEAST
                            t1 = NORTH
                            t2 = EAST
                        Case SOUTHEAST
                            t1 = EAST
                            t2 = SOUTH
                        Case SOUTHWEST
                            t1 = SOUTH
                            t2 = WEST
                        Case NORTHWEST
                            t1 = WEST
                            t2 = NORTH
                    End Select
                    
                    'Do the alternate movement
                    If NPC_MoveChar(NPCIndex, t1) = 0 Then
                        If NPC_MoveChar(NPCIndex, t2) = 0 Then
                            NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000   'Nowhere to go, so wait a while to try again
                        End If
                    End If
                
                End If
                    
                Exit Sub
    
            End If
            
        '*** Attack the nearest player, and run from them ***
        Case 4

            'Look for a near-by character
            i = NPC_AI_ClosestChar(NPCIndex, NPCList(NPCIndex).AttackRange \ 2, NPCList(NPCIndex).AttackRange \ 2, 0)
            
            If i = 0 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000
                Exit Sub
            Else
                
                'Get the position of the character
                Select Case CharList(i).CharType
                    Case CharType_PC: tPos = UserList(CharList(i).Index).Pos
                    Case CharType_NPC: tPos = NPCList(CharList(i).Index).Pos
                End Select
            
                'Run away (movement)
                If Server_RectDistance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, tPos.X, tPos.Y, 3, 3) Then
                    tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, tPos)
                    Select Case tHeading
                        Case NORTH: tHeading = SOUTH
                        Case NORTHEAST: tHeading = SOUTHWEST
                        Case EAST: tHeading = WEST
                        Case SOUTHEAST: tHeading = NORTHWEST
                        Case SOUTH: tHeading = NORTH
                        Case SOUTHWEST: tHeading = NORTHEAST
                        Case WEST: tHeading = EAST
                        Case NORTHWEST: tHeading = SOUTHEAST
                    End Select
                    If NPC_MoveChar(NPCIndex, tHeading) = 0 Then
                        Select Case tHeading
                            Case NORTH
                                t1 = NORTHEAST
                                t2 = NORTHWEST
                                t3 = EAST
                                t4 = WEST
                            Case EAST
                                t1 = NORTHEAST
                                t2 = SOUTHEAST
                                t3 = NORTH
                                t4 = SOUTH
                            Case SOUTH
                                t1 = SOUTHWEST
                                t2 = SOUTHEAST
                                t3 = WEST
                                t4 = EAST
                            Case WEST
                                t1 = SOUTHWEST
                                t2 = NORTHWEST
                                t3 = SOUTH
                                t4 = NORTH
                            Case NORTHEAST
                                t1 = NORTH
                                t2 = EAST
                                t3 = NORTHWEST
                                t4 = SOUTHEAST
                            Case SOUTHEAST
                                t1 = EAST
                                t2 = SOUTH
                                t3 = SOUTHWEST
                                t4 = NORTHEAST
                            Case SOUTHWEST
                                t1 = SOUTH
                                t2 = WEST
                                t3 = SOUTHEAST
                                t4 = NORTHWEST
                            Case NORTHWEST
                                t1 = WEST
                                t2 = NORTH
                                t3 = NORTHEAST
                                t4 = SOUTHWEST
                        End Select
                        If NPC_MoveChar(NPCIndex, t1) = 0 Then
                            If NPC_MoveChar(NPCIndex, t2) = 0 Then
                                If NPC_MoveChar(NPCIndex, t3) = 0 Then
                                    If NPC_MoveChar(NPCIndex, t4) = 0 Then  'If this doesn't happen, then we're out of stuff to do, so attack
                                        Select Case CharList(i).CharType
                                            Case CharType_PC: NPC_AI_Attack NPCIndex
                                            Case CharType_NPC: NPC_AI_AttackNPC NPCIndex
                                        End Select
                                        NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Exit Sub
                End If
                    
                'Attack
                Select Case CharList(i).CharType
                    Case CharType_PC: b = NPC_AI_Attack(NPCIndex)
                    Case CharType_NPC: b = NPC_AI_AttackNPC(NPCIndex)
                End Select
                If b Then
                    NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
                    Exit Sub
                End If
                
                'Nothing to do, so put on a delay
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000
                
            End If
            
        '*** Heal the nearest NPC, and run from users ***
        Case 5
            
            'Look for a user
            i = NPC_AI_ClosestPC(NPCIndex, (ScreenWidth \ 64) - 1, (ScreenHeight \ 64) - 1)
            If i > 0 Then
                tPos = UserList(i).Pos
            
            'Look for user slaves
            Else
                i = NPC_AI_ClosestNPC(NPCIndex, (ScreenWidth \ 64) - 1, (ScreenHeight \ 64) - 1, 0)
                If i > 0 Then tPos = NPCList(i).Pos
            End If
            
            If i = 0 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000
                Exit Sub
            Else
                
                'Run away (movement)
                If Server_RectDistance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, tPos.X, tPos.Y, 3, 3) Then
                    tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, tPos)
                    Select Case tHeading
                        Case NORTH: tHeading = SOUTH
                        Case NORTHEAST: tHeading = SOUTHWEST
                        Case EAST: tHeading = WEST
                        Case SOUTHEAST: tHeading = NORTHWEST
                        Case SOUTH: tHeading = NORTH
                        Case SOUTHWEST: tHeading = NORTHEAST
                        Case WEST: tHeading = EAST
                        Case NORTHWEST: tHeading = SOUTHEAST
                    End Select
                    If NPC_MoveChar(NPCIndex, tHeading) = 0 Then
                        Select Case tHeading
                            Case NORTH
                                t1 = NORTHEAST
                                t2 = NORTHWEST
                                t3 = EAST
                                t4 = WEST
                            Case EAST
                                t1 = NORTHEAST
                                t2 = SOUTHEAST
                                t3 = NORTH
                                t4 = SOUTH
                            Case SOUTH
                                t1 = SOUTHWEST
                                t2 = SOUTHEAST
                                t3 = WEST
                                t4 = EAST
                            Case WEST
                                t1 = SOUTHWEST
                                t2 = NORTHWEST
                                t3 = SOUTH
                                t4 = NORTH
                            Case NORTHEAST
                                t1 = NORTH
                                t2 = EAST
                                t3 = NORTHWEST
                                t4 = SOUTHEAST
                            Case SOUTHEAST
                                t1 = EAST
                                t2 = SOUTH
                                t3 = SOUTHWEST
                                t4 = NORTHEAST
                            Case SOUTHWEST
                                t1 = SOUTH
                                t2 = WEST
                                t3 = SOUTHEAST
                                t4 = NORTHWEST
                            Case NORTHWEST
                                t1 = WEST
                                t2 = NORTH
                                t3 = NORTHEAST
                                t4 = SOUTHWEST
                        End Select
                        If NPC_MoveChar(NPCIndex, t1) = 0 Then
                            If NPC_MoveChar(NPCIndex, t2) = 0 Then
                                If NPC_MoveChar(NPCIndex, t3) = 0 Then
                                    If NPC_MoveChar(NPCIndex, t4) = 0 Then  'If this doesn't happen, then we're out of stuff to do, so attack
                                        NPC_AI_Attack NPCIndex
                                        NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Exit Sub
                End If
                    
                'Heal
                If NPCList(NPCIndex).Counters.ActionDelay < timeGetTime Then
                    If NPCList(NPCIndex).Counters.SpellExhaustion < timeGetTime Then
                
                        'Loop through the NPCs in range
                        For tX = 1 To MaxServerDistanceX - 1
                            For tY = 1 To MaxServerDistanceY - 1
                                For X = NPCList(NPCIndex).Pos.X - tX To NPCList(NPCIndex).Pos.X + tX Step tX
                                    For Y = NPCList(NPCIndex).Pos.Y - tY To NPCList(NPCIndex).Pos.Y + tY Step tY
                                        If X > 1 Then
                                            If X < MapInfo(NPCList(NPCIndex).Pos.Map).Width Then
                                                If Y > 1 Then
                                                    If Y < MapInfo(NPCList(NPCIndex).Pos.Map).Height Then
                                                        If MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex > 0 Then
                                                            i = MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex
                                                            If NPCList(i).OwnerIndex = 0 Then   'Don't heal summoned NPCs
                                                                If NPCList(i).BaseStat(SID.MinHP) < NPCList(i).ModStat(SID.MaxHP) Then
                                                                    If Skill_Heal_NPCtoNPC(NPCIndex, i) Then
                                                                        'Skill was casted, abort
                                                                        Exit Sub
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next Y
                                Next X
                            Next tY
                        Next tX
                        NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000
                        
                    End If
                End If
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 750
                
            End If
            
        '*** Banker ***
        'Case 6
            'This NPC has no AI here - the only reference to the banker AI is in the clicking events.
            ' Just do a search for "ai = 6" to find it.
            
        '*** Summoned Melee Suicidal-Attacker ***
        Case 7

            'This routine is for summoned NPCs only!
            If NPCList(NPCIndex).OwnerIndex = 0 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 5000
                Exit Sub
            End If

            'Attack
            b = NPC_AI_AttackNPC(NPCIndex, NPCList(NPCIndex).OwnerIndex)
            If b = 1 Then
                NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
                Exit Sub
            End If
        
            'Check for a near-by NPC to attack
            i = NPC_AI_ClosestNPC(NPCIndex, 5, 5, NPCList(NPCIndex).OwnerIndex)
            If i > 0 Then tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, NPCList(i).Pos) Else tHeading = 0

            'Move towards the owner if no nearby NPCs to move to to attack were found
            If tHeading = 0 Then
                If Abs(CInt(NPCList(NPCIndex).Pos.X) - CInt(UserList(NPCList(NPCIndex).OwnerIndex).Pos.X)) < 2 Then
                    If Abs(CInt(NPCList(NPCIndex).Pos.Y) - CInt(UserList(NPCList(NPCIndex).OwnerIndex).Pos.Y)) < 2 Then
                        NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 500
                        Exit Sub
                    End If
                End If
                tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, UserList(NPCList(NPCIndex).OwnerIndex).Pos)
            End If

            'Move towards the retrieved position
            If NPC_MoveChar(NPCIndex, tHeading) = 0 Then
            
                'Move towards alternate positions (the two directions that surround the selected direction)
                Select Case tHeading
                    Case NORTH
                        t1 = NORTHEAST
                        t2 = NORTHWEST
                    Case EAST
                        t1 = NORTHEAST
                        t2 = SOUTHEAST
                    Case SOUTH
                        t1 = SOUTHWEST
                        t2 = SOUTHEAST
                    Case WEST
                        t1 = SOUTHWEST
                        t2 = NORTHWEST
                    Case NORTHEAST
                        t1 = NORTH
                        t2 = EAST
                    Case SOUTHEAST
                        t1 = EAST
                        t2 = SOUTH
                    Case SOUTHWEST
                        t1 = SOUTH
                        t2 = WEST
                    Case NORTHWEST
                        t1 = WEST
                        t2 = NORTH
                End Select
                
                'Do the alternate movement
                If NPC_MoveChar(NPCIndex, t1) = 0 Then
                    If NPC_MoveChar(NPCIndex, t2) = 0 Then
                        NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + 1000   'Nowhere to go, so wait a while to try again
                    End If
                End If
            
            End If

    End Select

End Sub

Public Function NPC_AI_ClosestChar(ByVal NPCIndex As Integer, ByVal SearchX As Byte, ByVal SearchY As Byte, Optional ByVal NotSlaveOfUserIndex As Integer = 0, Optional ByVal IsAttackable As Byte = 1, Optional ByVal IsHostile As Byte = 1) As Integer

'*****************************************************************
'Return the index of the closest player character (NPC or PC)
'Note - the char index is returned, not PC or NPC index!
'*****************************************************************
Dim tX As Integer
Dim tY As Integer
Dim X As Integer
Dim Y As Integer

    Log "Call NPC_AI_ClosestChar(" & NPCIndex & "," & SearchX & "," & SearchY & ")", CodeTracker '//\\LOGLINE//\\
    
    'Check for a valid, active NPC
    If NPCList(NPCIndex).flags.NPCAlive = 0 Then Exit Function
    If NPCList(NPCIndex).flags.NPCActive = 0 Then Exit Function
    
    'Expand the search range
    For tX = 1 To SearchX
        For tY = 1 To SearchY
            'Loop through the search area (only look on the outside of the search rectangle to prevent checking the same thing multiple times)
            For X = NPCList(NPCIndex).Pos.X - tX To NPCList(NPCIndex).Pos.X + tX Step tX
                For Y = NPCList(NPCIndex).Pos.Y - tY To NPCList(NPCIndex).Pos.Y + tY Step tY
                    'Make sure tile is legal
                    If X >= 1 Then
                        If X <= MapInfo(NPCList(NPCIndex).Pos.Map).Width Then
                            If Y >= 1 Then
                                If Y <= MapInfo(NPCList(NPCIndex).Pos.Map).Height Then
                                
                                    'Look for a npc
                                    If MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex > 0 Then
                                        If MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex <> NPCIndex Then
 
                                            'Perform special checks
                                            If IsAttackable Then
                                                If NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).Attackable = 0 Then GoTo NextChar
                                            End If
                                            If IsHostile Then
                                                If NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).Hostile = 0 Then GoTo NextChar
                                            End If
                                            If NotSlaveOfUserIndex > -1 Then
                                                If NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).OwnerIndex = NotSlaveOfUserIndex Then GoTo NextChar
                                            End If
   
                                            'Look for a NPC
                                            NPC_AI_ClosestChar = NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).Char.CharIndex
                                            Log "Rtrn NPC_AI_ClosestChar = " & NPC_AI_ClosestChar, CodeTracker '//\\LOGLINE//\\
                                            Exit Function
                                            
                                        End If
                                    
                                    'Look for a PC
                                    ElseIf MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).UserIndex > 0 Then
                                        NPC_AI_ClosestChar = UserList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).UserIndex).Char.CharIndex
                                        Log "Rtrn NPC_AI_ClosestChar = " & NPC_AI_ClosestChar, CodeTracker '//\\LOGLINE//\\
                                        Exit Function
                                    End If
                                    
                                End If
                            End If
                        End If
                    End If

NextChar:
                    
                Next Y
            Next X
        Next tY
    Next tX
    
    Log "Rtrn NPC_AI_ClosestChar = " & NPC_AI_ClosestChar, CodeTracker '//\\LOGLINE//\\

End Function

Public Function NPC_AI_ClosestPC(ByVal NPCIndex As Integer, ByVal SearchX As Byte, ByVal SearchY As Byte) As Integer

'*****************************************************************
'Return the index of the closest player character (PC)
'*****************************************************************
Dim tX As Integer
Dim tY As Integer
Dim X As Integer
Dim Y As Integer

    Log "Call NPC_AI_ClosestPC(" & NPCIndex & "," & SearchX & "," & SearchY & ")", CodeTracker '//\\LOGLINE//\\
    
    'Check for a valid, active NPC
    If NPCList(NPCIndex).flags.NPCAlive = 0 Then Exit Function
    If NPCList(NPCIndex).flags.NPCActive = 0 Then Exit Function
    
    'Expand the search range
    For tX = 1 To SearchX
        For tY = 1 To SearchY
            'Loop through the search area (only look on the outside of the search rectangle to prevent checking the same thing multiple times)
            For X = NPCList(NPCIndex).Pos.X - tX To NPCList(NPCIndex).Pos.X + tX Step tX
                For Y = NPCList(NPCIndex).Pos.Y - tY To NPCList(NPCIndex).Pos.Y + tY Step tY
                    'Make sure tile is legal
                    If X >= 1 Then
                        If X <= MapInfo(NPCList(NPCIndex).Pos.Map).Width Then
                            If Y >= 1 Then
                                If Y <= MapInfo(NPCList(NPCIndex).Pos.Map).Height Then
                                    'Look for a user
                                    If MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).UserIndex > 0 Then
                                        NPC_AI_ClosestPC = MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).UserIndex
                                        Log "Rtrn NPC_AI_ClosestPC = " & NPC_AI_ClosestPC, CodeTracker '//\\LOGLINE//\\
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
    
    Log "Rtrn NPC_AI_ClosestPC = " & NPC_AI_ClosestPC, CodeTracker '//\\LOGLINE//\\

End Function

Public Function NPC_AI_ClosestNPC(ByVal NPCIndex As Integer, ByVal SearchX As Byte, ByVal SearchY As Byte, Optional ByVal NotSlaveOfUserIndex As Integer = 0, Optional ByVal IsAttackable As Byte = 1, Optional ByVal IsHostile As Byte = 1) As Integer

'*****************************************************************
'Return the index of the closest player character (PC)
'*****************************************************************
Dim tX As Integer
Dim tY As Integer
Dim X As Integer
Dim Y As Integer

    Log "Call NPC_AI_ClosestNPC(" & NPCIndex & "," & SearchX & "," & SearchY & ")", CodeTracker '//\\LOGLINE//\\
    
    'Expand the search range
    For tX = 1 To SearchX
        For tY = 1 To SearchY
        
            'Loop through the search area (only look on the outside of the search rectangle to prevent checking the same thing multiple times)
            For X = NPCList(NPCIndex).Pos.X - tX To NPCList(NPCIndex).Pos.X + tX Step tX
                For Y = NPCList(NPCIndex).Pos.Y - tY To NPCList(NPCIndex).Pos.Y + tY Step tY
                
                    'Make sure tile is legal
                    If X >= 1 Then
                        If X <= MapInfo(NPCList(NPCIndex).Pos.Map).Width Then
                            If Y >= 1 Then
                                If Y <= MapInfo(NPCList(NPCIndex).Pos.Map).Height Then
                                
                                    'Look for a npc
                                    If MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex > 0 Then
                                        If MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex <> NPCIndex Then
 
                                            'Perform special checks
                                            If IsAttackable Then
                                                If NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).Attackable = 0 Then GoTo NextNPC
                                            End If
                                            If IsHostile Then
                                                If NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).Hostile = 0 Then GoTo NextNPC
                                            End If
                                            If NotSlaveOfUserIndex > -1 Then
                                                If NPCList(MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex).OwnerIndex = NotSlaveOfUserIndex Then GoTo NextNPC
                                            End If
                                            
                                            'We found our match!
                                            NPC_AI_ClosestNPC = MapInfo(NPCList(NPCIndex).Pos.Map).Data(X, Y).NPCIndex
                                            Log "Rtrn NPC_AI_ClosestNPC = " & NPC_AI_ClosestNPC, CodeTracker '//\\LOGLINE//\\
                                            Exit Function

                                        End If
                                    End If
                                    
                                End If
                            End If
                        End If
                    End If

NextNPC:

                Next Y
            Next X
        Next tY
    Next tX
    
    Log "Rtrn NPC_AI_ClosestNPC = " & NPC_AI_ClosestNPC, CodeTracker '//\\LOGLINE//\\

End Function

Private Function NPC_AttackNPC(ByVal NPCIndex As Integer, ByVal TargetIndex As Integer) As Long

'*****************************************************************
'Have a NPC attack a NPC
'*****************************************************************
Dim HitRate As Long
Dim Hit As Long

    Log "Call NPC_AttackNPC(" & NPCIndex & "," & TargetIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for an action delay
    If NPCList(NPCIndex).Counters.ActionDelay > timeGetTime Then
        Log "NPC_AttackNPC: NPC action delay > timeGetTime - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'Set the action delay
    NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
    
    'Update the hit rate
    HitRate = NPCList(NPCIndex).ModStat(SID.Agi) + 50 'Remember, AGI for NPCs is the hit rate!

    'Calculate if they hit
    If Server_RandomNumber(1, 100) >= (HitRate - NPCList(TargetIndex).ModStat(SID.Agi)) Then
        Log "NPC_AttackNPC: NPC's attack missed", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 5
        ConBuf.Put_Byte DataCode.Server_SetCharDamage
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        ConBuf.Put_Integer -1
        Data_Send ToNPCArea, TargetIndex, ConBuf.Get_Buffer, NPCList(TargetIndex).Pos.Map
        Exit Function
    End If

    'Calculate hit
    Hit = Server_RandomNumber(NPCList(NPCIndex).ModStat(SID.MinHIT), NPCList(NPCIndex).ModStat(SID.MaxHIT))
    Hit = Hit - (NPCList(TargetIndex).ModStat(SID.DEF) \ 2)
    If Hit < 1 Then Hit = 1
    Log "NPC_AttackNPC: Hit value = " & Hit, CodeTracker '//\\LOGLINE//\\

    'Return the value
    NPC_AttackNPC = Hit
    
End Function

Private Sub NPC_AttackUser_ApplyDamage(ByVal NPCIndex As Integer, ByVal UserIndex As Integer, ByVal Hit As Long)

'*****************************************************************
'Applies damage on a user
'*****************************************************************

    'Hit user
    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = UserList(UserIndex).Stats.BaseStat(SID.MinHP) - Hit

    'Display damage
    ConBuf.PreAllocate 5
    ConBuf.Put_Byte DataCode.Server_SetCharDamage
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Integer Hit
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

    'User Die
    If UserList(UserIndex).Stats.BaseStat(SID.MinHP) <= 0 Then
        Log "NPC_AttackUser_ApplyDamage: NPC's attack killed user", CodeTracker '//\\LOGLINE//\\
    
        'Kill user
        ConBuf.PreAllocate 3 + Len(NPCList(NPCIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 73
        ConBuf.Put_String NPCList(NPCIndex).Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        User_Kill UserIndex
        
    End If

End Sub

Private Function NPC_AttackUser(ByVal NPCIndex As Integer, ByVal UserIndex As Integer) As Long

'*****************************************************************
'Have a NPC attack a User
'*****************************************************************
Dim HitRate As Long
Dim Hit As Long

    Log "Call NPC_AttackUser(" & NPCIndex & "," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for an action delay
    If NPCList(NPCIndex).Counters.ActionDelay > timeGetTime Then
        Log "NPC_AttackUser: NPC action delay > timeGetTime - aborting", CodeTracker '//\\LOGLINE//\\
        NPC_AttackUser = -1
        Exit Function
    End If
    
    'Don't allow if not logged in completely
    If UserList(UserIndex).flags.UserLogged = 0 Then
        Log "NPC_AttackUser: User " & UserIndex & " (" & UserList(UserIndex).Name & ") not logged in - aborting", CodeTracker '//\\LOGLINE//\\
        NPC_AttackUser = -1
        Exit Function
    End If

    'Set the action delay
    NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + NPCDelayFight
    
    'Update the hit rate
    HitRate = NPCList(NPCIndex).ModStat(SID.Agi) + 50 'Remember, AGI for NPCs is the hit rate!

    'Calculate if they hit
    If Server_RandomNumber(1, 100) >= (HitRate - UserList(UserIndex).Stats.ModStat(SID.Agi)) Then
        Log "NPC_AttackUser: NPC's attack missed", CodeTracker '//\\LOGLINE//\\
        NPC_AttackUser = 0
        Exit Function
    End If

    'Calculate hit
    Hit = Server_RandomNumber(NPCList(NPCIndex).ModStat(SID.MinHIT), NPCList(NPCIndex).ModStat(SID.MaxHIT))
    Hit = Hit - (UserList(UserIndex).Stats.ModStat(SID.DEF) \ 2)
    If Hit < 1 Then Hit = 1
    
    'Return the value
    NPC_AttackUser = Hit

End Function

Private Sub NPC_ChangeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal NPCIndex As Integer, Optional ByVal Body As Integer = -1, Optional ByVal Head As Integer = -1, Optional ByVal Weapon As Integer = -1, Optional ByVal Hair As Integer = -1, Optional ByVal Wings As Integer = -1)

'*****************************************************************
'Changes a NPC char's head,body and heading
'*****************************************************************
Dim ChangeFlags As Byte
Dim FlagSizes As Byte

    Log "Call NPC_ChangeChar(" & sndRoute & "," & sndIndex & "," & NPCIndex & "," & Body & "," & Head & "," & Weapon & "," & Hair & "," & Wings & ")", CodeTracker  '//\\LOGLINE//\\

    'Check for a valid NPC
    If NPCIndex <= 0 Then
        Log "NPC_ChangeChar: NPCIndex <= 0 - aborting", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    If NPCIndex > LastNPC Then
        Log "NPC_ChangeChar: NPCIndex > LastNPC - aborting", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for changed values
    With NPCList(NPCIndex).Char
        If Body > -1 Then
            If .Body <> Body Then
                .Body = Body
                ChangeFlags = ChangeFlags Or 1
                FlagSizes = FlagSizes + 2
            End If
        End If
        If Head > -1 Then
            If .Head <> Head Then
                .Head = Head
                ChangeFlags = ChangeFlags Or 2
                FlagSizes = FlagSizes + 2
            End If
        End If
        If Weapon > -1 Then
            If .Weapon <> Weapon Then
                .Weapon = Weapon
                ChangeFlags = ChangeFlags Or 4
                FlagSizes = FlagSizes + 2
            End If
        End If
        If Hair > -1 Then
            If .Hair <> Hair Then
                .Hair = Hair
                ChangeFlags = ChangeFlags Or 8
                FlagSizes = FlagSizes + 2
            End If
        End If
        If Wings > -1 Then
            If .Wings <> Wings Then
                .Wings = Wings
                ChangeFlags = ChangeFlags Or 16
                FlagSizes = FlagSizes + 2
            End If
        End If
    End With
    
    'Make sure there is a packet to send
    If ChangeFlags = 0 Then Exit Sub
    
    'Create the packet
    ConBuf.PreAllocate 4 + FlagSizes
    ConBuf.Put_Byte DataCode.Server_ChangeChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Byte ChangeFlags
    If ChangeFlags And 1 Then ConBuf.Put_Integer Body
    If ChangeFlags And 2 Then ConBuf.Put_Integer Head
    If ChangeFlags And 8 Then ConBuf.Put_Integer Weapon
    If ChangeFlags And 16 Then ConBuf.Put_Integer Hair
    If ChangeFlags And 32 Then ConBuf.Put_Integer Wings
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map, PP_ChangeChar

End Sub

Public Sub NPC_Close(ByVal NPCIndex As Integer, Optional ByVal CleanArray As Byte = 1)

'*****************************************************************
'Closes a NPC
'*****************************************************************

    Log "Call NPC_Close(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Close down the NPC
    NPCList(NPCIndex).flags.NPCActive = 0
    CharList(NPCList(NPCIndex).Char.CharIndex).Index = 0
    CharList(NPCList(NPCIndex).Char.CharIndex).CharType = 0

    'Clean up the NPC array
    If NPCIndex = LastNPC Then
        If CleanArray Then NPC_CleanArray
    End If
    
    'Update number of NPCs
    If NumNPCs <> 0 Then NumNPCs = NumNPCs - 1
    Log "NPC_Close: NumNPCs = " & NumNPCs, CodeTracker '//\\LOGLINE//\\

End Sub

Public Sub NPC_CleanArray()

'*****************************************************************
'Cleans the NPCList array to free memory
'*****************************************************************
Dim t As Integer    'Holds the unaltered value of LastNPC
    
    'Prevent crashing
    If LastNPC = 0 Then
        Erase NPCList
        Exit Sub
    End If

    'Store the LastNPC value for comparison later
    t = LastNPC

    'Loop through the NPCs from the last NPC backwards to find the number of slots we can clear
    Do Until NPCList(LastNPC).flags.NPCActive = 1
        LastNPC = LastNPC - 1
        If LastNPC = 0 Then Exit Do
    Loop
    
    'Check if the LastNPC value has changed
    If t <> LastNPC Then
    
        'Resize the array (or erase)
        If LastNPC <> 0 Then
            ReDim Preserve NPCList(1 To LastNPC)
        Else
            Erase NPCList()
        End If
        
    End If

End Sub

Public Sub NPC_Heal(ByVal NPCIndex As Integer, ByVal Value As Integer)

'*****************************************************************
'Raise a NPC's HP - ONLY USE THIS SUB TO RAISE NPC HP
'*****************************************************************
Dim HPA As Byte
Dim HPB As Byte

    'Get the pre-healing percentage
    HPA = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)

    'Raise the HP
    NPCList(NPCIndex).BaseStat(SID.MinHP) = NPCList(NPCIndex).BaseStat(SID.MinHP) + Value
    
    'Don't go over the limit
    If NPCList(NPCIndex).BaseStat(SID.MinHP) > NPCList(NPCIndex).ModStat(SID.MaxHP) Then NPCList(NPCIndex).BaseStat(SID.MinHP) = NPCList(NPCIndex).ModStat(SID.MaxHP)

    'Check to update health percentage client-side
    If NPCList(NPCIndex).BaseStat(SID.MinHP) > 0 Then
        HPB = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
        If HPA <> HPB Then
            ConBuf.PreAllocate 4
            ConBuf.Put_Byte DataCode.Server_CharHP
            ConBuf.Put_Byte HPB
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
        End If
    End If

End Sub

Public Sub NPC_Damage(ByVal NPCIndex As Integer, ByVal UserIndex As Integer, ByVal Damage As Long, Optional ByVal AttackerCharIndex As Integer = 0)

'*****************************************************************
'Do damage to a NPC - ONLY USE THIS SUB TO HURT NPCS
'*****************************************************************
Dim NewSlot As Byte
Dim NewX As Byte
Dim NewY As Byte
Dim HPA As Byte         'HP percentage before reducing hp
Dim HPB As Byte         'HP percentage after reducing hp
Dim i As Integer

    Log "Call NPC_Damage(" & NPCIndex & "," & UserIndex & "," & Damage & ")", CodeTracker  '//\\LOGLINE//\\

    'Check if the NPC can be attacked
    If NPCList(NPCIndex).Attackable = 0 Then
        Log "NPC_Damage: Attackable = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'If NPC has no health, they can not be attacked
    If NPCList(NPCIndex).ModStat(SID.MaxHP) = 0 Then
        Log "NPC_Damage: ModStat MaxHP = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If NPCList(NPCIndex).BaseStat(SID.MaxHP) = 0 Then
        Log "NPC_Damage: BaseStat MaxHP = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Make sure the NPC isn't the user's slave (don't damage your slaves)
    If UserIndex > 0 Then
        If NPCList(NPCIndex).OwnerIndex = UserIndex Then Exit Sub
    End If
    
    'Get the pre-damage percentage
    HPA = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)

    'Lower the NPC's life
    NPCList(NPCIndex).BaseStat(SID.MinHP) = NPCList(NPCIndex).BaseStat(SID.MinHP) - Damage
 
    'Check to update health percentage client-side
    If NPCList(NPCIndex).BaseStat(SID.MinHP) > 0 Then
        HPB = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
        If HPA <> HPB Then
            ConBuf.PreAllocate 4
            ConBuf.Put_Byte DataCode.Server_CharHP
            ConBuf.Put_Byte HPB
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
        End If
    End If

    'Display the damage on the client screen
    ConBuf.PreAllocate 5
    ConBuf.Put_Byte DataCode.Server_SetCharDamage
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Integer Damage
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

    'Check if the NPC died
    If NPCList(NPCIndex).BaseStat(SID.MinHP) <= 0 Then
        Log "NPC_Damage: NPC killed", CodeTracker '//\\LOGLINE//\\
    
        'If the NPC was killed by a user
        If UserIndex > 0 Then
        
            Log "NPC_Damage: It was a user who killed the NPC", CodeTracker '//\\LOGLINE//\\

            'Check on quests
            For i = 1 To MaxQuests
                If UserList(UserIndex).Quest(i) > 0 Then
                    If QuestData(UserList(UserIndex).Quest(i)).FinishReqNPC = NPCList(NPCIndex).NPCNumber Then
                        Log "NPC_Damage: User killed a NPC required for a quest", CodeTracker '//\\LOGLINE//\\

                        'User must kill at least one more of the NPC
                        If UserList(UserIndex).QuestStatus(i).NPCKills < QuestData(UserList(UserIndex).Quest(i)).FinishReqNPCAmount Then
                            UserList(UserIndex).QuestStatus(i).NPCKills = UserList(UserIndex).QuestStatus(i).NPCKills + 1
                            ConBuf.PreAllocate 7 + Len(NPCName(QuestData(UserList(UserIndex).Quest(i)).FinishReqNPC))
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 74
                            ConBuf.Put_Integer UserList(UserIndex).QuestStatus(i).NPCKills
                            ConBuf.Put_Integer QuestData(UserList(UserIndex).Quest(i)).FinishReqNPCAmount
                            ConBuf.Put_String NPCName(QuestData(UserList(UserIndex).Quest(i)).FinishReqNPC)
                            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                            
                        End If

                    End If
                End If
            Next i
            
            'Check if the user is part of a group
            If UserList(UserIndex).GroupIndex > 0 Then
            
                'Split up the exp and gold among the group
                Group_EXPandGold UserIndex, UserList(UserIndex).GroupIndex, NPCList(NPCIndex).GiveEXP, NPCList(NPCIndex).GiveGLD
            
            Else
    
                'Give EXP and gold to just the user
                UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + NPCList(NPCIndex).GiveGLD
                User_RaiseExp UserIndex, NPCList(NPCIndex).GiveEXP
            
            End If
            
            'Display kill message to the user
            ConBuf.PreAllocate 3 + Len(NPCList(NPCIndex).Name)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 75
            ConBuf.Put_String NPCList(NPCIndex).Name
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'If the NPC was attacked by a NPC slave to a user
        Else
            If AttackerCharIndex > 0 Then
                If CharList(AttackerCharIndex).CharType = CharType_NPC Then
                    If NPCList(CharList(AttackerCharIndex).Index).OwnerIndex > 0 Then
                        'The NPC is owned by a user, so give the EXP to the user (or group) like above
                        i = NPCList(CharList(AttackerCharIndex).Index).OwnerIndex
                        If i > 0 Then
                            If i <= LastUser Then
                                If UserList(i).GroupIndex > 0 Then
                                    Group_EXPandGold i, UserList(i).GroupIndex, NPCList(NPCIndex).GiveEXP, NPCList(NPCIndex).GiveGLD
                                Else
                                    User_RaiseExp i, NPCList(NPCIndex).GiveEXP
                                    UserList(i).Stats.BaseStat(SID.Gold) = UserList(i).Stats.BaseStat(SID.Gold) + NPCList(NPCIndex).GiveGLD
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
        'Drop items
        If NPCList(NPCIndex).NumDropItems > 0 Then
            For i = 1 To NPCList(NPCIndex).NumDropItems
                If NPCList(NPCIndex).DropRate(i) > (Rnd * 100) Then
                    Log "NPC_Damage: Item dropped (" & NPCList(NPCIndex).DropRate(i) & ")", CodeTracker '//\\LOGLINE//\\
                    
                    'Get the closest available position to put the item
                    Obj_ClosestFreeSpot NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, NewX, NewY, NewSlot
                    
                    'Make sure the position is valid
                    If NewX = 0 Then
                        
                        'If this object is invalid, so will the rest of them be, so just skip them all :(
                        Log "NPC_Damage: No valid item drop spot found - current and following loot will not appear", CodeTracker '//\\LOGLINE//\\
                        Exit For
                    
                    End If
                
                    'Create the object
                    Obj_Make NPCList(NPCIndex).DropItems(i), NewSlot, NPCList(NPCIndex).Pos.Map, NewX, NewY
                
                End If
            Next i
        End If

        'Kill off the NPC
        NPC_Kill NPCIndex

    End If

End Sub

Sub NPC_EraseChar(ByVal NPCIndex As Integer)

'*****************************************************************
'Erase a character
'*****************************************************************

    Log "Call NPC_EraseChar(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Remove from list
    CharList(NPCList(NPCIndex).Char.CharIndex).Index = 0
    CharList(NPCList(NPCIndex).Char.CharIndex).CharType = 0
    
    'Remove from map
    MapInfo(NPCList(NPCIndex).Pos.Map).Data(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0

    'Send erase command to clients
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.Server_EraseChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    Data_Send ToMap, 0, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

    'Clear the variables
    NPCList(NPCIndex).Char.CharIndex = 0
    NPCList(NPCIndex).flags.NPCAlive = 0

    'Set at the respawn spot
    NPCList(NPCIndex).Pos.Map = NPCList(NPCIndex).StartPos.Map
    NPCList(NPCIndex).Pos.X = NPCList(NPCIndex).StartPos.X
    NPCList(NPCIndex).Pos.Y = NPCList(NPCIndex).StartPos.Y

End Sub

Public Sub NPC_Kill(ByVal NPCIndex As Integer)

'*****************************************************************
'Kill a NPC
'*****************************************************************
Dim Slot As Byte
Dim i As Byte

    Log "Call NPC_Kill(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'If thralled, remove them
    If NPCList(NPCIndex).flags.Thralled Then
        NPCList(NPCIndex).flags.NPCActive = 0
        
        'If they were bounded as a slave (summon) NPC, change the user's summon count
        If NPCList(NPCIndex).OwnerIndex > 0 Then

            With UserList(NPCList(NPCIndex).OwnerIndex)

                'Find what slot the NPC was occupying
                For i = 1 To .NumSlaves
                    If .SlaveNPCIndex(i) = NPCIndex Then Exit For
                Next i
                Slot = i
                
                'Clean up the array
                If .NumSlaves = 1 Then
                    Erase .SlaveNPCIndex
                ElseIf Slot = .NumSlaves Then
                    ReDim Preserve .SlaveNPCIndex(1 To Slot - 1)
                Else
                    For i = Slot To .NumSlaves - 1
                        .SlaveNPCIndex(i) = .SlaveNPCIndex(i + 1)
                    Next i
                End If
                .NumSlaves = .NumSlaves - 1
                
            End With
            
        End If
                
    Else
        'Set death time for respawn wait
        NPCList(NPCIndex).Counters.RespawnCounter = timeGetTime + NPCList(NPCIndex).RespawnWait
    End If
    
    'Set health back to 100%
    NPCList(NPCIndex).BaseStat(SID.MinHP) = NPCList(NPCIndex).ModStat(SID.MaxHP)

    'Erase it from map
    NPC_EraseChar NPCIndex

End Sub

Public Sub NPC_MakeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal NPCIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Makes and places a NPC character
'*****************************************************************
Dim SndHP As Byte
Dim SndMP As Byte

    Log "Call NPC_MakeChar(" & sndRoute & "," & sndIndex & "," & NPCIndex & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Place character on map
    MapInfo(Map).Data(X, Y).NPCIndex = NPCIndex

    'Set alive flag
    NPCList(NPCIndex).flags.NPCAlive = 1

    'Set the hp/mp to send
    If NPCList(NPCIndex).ModStat(SID.MaxHP) > 0 Then SndHP = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
    If NPCList(NPCIndex).ModStat(SID.MaxMAN) > 0 Then SndMP = CByte((NPCList(NPCIndex).BaseStat(SID.MinMAN) / NPCList(NPCIndex).ModStat(SID.MaxMAN)) * 100)

    'NPCs wont be created with active spells
    ZeroMemory NPCList(NPCIndex).Skills, Len(NPCList(NPCIndex).Skills)

    'Send make character command to clients
    ConBuf.PreAllocate 22 + Len(NPCList(NPCIndex).Name)
    ConBuf.Put_Byte DataCode.Server_MakeChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Body
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Head
    ConBuf.Put_Byte NPCList(NPCIndex).Char.Heading
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Byte X
    ConBuf.Put_Byte Y
    ConBuf.Put_Byte NPCList(NPCIndex).BaseStat(SID.Speed)   'We dont use modstat on speed since for one it may not have been updated
    ConBuf.Put_String NPCList(NPCIndex).Name                ' yet, along with theres nothing to mod the stat
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Weapon
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Hair
    ConBuf.Put_Integer NPCList(NPCIndex).Char.Wings
    ConBuf.Put_Byte SndHP
    ConBuf.Put_Byte SndMP
    ConBuf.Put_Byte NPCList(NPCIndex).ChatID
    If NPCList(NPCIndex).OwnerIndex > 0 Then
        ConBuf.Put_Byte ClientCharType_Slave
        ConBuf.Put_Integer UserList(NPCList(NPCIndex).OwnerIndex).Char.CharIndex
    Else
        ConBuf.Put_Byte ClientCharType_NPC
    End If
    
    'Send the NPC
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, Map

End Sub

Public Function NPC_MoveChar(ByVal NPCIndex As Integer, ByVal nHeading As Byte) As Byte

'*****************************************************************
'Moves a NPC from one tile to another
'*****************************************************************

Dim nPos As WorldPos

    Log "Call NPC_MoveChar(" & NPCIndex & "," & nHeading & ")", CodeTracker '//\\LOGLINE//\\
    
    'Move
    nPos = NPCList(NPCIndex).Pos
    Server_HeadToPos nHeading, nPos

    'Move if legal pos
    If Server_LegalPos(NPCList(NPCIndex).Pos.Map, nPos.X, nPos.Y, nHeading, True) Then

        'Set the move delay (we set the lag buffer to 0 since NPCs don't lag)
        NPCList(NPCIndex).Counters.ActionDelay = timeGetTime + Server_WalkTimePerTile(NPCList(NPCIndex).ModStat(SID.Speed))

        'Send the movement update packet
        ConBuf.PreAllocate 6
        ConBuf.Put_Byte DataCode.Server_MoveChar
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Byte nPos.X
        ConBuf.Put_Byte nPos.Y
        ConBuf.Put_Byte nHeading
        Data_Send ToNPCMove, NPCIndex, ConBuf.Get_Buffer

        'Update map and user pos
        MapInfo(NPCList(NPCIndex).Pos.Map).Data(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0
        NPCList(NPCIndex).Pos = nPos
        NPCList(NPCIndex).Char.Heading = nHeading
        NPCList(NPCIndex).Char.HeadHeading = nHeading
        MapInfo(NPCList(NPCIndex).Pos.Map).Data(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = NPCIndex
        
        'NPC moved, return the flag
        NPC_MoveChar = 1

    End If
    
    Log "Rtrn NPC_MoveChar = " & NPC_MoveChar, CodeTracker '//\\LOGLINE//\\

End Function

Public Function NPC_NextOpen() As Integer

'*****************************************************************
'Finds the next open NPC Index in NPCList
'*****************************************************************

Dim LoopC As Long

    Log "Call NPC_NextOpen", CodeTracker '//\\LOGLINE//\\

    Do
        LoopC = LoopC + 1
        If LoopC > LastNPC Then
            LoopC = LastNPC + 1
            Exit Do
        End If
        If LoopC > MaxNPCs Then
            LoopC = 0
            Exit Do
        End If
    Loop While NPCList(LoopC).flags.NPCActive = 1

    NPC_NextOpen = LoopC
    
    Log "Rtrn NPC_NextOpen = " & NPC_NextOpen, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub NPC_Spawn(ByVal NPCIndex As Integer, Optional ByVal BypassUpdate As Byte = 0)

'*****************************************************************
'Places a NPC that has been Opened
'*****************************************************************

Dim TempPos As WorldPos
Dim CharIndex As Integer

    Log "Call NPC_Spawn(" & NPCIndex & "," & BypassUpdate & ")", CodeTracker '//\\LOGLINE//\\

    'Give it a char index if needed
    If NPCList(NPCIndex).Char.CharIndex = 0 Then
        CharIndex = Server_NextOpenCharIndex
        NPCList(NPCIndex).Char.CharIndex = CharIndex
        CharList(CharIndex).Index = NPCIndex
        CharList(CharIndex).CharType = CharType_NPC
    End If

    'Find a place to put npc
    Server_ClosestLegalPos NPCList(NPCIndex).StartPos, TempPos
    If Not Server_LegalPos(TempPos.Map, TempPos.X, TempPos.Y, 0) Then
        Log "NPC_Spawn: No legal pos found", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Set vars
    NPCList(NPCIndex).Pos = TempPos
    NPCList(NPCIndex).flags.NPCAlive = 1

    'Make NPC Char
    If BypassUpdate = 0 Then
        If MapInfo(TempPos.Map).NumUsers > 0 Then
            NPC_MakeChar ToMap, MapUsers(TempPos.Map).Index(1), NPCIndex, TempPos.Map, TempPos.X, TempPos.Y
        End If
    End If
    
End Sub
