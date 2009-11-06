Attribute VB_Name = "NPCs"
Option Explicit

Public Sub NPC_UpdateModStats(ByVal NPCIndex As Integer)
Dim i As Long

    Log "Call NPC_UpdateModStats(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    With NPCList(NPCIndex)
    
        'Copy the base stats to the mod stats (we can use copymemory since we dont have to give item bonuses)
        CopyMemory .ModStat(1), .BaseStat(1), NumStats * 4  '* 4 since we are using longs (4 bytes)
            
        'War curse
        If .Skills.WarCurse > 0 Then
            Log "NPC_UpdateModStats: Updating modstats with effects from WarCurse", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.Agi) = .ModStat(SID.Agi) - (.Skills.WarCurse * 0.25)
            .ModStat(SID.DEF) = .ModStat(SID.DEF) - (.Skills.WarCurse * 0.25)
            .ModStat(SID.Str) = .ModStat(SID.Str) - (.Skills.WarCurse * 0.25)
            .ModStat(SID.Mag) = .ModStat(SID.Mag) - (.Skills.WarCurse * 0.25)
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - (.Skills.WarCurse * 0.25)
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - (.Skills.WarCurse * 0.25)
            .ModStat(SID.WeaponSkill) = .ModStat(SID.WeaponSkill) - (.Skills.WarCurse * 0.25)
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
            .ModStat(SID.Agi) = .ModStat(SID.Agi) + .Skills.Bless * 0.5
            .ModStat(SID.Mag) = .ModStat(SID.Mag) + .Skills.Bless * 0.5
            .ModStat(SID.Str) = .ModStat(SID.Str) + .Skills.Bless * 0.5
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + .Skills.Bless * 0.25
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + .Skills.Bless * 0.25
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + .Skills.Bless * 0.25
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
Dim NewHeading As Byte
Dim nPos As WorldPos
Dim Angle As Single
Dim X As Long

    Log "Call NPC_AI_Attack(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    '*** Melee attacking ***
    If NPCList(NPCIndex).AttackRange <= 1 Then
        
        'Check in all directions
        For HeadingLoop = NORTH To NORTHWEST
            nPos = NPCList(NPCIndex).Pos
            Server_HeadToPos HeadingLoop, nPos

            'If a legal pos and a user is found attack
            If MapData(nPos.Map, nPos.X, nPos.Y).UserIndex > 0 Then

                'Face the NPC to the target and tell everyone in the PC area to show the attack animation
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Rotate
                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                ConBuf.Put_Byte HeadingLoop
                ConBuf.Put_Byte DataCode.User_Attack
                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer

                'Attack
                NPC_AttackUser NPCIndex, MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                NPC_AI_Attack = 1
                Exit Function

            End If
            
        Next HeadingLoop
        
    '*** Ranged attacking ***
    Else
        
        'Check for the closest user
        X = NPC_AI_ClosestPC(NPCIndex, 10, 8)
        If X > 0 Then
            
            'Check for a valid range
            If Server_Distance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(X).Pos.X, UserList(X).Pos.Y) <= NPCList(NPCIndex).AttackRange Then

                'Get the new heading
                Angle = Engine_GetAngle(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(X).Pos.X, UserList(X).Pos.Y)
                If Angle >= 337.5 Or Angle <= 22.5 Then '337.5 to 22.5
                    NewHeading = NORTH
                ElseIf Angle <= 67.5 Then   '22.5 to 67.5
                    NewHeading = NORTHEAST
                ElseIf Angle <= 112.5 Then  '67.5 to 112.5
                    NewHeading = EAST
                ElseIf Angle <= 157.5 Then  '112.5 to 157.5
                    NewHeading = SOUTHEAST
                ElseIf Angle <= 202.5 Then  '157.5 to 202.5
                    NewHeading = SOUTH
                ElseIf Angle <= 247.5 Then  '202.5 to 247.5
                    NewHeading = SOUTHWEST
                ElseIf Angle <= 292.5 Then  '247.5 to 292.5
                    NewHeading = WEST
                Else                        '292.5 to 337.5
                    NewHeading = NORTHWEST
                End If

                'Face the NPC to the target
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Rotate
                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                ConBuf.Put_Byte NewHeading
                Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
   
                'Attack the user
                NPC_AttackUser NPCIndex, X
                NPC_AI_Attack = 1
                Exit Function
                
            End If
        End If
    End If

End Function

Sub NPC_AI(ByVal NPCIndex As Integer)

'*****************************************************************
'Moves NPC based on it's .movement value
'*****************************************************************
Dim tHeading As Byte
Dim t1 As Byte
Dim t2 As Byte
Dim Y As Long
Dim X As Long
Dim b As Byte
Dim tX As Long
Dim tY As Long
Dim i As Integer

    Log "Call NPC_AI(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Do nothing if no players are on the map
    If MapInfo(NPCList(NPCIndex).Pos.Map).NumUsers = 0 Then
        Log "NPC_AI: NPC's map has no users - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Update the action delay counter
    If NPCList(NPCIndex).Flags.ActionDelay > 0 Then
        Log "NPC_AI: NPC's action delay > 0 - aborting", CodeTracker '//\\LOGLINE//\\
        NPCList(NPCIndex).Flags.ActionDelay = NPCList(NPCIndex).Flags.ActionDelay - Elapsed
        Exit Sub
    End If

    'Movement
    Select Case NPCList(NPCIndex).AI
    
        '*** Random movement ***
        Case 2
        
            'Attack
            If NPCList(NPCIndex).Hostile Then b = NPC_AI_Attack(NPCIndex)
            If b = 1 Then
                NPCList(NPCIndex).Flags.ActionDelay = NPCDelayFight
                Exit Sub
            End If
            
            'Move
            NPC_MoveChar NPCIndex, Int(Rnd * 8) + 1

        '*** Go towards nearby players - simple/fast AI ***
        Case 3
    
            'Attack
            If NPCList(NPCIndex).Hostile Then b = NPC_AI_Attack(NPCIndex)
            If b = 1 Then
                NPCList(NPCIndex).Flags.ActionDelay = NPCDelayFight
                Exit Sub
            End If
            
            'Look for a user
            X = NPC_AI_ClosestPC(NPCIndex, 10, 8)
            
            'If no users are nearby, then put a moderate delay to lighten the CPU load
            If X = 0 Then
                NPCList(NPCIndex).Flags.ActionDelay = 1000
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
                        Log "NPC_AI: Using alternate movement method for AI 3", CodeTracker '//\\LOGLINE//\\
                        NPC_MoveChar NPCIndex, t2   'If this doesn't happen, then we're out of stuff to do
                    End If
                
                End If
                    
                Exit Sub
    
            End If
            
        '*** Attack the nearest player, and run from them ***
        Case 4
            
            'Look for a user
            X = NPC_AI_ClosestPC(NPCIndex, 10, 8)
            If X = 0 Then
                NPCList(NPCIndex).Flags.ActionDelay = 1000
                Exit Sub
            Else
            
                'Run away (movement)
                If Server_RectDistance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(X).Pos.X, UserList(X).Pos.Y, 3, 3) Then
                    tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, UserList(X).Pos)
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
                        If NPC_MoveChar(NPCIndex, t1) = 0 Then
                            Log "NPC_AI: Using alternate movement method for AI 4", CodeTracker '//\\LOGLINE//\\
                            NPC_MoveChar NPCIndex, t2   'If this doesn't happen, then we're out of stuff to do
                        End If
                    End If
                    Exit Sub
                End If
                    
                'Attack
                b = NPC_AI_Attack(NPCIndex)
                If b Then
                    NPCList(NPCIndex).Flags.ActionDelay = NPCDelayFight
                    Exit Sub
                End If
                
            End If
            
        '*** Heal the nearest NPC, and run from users ***
        Case 5
            
            'Look for a user
            X = NPC_AI_ClosestPC(NPCIndex, 10, 8)
            If X = 0 Then
                NPCList(NPCIndex).Flags.ActionDelay = 1000
                Exit Sub
            Else
            
                'Run away (movement)
                If Server_RectDistance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(X).Pos.X, UserList(X).Pos.Y, 3, 3) Then
                    tHeading = Server_FindDirection(NPCList(NPCIndex).Pos, UserList(X).Pos)
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
                        If NPC_MoveChar(NPCIndex, t1) = 0 Then
                            Log "NPC_AI: Using alternate movement method for AI 4", CodeTracker '//\\LOGLINE//\\
                            NPC_MoveChar NPCIndex, t2   'If this doesn't happen, then we're out of stuff to do
                        End If
                    End If
                    Exit Sub
                End If
                    
                'Heal
                If NPCList(NPCIndex).Flags.ActionDelay <= 0 Then
                    If NPCList(NPCIndex).Counters.SpellExhaustion <= 0 Then
                
                        'Loop through the NPCs in range
                        For tX = 1 To MaxServerDistanceX - 1
                            For tY = 1 To MaxServerDistanceY - 1
                                For X = NPCList(NPCIndex).Pos.X - tX To NPCList(NPCIndex).Pos.X + tX Step tX
                                    For Y = NPCList(NPCIndex).Pos.Y - tY To NPCList(NPCIndex).Pos.Y + tY Step tY
                                        If X > MinXBorder Then
                                            If X < MaxXBorder Then
                                                If Y > MinYBorder Then
                                                    If Y < MaxYBorder Then
                                                        If MapData(NPCList(NPCIndex).Pos.Map, X, Y).NPCIndex > 0 Then
                                                            i = MapData(NPCList(NPCIndex).Pos.Map, X, Y).NPCIndex
                                                            If NPCList(i).BaseStat(SID.MinHP) < NPCList(i).ModStat(SID.MaxHP) Then
                                                                If Skill_Heal_NPCtoNPC(NPCIndex, i) Then
                                                                    NPCList(NPCIndex).Flags.ActionDelay = Heal_Exhaust
                                                                    Exit Sub
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
                        NPCList(NPCIndex).Flags.ActionDelay = 2000
                        
                    End If
                End If
                NPCList(NPCIndex).Flags.ActionDelay = 250
                
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

    Log "Call NPC_AI_ClosestPC(" & NPCIndex & "," & SearchX & "," & SearchY & ")", CodeTracker '//\\LOGLINE//\\
    
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

Public Function NPC_AI_ClosestNPC(ByVal NPCIndex As Integer, ByVal SearchX As Byte, ByVal SearchY As Byte) As Integer

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
                    If X > MinXBorder Then
                        If X < MaxXBorder Then
                            If Y > MinYBorder Then
                                If Y < MaxYBorder Then
                                    'Look for a npc
                                    If MapData(NPCList(NPCIndex).Pos.Map, X, Y).NPCIndex > 0 Then
                                        NPC_AI_ClosestNPC = MapData(NPCList(NPCIndex).Pos.Map, X, Y).NPCIndex
                                        Log "Rtrn NPC_AI_ClosestNPC = " & NPC_AI_ClosestNPC, CodeTracker '//\\LOGLINE//\\
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
    
    Log "Rtrn NPC_AI_ClosestNPC = " & NPC_AI_ClosestNPC, CodeTracker '//\\LOGLINE//\\

End Function

Sub NPC_AttackUser(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)

'*****************************************************************
'Have a NPC attack a User
'*****************************************************************

Dim Hit As Integer

    Log "Call NPC_AttackUser(" & NPCIndex & "," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for an action delay
    If NPCList(NPCIndex).Flags.ActionDelay > 0 Then
        Log "NPC_AttackUser: NPC action delay > 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Don't allow if switching maps
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Log "NPC_AttackUser: NPC switching maps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Don't allow if not logged in completely
    If UserList(UserIndex).Flags.UserLogged = 0 Then
        Log "NPC_AttackUser: User " & UserIndex & " (" & UserList(UserIndex).Name & ") not logged in - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Set the action delay
    NPCList(NPCIndex).Flags.ActionDelay = NPCDelayFight
    
    'Create the sound effect and make the attack grh
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte SOUND_SWING
    ConBuf.Put_Byte NPCList(NPCIndex).Pos.X
    ConBuf.Put_Byte NPCList(NPCIndex).Pos.Y
    If NPCList(NPCIndex).AttackGrh > 0 Then
        If NPCList(NPCIndex).AttackRange > 1 Then
            ConBuf.Put_Byte DataCode.Server_MakeProjectile
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            ConBuf.Put_Long NPCList(NPCIndex).AttackGrh
            ConBuf.Put_Byte NPCList(NPCIndex).ProjectileRotateSpeed
        Else
            ConBuf.Put_Byte DataCode.Server_MakeSlash
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            ConBuf.Put_Long NPCList(NPCIndex).AttackGrh
        End If
    End If
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

    'Calculate if they hit
    If Server_RandomNumber(1, 100) >= ((NPCList(NPCIndex).ModStat(SID.WeaponSkill) + 50) - UserList(UserIndex).Stats.ModStat(SID.Agi)) Then
        Log "NPC_AttackUser: NPC's attack missed", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetCharDamage
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Integer -1
        Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
        Exit Sub
    End If

    'Calculate hit
    Hit = Server_RandomNumber(NPCList(NPCIndex).ModStat(SID.MinHIT), NPCList(NPCIndex).ModStat(SID.MaxHIT))
    Hit = Hit - (UserList(UserIndex).Stats.ModStat(SID.DEF) \ 2)
    If Hit < 1 Then Hit = 1
    Log "NPC_AttackUser: Hit value = " & Hit, CodeTracker '//\\LOGLINE//\\

    'Hit user
    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = UserList(UserIndex).Stats.BaseStat(SID.MinHP) - Hit

    'Display damage
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_SetCharDamage
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Integer Hit
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

    'User Die
    If UserList(UserIndex).Stats.BaseStat(SID.MinHP) <= 0 Then
        Log "NPC_AttackUser: NPC's attack killed user", CodeTracker '//\\LOGLINE//\\
    
        'Kill user
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 73
        ConBuf.Put_String NPCList(NPCIndex).Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        User_Kill UserIndex
        
    End If

End Sub

Sub NPC_ChangeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal NPCIndex As Integer, Optional ByVal Body As Integer = -1, Optional ByVal Head As Integer = -1, Optional ByVal Heading As Byte = 0, Optional ByVal Weapon As Integer = -1, Optional ByVal Hair As Integer = -1, Optional ByVal Wings As Integer = -1)

'*****************************************************************
'Changes a NPC char's head,body and heading
'*****************************************************************
Dim ChangeFlags As Byte

    Log "Call NPC_ChangeChar(" & sndRoute & "," & sndIndex & "," & NPCIndex & "," & Body & "," & Head & "," & Heading & "," & Weapon & "," & Hair & "," & Wings & ")", CodeTracker '//\\LOGLINE//\\

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
            If .Body <> Body Then .Body = Body
            ChangeFlags = ChangeFlags Or 1
        End If
        If Head > -1 Then
            If .Head <> Head Then .Head = Head
            ChangeFlags = ChangeFlags Or 2
        End If
        If Heading > 0 Then
            If .Heading <> Heading Then .Heading = Heading
            ChangeFlags = ChangeFlags Or 4
        End If
        If Weapon > -1 Then
            If .Weapon <> Weapon Then .Weapon = Weapon
            ChangeFlags = ChangeFlags Or 8
        End If
        If Hair > -1 Then
            If .Hair <> Hair Then .Hair = Hair
            ChangeFlags = ChangeFlags Or 16
        End If
        If Wings > -1 Then
            If .Wings <> Wings Then .Wings = Wings
            ChangeFlags = ChangeFlags Or 32
        End If
    End With
    
    'Make sure there is a packet to send
    If ChangeFlags = 0 Then Exit Sub
    
    'Create the packet
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_ChangeChar
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    ConBuf.Put_Byte ChangeFlags
    If ChangeFlags And 1 Then ConBuf.Put_Integer Body
    If ChangeFlags And 2 Then ConBuf.Put_Integer Head
    If ChangeFlags And 4 Then ConBuf.Put_Byte Heading
    If ChangeFlags And 8 Then ConBuf.Put_Integer Weapon
    If ChangeFlags And 16 Then ConBuf.Put_Integer Hair
    If ChangeFlags And 32 Then ConBuf.Put_Integer Wings
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map, PP_ChangeChar

End Sub

Sub NPC_Close(ByVal NPCIndex As Integer)

'*****************************************************************
'Closes a NPC
'*****************************************************************

    Log "Call NPC_Close(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    NPCList(NPCIndex).Flags.NPCActive = 0

    'Update LastNPC
    Log "NPC_Close: Updating LastNPC", CodeTracker '//\\LOGLINE//\\
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
    Log "NPC_Close: NumNPCs = " & NumNPCs, CodeTracker '//\\LOGLINE//\\

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
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_CharHP
            ConBuf.Put_Byte HPB
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
        End If
    End If

End Sub

Public Sub NPC_Damage(NPCIndex As Integer, UserIndex As Integer, Damage As Integer)

'*****************************************************************
'Do damage to a NPC - ONLY USE THIS SUB TO HURT NPCS
'*****************************************************************
Dim NewSlot As Byte
Dim NewX As Byte
Dim NewY As Byte
Dim HPA As Byte         'HP percentage before reducing hp
Dim HPB As Byte         'HP percentage after reducing hp
Dim i As Integer

    Log "Call NPC_Damage(" & NPCIndex & "," & UserIndex & "," & Damage & ")", CodeTracker '//\\LOGLINE//\\

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

    'Get the pre-damage percentage
    HPA = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)

    'Lower the NPC's life
    NPCList(NPCIndex).BaseStat(SID.MinHP) = NPCList(NPCIndex).BaseStat(SID.MinHP) - Damage

    'Check to update health percentage client-side
    If NPCList(NPCIndex).BaseStat(SID.MinHP) > 0 Then
        HPB = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
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
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 74
                            ConBuf.Put_Integer UserList(UserIndex).QuestStatus(i).NPCKills
                            ConBuf.Put_Integer QuestData(UserList(UserIndex).Quest(i)).FinishReqNPCAmount
                            
                            'Get the NPC's name from the database
                            DB_RS.Open "SELECT name FROM npcs WHERE `id`='" & QuestData(UserList(UserIndex).Quest(i)).FinishReqNPC & "'", DB_Conn, adOpenStatic, adLockOptimistic
                            ConBuf.Put_String DB_RS!Name
                            DB_RS.Close
                            
                            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                            
                        End If

                    End If
                End If
            Next i

            'Give EXP and gold
            User_RaiseExp UserIndex, NPCList(NPCIndex).GiveEXP
            UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + NPCList(NPCIndex).GiveGLD

            'Display kill message to the user
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 75
            ConBuf.Put_String NPCList(NPCIndex).Name
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            
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
    
    Log "Call NPC_Kill(" & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Set health back to 100%
    NPCList(NPCIndex).BaseStat(SID.MinHP) = NPCList(NPCIndex).ModStat(SID.MaxHP)

    'Erase it from map
    NPC_EraseChar NPCIndex

    'Set death time for respawn wait
    NPCList(NPCIndex).Counters.RespawnCounter = timeGetTime

End Sub

Sub NPC_MakeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal NPCIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Makes and places a NPC character
'*****************************************************************
Dim SndHP As Byte
Dim SndMP As Byte

    Log "Call NPC_MakeChar(" & sndRoute & "," & sndIndex & "," & NPCIndex & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

'Place character on map

    MapData(Map, X, Y).NPCIndex = NPCIndex

    'Set alive flag
    NPCList(NPCIndex).Flags.NPCAlive = 1

    'Set the hp/mp to send
    If NPCList(NPCIndex).ModStat(SID.MaxHP) > 0 Then SndHP = CByte((NPCList(NPCIndex).BaseStat(SID.MinHP) / NPCList(NPCIndex).ModStat(SID.MaxHP)) * 100)
    If NPCList(NPCIndex).ModStat(SID.MaxMAN) > 0 Then SndMP = CByte((NPCList(NPCIndex).BaseStat(SID.MinMAN) / NPCList(NPCIndex).ModStat(SID.MaxMAN)) * 100)

    'Send make character command to clients
    ConBuf.Clear
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

    'NPCs wont be created with active spells
    ZeroMemory NPCList(NPCIndex).Skills, Len(NPCList(NPCIndex).Skills)

    'Send the NPC
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, Map

End Sub

Function NPC_MoveChar(ByVal NPCIndex As Integer, ByVal nHeading As Byte) As Byte

'*****************************************************************
'Moves a NPC from one tile to another
'*****************************************************************

Dim nPos As WorldPos

    Log "Call NPC_MoveChar(" & NPCIndex & "," & nHeading & ")", CodeTracker '//\\LOGLINE//\\
    
    'Move
    nPos = NPCList(NPCIndex).Pos
    Server_HeadToPos nHeading, nPos

    'Move if legal pos
    If Server_LegalPos(NPCList(NPCIndex).Pos.Map, nPos.X, nPos.Y, nHeading) = True Then

        'Set the move delay (we set the lag buffer to 0 since NPCs don't lag)
        NPCList(NPCIndex).Flags.ActionDelay = Server_WalkTimePerTile(NPCList(NPCIndex).ModStat(SID.Speed), 0)

        'Send the movement update packet
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MoveChar
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Byte nPos.X
        ConBuf.Put_Byte nPos.Y
        ConBuf.Put_Byte nHeading
        Data_Send ToNPCMove, NPCIndex, ConBuf.Get_Buffer

        'Update map and user pos
        MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = 0
        NPCList(NPCIndex).Pos = nPos
        NPCList(NPCIndex).Char.Heading = nHeading
        NPCList(NPCIndex).Char.HeadHeading = nHeading
        MapData(NPCList(NPCIndex).Pos.Map, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y).NPCIndex = NPCIndex
        
        'NPC moved, return the flag
        NPC_MoveChar = 1

    End If
    
    Log "Rtrn NPC_MoveChar = " & NPC_MoveChar, CodeTracker '//\\LOGLINE//\\

End Function

Function NPC_NextOpen() As Integer

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
    Loop While NPCList(LoopC).Flags.NPCActive = 1

    NPC_NextOpen = LoopC
    
    Log "Rtrn NPC_NextOpen = " & NPC_NextOpen, CodeTracker '//\\LOGLINE//\\

End Function

Sub NPC_Spawn(ByVal NPCIndex As Integer, Optional ByVal BypassUpdate As Byte = 0)

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
    NPCList(NPCIndex).Flags.NPCAlive = 1

    'Make NPC Char
    If Not BypassUpdate Then
        If UBound(MapUsers(TempPos.Map).Index) > 0 Then
            NPC_MakeChar ToMap, MapUsers(TempPos.Map).Index(1), NPCIndex, TempPos.Map, TempPos.X, TempPos.Y
        End If
    End If
    
End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:46)  Decl: 13  Code: 684  Total: 697 Lines
':) CommentOnly: 126 (18.1%)  Commented: 25 (3.6%)  Empty: 166 (23.8%)  Max Logic Depth: 13
