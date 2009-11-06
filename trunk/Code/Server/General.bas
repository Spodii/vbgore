Attribute VB_Name = "General"
Option Explicit

'How much time between the server loops - this is to let some slack on the CPU as to not overwork it
' The server will stop sleeping if the elapsed time for the loop is > this value. It is suggested
' you don't change this value lower than 5 (unless you hate your server computer and want it to die).
Private Const GameLoopTime As Currency = 15

'Adjust these values accordingly depending on how often you want routines to update
'Low values = faster updating (smoother gameplay), but more CPU usage
Private Const UpdateRate_UserStats As Currency = 400        'Updating user stats on the client
Private Const UpdateRate_UserRecover As Currency = 3000     'Recovering the user's stats (HP, MP, etc)
Private Const UpdateRate_UserCounters As Currency = 200     'Updating user counters (aggressive face, spells, exhaustion, etc)
Private Const UpdateRate_UserSendBuffer As Currency = 50    'Check to send the user's buffer
Private Const UpdateRate_NPCAI As Currency = 50             'Updating NPC AI
Private Const UpdateRate_NPCCounters As Currency = 200      'Updating NPC counters
Private Const UpdateRate_Maps As Currency = 30000           'Updating map ground objects / unloading maps from memory
Private Const UpdateRate_Bandwidth As Currency = 1000       'Updating bandwidth in/out information

Private LastUpdate_UserStats As Currency
Private LastUpdate_UserRecover As Currency
Private LastUpdate_UserCounters As Currency
Private LastUpdate_UserSendBuffer As Currency
Private LastUpdate_NPCAI As Currency
Private LastUpdate_NPCCounters As Currency
Private LastUpdate_Maps As Currency
Private LastUpdate_Bandwidth As Currency

'To save excessive looping, flags are set to go with the next loop instead of a loop in their own
Private UpdateUserStats As Byte     'If the user stats will update
Private RecoverUserStats As Byte    'If the user stats will recover
Private UpdateUserCounters As Byte  'If the user counters will be updated
Private SendUserBuffer As Byte      'If the user's buffer will be checked to be sent
Private UpdateNPCAI As Byte         'Call the NPC AI routine
Private UpdateNPCCounters As Byte   'Update the NPC's counters

'Sleep API - used to "sleep" the process and free the CPU usage
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Server_Update()

'*****************************************************************
'Primary update unit - looks for routines to update
'*****************************************************************
Dim UpdateUsers As Byte 'We only update users if one of the user counters go off
Dim UpdateNPCs As Byte  'Same as above, but with NPCs
Dim LoopStartTime As Currency   'Time at the start of the loop (to find the elapsed time)
Dim Elapsed As Currency         'Time elapsed through the loop
    
    'Set the server as running
    ServerRunning = 1

    'Loop until ServerRunning = 0
    Do While ServerRunning
    
        'Get the start time so we know how long the loop took
        LoopStartTime = CurrentTime

        '*** Check for updating flags ***
        
        'User stats (updating client-side view)
        If LastUpdate_UserStats + UpdateRate_UserStats < CurrentTime Then
            UpdateUserStats = 1
            LastUpdate_UserStats = CurrentTime
            UpdateUsers = 1
        End If
        
        'User stat recovery (raising HP, MP, SP, etc)
        If LastUpdate_UserRecover + UpdateRate_UserRecover < CurrentTime Then
            RecoverUserStats = 1
            LastUpdate_UserRecover = CurrentTime
            UpdateUsers = 1
        End If
        
        'User counters (aggressive face, spells, spell exhaustion, etc)
        If LastUpdate_UserCounters + UpdateRate_UserCounters < CurrentTime Then
            UpdateUserCounters = 1
            LastUpdate_UserCounters = CurrentTime
            UpdateUsers = 1
        End If
        
        'Sending the packet buffer
        If LastUpdate_UserSendBuffer + UpdateRate_UserSendBuffer < CurrentTime Then
            SendUserBuffer = 1
            LastUpdate_UserSendBuffer = CurrentTime
            UpdateUsers = 1
        End If
        
        'NPC AI
        If LastUpdate_NPCAI + UpdateRate_NPCAI < CurrentTime Then
            UpdateNPCAI = 1
            LastUpdate_NPCAI = CurrentTime
            UpdateNPCs = 1
        End If
        
        'NPC counters
        If LastUpdate_NPCCounters + UpdateRate_NPCCounters < CurrentTime Then
            UpdateNPCCounters = 1
            LastUpdate_NPCCounters = CurrentTime
            UpdateNPCs = 1
        End If
        
        '*** Check for actual updating routines ***
        
        'Update users if one of the flags have gone off
        If UpdateUsers Then Server_Update_Users
        
        'General NPC information
        If UpdateNPCs Then Server_Update_NPCs
        
        'Map updating
        If LastUpdate_Maps + UpdateRate_Maps < CurrentTime Then
            Server_Update_Maps
            LastUpdate_Maps = CurrentTime
        End If
        
        'Bandwidth report updating
        If CalcTraffic Then
            If LastUpdate_Bandwidth + UpdateRate_Bandwidth < CurrentTime Then
                LastUpdate_Bandwidth = CurrentTime
                Server_Update_Bandwidth
            End If
        End If

        '*** Cooldown ***
        
        'Let other events happen (this is required for the socket to get packets, so don't try removing it to save time)
        DoEvents
        
        'Check if we have enough time to sleep
        Elapsed = CurrentTime - LoopStartTime
        If Elapsed < GameLoopTime Then
            If Elapsed >= 0 Then    'Make sure nothing weird happens, causing for a huge sleep time
                Sleep Int(GameLoopTime - Elapsed)
            End If
        End If
        
    Loop
        
End Sub

Private Sub Server_Update_Bandwidth()

'*****************************************************************
'Updates the bandwidth usage variables
'*****************************************************************

    'Turn bytes into kilobytes
    If DataIn > 1024 Then
        Do While DataIn > 1024
            DataIn = DataIn - 1024
            DataKBIn = DataKBIn + 1
        Loop
    End If

    If DataOut > 1024 Then
        Do While DataOut > 1024
            DataOut = DataOut - 1024
            DataKBOut = DataKBOut + 1
        Loop
    End If

    'Ignore any errors (since they're most likely due to strange times with CurrentTime)
    On Error Resume Next
    
        'Display statistics (KB)
        frmMain.BytesInTxt.Text = Round((DataKBIn + (DataIn / 1024)) / ((CurrentTime - ServerStartTime) * 0.001), 6)
        frmMain.BytesOutTxt.Text = Round((DataKBOut + (DataOut / 1024)) / ((CurrentTime - ServerStartTime) * 0.001), 6)
    
        'Display statistics (Bytes)
        'frmMain.BytesInTxt.Text = Round(((DataKBIn * 1024) + DataIn) / ((CurrentTime - ServerStartTime) / 1000), 6)
        'frmMain.BytesOutTxt.Text = Round(((DataKBOut * 1024) + DataOut) / ((CurrentTime - ServerStartTime) / 1000), 6)

    On Error GoTo 0

End Sub

Private Sub Server_Update_NPCs()

'*****************************************************************
'Updates the NPCs
'*****************************************************************
Dim NPCIndex As Integer

    'Update NPCs
    For NPCIndex = 1 To LastNPC

        'Make sure NPC is active
        If NPCList(NPCIndex).Flags.NPCActive Then

            'See if npc is alive
            If NPCList(NPCIndex).Flags.NPCAlive Then

                'Only update npcs in user populated maps
                If MapInfo(NPCList(NPCIndex).Pos.Map).NumUsers Then
                
                    'Check to update mod stats
                    If NPCList(NPCIndex).Flags.UpdateStats Then
                        NPCList(NPCIndex).Flags.UpdateStats = 0
                        NPC_UpdateModStats NPCIndex
                    End If
                    
                    '*** Update counters ***
                    If UpdateNPCCounters Then   'Update aggressive-face timer
                        If NPCList(NPCIndex).Counters.AggressiveCounter > 0 Then
                            If NPCList(NPCIndex).Counters.AggressiveCounter < CurrentTime Then
                                NPCList(NPCIndex).Counters.AggressiveCounter = 0
                                ConBuf.Clear
                                ConBuf.Put_Byte DataCode.User_AggressiveFace
                                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                ConBuf.Put_Byte 0
                                Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            End If
                        End If                  'Update warcurse time
                        If NPCList(NPCIndex).Skills.WarCurse > 0 Then
                            If NPCList(NPCIndex).Counters.WarCurseCounter < CurrentTime Then
                                NPCList(NPCIndex).Counters.WarCurseCounter = 0
                                NPCList(NPCIndex).Skills.WarCurse = 0
                                ConBuf.Clear
                                ConBuf.Put_Byte DataCode.Server_Message
                                ConBuf.Put_Byte 1
                                ConBuf.Put_String NPCList(NPCIndex).Name
                                Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
                                ConBuf.Clear
                                ConBuf.Put_Byte DataCode.Server_IconWarCursed
                                ConBuf.Put_Byte 0
                                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            End If
                        End If                  'Update spell exhaustion
                        If NPCList(NPCIndex).Counters.SpellExhaustion > 0 Then
                            If NPCList(NPCIndex).Counters.SpellExhaustion < CurrentTime Then
                                NPCList(NPCIndex).Counters.SpellExhaustion = 0
                                ConBuf.Clear
                                ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
                                ConBuf.Put_Byte 0
                                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            End If
                        End If
                    End If

                    '*** NPC AI ***
                    If UpdateNPCAI Then
                        If NPCList(NPCIndex).Counters.ActionDelay < CurrentTime Then NPC_AI NPCIndex
                    End If

                End If

            Else
                
                '*** Respawn NPC ***
                'Check if it's time to respawn
                If NPCList(NPCIndex).Counters.RespawnCounter < CurrentTime Then NPC_Spawn NPCIndex

            End If
            
        End If
        
    Next NPCIndex
    
    'Clear the update flags
    UpdateNPCAI = 0
    UpdateNPCCounters = 0

End Sub

Private Sub Server_Update_Users()

'*****************************************************************
'Updates the users
'*****************************************************************
Dim UserIndex As Integer

    'Loop through all the users
    For UserIndex = 1 To LastUser

        'Make sure user is logged on
        If UserList(UserIndex).Flags.UserLogged Then

            '*** Disconnection timers ***
            'Check if it has been idle for too long
            If UserList(UserIndex).Counters.IdleCount <= CurrentTime - IdleLimit Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 85
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Server_CloseSocket UserIndex
                GoTo NextUser   'Skip to the next user
            End If
            
            'Check if the user was possible disconnected (or extremely laggy)
            If UserList(UserIndex).Counters.LastPacket <= CurrentTime - LastPacket Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 85
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Server_CloseSocket UserIndex
                GoTo NextUser   'Skip to the next user
            End If
            
            '*** Recover stats ***
            If RecoverUserStats Then    'HP
                If UserList(UserIndex).Stats.BaseStat(SID.MinHP) < UserList(UserIndex).Stats.ModStat(SID.MaxHP) Then
                    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = UserList(UserIndex).Stats.BaseStat(SID.MinHP) + 1 + UserList(UserIndex).Stats.ModStat(SID.Str) * 0.5
                End If                  'SP
                If UserList(UserIndex).Stats.BaseStat(SID.MinSTA) < UserList(UserIndex).Stats.ModStat(SID.MaxSTA) Then
                    UserList(UserIndex).Stats.BaseStat(SID.MinSTA) = UserList(UserIndex).Stats.BaseStat(SID.MinSTA) + 1 + UserList(UserIndex).Stats.ModStat(SID.Agi) * 0.5
                End If                  'MP
                If UserList(UserIndex).Stats.BaseStat(SID.MinMAN) < UserList(UserIndex).Stats.ModStat(SID.MaxMAN) Then
                    UserList(UserIndex).Stats.BaseStat(SID.MinMAN) = UserList(UserIndex).Stats.BaseStat(SID.MinMAN) + 1 + UserList(UserIndex).Stats.ModStat(SID.Mag) * 0.5
                End If
            End If

            '*** Update the counters ***
            If UpdateUserCounters Then  'Bless
                If UserList(UserIndex).Counters.BlessCounter > 0 Then
                    If UserList(UserIndex).Counters.BlessCounter < CurrentTime Then
                        UserList(UserIndex).Skills.Bless = 0
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_IconBlessed
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Protection
                If UserList(UserIndex).Counters.ProtectCounter > 0 Then
                    If UserList(UserIndex).Counters.ProtectCounter < CurrentTime Then
                        UserList(UserIndex).Skills.Protect = 0
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_IconProtected
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Strengthen
                If UserList(UserIndex).Counters.StrengthenCounter > 0 Then
                    If UserList(UserIndex).Counters.StrengthenCounter < CurrentTime Then
                        UserList(UserIndex).Skills.Strengthen = 0
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_IconStrengthened
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Spell exhaustion
                If UserList(UserIndex).Counters.SpellExhaustion > 0 Then
                    If UserList(UserIndex).Counters.SpellExhaustion < CurrentTime Then
                        UserList(UserIndex).Counters.SpellExhaustion = 0
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Aggressive face
                If UserList(UserIndex).Counters.AggressiveCounter > 0 Then
                    If UserList(UserIndex).Counters.AggressiveCounter < CurrentTime Then
                        UserList(UserIndex).Counters.AggressiveCounter = 0
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.User_AggressiveFace
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        ConBuf.Put_Byte 0
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If
            End If
            
            '*** Send queued packet buffer ***
            If SendUserBuffer Then

                'Check if the packet wait time has passed
                If UserList(UserIndex).HasBuffer Then
                    If UserList(UserIndex).PacketWait < CurrentTime Then
    
                        'Send the packet buffer to the user
                        If UserList(UserIndex).PPValue = PP_High Then
                            
                            'High priority - send asap
                            Data_Send_Buffer UserIndex
                            
                        ElseIf UserList(UserIndex).PPValue = PP_Low Then
                            
                            'Low priority - check counter for sending
                            If UserList(UserIndex).PPCount < CurrentTime Then Data_Send_Buffer UserIndex
                        
                        End If
                        
                    End If
                End If
                
            End If
            
            '*** Update user stats (on client-side) ***
            If UpdateUserStats Then UserList(UserIndex).Stats.SendUpdatedStats
            
        End If

NextUser:

    Next UserIndex
    
    'Clear the update flags
    UpdateUserStats = 0
    RecoverUserStats = 0
    UpdateUserCounters = 0
    SendUserBuffer = 0

End Sub

Private Sub Server_Update_Maps()

'*****************************************************************
'Updates all the maps (removes objects / unloads maps from memory)
'*****************************************************************
Dim ObjIndex As Byte    'Slot of the object on the tile
Dim MapIndex As Long    'Index of the map being looped through
Dim X As Byte   'Co-ordinates of the tile being checked
Dim Y As Byte

    'Loop through all the maps
    For MapIndex = 1 To NumMaps
        
        'Make sure the map is in use before checking
        If MapInfo(MapIndex).NumUsers > 0 Then
                
            'The map has users on it, so check through the tiles in-bounds
            For X = MinXBorder To MaxXBorder
                For Y = MinYBorder To MaxYBorder
                    
                    '*** Removing old objects ***
                    'Check if an object exists on the tile - loop through all on there
                    If MapInfo(MapIndex).ObjTile(X, Y).NumObjs > 0 Then
                        For ObjIndex = 1 To MapInfo(MapIndex).ObjTile(X, Y).NumObjs
                            
                            'Check if it is time to remove the object
                            If MapInfo(MapIndex).ObjTile(X, Y).ObjLife(ObjIndex) < CurrentTime - GroundObjLife Then
                                Obj_Erase MapInfo(MapIndex).ObjTile(X, Y).ObjInfo(ObjIndex).Amount, ObjIndex, MapIndex, X, Y
                            End If
                            
                        Next ObjIndex
                    End If
                    
                Next Y
            Next X
            
        Else
            
            '*** Unloading maps from memory ***
            'The map is empty, check if it is being counted down to being unloaded
            If MapInfo(MapIndex).UnloadTimer > 0 Then Unload_Map MapIndex
        
        End If
        
    Next MapIndex

End Sub

Private Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte

'*****************************************************************
'Check if two lines intersect (return 1 if true)
'*****************************************************************

Dim m1 As Single
Dim M2 As Single
Dim B1 As Single
Dim B2 As Single
Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    B1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    B2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If B2 = B1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((B2 - B1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
        
    End If
    
End Function

Public Function Engine_ClearPath(ByVal Map As Integer, ByVal CharX As Long, ByVal CharY As Long, ByVal TargetX As Long, ByVal TargetY As Long) As Byte

'***************************************************
'Check if the path is clear from the user to the target of blocked tiles
'For the line-rect collision, we pretend that each tile is 2 units wide so we can give them a width of 1 to center things
'***************************************************
Dim X As Long
Dim Y As Long

    '****************************************
    '***** Target is on top of the user *****
    '****************************************
    
    'If the target position = user position, we must be targeting ourself, so nothing can be blocking us from us (I hope o.O )
    If CharX = TargetX Then
        If CharY = TargetY Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If

    '********************************************
    '***** Target is right next to the user *****
    '********************************************
    
    'Target is at one of the 4 diagonals of the user
    If Abs(CharX - TargetX) = 1 Then
        If Abs(CharY - TargetY) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Target is above or below the user
    If CharX = TargetX Then
        If Abs(CharY - TargetY) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Target is to the left or right of the user
    If CharY = TargetY Then
        If Abs(CharX - TargetX) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    '********************************************
    '***** Target is diagonal from the user *****
    '********************************************
    
    'Check if the target is diagonal from the user - only do the following checks if diagonal from the target
    If Abs(CharX - TargetX) = Abs(CharY - TargetY) Then

        If CharX > TargetX Then
                        
            'Diagonal to the top-left
            If CharY > TargetY Then
                For X = TargetX To CharX - 1
                    For Y = TargetY To CharY - 1
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            
            'Diagonal to the bottom-left
            Else
                For X = TargetX To CharX - 1
                    For Y = CharY + 1 To TargetY
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            End If

        End If
        
        If CharX < TargetX Then
        
            'Diagonal to the top-right
            If CharY > TargetY Then
                For X = CharX + 1 To TargetX
                    For Y = TargetY To CharY - 1
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
                
            'Diagonal to the bottom-right
            Else
                For X = CharX + 1 To TargetX
                    For Y = CharY + 1 To TargetY
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            End If
        
        End If
    
        Engine_ClearPath = 1
        Exit Function
    
    End If

    '*******************************************************************
    '***** Target is directly vertical or horizontal from the user *****
    '*******************************************************************
    
    'Check if target is directly above the user
    If CharX = TargetX Then 'Check if x values are the same (straight line between the two)
        If CharY > TargetY Then
            For Y = TargetY + 1 To CharY - 1
                If MapInfo(Map).Data(CharX, Y).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next Y
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly below the user
    If CharX = TargetX Then
        If CharY < TargetY Then
            For Y = CharY + 1 To TargetY - 1
                If MapInfo(Map).Data(CharX, Y).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next Y
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly to the left of the user
    If CharY = TargetY Then
        If CharX > TargetX Then
            For X = TargetX + 1 To CharX - 1
                If MapInfo(Map).Data(X, CharY).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next X
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly to the right of the user
    If CharY = TargetY Then
        If CharX < TargetX Then
            For X = CharX + 1 To TargetX - 1
                If MapInfo(Map).Data(X, CharY).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next X
            Engine_ClearPath = 1
            Exit Function
        End If
    End If

    '*******************************************************************
    '***** Target is directly vertical or horizontal from the user *****
    '*******************************************************************
    
    
    If CharY > TargetY Then
    
        'Check if the target is to the top-left of the user
        If CharX > TargetX Then
            For X = TargetX To CharX
                For Y = TargetY To CharY
                    'We must do * 2 on the tiles so we can use +1 to get the center (its like * 32 and + 16 - this does the same affect)
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, CharX * 2 + 1, CharY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
            Engine_ClearPath = 1
            Exit Function
    
        'Check if the target is to the top-right of the user
        Else
            For X = CharX To TargetX
                For Y = TargetY To CharY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, CharX * 2 + 1, CharY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        End If
        
    Else
    
        'Check if the target is to the bottom-left of the user
        If CharX > TargetX Then
            For X = TargetX To CharX
                For Y = CharY To TargetY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, CharX * 2 + 1, CharY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        
        'Check if the target is to the bottom-right of the user
        Else
            For X = CharX To TargetX
                For Y = CharY To TargetY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, CharX * 2 + 1, CharY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapInfo(Map).Data(X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        End If
    
    End If
    
    Engine_ClearPath = 1

End Function

Private Function Engine_Collision_LineRect(ByVal SX As Long, ByVal SY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Byte

'*****************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'*****************************************************************

    'Top line
    If Engine_Collision_Line(SX, SY, SX + SW, SY, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(SX + SW, SY, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(SX, SY + SH, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(SX, SY, SX, SY + SW, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

End Function

Private Function Engine_Collision_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte

'*****************************************************************
'Find if a value is between two other values (used for line collision)
'*****************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If Value >= Bound1 Then
            If Value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
End Function

Public Function ByteArrayToStr(ByRef ByteArray() As Byte) As String

'*****************************************************************
'Take a byte array and print it out in a readable string
'Example output: 084[T] 086[V] 088[X] 090[Z] 092[\] 094[^]
'*****************************************************************

On Error GoTo ErrOut

Dim Char As String
Dim i As Long
    
    Log "ByteArrayToStr: ByteArray LBound() = " & LBound(ByteArray) & " UBound() = " & UBound(ByteArray), CodeTracker '//\\LOGLINE//\\
    For i = LBound(ByteArray) To UBound(ByteArray)
        If ByteArray(i) > 32 Then Char = Chr$(ByteArray(i)) Else Char = " "
        If ByteArray(i) >= 100 Then
            ByteArrayToStr = ByteArrayToStr & ByteArray(i) & "[" & Char & "] "
        ElseIf ByteArray(i) >= 10 Then
            ByteArrayToStr = ByteArrayToStr & "0" & ByteArray(i) & "[" & Char & "] "
        Else
            ByteArrayToStr = ByteArrayToStr & "00" & ByteArray(i) & "[" & Char & "] "
        End If
    Next i
    ByteArrayToStr = Left$(ByteArrayToStr, Len(ByteArrayToStr) - 1)
    
'If there was an error, we were probably passed an erased ByteArray
ErrOut:

    Log "ByteArrayToStr: Unknown error in routine!", CriticalError '//\\LOGLINE//\\
    
End Function

Public Function Server_WalkTimePerTile(ByVal Speed As Long, Optional ByVal LagBuffer As Integer = 350) As Long
'*****************************************************************
'Takes a speed value and returns the time it takes to walk a tile
'To fine the value:
'(Speed + 4) * BaseWalkSpeed = Pixels/second
'Pixels/sec / 32 = Tiles/sec
'1000 / Tiles/sec = Seconds per tile - how long it takes to walk by one tile
'*****************************************************************

    Log "Call Server_WalkTimePerTile(" & Speed & ")", CodeTracker '//\\LOGLINE//\\

    '4 = The client works off a base value of 4 for speed, so the speed is calculated as 4 + Speed in the client
    '11 = BaseWalkSpeed - how fast we move in pixels/sec
    '32 = The size of a tile
    '1000 = Miliseconds in a second
    'LagBuffer = We have to give some slack for network lag and client lag - raise this value if people skip too much
    '     and lower it if people are speedhacking and getting too much extra speed
    Server_WalkTimePerTile = 1000 / (((Speed + 4) * 11) / 32) - LagBuffer
    
    'Make sure the lag buffer doesn't overshoot the value into the negatives
    If Server_WalkTimePerTile < 0 Then Server_WalkTimePerTile = 0
    
    Log "Rtrn Server_WalkTimePerSecond = " & Server_WalkTimePerTile, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_UserExist(ByVal UserName As String) As Boolean
'*****************************************************************
'Checks the database for if a user exists by the specified name
'*****************************************************************

    Log "Call Server_UserExist(" & UserName & ")", CodeTracker '//\\LOGLINE//\\

    'Make the query
    DB_RS.Open "SELECT name FROM users WHERE `name`='" & UserName & "'", DB_Conn, adOpenStatic, adLockOptimistic

    'If End Of File = true, then the user doesn't exist
    If DB_RS.EOF = True Then Server_UserExist = False Else Server_UserExist = True

    'Close the recordset
    DB_RS.Close
    
    Log "Rtrn Server_UserExist = " & Server_UserExist, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_LegalString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for illegal characters in the string (string wrapper for Server_LegalCharacter)
'*****************************************************************
Dim b() As Byte
Dim i As Long

    Log "Call Server_LegalString(" & CheckString & ")", CodeTracker '//\\LOGLINE//\\

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then
        Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If LenB(CheckString) < 1 Then
        Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Copy the string to a byte array
    b() = StrConv(CheckString, vbFromUnicode)

    'Loop through the string
    For i = 0 To UBound(b)
        
        'Check the values
        If Server_LegalCharacter(b(i)) = False Then
            Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
        
    Next i
    
    'If we have made it this far, then all is good
    Server_LegalString = True
    
    Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Server_LegalString = False
    
    Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_ValidString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for valid characters in the string (string wrapper for Server_ValidCharacter)
'Make sure to update on the client, too!
'*****************************************************************
Dim b() As Byte
Dim i As Long

    Log "Call Server_ValidString(" & CheckString & ")", CodeTracker '//\\LOGLINE//\\

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then
        Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If LenB(CheckString) < 1 Then
        Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Copy the string to a byte array
    b() = StrConv(CheckString, vbFromUnicode)

    'Loop through the string
    For i = 0 To UBound(b)
        
        'Check the values
        If Not Server_ValidCharacter(b(i)) Then
            Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
        
    Next i
    
    'If we have made it this far, then all is good
    Server_ValidString = True
    
    Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Server_ValidString = False
    
    Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\

End Function

Private Function Server_ValidCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is used for chat/etc)
'Make sure you update the client's Game_ValidCharacter, too!
'*****************************************************************

    Log "Call Server_ValidCharacter(" & KeyAscii & ")", CodeTracker '//\\LOGLINE//\\

    If KeyAscii >= 32 Then Server_ValidCharacter = True

End Function

Public Function Server_LegalCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is for username/pass)
'Make sure you update the client's Game_LegalCharacter, too!
'*****************************************************************

    Log "Call Server_LegalCharacter(" & KeyAscii & ")", CodeTracker '//\\LOGLINE//\\
    
    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 Then
        If KeyAscii <= 57 Then
            Server_LegalCharacter = True
            Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If
    
    'Allow letters A to Z
    If KeyAscii >= 65 Then
        If KeyAscii <= 90 Then
            Server_LegalCharacter = True
            Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If
    
    'Allow letters a to z
    If KeyAscii >= 97 Then
        If KeyAscii <= 122 Then
            Server_LegalCharacter = True
            Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If
    
    'Allow foreign characters
    If KeyAscii >= 128 Then
        If KeyAscii <= 168 Then
            Server_LegalCharacter = True
            Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If
    
    Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
    
Exit Function

ErrOut:

    Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
    
End Function

Public Function Server_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single

'*****************************************************************
'Finds the distance between two points
'*****************************************************************

    Log "Call Server_Distance(" & x1 & "," & Y1 & "," & x2 & "," & Y2 & ")", CodeTracker '//\\LOGLINE//\\

    Server_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))
    
    Log "Rtrn Server_Distance = " & Server_Distance, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_RectDistance(ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal MaxXDist As Long, ByVal MaxYDist As Long) As Byte

'*****************************************************************
'Check if two tile points are in the same screen
'*****************************************************************

    Log "Call Server_RectDistance(" & x1 & "," & Y1 & "," & x2 & "," & Y2 & "," & MaxXDist & "," & MaxYDist & ")", CodeTracker '//\\LOGLINE//\\

    If Abs(x1 - x2) < MaxXDist + 1 Then
        If Abs(Y1 - Y2) < MaxYDist + 1 Then
            Server_RectDistance = True
        End If
    End If
    
    Log "Rtrn Server_RectDistance = " & Server_RectDistance, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
On Error GoTo ErrOut
    
    Log "Call Server_FileExist(" & File & "," & FileType & ")", CodeTracker '//\\LOGLINE//\\

    If Dir$(File, FileType) <> "" Then Server_FileExist = True
    
    Log "Rtrn Server_FileExist = " & Server_FileExist, CodeTracker '//\\LOGLINE//\\

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Server_FileExist = False
    Log "Rtrn Server_FileExist = " & Server_FileExist, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Integer

'*****************************************************************
'Find a Random number between a range
'*****************************************************************

    Log "Call Server_RandomNumber(" & LowerBound & "," & UpperBound & ")", CodeTracker '//\\LOGLINE//\\

    Server_RandomNumber = Fix((UpperBound - LowerBound + 1) * Rnd) + LowerBound
    
    Log "Rtrn Server_RandomNumber = " & Server_RandomNumber, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub Server_RefreshUserListBox()

'*****************************************************************
'Refreshes the User list box
'*****************************************************************

Dim LoopC As Long

    Log "Call Server_RefreshUserListBox", CodeTracker '//\\LOGLINE//\\

    If LastUser < 0 Then
        Log "Server_RefreshUserListBox: No users to update", CodeTracker '//\\LOGLINE//\\
        frmMain.Userslst.Clear
        Exit Sub
    End If

    frmMain.Userslst.Clear
    CurrConnections = 0
    Log "Server_RefreshUserListBox: Updating " & LastUser & " users", CodeTracker '//\\LOGLINE//\\
    For LoopC = 1 To LastUser
        If LenB(UserList(LoopC).Name) Then
            frmMain.Userslst.AddItem UserList(LoopC).Name
            CurrConnections = CurrConnections + 1
        End If
    Next LoopC
    TrayModify ToolTip, "Game Server: " & CurrConnections & " connections"

End Sub

Public Function CurrentTime() As Currency

'*****************************************************************
'Wrapper for returning the system time in the format of a function
'instead of a passing a variable as ByRef - much easier this way to use.
'*****************************************************************

    GetSystemTime CurrentTime

End Function
