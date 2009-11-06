Attribute VB_Name = "General"
Option Explicit

'How much time between the server loops - this is to let some slack on the CPU as to not overwork it
' The server will stop sleeping if the elapsed time for the loop is > this value. It is suggested
' you don't change this value lower than 5 (unless you hate your server computer and want it to die).
' Try to keep this value below the lowest common denominator of all the timers (in this case, 50).
Private Const GameLoopTime As Long = 10

'Adjust these values accordingly depending on how often you want routines to update
'Low values = faster updating (smoother gameplay), but more CPU usage
Private Const UpdateRate_UserStats As Long = 400        'Updating user stats on the client
Private Const UpdateRate_UserRecover As Long = 3000     'Recovering the user's stats (HP, MP, etc)
Private Const UpdateRate_UserCounters As Long = 200     'Updating user counters (aggressive face, spells, exhaustion, etc)
Private Const UpdateRate_UserSendBuffer As Long = 50    'Check to send the user's buffer
Private Const UpdateRate_NPCAI As Long = 50             'Updating NPC AI
Private Const UpdateRate_NPCCounters As Long = 200      'Updating NPC counters
Private Const UpdateRate_Maps As Long = 30000           'Updating map ground objects (to remove them) / unloading maps from memory
Private Const UpdateRate_Bandwidth As Long = 1000       'Updating bandwidth in/out information
Private Const UpdateRate_UnloadObjs As Long = 120000    'Unloading objects from memory

Private LastUpdate_UserStats As Long
Private LastUpdate_UserRecover As Long
Private LastUpdate_UserCounters As Long
Private LastUpdate_UserSendBuffer As Long
Private LastUpdate_NPCAI As Long
Private LastUpdate_NPCCounters As Long
Private LastUpdate_Maps As Long
Private LastUpdate_Bandwidth As Long
Private LastUpdate_ServerFPS As Long    'For DEBUG_MapFPS
Private LastUpdate_UnloadObjs As Long

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
Dim LoopStartTime As Long       'Time at the start of the update loop
Dim UpdateUsers As Byte         'We only update users if one of the user counters go off
Dim UpdateNPCs As Byte          'Same as above, but with NPCs
Dim Elapsed As Long             'Time elapsed through the loop
Dim FPS As Long                 'Used for DEBUG_MapFPS

    'Set the server as running
    ServerRunning = 1

    'Loop until ServerRunning = 0
    Do While ServerRunning
    
        'Make sure that the system's clock didn't reset (check the sub for more details)
        ValidateTime
        
        'Get the start time so we know how long the loop took
        LoopStartTime = timeGetTime

        '*** Unload ***
        'Note that we have to put this in the loop in case the socket fails to unload
        'The socket is going to fail to unload once if theres connections made to it
        'Check if we're unloading the server
        If UnloadServer Then
            
            'Close the server
            Server_Unload
            
        End If

        '*** Check for updating flags ***
        
        'User stats (updating client-side view)
        If LastUpdate_UserStats + UpdateRate_UserStats < LoopStartTime Then
            UpdateUserStats = 1
            LastUpdate_UserStats = LoopStartTime
            UpdateUsers = 1
        End If
        
        'User stat recovery (raising HP, MP, SP, etc)
        If LastUpdate_UserRecover + UpdateRate_UserRecover < LoopStartTime Then
            RecoverUserStats = 1
            LastUpdate_UserRecover = LoopStartTime
            UpdateUsers = 1
        End If
        
        'User counters (aggressive face, spells, spell exhaustion, etc)
        If LastUpdate_UserCounters + UpdateRate_UserCounters < LoopStartTime Then
            UpdateUserCounters = 1
            LastUpdate_UserCounters = LoopStartTime
            UpdateUsers = 1
        End If
        
        'Sending the packet buffer
        If LastUpdate_UserSendBuffer + UpdateRate_UserSendBuffer < LoopStartTime Then
            SendUserBuffer = 1
            LastUpdate_UserSendBuffer = LoopStartTime
            UpdateUsers = 1
        End If
        
        'NPC AI
        If LastUpdate_NPCAI + UpdateRate_NPCAI < LoopStartTime Then
            UpdateNPCAI = 1
            LastUpdate_NPCAI = LoopStartTime
            UpdateNPCs = 1
        End If
        
        'NPC counters
        If LastUpdate_NPCCounters + UpdateRate_NPCCounters < LoopStartTime Then
            UpdateNPCCounters = 1
            LastUpdate_NPCCounters = LoopStartTime
            UpdateNPCs = 1
        End If
        
        'Object unloading
        If LastUpdate_UnloadObjs + UpdateRate_UnloadObjs < LoopStartTime Then
            LastUpdate_UnloadObjs = LoopStartTime
            ObjData.CheckObjUnloading
        End If
        
        '*** Check for actual updating routines ***
        
        'Update users if one of the flags have gone off
        If UpdateUsers Then Server_Update_Users
        
        'General NPC information
        If UpdateNPCs Then Server_Update_NPCs
        
        'Map updating
        If LastUpdate_Maps + UpdateRate_Maps < LoopStartTime Then
            Server_Update_Maps
            LastUpdate_Maps = LoopStartTime
        End If
        
        'Bandwidth report updating
        If CalcTraffic Then
            If LastUpdate_Bandwidth + UpdateRate_Bandwidth < LoopStartTime Then
                LastUpdate_Bandwidth = LoopStartTime
                Server_Update_Bandwidth
            End If
        End If

        '*** Cooldown ***
        
        'Let other events happen (this is required for the socket to get packets, so don't try removing it to save time)
        DoEvents
        
        'Check if we have enough time to sleep
        Elapsed = timeGetTime - LoopStartTime
        If Elapsed < GameLoopTime Then
            If Elapsed >= 0 Then    'Make sure nothing weird happens, causing for a huge sleep time
                Sleep Int(GameLoopTime - Elapsed)
            End If
        End If

        '*** Update FPS ***
        If DEBUG_MapFPS Then
            FPS = FPS + 1
            If LastUpdate_ServerFPS + 1000 < timeGetTime Then
                FPSIndex = FPSIndex + 1
                
                'Check to make the array larger
                If ServerFPSUbound < FPSIndex Then
                    ServerFPSUbound = FPSIndex + 60 'Allocate a minute at a time
                    ReDim Preserve ServerFPS(1 To ServerFPSUbound) As ServerFPS
                End If
                
                'This basically adjusts it if the time is not exactly 1000ms
                ServerFPS(FPSIndex).FPS = Round(FPS * (1000 / (timeGetTime - LastUpdate_ServerFPS)))
                
                'Store the users and NPC values
                ServerFPS(FPSIndex).Users = NumUsers
                ServerFPS(FPSIndex).NPCs = NumNPCs
                
                'Clear the FPS
                FPS = 0
                
                'Set the last time the FPS was updated to now
                LastUpdate_ServerFPS = timeGetTime
                
            End If
        End If
        
    Loop
    
    'If for some reason the loop stops, unload the server
    Server_Unload
        
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
    
    'Update the tooltip
    TrayModify ToolTip, Server_BuildToolTipString

End Sub

Private Sub Server_Update_NPCs()

'*****************************************************************
'Updates the NPCs
'*****************************************************************
Dim NPCIndex As Integer

    'Update NPCs
    For NPCIndex = 1 To LastNPC

        'Make sure NPC is active
        If NPCList(NPCIndex).flags.NPCActive Then

            'See if npc is alive
            If NPCList(NPCIndex).flags.NPCAlive Then

                'Only update npcs in user populated maps
                If MapInfo(NPCList(NPCIndex).Pos.Map).NumUsers Then
                
                    'Check to update mod stats
                    If NPCList(NPCIndex).flags.UpdateStats Then
                        NPCList(NPCIndex).flags.UpdateStats = 0
                        NPC_UpdateModStats NPCIndex
                    End If
                    
                    '*** Update counters ***
                    If UpdateNPCCounters Then   'Update aggressive-face timer
                        If NPCList(NPCIndex).Counters.AggressiveCounter > 0 Then
                            If NPCList(NPCIndex).Counters.AggressiveCounter < timeGetTime Then
                                NPCList(NPCIndex).Counters.AggressiveCounter = 0
                                ConBuf.PreAllocate 4
                                ConBuf.Put_Byte DataCode.User_AggressiveFace
                                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                ConBuf.Put_Byte 0
                                Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            End If
                        End If                  'Update warcurse time
                        If NPCList(NPCIndex).Skills.WarCurse > 0 Then
                            If NPCList(NPCIndex).Counters.WarCurseCounter < timeGetTime Then
                                NPCList(NPCIndex).Counters.WarCurseCounter = 0
                                NPCList(NPCIndex).Skills.WarCurse = 0
                                ConBuf.PreAllocate 3 + Len(NPCList(NPCIndex).Name)
                                ConBuf.Put_Byte DataCode.Server_Message
                                ConBuf.Put_Byte 1
                                ConBuf.Put_String NPCList(NPCIndex).Name
                                Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
                                ConBuf.PreAllocate 4
                                ConBuf.Put_Byte DataCode.Server_IconWarCursed
                                ConBuf.Put_Byte 0
                                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            End If
                        End If                  'Update spell exhaustion
                        If NPCList(NPCIndex).Counters.SpellExhaustion > 0 Then
                            If NPCList(NPCIndex).Counters.SpellExhaustion < timeGetTime Then
                                NPCList(NPCIndex).Counters.SpellExhaustion = 0
                                ConBuf.PreAllocate 4
                                ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
                                ConBuf.Put_Byte 0
                                ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            End If
                        End If
                    End If

                    '*** NPC AI ***
                    If UpdateNPCAI Then
                        If NPCList(NPCIndex).Counters.ActionDelay < timeGetTime Then NPC_AI NPCIndex
                    End If

                End If

            Else
                
                '*** Respawn NPC ***
                'Check if it's time to respawn
                If NPCList(NPCIndex).Counters.RespawnCounter < timeGetTime Then NPC_Spawn NPCIndex

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
        If UserList(UserIndex).flags.UserLogged Then

            '*** Disconnection timers ***
            'Check if it has been idle for too long
            If UserList(UserIndex).Counters.IdleCount <= timeGetTime - IdleLimit Then
                Data_Send ToIndex, UserIndex, cMessage(85).Data
                Server_CloseSocket UserIndex
                GoTo NextUser   'Skip to the next user
            End If
            
            'Check if the user was possible disconnected (or extremely laggy)
            If UserList(UserIndex).Counters.LastPacket <= timeGetTime - LastPacket Then
                Data_Send ToIndex, UserIndex, cMessage(85).Data
                Server_CloseSocket UserIndex
                GoTo NextUser   'Skip to the next user
            End If
            
            '*** Recover stats ***
            If RecoverUserStats Then    'HP
                With UserList(UserIndex).Stats
                    If .BaseStat(SID.MinHP) < .ModStat(SID.MaxHP) Then
                        .BaseStat(SID.MinHP) = .BaseStat(SID.MinHP) + 1 + (.ModStat(SID.str) * 0.5)
                    End If                  'SP
                    If .BaseStat(SID.MinSTA) < .ModStat(SID.MaxSTA) Then
                        .BaseStat(SID.MinSTA) = .BaseStat(SID.MinSTA) + 1 + (.ModStat(SID.Agi) * 0.5)
                    End If                  'MP
                    If .BaseStat(SID.MinMAN) < .ModStat(SID.MaxMAN) Then
                        .BaseStat(SID.MinMAN) = .BaseStat(SID.MinMAN) + 1 + (.ModStat(SID.Mag) * 0.5)
                    End If
                End With
            End If

            '*** Update the counters ***
            If UpdateUserCounters Then  'Bless
                If UserList(UserIndex).Counters.BlessCounter > 0 Then
                    If UserList(UserIndex).Counters.BlessCounter < timeGetTime Then
                        UserList(UserIndex).Skills.Bless = 0
                        ConBuf.PreAllocate 4
                        ConBuf.Put_Byte DataCode.Server_IconBlessed
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Protection
                If UserList(UserIndex).Counters.ProtectCounter > 0 Then
                    If UserList(UserIndex).Counters.ProtectCounter < timeGetTime Then
                        UserList(UserIndex).Skills.Protect = 0
                        ConBuf.PreAllocate 4
                        ConBuf.Put_Byte DataCode.Server_IconProtected
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Strengthen
                If UserList(UserIndex).Counters.StrengthenCounter > 0 Then
                    If UserList(UserIndex).Counters.StrengthenCounter < timeGetTime Then
                        UserList(UserIndex).Skills.Strengthen = 0
                        ConBuf.PreAllocate 4
                        ConBuf.Put_Byte DataCode.Server_IconStrengthened
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Spell exhaustion
                If UserList(UserIndex).Counters.SpellExhaustion > 0 Then
                    If UserList(UserIndex).Counters.SpellExhaustion < timeGetTime Then
                        UserList(UserIndex).Counters.SpellExhaustion = 0
                        ConBuf.PreAllocate 4
                        ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    End If
                End If                  'Aggressive face
                If UserList(UserIndex).Counters.AggressiveCounter > 0 Then
                    If UserList(UserIndex).Counters.AggressiveCounter < timeGetTime Then
                        UserList(UserIndex).Counters.AggressiveCounter = 0
                        ConBuf.PreAllocate 4
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
                    If UserList(UserIndex).PacketWait < timeGetTime Then
    
                        'Send the packet buffer to the user
                        If UserList(UserIndex).PPValue = PP_High Then
                            
                            'High priority - send asap
                            Data_Send_Buffer UserIndex
                            
                        ElseIf UserList(UserIndex).PPValue = PP_Low Then
                            
                            'Low priority - check counter for sending
                            If UserList(UserIndex).PPCount < timeGetTime Then Data_Send_Buffer UserIndex
                        
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
                            If MapInfo(MapIndex).ObjTile(X, Y).ObjLife(ObjIndex) < timeGetTime - GroundObjLife Then
                                Obj_Erase MapInfo(MapIndex).ObjTile(X, Y).ObjInfo(ObjIndex).Amount, ObjIndex, MapIndex, X, Y
                            End If
                            
                        Next ObjIndex
                    End If
                    
                Next Y
            Next X
            
        Else
            
            '*** Unloading maps from memory ***
            'The map is empty, see if it needs to be unloaded (don't worry if it is already unloaded)
            Unload_Map MapIndex
        
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

Public Function Server_WalkTimePerTile(ByVal Speed As Long, Optional ByVal LagBuffer As Integer = 250) As Long
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
    '1/32 = The size of a tile
    '1000 = Miliseconds in a second
    'LagBuffer = We have to give some slack for network lag and client lag - raise this value if people skip too much
    '     and lower it if people are speedhacking and getting too much extra speed
    'Server_WalkTimePerTile = 1000 / (((Speed + 4) * 11) / 32) - LagBuffer
    Server_WalkTimePerTile = (1000 / ((Speed + 4) * 0.34375)) - LagBuffer
    
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
    Server_UserExist = Not DB_RS.EOF
    
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
'Check if two tile points are in the same area
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

Public Function Server_RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long

'*****************************************************************
'Find a Random number between a range
'*****************************************************************

    Log "Call Server_RandomNumber(" & LowerBound & "," & UpperBound & ")", CodeTracker '//\\LOGLINE//\\

    Server_RandomNumber = Fix((UpperBound - LowerBound + 1) * Rnd) + LowerBound
    
    Log "Rtrn Server_RandomNumber = " & Server_RandomNumber, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_BuildToolTipString() As String

'*****************************************************************
'Builds the tooltip string
'*****************************************************************
Dim kBpsIn As Single
Dim kBpsOut As Single

    'Get the number of connections
    Server_UpdateConnections

    'Display statistics (Kilobytes)
    On Error Resume Next
        kBpsIn = Round((DataKBIn + (DataIn * 0.0009765625)) / ((timeGetTime - ServerStartTime) * 0.001), 5)
        kBpsOut = Round((DataKBOut + (DataOut * 0.0009765625)) / ((timeGetTime - ServerStartTime) * 0.001), 5)
    On Error GoTo 0

    'Display statistics (Bytes)
    'kBpsIn = Round(((DataKBIn * 1024) + DataIn) / ((timeGetTime - ServerStartTime) / 1000), 5)
    'kBpsOut = Round(((DataKBOut * 1024) + DataOut) / ((timeGetTime - ServerStartTime) / 1000), 5)
    
    'Build the string
    Server_BuildToolTipString = "Connections: " & CurrConnections & vbNewLine & _
                                "kBps in: " & kBpsIn & vbNewLine & _
                                "kBps out: " & kBpsOut

End Function

Public Sub Server_UpdateConnections()

'*****************************************************************
'Find the number of users connected
'*****************************************************************

Dim LoopC As Long

    Log "Call Server_UpdateConnections", CodeTracker '//\\LOGLINE//\\

    'Clear the connections
    CurrConnections = 0

    'No users
    If LastUser <= 0 Then
        Log "Server_UpdateConnections: No users to update", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Loop through all the users
    Log "Server_UpdateConnections: Updating " & LastUser & " users", CodeTracker '//\\LOGLINE//\\
    For LoopC = 1 To LastUser
        If LenB(UserList(LoopC).Name) Then
            If UserList(LoopC).flags.UserLogged Then
                CurrConnections = CurrConnections + 1
            End If
        End If
    Next LoopC

End Sub

Public Sub ValidateTime()

'*****************************************************************
'This will validate that the timer hasn't rolled over
'If the timer does roll over, everything time-based will go out of
'wack, so we just turn off the server and let it reset
'This only happens after the server computer is on for 596.5 hours
'after turning on, then every 1193 hours after that
'*****************************************************************

    'Check if there was a roll-over (current time < last check)
    If timeGetTime < LastTimeGetTime Then UnloadServer = 1
    
    'Set the last check to now since we just checked it
    LastTimeGetTime = timeGetTime
    
End Sub

Public Function Server_IPisBanned(ByVal IP As String, ByRef ReturnReason As String) As Boolean

'*****************************************************************
'Returns whether an IP is banned or not
'*****************************************************************

    'Make the database query
    DB_RS.Open "SELECT * FROM banned_ips WHERE `ip`='" & IP & "'", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Get the value
    Server_IPisBanned = Not DB_RS.EOF
    
    'Return the reason
    If Server_IPisBanned Then ReturnReason = DB_RS!Reason

    'Close the database recordset
    DB_RS.Close

End Function

Public Sub Server_ConnectToServer(ByVal ServerIndex As Byte)

'*****************************************************************
'Connects this server to another server (only used for multiple servers)
'*****************************************************************

    'Make sure it isn't THIS server
    If ServerIndex = ServerID Then Exit Sub

    Select Case frmMain.ServerSocket(ServerIndex).State
    
        'Make sure the socket state is valid (not Error, Disconnected or Closing)
        Case sckConnected: Exit Sub
        Case sckConnecting: Exit Sub
        Case sckListening: Exit Sub

    End Select
    
    'Make the connection
    frmMain.ServerSocket(ServerIndex).Close
    frmMain.ServerSocket(ServerIndex).LocalPort = 0
    frmMain.ServerSocket(ServerIndex).Connect
    DoEvents
    
End Sub

Public Sub Server_Unload()

'*****************************************************************
'Unload the server and all the variables
'*****************************************************************
Dim Cancel As Byte
Dim FileNum As Byte
Dim LoopC As Long
Dim s As String

    On Error Resume Next

    Log "Call Server_Unload()", CodeTracker '//\\LOGLINE//\\
    
    'Close down the socket
    GOREsock_ShutDown
    GOREsock_UnHook
    If GOREsock_Loaded Then
        GOREsock_Terminate
        Cancel = 1
    End If
    
    If Cancel <> 1 Then
        
        'Stop the server loop
        ServerRunning = 0
        
        'Remove from system tray
        TrayDelete
        
        'Kill the database connection
        DB_Conn.Close
    
        'Save the debug files
        If DEBUG_RecordPacketsOut Then Save_PacketsOut
        If DEBUG_RecordPacketsIn Then Save_PacketsIn
        If DEBUG_MapFPS Then Save_FPS
        
        'Kill the temp files
        Kill ServerTempPath & "*"
        
        'Close the log files                                                                                            '//\\LOGLINE//\\
        If LogFileNumGeneral Then Close #LogFileNumGeneral                                                              '//\\LOGLINE//\\
        If LogFileNumCodeTracker Then Close #LogFileNumCodeTracker                                                      '//\\LOGLINE//\\
        If LogFileNumPacketIn Then Close #LogFileNumPacketIn                                                            '//\\LOGLINE//\\
        If LogFileNumPacketOut Then Close #LogFileNumPacketOut                                                          '//\\LOGLINE//\\
        If LogFileNumCriticalError Then Close #LogFileNumCriticalError                                                  '//\\LOGLINE//\\
        If LogFileNumInvalidPacketData Then Close #LogFileNumInvalidPacketData                                          '//\\LOGLINE//\\
    
        'Deallocate all arrays to avoid memory leaks
        Erase UserList
        Erase NPCList
        Erase MapInfo
        Erase CharList
        Erase NPCName
        Erase QuestData
        Erase HelpBuffer
        Erase DebugPacketsOut
        Erase DebugPacketsIn
        Erase MapUsers
        For LoopC = 1 To NumMaps
            Erase MapUsers(LoopC).Index
        Next LoopC
        Erase MapUsers
        Set ObjData = Nothing
        
        'Unload the form
        Unload frmMain
        
        'Close everything down
        End
    
    End If

End Sub
