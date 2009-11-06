Attribute VB_Name = "GameLogic"
Option Explicit

Sub Obj_Erase(ByVal Num As Integer, ByVal ObjSlot As Byte, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Erase a object
'*****************************************************************

    Log "Call Obj_Erase(" & Num & "," & ObjSlot & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Check for a valid index
    If ObjSlot > MapData(Map, X, Y).NumObjs Then
        Log "Obj_Erase: Invalid ObjSlot specified (" & ObjSlot & ")", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check to erase every object
    If Num = -1 Then Num = MapData(Map, X, Y).ObjInfo(ObjSlot).Amount

    'Remove the amount
    Log "Obj_Erase: Removing " & Num & " objects from (" & Map & "," & X & "," & Y & ") - current amount = " & MapData(Map, X, Y).ObjInfo(ObjSlot).Amount, CodeTracker '//\\LOGLINE//\\
    MapData(Map, X, Y).ObjInfo(ObjSlot).Amount = MapData(Map, X, Y).ObjInfo(ObjSlot).Amount - Num
    
    'Check if they are all gone
    If MapData(Map, X, Y).ObjInfo(ObjSlot).Amount <= 0 Then
        Log "Obj_Erase: Erasing object from client screens at (" & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\
        MapData(Map, X, Y).ObjInfo(ObjSlot).ObjIndex = 0
        MapData(Map, X, Y).ObjInfo(ObjSlot).Amount = 0
        MapData(Map, X, Y).ObjLife(ObjSlot) = 0
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_EraseObject
        ConBuf.Put_Byte CByte(X)
        ConBuf.Put_Byte CByte(Y)
        Data_Send ToMap, 0, ConBuf.Get_Buffer, Map
    End If

End Sub

Function Obj_ClosestFreeSpot(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByRef NewX As Byte, ByRef NewY As Byte, ByRef NewSlot As Byte)

'*****************************************************************
'Find the closest place to put an object
'*****************************************************************
Dim lX As Byte
Dim lY As Byte

    Log "Call Obj_ClosestFreeSpot(" & Map & "," & X & "," & Y & "," & NewX & "," & NewY & "," & NewSlot & ")", CodeTracker '//\\LOGLINE//\\
    
    'Check the defined location
    If Not (MapData(Map, X, Y).Blocked And BlockedAll) Then
        If MapData(Map, X, Y).NumObjs < MaxObjsPerTile Then
            
            'Spot is useable
            NewX = X
            NewY = Y
            NewSlot = MapData(Map, X, Y).NumObjs + 1
            Log "Rtrn Obj_ClosestFreeSpot = " & Obj_ClosestFreeSpot, CodeTracker '//\\LOGLINE//\\
            Exit Function
            
        End If
    End If
    
    'Primary spot didn't work, so loop around it and check if those work
    If X > 0 Then
        If Y > 0 Then
            For lX = X - 1 To X + 1
                For lY = Y - 1 To Y + 1
                    If lX > MinXBorder Then
                        If lX < MaxXBorder Then
                            If lY > MinYBorder Then
                                If lY < MaxYBorder Then
                                    If MapData(Map, lX, lY).Blocked = 0 Then
                                        If MapData(Map, lX, lY).NumObjs < MaxObjsPerTile Then
                                            
                                            'Spot is useable
                                            NewX = lX
                                            NewY = lY
                                            NewSlot = MapData(Map, lX, lY).NumObjs + 1
                                            Log "Rtrn Obj_ClosestFreeSpot = " & Obj_ClosestFreeSpot, CodeTracker '//\\LOGLINE//\\
                                            Exit Function
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next lY
            Next lX
        Else '//\\LOGLINE//\\
            Log "Obj_ClosestFreeSpot: X value is zero, can not subtract 1! Crash avoided!", CriticalError '//\\LOGLINE//\\
        End If
    Else    '//\\LOGLINE//\\
        Log "Obj_ClosestFreeSpot: X value is zero, can not subtract 1! Crash avoided!", CriticalError '//\\LOGLINE//\\
    End If
    
    Log "Rtrn Obj_ClosestFreeSpot = " & Obj_ClosestFreeSpot, CodeTracker '//\\LOGLINE//\\

End Function

Sub Obj_Make(Obj As Obj, ByVal ObjSlot As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal BypassUpdate As Byte = 0)

'*****************************************************************
'Create an object
'*****************************************************************

    Log "Call Obj_Make(N/A," & ObjSlot & "," & Map & "," & X & "," & Y & "," & BypassUpdate & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the ObjIndex isn't too high
    If ObjSlot > MaxObjsPerTile Then
        Log "Obj_Make: ObjSlot value too high (" & ObjSlot & ")", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If

    'Resize the object array to fit the slot
    If ObjSlot > MapData(Map, X, Y).NumObjs Then
        ReDim Preserve MapData(Map, X, Y).ObjInfo(1 To ObjSlot)
        ReDim Preserve MapData(Map, X, Y).ObjLife(1 To ObjSlot)
        MapData(Map, X, Y).NumObjs = ObjSlot
    End If
    
    'Add the object to the map slot
    MapData(Map, X, Y).ObjInfo(ObjSlot) = Obj
    MapData(Map, X, Y).ObjLife(ObjSlot) = timeGetTime
    
    'Clean the map tile just in case
    Obj_CleanMapTile Map, X, Y
    
    'Send the update to everyone on the map
    If BypassUpdate = 0 Then
        Log "Obj_Make: Updating object information with packet Server_MakeObject", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MakeObject
        ConBuf.Put_Long ObjData(Obj.ObjIndex).GrhIndex
        ConBuf.Put_Byte X
        ConBuf.Put_Byte Y
        Data_Send ToMap, 0, ConBuf.Get_Buffer, Map
    End If
    
End Sub

Sub Quest_SendReqString(ByVal UserIndex As Integer, ByVal QuestID As Integer)

'*****************************************************************
'Builds the string that says what is required for the quest
'*****************************************************************
Dim MessageID As Byte
Dim TempNPCName As String
Dim S As String

    Log "Call Quest_SendReqString(" & UserIndex & "," & QuestID & ")", CodeTracker '//\\LOGLINE//\\

    'Get the target NPC's name if there is one - to do this, we have to open up the NPC file since we dont store the "defaults" like we do with objects/quests/etc
    If QuestData(QuestID).FinishReqNPC Then
        DB_RS.Open "SELECT name FROM npcs WHERE id=" & QuestData(QuestID).FinishReqNPC
        TempNPCName = DB_RS!Name
        DB_RS.Close
    End If

    'Figure out the structure of our quest for the language file
    '9 = NPC only, 10 = Object only, 11 = NPC and Object
    If QuestData(QuestID).FinishReqNPC Then     'Needs NPC
        If QuestData(QuestID).FinishReqObj Then
            MessageID = 11                      'Needs object
        Else
            MessageID = 9                       'Doesn't need object
        End If
    Else
        If QuestData(QuestID).FinishReqObj Then
            MessageID = 11                      'Needs object
        Else
            'No object or NPC requirement found! Stupid quests dont deserve to be talked about
            Log "Quest_SendReqString: Error in Quest by ID " & QuestID & " - quest has no requirements!", CriticalError '//\\LOGLINE//\\
            Exit Sub
        End If
    End If

    'Set the general packet setup
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte MessageID

    'Build the packet according to the MessageID
    Select Case MessageID
        Case 9
            'NPC only
            ConBuf.Put_Integer QuestData(QuestID).FinishReqNPCAmount
            ConBuf.Put_String TempNPCName
        Case 10
            'Object only
            ConBuf.Put_Integer QuestData(QuestID).FinishReqObjAmount
            ConBuf.Put_String ObjData(QuestData(QuestID).FinishReqObj).Name
        Case 11
            'NPC and object
            ConBuf.Put_Integer QuestData(QuestID).FinishReqNPCAmount
            ConBuf.Put_String TempNPCName
            ConBuf.Put_Integer QuestData(QuestID).FinishReqObjAmount
            ConBuf.Put_String ObjData(QuestData(QuestID).FinishReqObj).Name
    End Select
    
    'Send the data to the user
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Quest_CheckIfComplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal UserQuestSlot As Byte)

'*****************************************************************
'Checks if a quest is ready to be completed
'*****************************************************************

Dim Slot As Byte

    Log "Call Quest_CheckIfComplete(" & UserIndex & "," & NPCIndex & "," & UserQuestSlot & ")", CodeTracker '//\\LOGLINE//\\

    'Check NPC kills
    If QuestData(NPCList(NPCIndex).Quest).FinishReqNPC Then
        If UserList(UserIndex).QuestStatus(UserQuestSlot).NPCKills < QuestData(NPCList(NPCIndex).Quest).FinishReqNPCAmount Then
            Quest_SayIncomplete UserIndex, NPCIndex
            Exit Sub
        End If
    End If

    'Check objects
    If QuestData(NPCList(NPCIndex).Quest).FinishReqObj Then

        'Check through the user's slots looking for the object of the same index
        For Slot = 1 To MAX_INVENTORY_SLOTS
            If UserList(UserIndex).Object(Slot).ObjIndex = QuestData(NPCList(NPCIndex).Quest).FinishReqObj Then

                'Check the amount the user has
                If UserList(UserIndex).Object(Slot).Amount < QuestData(NPCList(NPCIndex).Quest).FinishReqObjAmount Then
                    Quest_SayIncomplete UserIndex, NPCIndex
                    Exit Sub
                Else
                    'We will set the slot to 0 so we can check if we ran out of slots or found the object we were looking for
                    Slot = 0
                    Exit For
                End If

                'And now we check! lawlz! ^_^
                If Slot <> 0 Then
                    Quest_SayIncomplete UserIndex, NPCIndex
                    Exit Sub
                End If

            End If
        Next Slot

    End If

    'Say the finishing text
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer

    'The user is done, give them the rewards
    'EXP reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp > 0 Then
        User_RaiseExp UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 3
        ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
    'Gold reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold > 0 Then
        UserList(UserIndex).Gold = UserList(UserIndex).Gold + QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 4
        ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
    'Object reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObj > 0 Then
        User_GiveObj UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObj, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObjAmount
    End If
    
    'Learn skill reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill > 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        If UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) = 1 Then
            'User already knew the skill
            ConBuf.Put_Byte 5
            ConBuf.Put_String Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Else
            'User learns the new skill
            ConBuf.Put_Byte 6
            ConBuf.Put_String Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            'Give the user the skill
            UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) = 1
            User_SendKnownSkills UserIndex
        End If
    End If

    'Add the quest to the user's finished quest list
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).Redoable Then

        'Only add a redoable quest in the list once
        Log "Quest_CheckIfComplete: Using InStr() operation", CodeTracker '//\\LOGLINE//\\
        If Not InStr(1, UserList(UserIndex).CompletedQuests & "-", "-" & UserList(UserIndex).Quest(UserQuestSlot) & "-") Then
            UserList(UserIndex).CompletedQuests = UserList(UserIndex).CompletedQuests & "-" & UserList(UserIndex).Quest(UserQuestSlot)
        End If

    Else

        'Add to the list
        UserList(UserIndex).CompletedQuests = UserList(UserIndex).CompletedQuests & "-" & UserList(UserIndex).Quest(UserQuestSlot)

    End If

    'Clear the quest slot so it can be used again
    UserList(UserIndex).QuestStatus(UserQuestSlot).NPCKills = 0
    UserList(UserIndex).Quest(UserQuestSlot) = 0

End Sub

Sub Quest_General(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Reacts to a user clicking a quest NPC
'*****************************************************************

Dim S As String
Dim i As Integer

    Log "Call Quest_General(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for valid values
    On Error GoTo ErrOut:
    If UserIndex <= 0 Then Exit Sub
    If UserIndex > LastUser Then Exit Sub
    If NPCIndex <= 0 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If NPCList(NPCIndex).Quest <= 0 Then Exit Sub
    On Error GoTo 0

    'Check if the user is currently involved in the quest
    For i = 1 To MaxQuests

        'If they are involved in a quest, then we will send it off to another sub
        If UserList(UserIndex).Quest(i) = NPCList(NPCIndex).Quest Then
            Quest_CheckIfComplete UserIndex, NPCIndex, i
            Exit Sub
        End If

    Next i

    'The user is not involved in this quest currently - check if they have already completed it
    Log "Quest_General: Using InStr() operation", CodeTracker '//\\LOGLINE//\\
    If InStr(1, UserList(UserIndex).CompletedQuests & "-", "-" & NPCList(NPCIndex).Quest & "-") Then

        'The user has completed this quest before, so check if it is redoable
        If QuestData(NPCList(NPCIndex).Quest).Redoable = 0 Then

            'The quest is not redoable, so sorry dude, no quest fo' j00
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 7
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub

        End If

    End If

    'The user has never done this quest before, so we make the NPC say whats up
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(NPCList(NPCIndex).Quest).StartTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer

    'Give the quest requirements
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 8
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Set the pending quest to the selected quest
    UserList(UserIndex).Flags.QuestNPC = NPCIndex
    
ErrOut:

End Sub

Sub Quest_SayIncomplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Make the targeted NPC say the "incomplete quest" text
'*****************************************************************

    Log "Call Quest_SayIncomplete(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Incomplete text
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(NPCList(NPCIndex).Quest).IncompleteTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer

    'Requirements text
    Quest_SendReqString UserIndex, NPCList(NPCIndex).Quest

End Sub

Function Server_CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean

'*****************************************************************
'Checks for a user with the same Name
'*****************************************************************

Dim LoopC As Long

    Log "Call Server_CheckForSameName(" & UserIndex & "," & Name & ")", CodeTracker '//\\LOGLINE//\\

    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged = 1 Then
            If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserIndex <> LoopC Then
                Server_CheckForSameName = True
                Log "Rtrn Server_CheckForSameName = " & Server_CheckForSameName, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
        End If
    Next LoopC

    Server_CheckForSameName = False
    
    Log "Rtrn Server_CheckForSameName = " & Server_CheckForSameName, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_CheckTargetedDistance(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Checks if a user is targeting a character in range
'*****************************************************************

Dim TargetID As Integer

    Log "Call Server_CheckTargetedDistance(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Set the target ID
    TargetID = UserList(UserIndex).Flags.TargetIndex
    Log "Server_CheckTargetedDistance: Target ID acquired (" & TargetID & ")", CodeTracker '//\\LOGLINE//\\

    Select Case UserList(UserIndex).Flags.Target

        'Self
        Case 0
            Server_CheckTargetedDistance = 1
            Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\
            Exit Function

        'User
        Case 1
        
            'Check the map
            If UserList(UserIndex).Pos.Map <> UserList(CharList(TargetID).Index).Pos.Map Then
                Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            
            'Check the X/Y position
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(CharList(TargetID).Index).Pos.X, UserList(CharList(TargetID).Index).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                Server_CheckTargetedDistance = 1
                Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
    
            'NPC
        Case 2
        
            'Check the map
            If UserList(UserIndex).Pos.Map <> NPCList(CharList(TargetID).Index).Pos.Map Then
                Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            
            'Check the X/Y position
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(CharList(TargetID).Index).Pos.X, NPCList(CharList(TargetID).Index).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                Server_CheckTargetedDistance = 1
                Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If

    End Select

    'Not in distance or nothing targeted, so tell the user it is not targeted
    If TargetID = 0 Or UserList(UserIndex).Flags.TargetIndex = 0 Then
        Log "Server_CheckTargetedDistance: Telling user nothing is targeted", CodeTracker '//\\LOGLINE//\\
        UserList(UserIndex).Flags.Target = 0
        UserList(UserIndex).Flags.TargetIndex = 0
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Target
        ConBuf.Put_Integer 0
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
    Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\

End Function

Sub Server_ClosestLegalPos(Pos As WorldPos, nPos As WorldPos)

'*****************************************************************
'Finds the closest legal tile to Pos and stores it in nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

    Log "Call Server_ClosestLegalPos([M:" & Pos.Map & " X:" & Pos.X & " Y:" & Pos.Y & "],N/A)", CodeTracker '//\\LOGLINE//\\

    'Set the new map
    nPos.Map = Pos.Map

    'Keep looping while the position is not legal
    Do While Not Server_LegalPos(Pos.Map, nPos.X, nPos.Y, 0)

        'If we have checked too much, then just leave
        If LoopC > 3 Then   'How many tiles in all directions to search
            Notfound = True
            Exit Do
        End If

        'Loop through the tiles
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                Log "Server_ClosestLegalPos: Checking map tile (" & nPos.Map & "," & tX & "," & tY & ")", CodeTracker '//\\LOGLINE//\\

                'Check if the position is legal
                If Server_LegalPos(nPos.Map, tX, tY, 0) = True Then
                    nPos.X = tX
                    nPos.Y = tY
                    tX = Pos.X + LoopC
                    tY = Pos.Y + LoopC
                End If

            Next tX
        Next tY

        'Check the next set of tiles
        LoopC = LoopC + 1

    Loop

    'If no position was found, return empty positions
    If Notfound Then
        nPos.X = 0
        nPos.Y = 0
    End If
    
    Log "Server_ClosestLegalPos: Returning position (" & nPos.Map & "," & nPos.X & "," & nPos.Y & ")", CodeTracker '//\\LOGLINE//\\

End Sub

Sub Server_DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Do any events on a tile
'*****************************************************************

Dim TempPos As WorldPos
Dim NewPos As WorldPos

    Log "Call Server_DoTileEvents(" & UserIndex & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Check for tile exit
    If MapData(Map, X, Y).TileExit.Map Then

        'Set the position values
        TempPos.X = MapData(Map, X, Y).TileExit.X
        TempPos.Y = MapData(Map, X, Y).TileExit.Y
        TempPos.Map = MapData(Map, X, Y).TileExit.Map
        Log "Server_DoTileEvents: Tile exist exists, warps to (" & TempPos.Map & "," & TempPos.X & "," & TempPos.Y & ")", CodeTracker '//\\LOGLINE//\\

        'Get the closest legal position
        Server_ClosestLegalPos TempPos, NewPos

        'If the position is legal, then warp the user there
        If Server_LegalPos(NewPos.Map, NewPos.X, NewPos.Y, 0) Then User_WarpChar UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y

    End If

End Sub

Function Server_FindDirection(Pos As WorldPos, Target As WorldPos) As Byte

'*****************************************************************
'Returns the direction in which the Target is from the Pos, 0 if equal
'*****************************************************************
Dim pX As Integer
Dim pY As Integer
Dim tX As Integer
Dim tY As Integer
Dim X As Integer
Dim Y As Integer

    Log "Call Server_FindDirection([M:" & Pos.Map & " X:" & Pos.X & " Y:" & Pos.Y & "],[M:" & Target.Map & " X:" & Target.X & " Y:" & Target.Y & "])", CodeTracker '//\\LOGLINE//\\

    'Put the bytes into integer variables (causes overflows for negatives, even if the return is an integer)
    pX = Pos.X
    pY = Pos.Y
    tX = Target.X
    tY = Target.Y
    
    'Find the difference
    X = pX - tX
    Y = pY - tY
    Log "Server_FindDirection: Position difference (X:" & X & " Y:" & Y & ") found", CodeTracker '//\\LOGLINE//\\

    'NE
    If X <= -1 Then
        If Y >= 1 Then
            Server_FindDirection = NORTHEAST
            Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If

    'NW
    If X >= 1 Then
        If Y >= 1 Then
            Server_FindDirection = NORTHWEST
            Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If

    'SW
    If X >= 1 And Y <= -1 Then
        Server_FindDirection = SOUTHWEST
        Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'SE
    If X <= -1 Then
        If Y <= -1 Then
            Server_FindDirection = SOUTHEAST
            Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    End If

    'South
    If Y <= -1 Then
        Server_FindDirection = SOUTH
        Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'north
    If Y >= 1 Then
        Server_FindDirection = NORTH
        Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'West
    If X >= 1 Then
        Server_FindDirection = WEST
        Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'East
    If X <= -1 Then
        Server_FindDirection = EAST
        Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\

End Function

Sub Server_HeadToPos(ByVal Head As Byte, ByRef Pos As WorldPos)

'*****************************************************************
'Takes Pos and moves it in heading direction
'*****************************************************************

Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

    Log "Call Server_HeadToPos(" & Head & ",[M:" & Pos.Map & " X:" & Pos.X & " Y:" & Pos.Y & "])", CodeTracker '//\\LOGLINE//\\

    X = Pos.X
    Y = Pos.Y

    If Head = NORTH Then
        nX = X
        nY = Y - 1
    End If

    If Head = SOUTH Then
        nX = X
        nY = Y + 1
    End If

    If Head = EAST Then
        nX = X + 1
        nY = Y
    End If

    If Head = WEST Then
        nX = X - 1
        nY = Y
    End If

    If Head = NORTHEAST Then
        nX = X + 1
        nY = Y - 1
    End If

    If Head = SOUTHEAST Then
        nX = X + 1
        nY = Y + 1
    End If

    If Head = SOUTHWEST Then
        nX = X - 1
        nY = Y + 1
    End If

    If Head = NORTHWEST Then
        nX = X - 1
        nY = Y - 1
    End If

    'return values
    Pos.X = nX
    Pos.Y = nY

End Sub

Function Server_InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

    Log "Call Server_InMapBounds(" & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    If X > MinXBorder Then
        If X < MaxXBorder Then
            If Y > MinYBorder Then
                If Y < MaxYBorder Then Server_InMapBounds = True
            End If
        End If
    End If
    
    Log "Rtrn Server_InMapBounds = " & Server_InMapBounds, CodeTracker '//\\LOGLINE//\\
    
End Function

Function Server_LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As Byte) As Boolean

'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
On Error GoTo ErrOut
    
    Log "Call Server_LegalPos(" & Map & "," & X & "," & Y & "," & Heading & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure it's a legal map
    If Map <= 0 Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If Map > NumMaps Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'Check to see if its out of bounds
    If X < MinXBorder Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If X > MaxXBorder Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If Y < MinYBorder Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If Y > MaxYBorder Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'Check if a character (User or NPC) is already at the tile
    If MapData(Map, X, Y).UserIndex > 0 Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If MapData(Map, X, Y).NPCIndex > 0 Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'Check to see if its blocked
    If MapData(Map, X, Y).Blocked = BlockedAll Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'Check the heading for directional blocking
    If Heading > 0 Then
        If MapData(Map, X, Y).Blocked And BlockedNorth Then
            If Heading = NORTH Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = NORTHEAST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = NORTHWEST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
        End If
        If MapData(Map, X, Y).Blocked And BlockedEast Then
            If Heading = EAST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = NORTHEAST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = SOUTHEAST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
        End If
        If MapData(Map, X, Y).Blocked And BlockedSouth Then
            If Heading = SOUTH Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = SOUTHEAST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = SOUTHWEST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
        End If
        If MapData(Map, X, Y).Blocked And BlockedWest Then
            If Heading = WEST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = NORTHWEST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
            If Heading = SOUTHWEST Then
                Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
                Exit Function
            End If
        End If
    End If

    'If we are still in this routine, then it must be a legal position
    Server_LegalPos = True
    
ErrOut:

    Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_NextOpenCharIndex() As Integer

'*****************************************************************
'Finds the next open CharIndex in Charlist
'*****************************************************************

Dim LoopC As Long
    
    Log "Call Server_NextOpenCharIndex", CodeTracker '//\\LOGLINE//\\
    
    'Check for the first char creation
    If LastChar = 0 Then
        ReDim CharList(0 To 1)
        LastChar = 1
        Server_NextOpenCharIndex = 1
        Log "Rtrn Server_NextOpenCharIndex = " & Server_NextOpenCharIndex, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Loop through the character slots
    Log "Server_NextOpenCharIndex: Starting loop (1 to " & LastChar & ")", CodeTracker '//\\LOGLINE//\\
    For LoopC = 1 To LastChar + 1

        'We need to create a new slot
        If LoopC > LastChar Then
            LastChar = LoopC
            Server_NextOpenCharIndex = LoopC
            ReDim Preserve CharList(0 To LastChar)
            Log "Rtrn Server_NextOpenCharIndex = " & Server_NextOpenCharIndex, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If

        'Re-use an old slot that is not being used
        If CharList(LoopC).Index = 0 Then
            Server_NextOpenCharIndex = LoopC
            Log "Rtrn Server_NextOpenCharIndex = " & Server_NextOpenCharIndex, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If

    Next LoopC
    
    Log "Rtrn Server_NextOpenCharIndex = " & Server_NextOpenCharIndex, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_SkillIDtoSkillName(ByVal SkillID As Byte) As String

'***************************************************
'Takes in a SkillID and returns the name of that skill
'***************************************************

    Log "Call Server_SkillIDtoSkillName(" & SkillID & ")", CodeTracker '//\\LOGLINE//\\

    Select Case SkillID
        Case SkID.Bless: Server_SkillIDtoSkillName = "Bless"
        Case SkID.IronSkin: Server_SkillIDtoSkillName = "Iron Skin"
        Case SkID.Strengthen: Server_SkillIDtoSkillName = "Strengthen"
        Case SkID.Warcry: Server_SkillIDtoSkillName = "Warcry"
        Case SkID.Protection: Server_SkillIDtoSkillName = "Protection"
        Case SkID.SpikeField: Server_SkillIDtoSkillName = "Spike Field"
        Case SkID.Heal: Server_SkillIDtoSkillName = "Heal"
        Case Else: Server_SkillIDtoSkillName = "Unknown Skill #" & SkillID
    End Select
    
    Log "Rtrn Server_SkillIDtoSkillName = " & Server_SkillIDtoSkillName, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub Server_WriteMail(WriterIndex As Integer, ReceiverName As String, Subject As String, Message As String, ObjIndexString As String, ObjAmountString As String)
Dim MailIndex As Long
Dim MailData As MailData
Dim TempSplit() As String
Dim TempSplit2() As String
Dim TempStr As String
Dim LoopC As Byte
Dim LoopX As Byte
Dim i As Long

    Log "Call Server_WriteMail(" & WriterIndex & "," & ReceiverName & "," & Subject & ",[" & Message & "],[" & ObjIndexString & "],[" & ObjAmountString & "])", CodeTracker '//\\LOGLINE//\\

    'Check for a writing delay
    If WriterIndex > 0 Then
        If UserList(WriterIndex).Counters.DelayTimeMail > timeGetTime Then
            'Not enough time has passed by - goodbye! :)
            Log "Server_WriteMail: Not enough time elapsed since last mail write for user " & WriterIndex & " (" & UserList(WriterIndex).Name & ")", CodeTracker '//\\LOGLINE//\\
            Exit Sub
        Else
            'Set the delay
            UserList(WriterIndex).Counters.DelayTimeMail = timeGetTime + DelayTimeMail
        End If
    End If
    
    'Check for a valid reciever
    If Server_UserExist(ReceiverName) = False Then
        If WriterIndex > 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 13
            ConBuf.Put_String ReceiverName
            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
        End If
        Exit Sub
    End If
    
    'Check for sending cost
    If MailCost > 0 Then
        If WriterIndex > 0 Then 'Prevent charging the server from sending mail (ID = -1 for server) ;)
            If UserList(WriterIndex).Stats.BaseStat(SID.Gold) < MailCost Then
            
                'Not enough money
                Log "Server_WriteMail: User " & WriterIndex & " (" & UserList(WriterIndex).Name & ") has not enough money to write mail (Gold: " & UserList(WriterIndex).Stats.BaseStat(SID.Gold) & ")", CodeTracker '//\\LOGLINE//\\
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 14
                ConBuf.Put_Long MailCost
                Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                Exit Sub
            
            End If
        End If
    End If

    'Get the next open mail ID
    DB_RS.Open "SELECT lastid FROM mail_lastid", DB_Conn, adOpenStatic, adLockOptimistic
    MailIndex = Val(DB_RS(0)) + 1
    DB_RS(0) = MailIndex    'Update the value in the database
    DB_RS.Update
    DB_RS.Close
    Log "Server_WriteMail: MailIndex acquired from `mail_lastid` table (" & MailIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Set up the mail type
    MailData.New = 1
    MailData.Message = Message
    MailData.RecieveDate = Date
    MailData.Subject = Subject
    If WriterIndex <> -1 Then MailData.WriterName = UserList(WriterIndex).Name Else MailData.WriterName = "Game Admin"

    'Split up the object index string
    TempSplit = Split(ObjIndexString, ",")
    For LoopC = 0 To UBound(TempSplit())
        MailData.Obj(LoopC + 1).ObjIndex = Val(TempSplit(LoopC))
    Next LoopC
    
    'Split up the object amount string
    TempSplit2 = Split(ObjAmountString, ",")
    For LoopC = 0 To UBound(TempSplit2())
        MailData.Obj(LoopC + 1).Amount = Val(TempSplit2(LoopC))
    Next LoopC

    'Check if the reciever is on
    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged Then
            If UCase$(UserList(LoopC).Name) = UCase$(ReceiverName) Then

                'Get the user's next open MailID slot
                LoopX = 0
                Do
                    LoopX = LoopX + 1
                    If LoopX > MaxMailPerUser Then
                        If WriterIndex <> -1 Then
                            'Message to the receiver
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 15
                            ConBuf.Put_String UserList(WriterIndex).Name
                            Data_Send ToIndex, LoopC, ConBuf.Get_Buffer
                            'Message to the sender
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 15
                            ConBuf.Put_String ReceiverName
                            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                            Exit Sub
                        Else
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 17
                            Data_Send ToIndex, LoopC, ConBuf.Get_Buffer
                            Exit Sub
                        End If
                    End If
                Loop While UserList(LoopC).MailID(LoopX) > 0

                'Add the mail directly to the user's MailID
                UserList(LoopC).MailID(LoopX) = MailIndex

                'Save the mail
                Save_Mail MailIndex, MailData

                'Display the recieve/sent messages
                If WriterIndex <> -1 Then
                    'Send message to sender
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 19
                    ConBuf.Put_String ReceiverName
                    Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                    'Send message to receiver
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 18
                    ConBuf.Put_String UserList(WriterIndex).Name
                    Data_Send ToIndex, LoopC, ConBuf.Get_Buffer, , PP_NewMail
                Else
                    'Send message to receiver that it was from the game admin
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 20
                    Data_Send ToIndex, LoopC, ConBuf.Get_Buffer, , PP_NewMail
                End If
                
                'Check for sending cost
                If MailCost > 0 Then
                    If WriterIndex > 0 Then
                        If UserList(WriterIndex).Stats.BaseStat(SID.Gold) < MailCost Then
                            Log "Server_WriteMail: Not enough gold to write message (Gold: " & UserList(WriterIndex).Stats.BaseStat(SID.Gold) & ")", CodeTracker '//\\LOGLINE//\\
                            Exit Sub    'This should never be reached, but just in case it does, leave this here
                        Else
                            UserList(WriterIndex).Stats.BaseStat(SID.Gold) = UserList(WriterIndex).Stats.BaseStat(SID.Gold) - MailCost
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 21
                            ConBuf.Put_Long MailCost
                            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                        End If
                    End If
                End If
                
                Exit Sub

            End If
        End If
    Next LoopC

    'The user is not on, so load up his character data and impliment it into the character
    DB_RS.Open "SELECT name,mail FROM users WHERE `name`='" & ReceiverName & "'", DB_Conn, adOpenStatic, adLockOptimistic
    TempStr = DB_RS!mail
    TempSplit = Split(TempStr, vbCrLf)
    If UBound(TempSplit) >= MaxMailPerUser Then 'No room for the mail
        Log "Server_WriteMail: User has too much mail alread - aborting", CodeTracker '//\\LOGLINE//\\
        DB_RS.Close
        If WriterIndex <> -1 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 16
            ConBuf.Put_String ReceiverName
            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
        End If
        Exit Sub
    Else    'Save the mail ID in the user
        If TempStr <> "" Then TempStr = TempStr & vbCrLf
        TempStr = TempStr & MailIndex
        DB_RS!mail = TempStr
        DB_RS.Update
        DB_RS.Close
    End If
    
    'Check for sending cost
    If MailCost > 0 Then
        If WriterIndex > 0 Then
            If UserList(WriterIndex).Stats.BaseStat(SID.Gold) < MailCost Then
                Log "Server_WriteMail: Not enough gold to write message (Gold: " & UserList(WriterIndex).Stats.BaseStat(SID.Gold) & ")", CodeTracker '//\\LOGLINE//\\
                Exit Sub    'This should never be reached, but just in case it does, leave this here
            Else
                UserList(WriterIndex).Stats.BaseStat(SID.Gold) = UserList(WriterIndex).Stats.BaseStat(SID.Gold) - MailCost
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 21
                ConBuf.Put_Long MailCost
                Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
            End If
        End If
    End If

    'Send the message of success
    If WriterIndex <> -1 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 19
        ConBuf.Put_String ReceiverName
        Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
    End If

    'Save the mail
    Save_Mail MailIndex, MailData
    
End Sub

Public Sub User_AddObjToInv(ByVal UserIndex As Integer, ByRef Object As Obj)
Dim LoopC As Long
Dim Map As Integer
Dim NewX As Byte
Dim NewY As Byte
Dim NewSlot As Byte

    Log "Call User_AddObjToInv(" & UserIndex & ",[I:" & Object.ObjIndex & " A:" & Object.Amount & "])", CodeTracker '//\\LOGLINE//\\
    
    'Look for a slot
    Log "User_AddObjToInv: Starting loop (1 to " & MAX_INVENTORY_SLOTS & ")", CodeTracker '//\\LOGLINE//\\
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Object(LoopC).ObjIndex = Object.ObjIndex Then
            If UserList(UserIndex).Object(LoopC).Amount + Object.Amount <= MAX_INVENTORY_OBJS Then
                UserList(UserIndex).Object(LoopC).Amount = UserList(UserIndex).Object(LoopC).Amount + Object.Amount
                Object.Amount = 0
                'Update this slot
                User_UpdateInv False, UserIndex, LoopC
                Exit Sub
            Else
                Object.Amount = Object.Amount - (MAX_INVENTORY_OBJS - UserList(UserIndex).Object(LoopC).Amount)
                UserList(UserIndex).Object(LoopC).Amount = MAX_INVENTORY_OBJS
                'Update this slot
                User_UpdateInv False, UserIndex, LoopC
            End If
        End If
    Next LoopC

    'Check if we have placed all items
    If Object.Amount Then
        'Look for an empty slot
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            If UserList(UserIndex).Object(LoopC).ObjIndex = 0 Then
                UserList(UserIndex).Object(LoopC).Amount = Object.Amount
                UserList(UserIndex).Object(LoopC).ObjIndex = UserList(UserIndex).Object(LoopC).ObjIndex
                Object.Amount = 0
                'Update this slot
                User_UpdateInv False, UserIndex, LoopC
                Exit Sub
            End If
        Next LoopC
    End If
    
    'No free slot was found, drop the object to the floor
    
    'Find the closest legal pos
    Obj_ClosestFreeSpot UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NewX, NewY, NewSlot
    
    'Make sure the spot is valid
    If NewX > 0 Then
        Obj_Make Object, NewSlot, UserList(UserIndex).Pos.Map, NewX, NewY
    Else '//\\LOGLINE//\\
        Log "User_AddObjToInv: No valid spot found to drop the object!", CodeTracker '//\\LOGLINE//\\
    End If

End Sub

Private Sub User_Attack_Ranged(ByVal UserIndex As Integer)

'*****************************************************************
'Begin an attack sequence (from ranged)
'*****************************************************************
Dim NewHeading As Byte
Dim TargetPos As WorldPos
Dim TargetIndex As Integer
Dim Angle As Single
    
    'Get the target index based on the NPCList() or UserList() arrays instead of CharList() value
    TargetIndex = CharList(UserList(UserIndex).Flags.TargetIndex).Index

    'Check if a NPC or PC
    Select Case UserList(UserIndex).Flags.Target
        Case 1  'PC
            With UserList(TargetIndex).Pos
                TargetPos.Map = .Map
                TargetPos.X = .X
                TargetPos.Y = .Y
            End With
            
        Case 2  'NPC
            
            'Check for a valid NPC
            If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
            
            With NPCList(TargetIndex).Pos
                TargetPos.Map = .Map
                TargetPos.X = .X
                TargetPos.Y = .Y
            End With
            
            
        Case Else
            Exit Sub
        
    End Select
    
    'Check for a valid distance
    If Server_Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, TargetPos.X, TargetPos.Y) > ObjData(UserList(UserIndex).WeaponEqpObjIndex).WeaponRange Then Exit Sub
    
    'Check for a valid target
    If Engine_ClearPath(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, TargetPos.X, TargetPos.Y) Then
                
        'Play attack sound
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_PlaySound3D
        ConBuf.Put_Byte SOUND_SWING
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
        
        'Get the new heading
        Angle = Engine_GetAngle(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, TargetPos.X, TargetPos.Y)
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
        
        Select Case UserList(UserIndex).Flags.Target
        
            'Attacking user
            Case 1

                'Send the data
                ConBuf.Clear
                If ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh Then
                    ConBuf.Put_Byte DataCode.Server_MakeProjectile
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
                    ConBuf.Put_Long ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh
                    ConBuf.Put_Byte ObjData(UserList(UserIndex).WeaponEqpObjIndex).ProjectileRotateSpeed
                End If
                ConBuf.Put_Byte DataCode.User_Rotate
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                ConBuf.Put_Byte NewHeading
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                User_AttackUser UserIndex, TargetIndex
            
            'Attacking NPC
            Case 2
                ConBuf.Clear
                If ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh Then
                    ConBuf.Put_Byte DataCode.Server_MakeProjectile
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
                    ConBuf.Put_Long ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh
                    ConBuf.Put_Byte ObjData(UserList(UserIndex).WeaponEqpObjIndex).ProjectileRotateSpeed
                End If
                ConBuf.Put_Byte DataCode.User_Rotate
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                ConBuf.Put_Byte NewHeading
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                User_AttackNPC UserIndex, TargetIndex
                
        End Select
        
    End If

End Sub

Sub User_Attack(ByVal UserIndex As Integer)

'*****************************************************************
'Begin a user attack sequence
'*****************************************************************

Dim AttackPos As WorldPos

    Log "Call User_Attack(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    On Error GoTo ErrOut
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Log "User_Attack: User switching maps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Stats.BaseStat(SID.MinSTA) <= 0 Then
        Log "User_Attack: MinSTA <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Counters.AttackCounter > timeGetTime - STAT_ATTACKWAIT Then
        Log "User_Attack: Not enough time elapsed since last attack - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0

    'Update counters
    UserList(UserIndex).Counters.AttackCounter = timeGetTime

    'Check for ranged attack
    If ObjData(UserList(UserIndex).WeaponEqpObjIndex).WeaponRange > 1 Then
        If UserList(UserIndex).Flags.TargetIndex = 0 Then Exit Sub
        User_Attack_Ranged UserIndex
        Exit Sub
    End If

    'Get tile user is attacking
    AttackPos = UserList(UserIndex).Pos
    Server_HeadToPos UserList(UserIndex).Char.Heading, AttackPos

    'Exit if not legal
    If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
        Log "User_Attack: Trying to attack an illegal position - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Look for user
    If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex > 0 Then
        Log "User_Attack: Found a user to attack", CodeTracker '//\\LOGLINE//\\

        'Play attack sound
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_PlaySound3D
        ConBuf.Put_Byte SOUND_SWING
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        If ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh Then
            ConBuf.Put_Byte DataCode.Server_MakeSlash
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            ConBuf.Put_Long ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh
        End If
        Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

        'Go to the user attacking user sub
        User_AttackUser UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
        Exit Sub

    End If

    'Look for NPC
    If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex > 0 Then
        Log "User_Attack: Found a NPC to attack", CodeTracker '//\\LOGLINE//\\
        If NPCList(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex).Attackable Then
            
            'If NPC has no health, they can not be attacked
            If NPCList(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex).ModStat(SID.MaxHP) = 0 Then
                Log "User_Attack: NPC's MaxHP = 0 - aborting", CodeTracker '//\\LOGLINE//\\
                Exit Sub
            End If
            If NPCList(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex).BaseStat(SID.MaxHP) = 0 Then
                Log "User_Attack: NPC's MaxHP = 0 - aborting", CodeTracker '//\\LOGLINE//\\
                Exit Sub
            End If

            'Play attack sound and create the graphic
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_PlaySound3D
            ConBuf.Put_Byte SOUND_SWING
            ConBuf.Put_Byte UserList(UserIndex).Pos.X
            ConBuf.Put_Byte UserList(UserIndex).Pos.Y
            If ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh Then
                ConBuf.Put_Byte DataCode.Server_MakeSlash
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                ConBuf.Put_Long ObjData(UserList(UserIndex).WeaponEqpObjIndex).UseGrh
            End If
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

            'Go to user attacking npc sub
            User_AttackNPC UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex

        Else
            Log "User_Attack: NPC is non-attackable", CodeTracker '//\\LOGLINE//\\

            'Can not attack the selected NPC, NPC is not attackable
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 2
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        End If
        Exit Sub
    End If
    
ErrOut:

End Sub

Sub User_AttackNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Have a User attack a NPC
'*****************************************************************

Dim HitSkill As Long    'User hit skill
Dim Hit As Integer      'Hit damage

    Log "Call User_AttackNPC(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Get the user hit skill
    HitSkill = UserList(UserIndex).Stats.ModStat(SID.Agi) * 2 + UserList(UserIndex).Stats.ModStat(SID.Str)

    'Display the attack
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Attack
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

    'Calculate if they hit
    If Server_RandomNumber(1, 100) >= ((HitSkill + 50) - NPCList(NPCIndex).ModStat(SID.Agi)) Then
        Log "User_AttackNPC: Attack chance did not pass, registering as a miss", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetCharDamage
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Integer -1
        Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Update aggressive-face
    If UserList(UserIndex).Counters.AggressiveCounter <= 0 Then
        Log "User_AttackNPC: Making the user aggressive-faced", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_AggressiveFace
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte 1
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If
    UserList(UserIndex).Counters.AggressiveCounter = AGGRESSIVEFACETIME

    'Calculate hit
    Hit = Server_RandomNumber(UserList(UserIndex).Stats.ModStat(SID.MinHIT), UserList(UserIndex).Stats.ModStat(SID.MaxHIT))
    Hit = Hit - (NPCList(NPCIndex).ModStat(SID.DEF) * 0.5)
    If Hit < 1 Then Hit = 1
    Log "User_AttackNPC: Hit (damage) value calculated (" & Hit & ")", CodeTracker '//\\LOGLINE//\\

    'Hurt the NPC
    NPC_Damage NPCIndex, UserIndex, Hit

End Sub

Sub User_AttackUser(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

'*****************************************************************
'Have a user attack a user
'*****************************************************************

Dim Hit As Integer

    Log "Call User_AttackUser(" & AttackerIndex & "," & VictimIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Don't allow if switchingmaps maps
    If UserList(VictimIndex).Flags.SwitchingMaps Then
        Log "User_AttackUser: Victim switching maps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Calculate hit
    Hit = Server_RandomNumber(UserList(AttackerIndex).Stats.ModStat(SID.MinHIT), UserList(AttackerIndex).Stats.ModStat(SID.MaxHIT))
    Hit = Hit - (UserList(VictimIndex).Stats.ModStat(SID.DEF) / 2)
    If Hit < 1 Then Hit = 1

    'Hit User
    UserList(VictimIndex).Stats.BaseStat(SID.MinHP) = UserList(VictimIndex).Stats.BaseStat(SID.MinHP) - Hit

    'Play the attack animation
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Attack
    ConBuf.Put_Integer UserList(AttackerIndex).Char.CharIndex
    Data_Send ToPCArea, AttackerIndex, ConBuf.Get_Buffer

    'User Die
    If UserList(VictimIndex).Stats.BaseStat(SID.MinHP) <= 0 Then
        Log "User_AttackUser: Killed the user", CodeTracker '//\\LOGLINE//\\

        'Kill user
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 22
        ConBuf.Put_String UserList(VictimIndex).Name
        Data_Send ToIndex, AttackerIndex, ConBuf.Get_Buffer

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 23
        ConBuf.Put_String UserList(AttackerIndex).Name
        Data_Send ToIndex, VictimIndex, ConBuf.Get_Buffer

        User_Kill VictimIndex

    End If

End Sub

Sub User_ChangeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal UserIndex As Integer, Optional ByVal Body As Integer = -1, Optional ByVal Head As Integer = -1, Optional ByVal Heading As Byte = 0, Optional ByVal Weapon As Integer = -1, Optional ByVal Hair As Integer = -1, Optional ByVal Wings As Integer = -1)

'*****************************************************************
'Changes a user char's head,body and heading
'*****************************************************************
Dim ChangeFlags As Byte

    Log "Call User_ChangeChar(" & sndRoute & "," & sndIndex & "," & UserIndex & "," & Body & "," & Head & "," & Heading & "," & Weapon & "," & Hair & "," & Wings & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    If UserIndex > MaxUsers Then
        Log "User_ChangeChar: UserIndex > MaxUsers - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex <= 0 Then
        Log "User_ChangeChar: UserIndex <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for changed values
    With UserList(UserIndex).Char
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
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte ChangeFlags
    If ChangeFlags And 1 Then ConBuf.Put_Integer Body
    If ChangeFlags And 2 Then ConBuf.Put_Integer Head
    If ChangeFlags And 4 Then ConBuf.Put_Byte Heading
    If ChangeFlags And 8 Then ConBuf.Put_Integer Weapon
    If ChangeFlags And 16 Then ConBuf.Put_Integer Hair
    If ChangeFlags And 32 Then ConBuf.Put_Integer Wings
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_ChangeChar

End Sub

Sub User_ChangeInv(ByVal UserIndex As Integer, ByVal Slot As Byte, Object As UserOBJ)

'*****************************************************************
'Changes a user's inventory
'*****************************************************************

    Log "Call User_ChangeInv(" & UserIndex & "," & Slot & ",[I:" & Object.ObjIndex & " A:" & Object.Amount & " E:" & Object.Equipped & "])", CodeTracker '//\\LOGLINE//\\

    If Object.ObjIndex < 0 Then
        Log "User_ChangeInv: ObjIndex < 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Object.ObjIndex > NumObjDatas Then
        Log "User_ChangeInv: ObjIndex > NumObjDatas - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    UserList(UserIndex).Object(Slot) = Object
    
    'Build the buffer
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_SetInventorySlot
    ConBuf.Put_Byte Slot
    ConBuf.Put_Long Object.ObjIndex
    
    'If the object has an index, then send the related information of the object
    'If index = 0, then we assume we are deleting it
    If Object.ObjIndex Then
        ConBuf.Put_String ObjData(Object.ObjIndex).Name
        ConBuf.Put_Long Object.Amount
        ConBuf.Put_Byte Object.Equipped
        ConBuf.Put_Long ObjData(Object.ObjIndex).GrhIndex
    End If
    
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub User_DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Drops a object from a User's slot
'*****************************************************************
Dim Obj As Obj
Dim i As Byte
Dim NewX As Byte
Dim NewY As Byte
Dim NewSlot As Byte

    Log "Call User_DropObj(" & UserIndex & "," & Slot & "," & Num & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    On Error GoTo ErrOut
    If UserList(UserIndex).Pos.Map <= 0 Then
        Log "User_DropObj: Map <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Map > NumMaps Then
        Log "User_DropObj: Map > NumMaps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.X < XMinMapSize Then
        Log "User_DropObj: User X < XMinMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.X > XMaxMapSize Then
        Log "User_DropObj: User X > XMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y < YMinMapSize Then
        Log "User_DropObj: User Y < YMinMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y > YMaxMapSize Then
        Log "User_DropObj: User Y > YMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Log "User_DropObj: User switching maps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Slot <= 0 Then
        Log "User_DropObj: Slot <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Slot > MAX_INVENTORY_SLOTS Then
        Log "User_DropObj: Slot > MAX_INVENTORY_SLOTS - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Object(Slot).Amount <= 0 Then
        Log "User_DropObj: Object amount <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Num > UserList(UserIndex).Object(Slot).Amount Then
        Log "User_DropObj: Requested drop amount > User's object amount - aborting", CodeTracker '//\\LOGLINE//\\
        Num = UserList(UserIndex).Object(Slot).Amount
    End If
    If Num <= 0 Then
        Log "User_DropObj: Requested drop amount <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    
    'Get the closest free slot available
    Obj_ClosestFreeSpot UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NewX, NewY, NewSlot

    'Check if the new spot is valid
    If NewX = 0 Then
        Log "User_DropObj: No free position could be found", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 24
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Set up the object
    Obj.ObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
    Obj.Amount = Num
    Obj_Make Obj, NewSlot, UserList(UserIndex).Pos.Map, NewX, NewY

    'Remove object from the user
    UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - Num
    If UserList(UserIndex).Object(Slot).Amount <= 0 Then
        Log "User_DropObj: User dropped all of the item, removing item from client", CodeTracker '//\\LOGLINE//\\
    
        'Unequip if the object is currently equipped
        If UserList(UserIndex).Object(Slot).Equipped = 1 Then User_RemoveInvItem UserIndex, Slot

        UserList(UserIndex).Object(Slot).ObjIndex = 0
        UserList(UserIndex).Object(Slot).Amount = 0
        UserList(UserIndex).Object(Slot).Equipped = 0
        
    End If

    User_UpdateInv False, UserIndex, Slot
    
    Exit Sub    '//\\LOGLINE//\\
    
ErrOut:

    Log "User_DropObj: Unexpected error in User_DropObj - GoTo ErrOut called!", CriticalError '//\\LOGLINE//\\

End Sub

Sub User_EraseChar(ByVal UserIndex As Integer)

'*****************************************************************
'Erase a character
'*****************************************************************

    Log "Call User_EraseChar(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    On Error GoTo ErrOut
    If UserList(UserIndex).Pos.Map <= 0 Then
        Log "User_EraseChar: Map <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Map > NumMaps Then
        Log "User_EraseChar: Map > NumMaps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    
    'Send erase command to clients
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_EraseChar
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    
    'Remove from list
    CharList(UserList(UserIndex).Char.CharIndex).Index = 0
    CharList(UserList(UserIndex).Char.CharIndex).CharType = 0
    
    'Update userlist
    UserList(UserIndex).Char.CharIndex = 0
    
    If UserList(UserIndex).Pos.X < XMinMapSize Then
        Log "User_EraseChar: User X < XMinMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.X > XMaxMapSize Then
        Log "User_EraseChar: User X > XMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y < YMinMapSize Then
        Log "User_EraseChar: User Y < YMinMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y > YMaxMapSize Then
        Log "User_EraseChar: User Y > YMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Remove from map
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    
    Exit Sub '//\\LOGLINE//\\
    
ErrOut:

    Log "User_EraseChar: Unexpected error in User_EraseChar - GoTo ErrOut called!", CriticalError '//\\LOGLINE//\\

End Sub

Sub Obj_CleanMapTile(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

'*****************************************************************
'Removes all the unused obj slots on a map tile
'Make sure you call this every time you remove an object from a tile!
'*****************************************************************
Dim NumObjs As Byte
Dim i As Long
Dim j As Long

    Log "Call Obj_CleanMapTile(" & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure we wern't given an empty map tile
    If MapData(Map, X, Y).NumObjs = 0 Then
        Log "Obj_CleanMapTile: NumObjs = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check through all the object slots
    For i = 1 To MapData(Map, X, Y).NumObjs
        If MapData(Map, X, Y).ObjInfo(i).ObjIndex > 0 Then
            If MapData(Map, X, Y).ObjInfo(i).Amount > 0 Then
                
                'Object found, so raise the count
                NumObjs = NumObjs + 1
                
                'Move down in the array if possible
                If i > 1 Then   'We can't sort any lower then 1, so don't even try it
                    
                    'Loop through all the previous object slots
                    For j = 1 To (i - 1)    '(i - 1) is since we don't to check the slot it is already on!
                    
                        'If the object slot is unused then
                        If MapData(Map, X, Y).ObjInfo(j).ObjIndex = 0 Or MapData(Map, X, Y).ObjInfo(j).Amount = 0 Then
                
                            'Scoot the item's keester down to that slot (swap the used into the unused)
                            MapData(Map, X, Y).ObjInfo(j) = MapData(Map, X, Y).ObjInfo(i)
                            MapData(Map, X, Y).ObjLife(j) = MapData(Map, X, Y).ObjLife(i)
                            
                            'Clear the old object
                            ZeroMemory MapData(Map, X, Y).ObjInfo(i), Len(MapData(Map, X, Y).ObjInfo(i))
                            MapData(Map, X, Y).ObjLife(i) = 0
                        
                        End If
                        
                    Next j
                    
                End If
                
            End If
        End If
    Next i
    
    'Once all that code above has gone through, NumObjs should have the number of valid objects
    ' and the first object slots should be used (unused at the end), so if redim the array by
    ' the NumObjs value, all we will cut off is the unused slots! :)
    If NumObjs > 0 Then
        Log "Obj_CleanMapTile: Resizing ObjInfo() array (1 To " & NumObjs & ")", CodeTracker '//\\LOGLINE//\\
        ReDim Preserve MapData(Map, X, Y).ObjInfo(1 To NumObjs)
        ReDim Preserve MapData(Map, X, Y).ObjLife(1 To NumObjs)
    Else
        'We have no slots at all used, so kill the whole damn thing
        Log "Obj_CleanMapTile: Erasing ObjInfo() array", CodeTracker '//\\LOGLINE//\\
        Erase MapData(Map, X, Y).ObjInfo
        Erase MapData(Map, X, Y).ObjLife
    End If
    
    'Assign the value to the map array for later usage
    MapData(Map, X, Y).NumObjs = NumObjs

End Sub

Sub User_GetObj(ByVal UserIndex As Integer)

'*****************************************************************
'Puts a object in a User's slot from the current User's position
'*****************************************************************
Dim ObjSlot As Byte
Dim Slot As Byte
Dim Map As Integer
Dim X As Byte
Dim Y As Byte
Dim i As Long

    Log "Call User_GetObj(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Assign the values to some smaller variables
    Map = UserList(UserIndex).Pos.Map
    X = UserList(UserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y

    'Check for invalid values
    On Error GoTo ErrOut
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Log "User_GetObj: User switching maps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Map <= 0 Then
        Log "User_GetObj: User map <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Map > NumMaps Then
        Log "User_GetObj: User map > NumMaps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If X < XMinMapSize Then
        Log "User_GetObj: User X < XMinMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If X > XMaxMapSize Then
        Log "User_GetObj: User X > XMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Y < YMinMapSize Then
        Log "User_GetObj: User Y < YMinMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Y > YMaxMapSize Then
        Log "User_GetObj: User Y > YMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    
    'No objects exist on the tile
    If MapData(Map, X, Y).NumObjs = 0 Then
        Log "User_GetObj: NumObjs on tile (" & Map & "," & X & "," & Y & ") = 0", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 25
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check for object on ground
    For i = 1 To MapData(Map, X, Y).NumObjs
        If MapData(Map, X, Y).ObjInfo(i).ObjIndex > 0 Then
            ObjSlot = i
            Exit For
        End If
    Next i
    Log "User_GetObj: ObjSlot = " & ObjSlot, CodeTracker '//\\LOGLINE//\\
    
    'For some reason, the NumObjs value is > 0 but there are no objects - no objs found
    If ObjSlot = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 25
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
        'Also request a cleaning of the map object array since it is obviously messy
        Obj_CleanMapTile Map, X, Y
        
        Exit Sub
    End If
    
    'Check to see if User already has the object type
    Slot = 1
    Do Until UserList(UserIndex).Object(Slot).ObjIndex = MapData(Map, X, Y).ObjInfo(ObjSlot).ObjIndex
        Slot = Slot + 1
        If Slot > MAX_INVENTORY_SLOTS Then
            Log "User_GetObj: Slot > MAX_INVENTORY_SLOTS", CodeTracker '//\\LOGLINE//\\
            Exit Do
        End If
    Loop

    'Else check if there is a empty slot
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 26
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        Loop
    End If

    'Fill object slot
    If UserList(UserIndex).Object(Slot).Amount + MapData(Map, X, Y).ObjInfo(ObjSlot).Amount <= MAX_INVENTORY_OBJS Then

        'Tell the user they recieved the items
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 27
        ConBuf.Put_Integer MapData(Map, X, Y).ObjInfo(ObjSlot).Amount
        ConBuf.Put_String ObjData(MapData(Map, X, Y).ObjInfo(ObjSlot).ObjIndex).Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'User takes all the items
        UserList(UserIndex).Object(Slot).ObjIndex = MapData(Map, X, Y).ObjInfo(ObjSlot).ObjIndex
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + MapData(Map, X, Y).ObjInfo(ObjSlot).Amount
        Obj_Erase MapData(Map, X, Y).ObjInfo(ObjSlot).Amount, ObjSlot, Map, X, Y

    Else
        'Over MAX_INV_OBJS
        If MapData(Map, X, Y).ObjInfo(Slot).Amount < UserList(UserIndex).Object(Slot).Amount Then
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 27
            ConBuf.Put_Integer Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + MapData(Map, X, Y).ObjInfo(Slot).Amount))
            ConBuf.Put_String ObjData(MapData(Map, X, Y).ObjInfo(Slot).ObjIndex).Name
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            MapData(Map, X, Y).ObjInfo(Slot).Amount = Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + MapData(Map, X, Y).ObjInfo(Slot).Amount))
        Else
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 27
            ConBuf.Put_Integer Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - MapData(Map, X, Y).ObjInfo(Slot).Amount)
            ConBuf.Put_String ObjData(MapData(Map, X, Y).ObjInfo(Slot).ObjIndex).Name
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            MapData(Map, X, Y).ObjInfo(Slot).Amount = Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - MapData(Map, X, Y).ObjInfo(Slot).Amount)
        End If
        UserList(UserIndex).Object(Slot).Amount = MAX_INVENTORY_OBJS
    End If

    'Update the user's inventory
    Call User_UpdateInv(False, UserIndex, Slot)

ErrOut:

End Sub

Sub User_GiveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer)

'*****************************************************************
'Give the user an object
'*****************************************************************

Dim Slot As Byte

    Log "Call User_GiveObj(" & UserIndex & "," & ObjIndex & "," & Amount & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    If UserIndex <= 0 Then
        Log "User_GiveObj: UserIndex <= 0", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If ObjIndex <= 0 Then
        Log "User_GiveObj: ObjIndex <= 0", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex > LastUser Then
        Log "User_GiveObj: UserIndex > LastUser", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If ObjIndex > NumObjDatas Then
        Log "User_GiveObj: ObjIndex > NumObjDatas", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Check to see if User already has object type
    Slot = 1
    Do Until UserList(UserIndex).Object(Slot).ObjIndex = ObjIndex
        Slot = Slot + 1
        If Slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop

    'Else check if there is a empty slot
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Object(Slot).ObjIndex = 0
            Log "User_GiveObj: Checking slot " & Slot, CodeTracker '//\\LOGLINE//\\
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Log "User_GiveObj: Slot > MAX_INVENTORY_SLOTS", CodeTracker '//\\LOGLINE//\\
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 26
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        Loop
    End If

    'Fill object slot
    If UserList(UserIndex).Object(Slot).Amount + Amount <= MAX_INVENTORY_OBJS Then

        'Tell the user they recieved the items
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 28
        ConBuf.Put_Integer Amount
        ConBuf.Put_String ObjData(ObjIndex).Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'User takes all the items
        UserList(UserIndex).Object(Slot).ObjIndex = ObjIndex
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + Amount

    Else
        'Over MAX_INV_OBJS
        If Amount < UserList(UserIndex).Object(Slot).Amount Then
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 28
            ConBuf.Put_Integer Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + Amount))
            ConBuf.Put_String ObjData(ObjIndex).Name
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Else
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 28
            ConBuf.Put_Integer Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - Amount)
            ConBuf.Put_String ObjData(ObjIndex).Name
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        End If
        UserList(UserIndex).Object(Slot).Amount = MAX_INVENTORY_OBJS
    End If

    'Update the user's inventory
    User_UpdateInv False, UserIndex, Slot

End Sub

Public Function User_IndexFromSox(inSox As Long) As Integer

'*****************************************************************
'Find UserIndex from given inSox ID
'*****************************************************************

Dim i As Long

    Log "Call User_IndexFromSox(" & inSox & ")", CodeTracker '//\\LOGLINE//\\

    Do
        i = i + 1
        If i > LastUser Then
            User_IndexFromSox = -1
            Log "Rtrn User_IndexFromSox = " & User_IndexFromSox, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    Loop While UserList(i).ConnID <> inSox
    User_IndexFromSox = i
    
    Log "Rtrn User_IndexFromSox = " & User_IndexFromSox, CodeTracker '//\\LOGLINE//\\

End Function

Sub User_Kill(ByVal UserIndex As Integer)

'*****************************************************************
'Kill a user
'*****************************************************************

Dim TempPos As WorldPos

    Log "Call User_Kill(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Set user health/mana/stamina back to full
    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = UserList(UserIndex).Stats.ModStat(SID.MaxHP)
    UserList(UserIndex).Stats.BaseStat(SID.MinMAN) = UserList(UserIndex).Stats.ModStat(SID.MaxMAN)
    UserList(UserIndex).Stats.BaseStat(SID.MinSTA) = UserList(UserIndex).Stats.ModStat(SID.MaxSTA)

    'Find a place to put user
    Call Server_ClosestLegalPos(ResPos, TempPos)
    If Server_LegalPos(TempPos.Map, TempPos.X, TempPos.Y, 0) = False Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 83
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        User_Close UserIndex
        Exit Sub
    End If
    
    'Remove the targeted NPC
    UserList(UserIndex).Flags.Target = 0
    UserList(UserIndex).Flags.TargetIndex = 0
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Target
    ConBuf.Put_Integer 0
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Warp him there
    User_WarpChar UserIndex, TempPos.Map, TempPos.X, TempPos.Y

End Sub

Sub User_LookAtTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Button As Byte)

'*****************************************************************
'Responds to the user clicking on a square
'*****************************************************************

Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim LoopC As Byte
Dim TempCharIndex As Integer
Dim TempIndex As Integer
Dim MsgData As MailData

    Log "Call User_LookAtTile(" & UserIndex & "," & Map & "," & X & "," & Y & "," & Button & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    On Error GoTo ErrOut
    If Not Server_InMapBounds(X, Y) Then
        Log "User_LookAtTile: Invalid tile looked at (X:" & X & " Y:" & Y & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex <= 0 Then
        Log "User_LookAtTile: UserIndex <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex > MaxUsers Then
        Log "User_LookAtTile: UserIndex > MaxUsers - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Map <= 0 Then
        Log "User_LookAtTile: Map <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Map > NumMaps Then
        Log "User_LookAtTile: Map > NumMaps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Log "User_LookAtTile: User witching maps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0

    '***** Right Click *****
    If Button = vbRightButton Then

        '*** Check for mailbox ***
        If MapData(Map, X, Y).Mailbox = 1 Then

            'Only check mail if right next to the mailbox
            If UserList(UserIndex).Pos.Map = Map Then
                If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y, 1, 1) Then

                    'Store the position of the mailbox for later reference in case user tries to use items away from mailbox
                    UserList(UserIndex).MailboxPos.Map = Map
                    UserList(UserIndex).MailboxPos.X = X
                    UserList(UserIndex).MailboxPos.Y = Y

                    'Resend all the mail
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_MailBox
                    For LoopC = 1 To MaxMailPerUser
                        If UserList(UserIndex).MailID(LoopC) > 0 Then
                            MsgData = Load_Mail(UserList(UserIndex).MailID(LoopC))
                            ConBuf.Put_Byte MsgData.New
                            ConBuf.Put_String MsgData.WriterName
                            ConBuf.Put_String CStr(MsgData.RecieveDate)
                            ConBuf.Put_String MsgData.Subject
                        End If
                    Next LoopC
                    ConBuf.Put_Byte 255 'The byte of value 255 states that we have reached the end, while 0 or 1 means it is a new message (states the "New" flag)
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    Exit Sub
                End If
            End If

            'User isn't next to the mailbox
            Log "User_LookAtTile: User not next to mailbox", CodeTracker '//\\LOGLINE//\\
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 29
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub

        End If

        '*** Check for Characters ***
        If Y + 1 <= YMaxMapSize Then
            If MapData(Map, X, Y + 1).UserIndex > 0 Then
                TempIndex = MapData(Map, X, Y + 1).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, X, Y + 1).NPCIndex > 0 Then
                TempIndex = MapData(Map, X, Y + 1).NPCIndex
                FoundChar = 2
            End If
        End If
        'Check for Character
        If FoundChar = 0 Then
            If MapData(Map, X, Y).UserIndex > 0 Then
                TempIndex = MapData(Map, X, Y).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, X, Y).NPCIndex > 0 Then
                TempIndex = MapData(Map, X, Y).NPCIndex
                FoundChar = 2
            End If
        End If
        'React to character
        If FoundChar = 1 Then
            If Len(UserList(TempIndex).Desc) > 1 Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 30
                ConBuf.Put_String UserList(TempIndex).Name
                ConBuf.Put_String UserList(TempIndex).Desc
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 31
                ConBuf.Put_String UserList(TempIndex).Name
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            FoundSomething = 1
        End If
        If FoundChar = 2 Then
            FoundSomething = 1
            '*** Check for NPC vendor ***
            If NPCList(TempIndex).NumVendItems > 0 Then
                User_TradeWithNPC UserIndex, TempIndex
                FoundSomething = 1
            Else
                '*** NPC not a vendor, give description ***
                If Len(NPCList(TempIndex).Name) > 1 Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 30
                    ConBuf.Put_String NPCList(TempIndex).Name
                    ConBuf.Put_String NPCList(TempIndex).Desc
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Else
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 31
                    ConBuf.Put_String NPCList(TempIndex).Name
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                End If
                '*** Quest NPC ***
                If NPCList(TempIndex).Quest > 0 Then Quest_General UserIndex, TempIndex
            End If
        End If

        '*** Check for object ***
        If MapData(Map, X, Y).NumObjs > 0 Then
            For LoopC = 1 To MapData(Map, X, Y).NumObjs
                If MapData(Map, X, Y).ObjInfo(LoopC).ObjIndex > 0 Then
                    'Check whether to use the singular or plural message
                    If MapData(Map, X, Y).ObjInfo(LoopC).Amount = 1 Then
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 32
                        ConBuf.Put_String ObjData(MapData(Map, X, Y).ObjInfo(LoopC).ObjIndex).Name
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    Else
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 86
                        ConBuf.Put_String ObjData(MapData(Map, X, Y).ObjInfo(LoopC).ObjIndex).Name
                        ConBuf.Put_Integer MapData(Map, X, Y).ObjInfo(LoopC).Amount
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    End If
                    FoundSomething = 1
                End If
            Next LoopC
            'If we went through that loop and FoundSomething = 0, then NumObjs > 0 while there are
            ' no objects, so we need to clean the array since it didn't get cleaned
            Obj_CleanMapTile Map, X, Y
        End If

        '*** Didn't find anything ***
        If FoundSomething = 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 33
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        End If

        '***** Left Click *****
    ElseIf Button = vbLeftButton Then

        '*** Look for NPC/Player to target ***
        If Y + 1 <= YMaxMapSize Then
            If MapData(Map, X, Y + 1).UserIndex > 0 Then
                TempCharIndex = UserList(MapData(Map, X, Y + 1).UserIndex).Char.CharIndex
                TempIndex = MapData(Map, X, Y + 1).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, X, Y + 1).NPCIndex > 0 Then
                TempCharIndex = NPCList(MapData(Map, X, Y + 1).NPCIndex).Char.CharIndex
                TempIndex = MapData(Map, X, Y + 1).NPCIndex
                FoundChar = 2
            End If
        End If
        If FoundChar = 0 Then
            If MapData(Map, X, Y).UserIndex > 0 Then
                TempCharIndex = UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex
                TempIndex = MapData(Map, X, Y).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, X, Y).NPCIndex > 0 Then
                TempCharIndex = NPCList(MapData(Map, X, Y).NPCIndex).Char.CharIndex
                TempIndex = MapData(Map, X, Y).NPCIndex
                FoundChar = 2
            End If
        End If

        'Validate distance
        If FoundChar = 0 Then
            If UserList(UserIndex).Flags.Target Then
                UserList(UserIndex).Flags.Target = 0
                UserList(UserIndex).Flags.TargetIndex = 0
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer 0
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            Exit Sub
        ElseIf FoundChar = 1 Then
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(TempIndex).Pos.X, UserList(TempIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                UserList(UserIndex).Flags.Target = 1
                UserList(UserIndex).Flags.TargetIndex = TempCharIndex
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer UserList(TempIndex).Char.CharIndex
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                If UserList(UserIndex).Flags.Target Then
                    UserList(UserIndex).Flags.Target = 0
                    UserList(UserIndex).Flags.TargetIndex = 0
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.User_Target
                    ConBuf.Put_Integer 0
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                End If
            End If
        ElseIf FoundChar = 2 Then
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(TempIndex).Pos.X, NPCList(TempIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                UserList(UserIndex).Flags.Target = 2
                UserList(UserIndex).Flags.TargetIndex = TempCharIndex
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer NPCList(TempIndex).Char.CharIndex
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                If UserList(UserIndex).Flags.Target Then
                    UserList(UserIndex).Flags.Target = 0
                    UserList(UserIndex).Flags.TargetIndex = 0
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.User_Target
                    ConBuf.Put_Integer 0
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                End If
            End If
        End If

    End If
    
ErrOut:

End Sub

Sub User_MakeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Makes and places a user's character
'*****************************************************************

Dim CharIndex As Integer

    Log "Call User_MakeChar(" & sndRoute & "," & sndIndex & "," & UserIndex & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

'Place character on map

    MapData(Map, X, Y).UserIndex = UserIndex

    'Give it a char if needed
    If UserList(UserIndex).Char.CharIndex = 0 Then
        CharIndex = Server_NextOpenCharIndex
        UserList(UserIndex).Char.CharIndex = CharIndex
        CharList(CharIndex).Index = UserIndex
        CharList(CharIndex).CharType = CharType_PC
    End If

    'Send make character command to clients
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_MakeChar
    ConBuf.Put_Integer UserList(UserIndex).Char.Body
    ConBuf.Put_Integer UserList(UserIndex).Char.Head
    ConBuf.Put_Byte UserList(UserIndex).Char.Heading
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte X
    ConBuf.Put_Byte Y
    ConBuf.Put_Byte UserList(UserIndex).Stats.ModStat(SID.Speed)
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_Integer UserList(UserIndex).Char.Weapon
    ConBuf.Put_Integer UserList(UserIndex).Char.Hair
    ConBuf.Put_Integer UserList(UserIndex).Char.Wings
    ConBuf.Put_Byte UserList(UserIndex).Stats.LastHPPercent
    ConBuf.Put_Byte UserList(UserIndex).Stats.LastMPPercent

    If UserList(UserIndex).Skills.Bless > 0 Then
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Else
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.Protect > 0 Then
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Else
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.IronSkin > 0 Then
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Else
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.Strengthen > 0 Then
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Else
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.WarCurse > 0 Then
        ConBuf.Put_Byte DataCode.Server_IconWarCursed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Else
        ConBuf.Put_Byte DataCode.Server_IconWarCursed
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If

    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, Map

End Sub

Sub User_MoveChar(ByVal UserIndex As Integer, ByVal nHeading As Byte, ByVal Running As Byte)

'*****************************************************************
'Moves a User from one tile to another
'*****************************************************************
Dim TempIndex As Integer
Dim nPos As WorldPos
Dim i As Long

    Log "Call User_MoveChar(" & UserIndex & "," & nHeading & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    If UserIndex <= 0 Then
        Log "User_MoveChar: UserIndex <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex > MaxUsers Then
        Log "User_MoveChar: UserIndex > MaxUsers - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If nHeading = 0 Then
        Log "User_MoveChar: nHeading = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If nHeading > 8 Then
        Log "User_MoveChar: nHeading > 8 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Update the move counter
    If timeGetTime < UserList(UserIndex).Counters.MoveCounter Then
        Log "User_MoveChar: Not enough time has elapsed for movement.", CodeTracker '//\\LOGLINE//\\

        'If the user is moving too fast, then put them back
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub

    End If
    
    'Check the running value
    If Running > 0 Then
        If UserList(UserIndex).Stats.BaseStat(SID.MinSTA) < RunningCost Then
            Running = 0
        Else
            Running = 1 'Make sure the value is 1, no higher
            UserList(UserIndex).Stats.BaseStat(SID.MinSTA) = UserList(UserIndex).Stats.BaseStat(SID.MinSTA) - RunningCost
        End If
    End If
    
    'Clear the pending quest NPC number
    UserList(UserIndex).Flags.QuestNPC = 0

    'Get the new time
    UserList(UserIndex).Counters.MoveCounter = timeGetTime + Server_WalkTimePerTile(UserList(UserIndex).Stats.ModStat(SID.Speed) + Running * RunningSpeed)

    'Get the new position
    nPos = UserList(UserIndex).Pos
    Server_HeadToPos nHeading, nPos

    'Move if legal pos
    If Server_LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, nHeading) = True Then
    
        'Send the movement
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MoveChar
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte nPos.X
        ConBuf.Put_Byte nPos.Y
        If Running Then
            ConBuf.Put_Byte nHeading Or 128
        Else
            ConBuf.Put_Byte nHeading
        End If
        Data_Send ToUserMove, UserIndex, ConBuf.Get_Buffer

        'Update map and user pos
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.Heading = nHeading
        UserList(UserIndex).Char.HeadHeading = nHeading
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex

        'If the user has movement packets in their buffer, force them to be sent asap since they could be queued packets
        ' from distant movements. If they are not queued packets, then they are going to be high priority anyways
        ' (unless you have changed the packet priorities lower, which for local movement, is not a good idea)
        If UserList(UserIndex).HasMovePacket Then UserList(UserIndex).PPValue = PP_High

        'Do tile events
        Server_DoTileEvents UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y

    End If

    'Make sure the user's position is correct
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_SetUserPosition
    ConBuf.Put_Byte UserList(UserIndex).Pos.X
    ConBuf.Put_Byte UserList(UserIndex).Pos.Y
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Function User_NameToIndex(ByVal Name As String) As Integer

'*****************************************************************
'Searches userlist for a name and return userindex
'*****************************************************************

Dim UserIndex As Integer

    Log "Call User_NameToIndex(" & Name & ")", CodeTracker '//\\LOGLINE//\\

    'Check for bad name
    If Len(Name) = 0 Then
        User_NameToIndex = 0
        Log "Rtrn User_NameToIndex = " & User_NameToIndex, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Find the user
    UserIndex = 1
    Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
        UserIndex = UserIndex + 1
        If UserIndex > LastUser Then
            Log "User_NameToIndex: UserIndex > LastUser", CodeTracker '//\\LOGLINE//\\
            UserIndex = 0
            Log "Rtrn User_NameToIndex = " & User_NameToIndex, CodeTracker '//\\LOGLINE//\\
            Exit Do
        End If
    Loop
    
    'Return the results
    User_NameToIndex = UserIndex
    
    Log "Rtrn User_NameToIndex = " & User_NameToIndex, CodeTracker '//\\LOGLINE//\\

End Function

Function User_NextOpen() As Integer

'*****************************************************************
'Finds the next open UserIndex in UserList
'*****************************************************************

    Log "Call User_NextOpen", CodeTracker '//\\LOGLINE//\\

    'Check if a user is logging into an empty server
    If LastUser = 0 Then
        User_NextOpen = 1
        LastUser = 1
        ReDim UserList(1 To 1)
        Log "Rtrn User_NextOpen = " & User_NextOpen, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If

    'Start with the first index
    User_NextOpen = 1

    'Find the next free user slot
    Do Until UserList(User_NextOpen).Flags.UserLogged = 0
        
        'Raise the value to check the next index
        User_NextOpen = User_NextOpen + 1
        
        'Check if we have gone over the maximum users allowed
        If User_NextOpen > MaxUsers Then
            Log "Rtrn User_NextOpen = " & User_NextOpen, CodeTracker '//\\LOGLINE//\\
            Exit Do
        End If
        
        'Check if we have exceeded the current array size (every slot is used)
        If User_NextOpen > UBound(UserList()) Then
            LastUser = User_NextOpen
            ReDim Preserve UserList(1 To User_NextOpen)
            Exit Do
        End If
        
    Loop

    Log "Rtrn User_NextOpen = " & User_NextOpen, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub User_RaiseExp(ByVal UserIndex As Integer, ByVal EXP As Long)

'*****************************************************************
'Raise the user's experience - this should be the only way exp is raised!
'*****************************************************************

Dim Levels As Integer

    Log "Call User_RaiseExp(" & UserIndex & "," & EXP & ")", CodeTracker '//\\LOGLINE//\\

    'Update the user's experience
    UserList(UserIndex).Stats.BaseStat(SID.EXP) = UserList(UserIndex).Stats.BaseStat(SID.EXP) + EXP

    'Loop as many times as needed to get every level gained in
    Do While UserList(UserIndex).Stats.BaseStat(SID.EXP) >= UserList(UserIndex).Stats.BaseStat(SID.ELU)
        Log "User_RaiseExp: User by index " & UserIndex & " (" & UserList(UserIndex).Name & ") leveled up", CodeTracker '//\\LOGLINE//\\
        
        'Set the number of levels gained
        Levels = Levels + 1
    
        'Raise stats / level requirements
        With UserList(UserIndex).Stats
            .BaseStat(SID.MinHIT) = .BaseStat(SID.MinHIT) + 1
            .BaseStat(SID.MaxHIT) = .BaseStat(SID.MaxHIT) + 1
            .BaseStat(SID.MaxHP) = .BaseStat(SID.MaxHP) + 10
            .BaseStat(SID.MaxMAN) = .BaseStat(SID.MaxMAN) + 10
            .BaseStat(SID.MaxSTA) = .BaseStat(SID.MaxSTA) + 10
            .BaseStat(SID.Points) = .BaseStat(SID.Points) + 5
            
            .BaseStat(SID.ELV) = .BaseStat(SID.ELV) + 1
            .BaseStat(SID.EXP) = .BaseStat(SID.EXP) - .BaseStat(SID.ELU)
            .BaseStat(SID.ELU) = .BaseStat(SID.ELV) * 5
        End With

    Loop

    'Check if needing to update from leveling
    If Levels = 1 Then
        Log "User_RaiseExp: User gained a level", CodeTracker '//\\LOGLINE//\\

        'Say the user's level raised
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 34
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    ElseIf Levels > 1 Then
        Log "User_RaiseExp: User gained multiple levels (" & Levels & ")", CodeTracker '//\\LOGLINE//\\

        'Say the user's level raised
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 35
        ConBuf.Put_Byte Levels
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    End If

End Sub

Sub User_RemoveInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Unequip a inventory item
'*****************************************************************

Dim Obj As ObjData

    Log "Call User_RemoveInvItem(" & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

'Set the object

    Obj = ObjData(UserList(UserIndex).Object(Slot).ObjIndex)

    'Get the object type
    Select Case Obj.ObjType

        'Check for weapon
    Case OBJTYPE_WEAPON
        Log "User_RemoveInvItem: Object type OBJTYPE_WEAPON", CodeTracker '//\\LOGLINE//\\
        
        'Update the weapon distance on the client
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_SetWeaponRange
        ConBuf.Put_Byte 0
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
        'Set the equipted variables
        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).WeaponEqpObjIndex = 0
        UserList(UserIndex).WeaponEqpSlot = 0
        UserList(UserIndex).Char.Weapon = 0
        UserList(UserIndex).WeaponType = Hand
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings

        'Check for armor
    Case OBJTYPE_ARMOR
        Log "User_RemoveInvItem: Object type OBJTYPE_ARMOR", CodeTracker '//\\LOGLINE//\\

        'Set the equipted variables
        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).ArmorEqpObjIndex = 0
        UserList(UserIndex).ArmorEqpSlot = 0
        UserList(UserIndex).Char.Body = 1
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings

        'Check for wings
    Case OBJTYPE_WINGS
        Log "User_RemoveInvItem: Object type OBJTYPE_WINGS", CodeTracker '//\\LOGLINE//\\
    
        'Set the equipted variables
        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).WingsEqpObjIndex = 0
        UserList(UserIndex).WingsEqpSlot = 0
        UserList(UserIndex).Char.Wings = 0
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings

    Case Else   '//\\LOGLINE//\\
        Log "User_RemoveInvItem: Unknown object type! Object type: " & Obj.ObjType, CriticalError '//\\LOGLINE//\\

    End Select
    
    'Force update of the modstats
    UserList(UserIndex).Stats.Update = 1

    'Update the user's stats
    User_UpdateInv False, UserIndex, Slot

End Sub

Public Sub User_SendKnownSkills(ByVal UserIndex As Integer)
Dim KnowSkillList() As Byte
Dim Index As Long   'Which KnowSkillList array index to use
Dim i As Byte

    Log "Call User_SendKnownSkills(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for a valid userindex
    If UserIndex <= 0 Then
        Log "User_SendKnownSkills: User index <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex > MaxUsers Then
        Log "User_SendKnownSkills: User index > Max users - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Size the knowskilllist
    Log "User_SendKnownSkills: ReDim KnowSkillList(1 To " & NumBytesForSkills & ")", CodeTracker '//\\LOGLINE//\\
    ReDim KnowSkillList(1 To NumBytesForSkills)

    'Compile all the known skills into a long (or array of longs if too many skills)
    For i = 1 To NumSkills
        
        'Check if the skill is known
        Log "User_SendKnownSkills: Checking if skill ID " & i & " is known", CodeTracker '//\\LOGLINE//\\
        If UserList(UserIndex).KnownSkills(i) Then
            
            'Find out which KnowSkillList array index to use
            Log "User_SendKnownSkills: Index value = " & Int((i - 1) / 8) + 1, CodeTracker '//\\LOGLINE//\\
            Index = Int((i - 1) / 8) + 1
            
            'Pack the information
            KnowSkillList(Index) = KnowSkillList(Index) Or (2 ^ (i - ((Index - 1) * 8) - 1))

        End If
            
    Next i

    'Send the information to the user
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_KnownSkills
    For i = 1 To NumBytesForSkills
        ConBuf.Put_Byte KnowSkillList(i)
    Next i
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Public Sub User_TradeWithNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Start a trade with a NPC
'*****************************************************************

Dim LoopC As Integer

    Log "Call User_TradeWithNPC(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

'Trade with a NPC

    If NPCList(NPCIndex).NumVendItems > 0 Then
        Log "User_TradeWithNPC: NumVendItems = " & NPCList(NPCIndex).NumVendItems, CodeTracker '//\\LOGLINE//\\

        'Check if close enough to trade with
        If Not Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, 6, 6) Then
            Log "User_TradeWithNPC: Can not trade - user too far away", CodeTracker '//\\LOGLINE//\\
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 36
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub
        End If

        Log "User_TradeWithNPC: Building vending items list", CodeTracker '//\\LOGLINE//\\
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Trade_StartNPCTrade
        ConBuf.Put_String NPCList(NPCIndex).Name
        ConBuf.Put_Integer NPCList(NPCIndex).NumVendItems
        For LoopC = 1 To NPCList(NPCIndex).NumVendItems
            ConBuf.Put_Long ObjData(NPCList(NPCIndex).VendItems(LoopC).ObjIndex).GrhIndex
            ConBuf.Put_String ObjData(NPCList(NPCIndex).VendItems(LoopC).ObjIndex).Name
            ConBuf.Put_Long ObjData(NPCList(NPCIndex).VendItems(LoopC).ObjIndex).Price
        Next LoopC
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        UserList(UserIndex).Flags.TradeWithNPC = NPCIndex
        Exit Sub
    End If

End Sub

Sub User_UpdateInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Updates a User's inventory
'*****************************************************************

Dim NullObj As UserOBJ
Dim LoopC As Long

    Log "Call User_UpdateInv(" & UpdateAll & "," & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

'Update one slot

    If UpdateAll = False Then
        'Update User inventory
        If UserList(UserIndex).Object(Slot).ObjIndex > 0 Then
            User_ChangeInv UserIndex, Slot, UserList(UserIndex).Object(Slot)
        Else
            User_ChangeInv UserIndex, Slot, NullObj
        End If
    Else
        'Update every slot
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            'Update User invetory
            If UserList(UserIndex).Object(LoopC).ObjIndex Then
                Call User_ChangeInv(UserIndex, LoopC, UserList(UserIndex).Object(LoopC))
            End If
        Next LoopC
    End If

End Sub

Sub User_UpdateMap(ByVal UserIndex As Integer)

'*****************************************************************
'Updates a user with the place of all chars in the Map
'*****************************************************************

Dim Map As Integer
Dim X As Byte
Dim Y As Byte
Dim i As Long

    Log "Call User_UpdateMap(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Map = UserList(UserIndex).Pos.Map

    'Send user char's pos
    Log "User_UpdateMap: For X = 1 to " & UBound(MapUsers(Map).Index()), CodeTracker '//\\LOGLINE//\\
    For X = 1 To UBound(MapUsers(Map).Index())
        Call User_MakeChar(ToIndex, UserIndex, MapUsers(Map).Index(X), Map, UserList(MapUsers(Map).Index(X)).Pos.X, UserList(MapUsers(Map).Index(X)).Pos.Y)
    Next X

    'Place chars and objects
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            'NPC update
            If MapData(Map, X, Y).NPCIndex Then NPC_MakeChar ToIndex, UserIndex, MapData(Map, X, Y).NPCIndex, Map, X, Y
            
            'Object update
            If MapData(Map, X, Y).NumObjs > 0 Then
                For i = 1 To MapData(Map, X, Y).NumObjs
                    If MapData(Map, X, Y).ObjInfo(i).ObjIndex Then
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_MakeObject
                        ConBuf.Put_Long ObjData(MapData(Map, X, Y).ObjInfo(i).ObjIndex).GrhIndex
                        ConBuf.Put_Byte X
                        ConBuf.Put_Byte Y
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, Map
                    End If
                Next i
            End If

        Next X
    Next Y

End Sub

Public Sub User_UpdateModStats(ByVal UserIndex As Integer)

'*****************************************************************
'Set the user's mod stats based on base stats and equipted items
'*****************************************************************

Dim WeaponObj As ObjData
Dim ArmorObj As ObjData
Dim WingsObj As ObjData
Dim i As Integer

    Log "Call User_UpdateModStats(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    If UserList(UserIndex).Flags.UserLogged = 0 Then
        Log "User_UpdateModStats: UserLogged = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Set the equipted items
    If UserList(UserIndex).WeaponEqpObjIndex > 0 Then WeaponObj = ObjData(UserList(UserIndex).WeaponEqpObjIndex)
    If UserList(UserIndex).ArmorEqpObjIndex > 0 Then ArmorObj = ObjData(UserList(UserIndex).ArmorEqpObjIndex)
    If UserList(UserIndex).WingsEqpObjIndex > 0 Then WingsObj = ObjData(UserList(UserIndex).WingsEqpObjIndex)

    With UserList(UserIndex).Stats

        'Equipted items
        For i = 1 To NumStats
            If i <> SID.MinHP Then
                If i <> SID.MinMAN Then
                    If i <> SID.MinSTA Then
                        If i <> SID.Gold Then
                            If i <> SID.Points Then
                                If i <> SID.EXP Then
                                    If i <> SID.ELU Then
                                        If i <> SID.ELV Then
                                            Log "User_UpdateModStats: Updating ModStat ID " & i, CodeTracker '//\\LOGLINE//\\
                                            .ModStat(i) = .BaseStat(i) + WeaponObj.AddStat(i) + ArmorObj.AddStat(i) + WingsObj.AddStat(i)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        
        'War curse
        If UserList(UserIndex).Skills.WarCurse > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell WarCurse", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.Agi) = .ModStat(SID.Agi) - (UserList(UserIndex).Skills.WarCurse * 0.25)
            .ModStat(SID.DEF) = .ModStat(SID.DEF) - (UserList(UserIndex).Skills.WarCurse * 0.25)
            .ModStat(SID.Str) = .ModStat(SID.Str) - (UserList(UserIndex).Skills.WarCurse * 0.25)
            .ModStat(SID.Mag) = .ModStat(SID.Mag) - (UserList(UserIndex).Skills.WarCurse * 0.25)
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - (UserList(UserIndex).Skills.WarCurse * 0.25)
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - (UserList(UserIndex).Skills.WarCurse * 0.25)
            .ModStat(SID.WeaponSkill) = .ModStat(SID.WeaponSkill) - (UserList(UserIndex).Skills.WarCurse * 0.25)
        End If
        
        'Strengthen
        If UserList(UserIndex).Skills.Strengthen > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell Strengthen", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + UserList(UserIndex).Skills.Strengthen
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + UserList(UserIndex).Skills.Strengthen
        End If
        
        'Protection
        If UserList(UserIndex).Skills.Protect > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell Protection", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + UserList(UserIndex).Skills.Protect
        End If
        
        'Bless
        If UserList(UserIndex).Skills.Bless > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell Bless", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.Agi) = .ModStat(SID.Agi) + UserList(UserIndex).Skills.Bless * 0.5
            .ModStat(SID.Mag) = .ModStat(SID.Mag) + UserList(UserIndex).Skills.Bless * 0.5
            .ModStat(SID.Str) = .ModStat(SID.Str) + UserList(UserIndex).Skills.Bless * 0.5
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + UserList(UserIndex).Skills.Bless * 0.25
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + UserList(UserIndex).Skills.Bless * 0.25
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + UserList(UserIndex).Skills.Bless * 0.25
        End If
        
        'Iron skin
        If UserList(UserIndex).Skills.IronSkin > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell Iron Skin", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + UserList(UserIndex).Skills.IronSkin * 2
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - UserList(UserIndex).Skills.IronSkin * 1.5
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - UserList(UserIndex).Skills.IronSkin * 1.5
        End If
        
    End With
    
End Sub

Sub User_UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Use/Equip a inventory item
'*****************************************************************

Dim Obj As ObjData

    Log "Call User_UseInvItem(" & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    On Error GoTo ErrOut
    If UserList(UserIndex).Flags.UserLogged = 0 Then
        Log "User_UseInvItem: UserLogged = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex > MaxUsers Then
        Log "User_UseInvItem: UserIndex > MaxUsers - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserIndex <= 0 Then
        Log "User_UseInvItem: UserIndex <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Slot > MAX_INVENTORY_SLOTS Then
        Log "User_UseInvItem: Slot > MAX_INVENTORY_SLOTS - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Slot <= 0 Then
        Log "User_UseInvItem: Slot <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Flags.SwitchingMaps Then
        Log "User_UseInvItem: SwitchingMaps = Yes - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Object(Slot).ObjIndex < 0 Then
        Log "User_UseInvItem: ObjIndex < 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Object(Slot).ObjIndex > NumObjDatas Then
        Log "User_UseInvItem: ObjIndex > NumObjDatas - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    
    Obj = ObjData(UserList(UserIndex).Object(Slot).ObjIndex)

    'Apply the replenish values
    With UserList(UserIndex).Stats
        .BaseStat(SID.MinHP) = .BaseStat(SID.MinHP) + (.ModStat(SID.MaxHP) * Obj.RepHPP) + Obj.RepHP
        .BaseStat(SID.MinMAN) = .BaseStat(SID.MinMAN) + (.ModStat(SID.MaxMAN) * Obj.RepMPP) + Obj.RepMP
        .BaseStat(SID.MinSTA) = .BaseStat(SID.MinSTA) + (.ModStat(SID.MaxSTA) * Obj.RepSPP) + Obj.RepSP
    End With

    Select Case Obj.ObjType
    
    Case OBJTYPE_USEONCE
        Log "User_UseInvItem: ObjType = OBJTYPE_USEONCE", CodeTracker '//\\LOGLINE//\\

        'Remove from inventory
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - 1
        If UserList(UserIndex).Object(Slot).Amount <= 0 Then UserList(UserIndex).Object(Slot).ObjIndex = 0
        
        'Set the paper-doll
        If Obj.SpriteHair <> -1 Then UserList(UserIndex).Char.Hair = Obj.SpriteHair
        If Obj.SpriteBody <> -1 Then UserList(UserIndex).Char.Body = Obj.SpriteBody
        If Obj.SpriteHead <> -1 Then UserList(UserIndex).Char.Head = Obj.SpriteHead
        If Obj.SpriteWeapon <> -1 Then UserList(UserIndex).Char.Weapon = Obj.SpriteWeapon
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings
        
    Case OBJTYPE_WEAPON
        Log "User_UseInvItem: ObjType = OBJTYPE_WEAPON", CodeTracker '//\\LOGLINE//\\

        'If currently equipped remove instead
        If UserList(UserIndex).Object(Slot).Equipped Then
            User_RemoveInvItem UserIndex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(UserIndex).WeaponEqpObjIndex > 0 Then User_RemoveInvItem UserIndex, UserList(UserIndex).WeaponEqpSlot

        'Equip
        UserList(UserIndex).Object(Slot).Equipped = 1
        UserList(UserIndex).WeaponEqpObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
        UserList(UserIndex).WeaponEqpSlot = Slot
        UserList(UserIndex).WeaponType = Obj.WeaponType
        
        'Update the weapon distance on the client
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_SetWeaponRange
        ConBuf.Put_Byte ObjData(UserList(UserIndex).WeaponEqpObjIndex).WeaponRange
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'Set the paper-doll
        If Obj.SpriteHair <> -1 Then UserList(UserIndex).Char.Hair = Obj.SpriteHair
        If Obj.SpriteBody <> -1 Then UserList(UserIndex).Char.Body = Obj.SpriteBody
        If Obj.SpriteHead <> -1 Then UserList(UserIndex).Char.Head = Obj.SpriteHead
        If Obj.SpriteWeapon <> -1 Then UserList(UserIndex).Char.Weapon = Obj.SpriteWeapon
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings

    Case OBJTYPE_ARMOR
        Log "User_UseInvItem: ObjType = OBJTYPE_ARMOR", CodeTracker '//\\LOGLINE//\\

        'If currently equipped remove instead
        If UserList(UserIndex).Object(Slot).Equipped Then
            User_RemoveInvItem UserIndex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(UserIndex).ArmorEqpObjIndex > 0 Then User_RemoveInvItem UserIndex, UserList(UserIndex).ArmorEqpSlot

        'Equip
        UserList(UserIndex).Object(Slot).Equipped = 1
        UserList(UserIndex).ArmorEqpObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
        UserList(UserIndex).ArmorEqpSlot = Slot

        'Set the paper-doll
        If Obj.SpriteHair <> -1 Then UserList(UserIndex).Char.Hair = Obj.SpriteHair
        If Obj.SpriteBody <> -1 Then UserList(UserIndex).Char.Body = Obj.SpriteBody
        If Obj.SpriteHead <> -1 Then UserList(UserIndex).Char.Head = Obj.SpriteHead
        If Obj.SpriteWeapon <> -1 Then UserList(UserIndex).Char.Weapon = Obj.SpriteWeapon
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings

    Case OBJTYPE_WINGS
        Log "User_UseInvItem: ObjType = OBJTYPE_WINGS", CodeTracker '//\\LOGLINE//\\
    
        'If currently equipped remove instead
        If UserList(UserIndex).Object(Slot).Equipped Then
            User_RemoveInvItem UserIndex, Slot
            Exit Sub
        End If

        'Remove old item if exists
        If UserList(UserIndex).WingsEqpObjIndex > 0 Then User_RemoveInvItem UserIndex, UserList(UserIndex).WingsEqpSlot

        'Equip
        UserList(UserIndex).Object(Slot).Equipped = 1
        UserList(UserIndex).WingsEqpObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
        UserList(UserIndex).WingsEqpSlot = Slot

        'Set the paper-doll
        If Obj.SpriteWings <> -1 Then UserList(UserIndex).Char.Wings = Obj.SpriteWings
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings

    Case Else
    
        'We have no idea what type of object it is! OMG!!!
        Log "User_UseInvItem: Unknown object type used! Object type: " & Obj.ObjType, CriticalError '//\\LOGLINE//\\

    End Select
    
    'Force update of the modstats
    UserList(UserIndex).Stats.Update = 1

    'Update user's stats and inventory
    User_UpdateInv False, UserIndex, Slot
    
ErrOut:

End Sub

Sub User_WarpChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal ForceSwitch As Boolean = False)

'*****************************************************************
'Warps user to another spot
'*****************************************************************

Dim OldMap As Integer
Dim LoopC As Long

    Log "Call User_WarpChar(" & UserIndex & "," & Map & "," & X & "," & Y & "," & ForceSwitch & ")", CodeTracker '//\\LOGLINE//\\

    OldMap = UserList(UserIndex).Pos.Map

    If OldMap <= 0 Then
        Log "User_WarpChar: OldMap <= 0", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If OldMap > NumMaps Then
        Log "User_WarpChar: OldMap > NumMaps", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    User_EraseChar UserIndex

    UserList(UserIndex).Pos.X = X
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.Map = Map

    If (OldMap <> Map) Or ForceSwitch = True Then
        Log "User_WarpChar: Switching maps", CodeTracker '//\\LOGLINE//\\
    
        'Check to update the database
        If MySQLUpdate_UserMap Then
            Log "User_WarpChar: Updating database with new map", CodeTracker '//\\LOGLINE//\\
            DB_RS.Open "SELECT * FROM users WHERE `name`='" & UserList(UserIndex).Name & "'", DB_Conn, adOpenStatic, adLockOptimistic
            DB_RS!pos_map = Map
            DB_RS.Update
            DB_RS.Close
        End If
        
        'Set switchingmap flag
        UserList(UserIndex).Flags.SwitchingMaps = 1

        'Tell client to try switching maps
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Map_LoadMap
        ConBuf.Put_Integer Map
        ConBuf.Put_Integer MapInfo(Map).MapVersion
        ConBuf.Put_Byte MapInfo(Map).Weather

        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        Log "User_WarpChar: MapInfo(" & Map & ").NumUsers = " & MapInfo(Map).NumUsers, CodeTracker '//\\LOGLINE//\\
        
        'Check if it's the first user on the map
        If MapInfo(Map).NumUsers = 1 Then
            ReDim MapUsers(Map).Index(1 To 1)
        Else
            ReDim Preserve MapUsers(Map).Index(1 To MapInfo(Map).NumUsers)
        End If
        MapUsers(Map).Index(MapInfo(Map).NumUsers) = UserIndex

        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers Then
            'Find current pos within connection group
            For LoopC = 1 To MapInfo(OldMap).NumUsers + 1
                If MapUsers(OldMap).Index(LoopC) = UserIndex Then Exit For
            Next LoopC
            'Move the rest of the list backwards
            For LoopC = LoopC To MapInfo(OldMap).NumUsers
                MapUsers(OldMap).Index(LoopC) = MapUsers(OldMap).Index(LoopC + 1)
            Next LoopC
            'Resize the list
            ReDim Preserve MapUsers(OldMap).Index(1 To MapInfo(OldMap).NumUsers)
        Else
            ReDim MapUsers(OldMap).Index(0)
        End If

        'Show Character to others
        User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y
    Else
        Log "User_WarpChar: Moving user, map is not changing", CodeTracker '//\\LOGLINE//\\
        User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_UserCharIndex
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:48)  Decl: 1  Code: 2408  Total: 2409 Lines
':) CommentOnly: 361 (15%)  Commented: 5 (0.2%)  Empty: 499 (20.7%)  Max Logic Depth: 7
