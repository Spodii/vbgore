Attribute VB_Name = "GameLogic"
Option Explicit

Public Function Obj_ValidObjForClass(ByVal Class As Byte, ByVal ObjIndex As Integer) As Byte

'*****************************************************************
'Checks if an object, by the object index, is useable by the passed class
'*****************************************************************
    
    'Check if theres a defined class requirement
    If ObjData.ClassReq(ObjIndex) > 0 Then
        
        'If Class AND ClassReq is true, then we meet the requirements
        If Not (Class And ObjData.ClassReq(ObjIndex)) Then Exit Function
        
    End If
    
    'Requirements met
    Obj_ValidObjForClass = 1

End Function

Public Sub Obj_Erase(ByVal Num As Integer, ByVal ObjSlot As Byte, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Erase a object
'*****************************************************************

    Log "Call Obj_Erase(" & Num & "," & ObjSlot & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Check for a valid index
    If ObjSlot > MapInfo(Map).ObjTile(X, Y).NumObjs Then
        Log "Obj_Erase: Invalid ObjSlot specified (" & ObjSlot & ")", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check to erase every object
    If Num = -1 Then Num = MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount

    'Remove the amount
    Log "Obj_Erase: Removing " & Num & " objects from (" & Map & "," & X & "," & Y & ") - current amount = " & MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount, CodeTracker '//\\LOGLINE//\\
    MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount = MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount - Num
    
    'Check if they are all gone
    If MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount <= 0 Then
        Log "Obj_Erase: Erasing object from client screens at (" & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\
        With MapInfo(Map).ObjTile(X, Y)
            .ObjInfo(ObjSlot).ObjIndex = 0
            .ObjInfo(ObjSlot).Amount = 0
            .ObjLife(ObjSlot) = 0
        End With
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_EraseObject
        ConBuf.Put_Byte CByte(X)
        ConBuf.Put_Byte CByte(Y)
        Data_Send ToMap, 0, ConBuf.Get_Buffer, Map
    End If

End Sub

Public Sub Obj_ClosestFreeSpot(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByRef NewX As Byte, ByRef NewY As Byte, ByRef NewSlot As Byte)

'*****************************************************************
'Find the closest place to put an object
'*****************************************************************
Dim lX As Byte
Dim lY As Byte

    Log "Call Obj_ClosestFreeSpot(" & Map & "," & X & "," & Y & "," & NewX & "," & NewY & "," & NewSlot & ")", CodeTracker '//\\LOGLINE//\\
    
    'Check the defined location
    If Not (MapInfo(Map).Data(X, Y).Blocked And BlockedAll) Then
        If MapInfo(Map).ObjTile(X, Y).NumObjs < MaxObjsPerTile Then
            
            'Spot is useable
            NewX = X
            NewY = Y
            NewSlot = MapInfo(Map).ObjTile(X, Y).NumObjs + 1
            Exit Sub
            
        End If
    End If
    
    'Primary spot didn't work, so loop around it and check if those work
    If X > 0 Then
        If Y > 0 Then
            For lX = X - 1 To X + 1
                For lY = Y - 1 To Y + 1
                    If lX > 1 Then
                        If lX < MapInfo(Map).Width Then
                            If lY > 1 Then
                                If lY < MapInfo(Map).Height Then
                                    If MapInfo(Map).Data(lX, lY).Blocked = 0 Then
                                        If MapInfo(Map).ObjTile(lX, lY).NumObjs < MaxObjsPerTile Then
                                            
                                            'Spot is useable
                                            NewX = lX
                                            NewY = lY
                                            NewSlot = MapInfo(Map).ObjTile(lX, lY).NumObjs + 1
                                            Exit Sub
                                            
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
    
End Sub

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
    If ObjSlot > MapInfo(Map).ObjTile(X, Y).NumObjs Then
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjInfo(1 To ObjSlot)
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjLife(1 To ObjSlot)
        MapInfo(Map).ObjTile(X, Y).NumObjs = ObjSlot
    End If
    
    'Add the object to the map slot
    MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot) = Obj
    MapInfo(Map).ObjTile(X, Y).ObjLife(ObjSlot) = timeGetTime
    
    'Clean the map tile just in case
    Obj_CleanMapTile Map, X, Y
    
    'Send the update to everyone on the map
    If BypassUpdate = 0 Then
        Log "Obj_Make: Updating object information with packet Server_MakeObject", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 7
        ConBuf.Put_Byte DataCode.Server_MakeObject
        ConBuf.Put_Long ObjData.GrhIndex(Obj.ObjIndex)
        ConBuf.Put_Byte X
        ConBuf.Put_Byte Y
        Data_Send ToMap, 0, ConBuf.Get_Buffer, Map
    End If
    
End Sub

Public Sub Quest_SendReqString(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Builds the string that says what is required for the quest
'*****************************************************************
Dim QuestID As Integer
Dim MessageID As Byte
Dim TempNPCName As String

    Log "Call Quest_SendReqString(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    QuestID = NPCList(NPCIndex).Quest
    'Get the target NPC's name if there is one - to do this, we have to open up the NPC file since we dont store the "defaults" like we do with objects/quests/etc
    If QuestData(QuestID).FinishReqNPC Then TempNPCName = NPCName(QuestData(QuestID).FinishReqNPC)

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

    'Build the packet according to the MessageID
    Select Case MessageID
        Case 9
            'NPC only
            ConBuf.PreAllocate 5 + Len(TempNPCName)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte MessageID
            ConBuf.Put_Integer QuestData(QuestID).FinishReqNPCAmount
            ConBuf.Put_String TempNPCName
        Case 10
            'Object only
            ConBuf.PreAllocate 5 + Len(ObjData.Name(QuestData(QuestID).FinishReqObj))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte MessageID
            ConBuf.Put_Integer QuestData(QuestID).FinishReqObjAmount
            ConBuf.Put_String ObjData.Name(QuestData(QuestID).FinishReqObj)
        Case 11
            'NPC and object
            ConBuf.PreAllocate 8 + Len(QuestData(QuestID).FinishReqNPCAmount) + Len(ObjData.Name(QuestData(QuestID).FinishReqObj))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte MessageID
            ConBuf.Put_Integer QuestData(QuestID).FinishReqNPCAmount
            ConBuf.Put_String TempNPCName
            ConBuf.Put_Integer QuestData(QuestID).FinishReqObjAmount
            ConBuf.Put_String ObjData.Name(QuestData(QuestID).FinishReqObj)
    End Select
    
    'Send the data to the user
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Private Sub Quest_CheckIfComplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal UserQuestSlot As Byte)

'*****************************************************************
'Checks if a quest is ready to be completed
'*****************************************************************
Dim Slot As Byte
Dim s As String
Dim i As Long

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
    ConBuf.PreAllocate 5 + Len(NPCList(NPCIndex).Name) + Len(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishTxt)
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

    'The user is done, give them the rewards
    'EXP reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp > 0 Then
        User_RaiseExp UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp
        ConBuf.PreAllocate 6
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 3
        ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
    'Gold reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold > 0 Then
        UserList(UserIndex).Gold = UserList(UserIndex).Gold + QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold
        ConBuf.PreAllocate 6
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
        If Skill_ValidSkillForClass(UserList(UserIndex).Class, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) Then
            If UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) = 1 Then
                'User already knew the skill
                s = Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill)
                ConBuf.PreAllocate 3 + Len(s)
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 5
                ConBuf.Put_String s
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                'User learns the new skill
                s = Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill)
                ConBuf.PreAllocate 3 + Len(s)
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 6
                ConBuf.Put_String s
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                'Give the user the skill
                UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) = 1
                User_SendKnownSkills UserIndex
            End If
        End If
    End If

    'Add the quest to the user's finished quest list
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).Redoable Then

        'Only add a redoable quest in the list once
        Log "Quest_CheckIfComplete: Using InStr() operation", CodeTracker '//\\LOGLINE//\\
        If UserList(UserIndex).NumCompletedQuests > 0 Then
            For i = 1 To UserList(UserIndex).NumCompletedQuests
                
                'If we find the quest in here, leave and don't add it to the list
                If UserList(UserIndex).CompletedQuests(i) = UserList(UserIndex).Quest(UserQuestSlot) Then Exit For
                
                'We hit the end of the loop and no match found, set the add to list flag and leave
                If i = UserList(UserIndex).NumCompletedQuests Then
                    i = -1
                    Exit For
                End If
                
            Next i
        Else
            
            'The user has never completed a quest before, so we have to add this one
            i = -1
        
        End If

    Else

        'Set the flag to add to the list
        i = -1

    End If

    'Add to the list if needed
    If i = -1 Then
        UserList(UserIndex).NumCompletedQuests = UserList(UserIndex).NumCompletedQuests + 1
        ReDim Preserve UserList(UserIndex).CompletedQuests(1 To UserList(UserIndex).NumCompletedQuests)
        UserList(UserIndex).CompletedQuests(UserList(UserIndex).NumCompletedQuests) = UserList(UserIndex).Quest(UserQuestSlot)
    End If
    
    'Clear the quest slot so it can be used again
    UserList(UserIndex).QuestStatus(UserQuestSlot).NPCKills = 0
    UserList(UserIndex).Quest(UserQuestSlot) = 0
    
    'Update the quest text
    Quest_SendText UserIndex, UserQuestSlot

End Sub

Private Sub Quest_General(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Reacts to a user clicking a quest NPC
'*****************************************************************
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
    Log "Quest_General: Checking if quest was already completed", CodeTracker '//\\LOGLINE//\\
    For i = 1 To UserList(UserIndex).NumCompletedQuests
        If UserList(UserIndex).CompletedQuests(i) = NPCList(NPCIndex).Quest Then
            
            'The user has completed this quest before, so check if it is redoable
            If QuestData(NPCList(NPCIndex).Quest).Redoable = 0 Then
                
                'The quest is not redoable, so sorry dude, no quest fo' j00
                Data_Send ToIndex, UserIndex, cMessage(7).Data
                Exit Sub
    
            End If
        
        End If
    Next i

    'The user has never done this quest before, so we make the NPC say whats up
    ConBuf.PreAllocate 7 + Len(NPCList(NPCIndex).Name) + Len(QuestData(NPCList(NPCIndex).Quest).StartTxt)
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(NPCList(NPCIndex).Quest).StartTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk Or DataCode.Comm_UseBubble
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

    'Give the quest requirements
    Data_Send ToIndex, UserIndex, cMessage(8).Data

    'Set the pending quest to the selected quest
    UserList(UserIndex).flags.QuestNPC = NPCIndex
    
ErrOut:

End Sub

Private Sub Quest_SayIncomplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Make the targeted NPC say the "incomplete quest" text
'*****************************************************************

    Log "Call Quest_SayIncomplete(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Incomplete text
    ConBuf.PreAllocate 7 + Len(NPCList(NPCIndex).Name) + Len(QuestData(NPCList(NPCIndex).Quest).IncompleteTxt)
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(NPCList(NPCIndex).Quest).IncompleteTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk Or DataCode.Comm_UseBubble
    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map

    'Requirements text
    Quest_SendReqString UserIndex, NPCIndex

End Sub

Function Server_CheckForSameName(ByVal UserIndex As Integer, ByVal sName As String) As Boolean

'*****************************************************************
'Checks for a user with the same Name
'*****************************************************************

Dim LoopC As Long

    Log "Call Server_CheckForSameName(" & UserIndex & "," & sName & ")", CodeTracker '//\\LOGLINE//\\

    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged = 1 Then
            If UCase$(UserList(LoopC).Name) = UCase$(sName) Then
                If UserIndex <> LoopC Then
                    Server_CheckForSameName = True
                    Log "Rtrn Server_CheckForSameName = " & Server_CheckForSameName, CodeTracker '//\\LOGLINE//\\
                    Exit Function
                End If
            End If
        End If
    Next LoopC

    Server_CheckForSameName = False
    
    Log "Rtrn Server_CheckForSameName = " & Server_CheckForSameName, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_CheckTargetedDistance(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Checks if a user is targeting a character in range
'*****************************************************************

Dim TargetID As Integer

    Log "Call Server_CheckTargetedDistance(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Set the target ID
    TargetID = UserList(UserIndex).flags.TargetIndex
    Log "Server_CheckTargetedDistance: Target ID acquired (" & TargetID & ")", CodeTracker '//\\LOGLINE//\\

    Select Case UserList(UserIndex).flags.Target

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
    If TargetID = 0 Or UserList(UserIndex).flags.TargetIndex = 0 Then
        Log "Server_CheckTargetedDistance: Telling user nothing is targeted", CodeTracker '//\\LOGLINE//\\
        UserList(UserIndex).flags.Target = 0
        UserList(UserIndex).flags.TargetIndex = 0
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.User_Target
        ConBuf.Put_Integer 0
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
    Log "Rtrn Server_CheckTargetedDistance = " & Server_CheckTargetedDistance, CodeTracker '//\\LOGLINE//\\

End Function

Public Sub Server_ClosestLegalPos(Pos As WorldPos, nPos As WorldPos)

'*****************************************************************
'Finds the closest legal tile to Pos and stores it in nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

    Log "Call Server_ClosestLegalPos([M:" & Pos.Map & " X:" & Pos.X & " Y:" & Pos.Y & "],N/A)", CodeTracker '//\\LOGLINE//\\

    'Check for a starting legal position
    If Server_LegalPos(Pos.Map, Pos.X, Pos.Y, 0) Then
        nPos = Pos
        Exit Sub
    End If

    'Set the new map
    nPos.Map = Pos.Map

    'Keep looping while the position is not legal
    Do

        'Loop through the tiles
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                Log "Server_ClosestLegalPos: Checking map tile (" & nPos.Map & "," & tX & "," & tY & ")", CodeTracker '//\\LOGLINE//\\

                'Check if the position is legal
                If Server_LegalPos(nPos.Map, tX, tY, 0) Then
                    nPos.X = tX
                    nPos.Y = tY
                    tX = Pos.X + LoopC
                    tY = Pos.Y + LoopC
                    Exit Sub
                End If

            Next tX
        Next tY

        'Check the next set of tiles
        LoopC = LoopC + 1
        
        'If we have checked too much, then just leave
        If LoopC > 5 Then   'How many tiles in all directions to search
            Notfound = True
            Exit Do
        End If

    Loop

    'If no position was found, return empty positions
    If Notfound Then
        nPos.X = 0
        nPos.Y = 0
    End If
    
    Log "Server_ClosestLegalPos: Returning position (" & nPos.Map & "," & nPos.X & "," & nPos.Y & ")", CodeTracker '//\\LOGLINE//\\

End Sub

Private Sub Server_DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Do any events on a tile
'*****************************************************************

Dim TempPos As WorldPos
Dim NewPos As WorldPos

    Log "Call Server_DoTileEvents(" & UserIndex & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Check for tile exit
    If MapInfo(Map).Data(X, Y).TileExitMap Then

        'Set the position values
        TempPos.X = MapInfo(Map).Data(X, Y).TileExitX
        TempPos.Y = MapInfo(Map).Data(X, Y).TileExitY
        TempPos.Map = MapInfo(Map).Data(X, Y).TileExitMap
        Log "Server_DoTileEvents: Tile exist exists, warps to (" & TempPos.Map & "," & TempPos.X & "," & TempPos.Y & ")", CodeTracker '//\\LOGLINE//\\

        'Check if it is a new map - if so, load the new map if needed
        If TempPos.Map <> Map Then Load_Maps_Temp TempPos.Map

        'Get the closest legal position
        If ServerMap(TempPos.Map) = ServerID Then
            Server_ClosestLegalPos TempPos, NewPos
            If Server_LegalPos(NewPos.Map, NewPos.X, NewPos.Y, 0) Then User_WarpChar UserIndex, NewPos.Map, NewPos.X, NewPos.Y
        Else
            NewPos = TempPos
            User_WarpChar UserIndex, NewPos.Map, NewPos.X, NewPos.Y
        End If
        
    End If

End Sub

Public Function Server_FindDirection(Pos As WorldPos, Target As WorldPos) As Byte

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

    'Same position
    If X = 0 Then
        If Y = 0 Then
            Server_FindDirection = NORTH
            Exit Function
        End If
    End If

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
    If X >= 1 Then
        If Y <= -1 Then
            Server_FindDirection = SOUTHWEST
            Log "Rtrn Server_FindDirection = " & Server_FindDirection, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
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

    'North
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

Public Sub Server_HeadToPos(ByVal Head As Byte, ByRef Pos As WorldPos)

'*****************************************************************
'Takes Pos and moves it in heading direction
'*****************************************************************

    Log "Call Server_HeadToPos(" & Head & ",[M:" & Pos.Map & " X:" & Pos.X & " Y:" & Pos.Y & "])", CodeTracker '//\\LOGLINE//\\

    Select Case Head
        Case NORTH
            If Pos.Y > 0 Then Pos.Y = Pos.Y - 1
        Case SOUTH
            If Pos.Y < 255 Then Pos.Y = Pos.Y + 1
        Case EAST
            If Pos.X < 255 Then Pos.X = Pos.X + 1
        Case WEST
            If Pos.X > 0 Then Pos.X = Pos.X - 1
        Case NORTHEAST
            If Pos.X < 255 Then Pos.X = Pos.X + 1
            If Pos.Y > 0 Then Pos.Y = Pos.Y - 1
        Case SOUTHEAST
            If Pos.X < 255 Then Pos.X = Pos.X + 1
            If Pos.Y < 255 Then Pos.Y = Pos.Y + 1
        Case SOUTHWEST
            If Pos.X > 0 Then Pos.X = Pos.X - 1
            If Pos.Y < 255 Then Pos.Y = Pos.Y + 1
        Case NORTHWEST
            If Pos.X > 0 Then Pos.X = Pos.X - 1
            If Pos.Y < 255 Then Pos.Y = Pos.Y - 1
    End Select

End Sub

Private Function Server_InMapBounds(ByVal Map As Long, ByVal X As Integer, ByVal Y As Integer) As Boolean

'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

    Log "Call Server_InMapBounds(" & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    If X >= 1 Then
        If X <= MapInfo(Map).Width Then
            If Y >= 1 Then
                If Y <= MapInfo(Map).Height Then Server_InMapBounds = True
            End If
        End If
    End If
    
    Log "Rtrn Server_InMapBounds = " & Server_InMapBounds, CodeTracker '//\\LOGLINE//\\
    
End Function

Public Function Server_LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As Byte) As Boolean

'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
Dim tmpBlocked As Byte

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
    If X < 1 Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If X > MapInfo(Map).Width Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If Y < 1 Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If Y > MapInfo(Map).Height Then
        Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    With MapInfo(Map).Data(X, Y)
    
        'Check if a character (User or NPC) is already at the tile
        If .UserIndex > 0 Then
            Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
        If .NPCIndex > 0 Then
            Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
        
        'Set the temp blocked variable
        tmpBlocked = .Blocked
        If tmpBlocked > 128 Then tmpBlocked = tmpBlocked - 128 'Remove the Blocked Attack flag
    
        'Check to see if its blocked
        If tmpBlocked = BlockedAll Then
            Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
    
    End With

    'Check the heading for directional blocking
    If Heading > 0 Then
        If tmpBlocked And BlockedNorth Then
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
        If tmpBlocked And BlockedEast Then
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
        If tmpBlocked And BlockedSouth Then
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
        If tmpBlocked And BlockedWest Then
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
    
    Exit Function
    
ErrOut:

    Log "Rtrn Server_LegalPos = " & Server_LegalPos, CodeTracker '//\\LOGLINE//\\

End Function

Public Function Server_NextOpenCharIndex() As Integer

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

Public Function Server_WriteMail(ByVal WriterIndex As Integer, ByVal ReceiverName As String, ByVal Subject As String, ByVal Message As String, ByRef Objs() As Obj) As Byte
Dim MailIndex As Long
Dim MailData As MailData
Dim NumObjs As Byte
Dim TempSplit() As String
Dim TempStr As String
Dim ServerIndex As Byte
Dim LoopC As Byte
Dim LoopX As Byte
Dim i As Long
Dim j As Byte

    Log "Call Server_WriteMail(" & WriterIndex & "," & ReceiverName & "," & Subject & ",[" & Message & "],[Objs() = N/A])", CodeTracker '//\\LOGLINE//\\

    'Check for a writing delay
    If WriterIndex > 0 Then
        If UserList(WriterIndex).Counters.DelayTimeMail > timeGetTime Then
            'Not enough time has passed by - goodbye! :)
            Log "Server_WriteMail: Not enough time elapsed since last mail write for user " & WriterIndex & " (" & UserList(WriterIndex).Name & ")", CodeTracker '//\\LOGLINE//\\
            Exit Function
        Else
            'Set the delay
            UserList(WriterIndex).Counters.DelayTimeMail = timeGetTime + DelayTimeMail
        End If
    End If
    
    'Check for a valid reciever
    If Not Server_UserExist(ReceiverName) Then
        If WriterIndex > 0 Then
            ConBuf.PreAllocate 3 + Len(ReceiverName)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 13
            ConBuf.Put_String ReceiverName
            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
        End If
        Exit Function
    End If
    
    'Check for sending cost
    If MailCost > 0 Then
        If WriterIndex > 0 Then 'Prevent charging the server from sending mail (ID = -1 for server) ;)
            If UserList(WriterIndex).Stats.BaseStat(SID.Gold) < MailCost Then
            
                'Not enough money
                Log "Server_WriteMail: User " & WriterIndex & " (" & UserList(WriterIndex).Name & ") has not enough money to write mail (Gold: " & UserList(WriterIndex).Stats.BaseStat(SID.Gold) & ")", CodeTracker '//\\LOGLINE//\\
                ConBuf.PreAllocate 6
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 14
                ConBuf.Put_Long MailCost
                Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                Exit Function
            
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

    'Get the number of objects
    On Error Resume Next
    NumObjs = UBound(Objs)
    If Err Then NumObjs = 0
    On Error GoTo 0

    'Put the objects into the MailData
    If NumObjs > 0 Then
        For i = 1 To NumObjs
            MailData.Obj(i) = Objs(i)
        Next i
    End If

    'Check if the reciever is on
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            If UCase$(UserList(LoopC).Name) = UCase$(ReceiverName) Then

                'Get the user's next open MailID slot
                LoopX = 0
                Do
                    LoopX = LoopX + 1
                    If LoopX > MaxMailPerUser Then
                        If WriterIndex <> -1 Then
                            'Message to the receiver
                            ConBuf.PreAllocate 3 + Len(UserList(WriterIndex).Name)
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 15
                            ConBuf.Put_String UserList(WriterIndex).Name
                            Data_Send ToIndex, LoopC, ConBuf.Get_Buffer
                            'Message to the sender
                            ConBuf.PreAllocate 3 + Len(ReceiverName)
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 15
                            ConBuf.Put_String ReceiverName
                            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                            Server_WriteMail = 1
                            Exit Function
                        Else
                            Data_Send ToIndex, LoopC, cMessage(17).Data
                            Exit Function
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
                    ConBuf.PreAllocate 3 + Len(ReceiverName)
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 19
                    ConBuf.Put_String ReceiverName
                    Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                    'Send message to receiver
                    ConBuf.PreAllocate 3 + Len(UserList(WriterIndex).Name)
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 18
                    ConBuf.Put_String UserList(WriterIndex).Name
                    Data_Send ToIndex, LoopC, ConBuf.Get_Buffer, , PP_NewMail
                Else
                    'Send message to receiver that it was from the game admin
                    Data_Send ToIndex, LoopC, cMessage(20).Data, , PP_NewMail
                End If
                
                'Check for sending cost
                If MailCost > 0 Then
                    If WriterIndex > 0 Then
                        If UserList(WriterIndex).Stats.BaseStat(SID.Gold) < MailCost Then
                            Log "Server_WriteMail: Not enough gold to write message (Gold: " & UserList(WriterIndex).Stats.BaseStat(SID.Gold) & ")", CodeTracker '//\\LOGLINE//\\
                            Exit Function 'This should never be reached, but just in case it does, leave this here
                        Else
                            UserList(WriterIndex).Stats.BaseStat(SID.Gold) = UserList(WriterIndex).Stats.BaseStat(SID.Gold) - MailCost
                            ConBuf.PreAllocate 6
                            ConBuf.Put_Byte DataCode.Server_Message
                            ConBuf.Put_Byte 21
                            ConBuf.Put_Long MailCost
                            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                        End If
                    End If
                End If
                
                Exit Function

            End If
        End If
    Next LoopC
        
    'Get the user's mail list and server from the database and store them
    DB_RS.Open "SELECT mail,server FROM users WHERE `name`='" & ReceiverName & "'", DB_Conn, adOpenStatic, adLockOptimistic
    ServerIndex = Val(DB_RS!server) 'Hold the server index
    TempStr = DB_RS!mail        'Hold the mail string
    TempSplit = Split(TempStr, vbNewLine)   'Hold the split up mail
    
    'Check if the user is on another server
    If NumServers > 1 Then

        'User is on another server, so send the message there
        If ServerIndex > 0 Then
        
            'Check if the user has too much mail already
            If UBound(TempSplit) >= MaxMailPerUser Then
                Log "Server_WriteMail: User has too much mail alread - aborting", CodeTracker '//\\LOGLINE//\\
                If WriterIndex <> -1 Then
                    ConBuf.PreAllocate 3 + Len(ReceiverName)
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 16
                    ConBuf.Put_String ReceiverName
                    Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                End If
                DB_RS.Close
                Exit Function
                
            'Send the message to the other server
            Else
                DB_RS.Close
                Server_ConnectToServer ServerIndex 'Confirm the connection
                If frmMain.ServerSocket(ServerIndex).State = sckConnected Then
                    ConBuf.PreAllocate 6 + Len(ReceiverName) + Len(MailData.Subject) + Len(MailData.Message) + (NumObjs * 4)
                    ConBuf.Put_Byte DataCode.Server_MailCompose
                    ConBuf.Put_String ReceiverName
                    ConBuf.Put_String MailData.WriterName
                    ConBuf.Put_String MailData.Subject
                    ConBuf.Put_StringEX MailData.Message
                    ConBuf.Put_Byte NumObjs
                    For j = 1 To NumObjs
                        ConBuf.Put_Integer MailData.Obj(j).ObjIndex
                        ConBuf.Put_Integer MailData.Obj(j).Amount
                    Next j
                    frmMain.ServerSocket(ServerIndex).SendData ConBuf.Get_Buffer
                    Save_Mail MailIndex, MailData
                    Server_WriteMail = 1
                    Exit Function
                Else
                    If WriterIndex <> -1 Then
                        ConBuf.PreAllocate 3 + Len(ReceiverName)
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 16
                        ConBuf.Put_String ReceiverName
                        Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                    End If
                    Exit Function
                End If
            End If
        End If
    End If

    'The user is not on, so load up his character data and impliment it into the character
    TempSplit = Split(TempStr, vbNewLine)
    If UBound(TempSplit) >= MaxMailPerUser Then 'No room for the mail
        Log "Server_WriteMail: User has too much mail alread - aborting", CodeTracker '//\\LOGLINE//\\
        DB_RS.Close
        If WriterIndex <> -1 Then
            ConBuf.PreAllocate 3 + Len(ReceiverName)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 16
            ConBuf.Put_String ReceiverName
            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
        End If
        Exit Function
    Else    'Save the mail ID in the user
        If LenB(TempStr) Then
            TempStr = TempStr & vbNewLine & MailIndex
        Else
            TempStr = TempStr & MailIndex
        End If
        DB_RS!mail = TempStr
        DB_RS.Update
        DB_RS.Close
    End If
    
    'Check for sending cost
    If MailCost > 0 Then
        If WriterIndex > 0 Then
            If UserList(WriterIndex).Stats.BaseStat(SID.Gold) < MailCost Then
                Log "Server_WriteMail: Not enough gold to write message (Gold: " & UserList(WriterIndex).Stats.BaseStat(SID.Gold) & ")", CodeTracker '//\\LOGLINE//\\
                Exit Function 'This should never be reached, but just in case it does, leave this here
            Else
                UserList(WriterIndex).Stats.BaseStat(SID.Gold) = UserList(WriterIndex).Stats.BaseStat(SID.Gold) - MailCost
                ConBuf.PreAllocate 6
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 21
                ConBuf.Put_Long MailCost
                Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
            End If
        End If
    End If

    'Send the message of success
    If WriterIndex <> -1 Then
        ConBuf.PreAllocate 3 + Len(ReceiverName)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 19
        ConBuf.Put_String ReceiverName
        Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
    End If

    'Save the mail
    Save_Mail MailIndex, MailData
    
    Server_WriteMail = 1

End Function

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
            If UserList(UserIndex).Object(LoopC).Amount + Object.Amount <= ObjData.Stacking(Object.ObjIndex) Then
                UserList(UserIndex).Object(LoopC).Amount = UserList(UserIndex).Object(LoopC).Amount + Object.Amount
                Object.Amount = 0
                'Update this slot
                User_UpdateInv False, UserIndex, LoopC
                Exit Sub
            Else
                Object.Amount = Object.Amount - (ObjData.Stacking(Object.ObjIndex) - UserList(UserIndex).Object(LoopC).Amount)
                UserList(UserIndex).Object(LoopC).Amount = ObjData.Stacking(Object.ObjIndex)
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

    'Check for a valid cached target
    If UserList(UserIndex).flags.TargetIndex < 1 Then Exit Sub
    If UserList(UserIndex).flags.TargetIndex > LastChar Then Exit Sub
    
    'Get the target index based on the NPCList() or UserList() arrays instead of CharList() value
    TargetIndex = CharList(UserList(UserIndex).flags.TargetIndex).Index
    
    'Check for a valid target index
    If TargetIndex = 0 Then Exit Sub

    'Check if a NPC or PC
    Select Case UserList(UserIndex).flags.Target
        Case CharType_PC  'PC
        
            'Check for a valid PC
            If UserList(TargetIndex).flags.UserLogged = 0 Then Exit Sub
            If UserList(TargetIndex).flags.Disconnecting = 1 Then Exit Sub
            
            With UserList(TargetIndex).Pos
                TargetPos.Map = .Map
                TargetPos.X = .X
                TargetPos.Y = .Y
            End With
            
        Case CharType_NPC  'NPC
            
            'Check for a valid NPC
            If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
            If NPCList(TargetIndex).flags.NPCAlive = 0 Then Exit Sub
            
            With NPCList(TargetIndex).Pos
                TargetPos.Map = .Map
                TargetPos.X = .X
                TargetPos.Y = .Y
            End With
            
            
        Case Else
            Exit Sub
        
    End Select
    
    'Check for a valid distance
    If Server_Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, TargetPos.X, TargetPos.Y) > ObjData.WeaponRange(UserList(UserIndex).WeaponEqpObjIndex) Then Exit Sub
    
    'Check for a valid target
    If Engine_ClearPath(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, TargetPos.X, TargetPos.Y) Then
    
        If UserList(UserIndex).WeaponEqpObjIndex > 0 Then
            
            'Play attack sound of the weapon (if the weapon makes an attack sound)
            If ObjData.UseSfx(UserList(UserIndex).WeaponEqpObjIndex) > 0 Then
                ConBuf.PreAllocate 4
                ConBuf.Put_Byte DataCode.Server_PlaySound3D
                ConBuf.Put_Byte ObjData.UseSfx(UserList(UserIndex).WeaponEqpObjIndex)
                ConBuf.Put_Byte UserList(UserIndex).Pos.X
                ConBuf.Put_Byte UserList(UserIndex).Pos.Y
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
            End If
            
        Else
        
            'Play the sound of no weapon attacking
            ConBuf.PreAllocate 4
            ConBuf.Put_Byte DataCode.Server_PlaySound3D
            ConBuf.Put_Byte UnequiptedSwingSfx  'Index of the unequipted swing sound
            ConBuf.Put_Byte UserList(UserIndex).Pos.X
            ConBuf.Put_Byte UserList(UserIndex).Pos.Y
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
                
        End If
        
        'Get the new heading
        NewHeading = Server_FindDirection(UserList(UserIndex).Pos, TargetPos)
        UserList(UserIndex).Char.Heading = NewHeading
        UserList(UserIndex).Char.HeadHeading = NewHeading
        
        Select Case UserList(UserIndex).flags.Target
        
            'Attacking user
            Case CharType_PC

                'Send the data
                If ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex) Then
                    ConBuf.PreAllocate 14   'All variables are used
                    ConBuf.Put_Byte DataCode.Server_MakeProjectile
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
                    ConBuf.Put_Long ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex)
                    ConBuf.Put_Byte ObjData.ProjectileRotateSpeed(UserList(UserIndex).WeaponEqpObjIndex)
                Else
                    ConBuf.PreAllocate 4    'Only the 3 variables below are used
                End If
                ConBuf.Put_Byte DataCode.User_Rotate
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                ConBuf.Put_Byte NewHeading
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                User_AttackUser UserIndex, TargetIndex
            
            'Attacking NPC
            Case CharType_NPC
                If ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex) Then
                    ConBuf.PreAllocate 14
                    ConBuf.Put_Byte DataCode.Server_MakeProjectile
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
                    ConBuf.Put_Long ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex)
                    ConBuf.Put_Byte ObjData.ProjectileRotateSpeed(UserList(UserIndex).WeaponEqpObjIndex)
                Else
                    ConBuf.PreAllocate 4
                End If
                ConBuf.Put_Byte DataCode.User_Rotate
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                ConBuf.Put_Byte NewHeading
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                User_AttackNPC UserIndex, TargetIndex
                
        End Select
        
    End If

End Sub

Public Sub User_Attack(ByVal UserIndex As Integer, ByVal Heading As Byte)

'*****************************************************************
'Begin a user attack sequence
'*****************************************************************

Dim AttackPos As WorldPos

    Log "Call User_Attack(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    On Error GoTo ErrOut
    If UserList(UserIndex).Stats.BaseStat(SID.MinSTA) <= 0 Then
        Log "User_Attack: MinSTA <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Counters.AttackCounter > timeGetTime - STAT_ATTACKWAIT Then
        Log "User_Attack: Not enough time elapsed since last attack - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    If Heading = 0 Then Exit Sub
    If Heading > 8 Then Exit Sub

    'Update counters
    UserList(UserIndex).Counters.AttackCounter = timeGetTime

    'Check for ranged attack
    If UserList(UserIndex).WeaponEqpObjIndex > 0 Then
        If ObjData.WeaponRange(UserList(UserIndex).WeaponEqpObjIndex) > 1 Then
            If UserList(UserIndex).flags.TargetIndex = 0 Then Exit Sub
            User_Attack_Ranged UserIndex
            Exit Sub
        End If
    End If
    
    'Get tile user is attacking
    AttackPos = UserList(UserIndex).Pos
    Server_HeadToPos Heading, AttackPos

    'Exit if not legal
    If AttackPos.X < 1 Or AttackPos.X > MapInfo(UserList(UserIndex).Pos.Map).Width Or AttackPos.Y <= 1 Or AttackPos.Y > MapInfo(UserList(UserIndex).Pos.Map).Height Then
        Log "User_Attack: Trying to attack an illegal position - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Change heading if not the same
    If UserList(UserIndex).Char.Heading <> Heading Then
        UserList(UserIndex).Char.Heading = Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte Heading
        Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
    End If

    'Look for user
    If MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).UserIndex > 0 Then
        Log "User_Attack: Found a user to attack", CodeTracker '//\\LOGLINE//\\

        'Play attack sound of the weapon (if the weapon makes an attack sound)
        If ObjData.UseSfx(UserList(UserIndex).WeaponEqpObjIndex) > 0 Then
            
            'User is using a weapon (search for sfx)
            ConBuf.PreAllocate 11
            ConBuf.Put_Byte DataCode.Server_PlaySound3D
            ConBuf.Put_Byte ObjData.UseSfx(UserList(UserIndex).WeaponEqpObjIndex)
            ConBuf.Put_Byte UserList(UserIndex).Pos.X
            ConBuf.Put_Byte UserList(UserIndex).Pos.Y
            ConBuf.Put_Byte DataCode.Server_MakeSlash
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            ConBuf.Put_Long ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex)
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
            
        Else
            
            'User is not using a weapon (don't search for a sfx)
            ConBuf.PreAllocate 11
            ConBuf.Put_Byte DataCode.Server_PlaySound3D
            ConBuf.Put_Byte UnequiptedSwingSfx  'Index of the unequipted swing sound
            ConBuf.Put_Byte UserList(UserIndex).Pos.X
            ConBuf.Put_Byte UserList(UserIndex).Pos.Y
            ConBuf.Put_Byte DataCode.Server_MakeSlash
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            ConBuf.Put_Long ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex)
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
            
        End If


        'Go to the user attacking user sub
        User_AttackUser UserIndex, MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).UserIndex
        Exit Sub

    End If

    'Look for NPC
    If MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).NPCIndex > 0 Then
        Log "User_Attack: Found a NPC to attack", CodeTracker '//\\LOGLINE//\\
        If NPCList(MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).NPCIndex).Attackable Then
            
            'If NPC has no health, they can not be attacked
            If NPCList(MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).NPCIndex).ModStat(SID.MaxHP) = 0 Then
                Log "User_Attack: NPC's MaxHP = 0 - aborting", CodeTracker '//\\LOGLINE//\\
                Exit Sub
            End If
            If NPCList(MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).NPCIndex).BaseStat(SID.MaxHP) = 0 Then
                Log "User_Attack: NPC's MaxHP = 0 - aborting", CodeTracker '//\\LOGLINE//\\
                Exit Sub
            End If
                
            'Check if the user's weapon has a swing graphic
            If ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex) Then
            
                'Play attack sound and create the graphic
                ConBuf.PreAllocate 11
                ConBuf.Put_Byte DataCode.Server_PlaySound3D
                ConBuf.Put_Byte ObjData.UseSfx(UserList(UserIndex).WeaponEqpObjIndex)
                ConBuf.Put_Byte UserList(UserIndex).Pos.X
                ConBuf.Put_Byte UserList(UserIndex).Pos.Y
                ConBuf.Put_Byte DataCode.Server_MakeSlash
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                ConBuf.Put_Long ObjData.UseGrh(UserList(UserIndex).WeaponEqpObjIndex)
                
            Else
            
                'Play sound only
                ConBuf.PreAllocate 4
                ConBuf.Put_Byte DataCode.Server_PlaySound3D
                ConBuf.Put_Byte UnequiptedSwingSfx  'Sound value of an unequipted swing
                ConBuf.Put_Byte UserList(UserIndex).Pos.X
                ConBuf.Put_Byte UserList(UserIndex).Pos.Y
                
            End If
            
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

            'Go to user attacking npc sub
            User_AttackNPC UserIndex, MapInfo(AttackPos.Map).Data(AttackPos.X, AttackPos.Y).NPCIndex

        Else
            Log "User_Attack: NPC is non-attackable", CodeTracker '//\\LOGLINE//\\

            'Can not attack the selected NPC, NPC is not attackable
            Data_Send ToIndex, UserIndex, cMessage(2).Data

        End If
        Exit Sub
    End If
    
ErrOut:

End Sub

Private Sub User_AttackNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Have a User attack a NPC
'*****************************************************************
Dim HitRate As Long 'User hit skill
Dim Hit As Long  'Hit damage

    Log "Call User_AttackNPC(" & UserIndex & "," & NPCIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Display the attack
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.User_Attack
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
    
    'Update the hit rate
    HitRate = UserList(UserIndex).Stats.ModStat(SID.Agi) + (UserList(UserIndex).Stats.ModStat(SID.Str) \ 4) + 50

    'Calculate if they hit
    If Server_RandomNumber(1, 100) >= (HitRate - NPCList(NPCIndex).ModStat(SID.Agi)) Then
        Log "User_AttackNPC: Attack chance did not pass, registering as a miss", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 5
        ConBuf.Put_Byte DataCode.Server_SetCharDamage
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Integer -1
        Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Update aggressive-face
    If UserList(UserIndex).Counters.AggressiveCounter <= 0 Then
        Log "User_AttackNPC: Making the user aggressive-faced", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_AggressiveFace
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte 1
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If
    UserList(UserIndex).Counters.AggressiveCounter = timeGetTime + AGGRESSIVEFACETIME

    'Calculate hit
    Hit = Server_RandomNumber(UserList(UserIndex).Stats.ModStat(SID.MinHIT), UserList(UserIndex).Stats.ModStat(SID.MaxHIT))
    Hit = Hit - (NPCList(NPCIndex).ModStat(SID.DEF) \ 2)
    If Hit < 1 Then Hit = 1
    Log "User_AttackNPC: Hit (damage) value calculated (" & Hit & ")", CodeTracker '//\\LOGLINE//\\

    'Hurt the NPC
    NPC_Damage NPCIndex, UserIndex, Hit

End Sub

Private Sub User_AttackUser(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

'*****************************************************************
'Have a user attack a user
'*****************************************************************
Dim HitRate As Long
Dim Hit As Long

    Log "Call User_AttackUser(" & AttackerIndex & "," & VictimIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Play the attack animation
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.User_Attack
    ConBuf.Put_Integer UserList(AttackerIndex).Char.CharIndex
    Data_Send ToPCArea, AttackerIndex, ConBuf.Get_Buffer
    
    'Update the hit rate
    HitRate = UserList(AttackerIndex).Stats.ModStat(SID.Agi) + (UserList(AttackerIndex).Stats.ModStat(SID.Str) \ 4) + 50

    'Calculate if they hit
    If Server_RandomNumber(1, 100) >= (HitRate - UserList(VictimIndex).Stats.ModStat(SID.Agi)) Then
        Log "User_AttackUser: Attack chance did not pass, registering as a miss", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 5
        ConBuf.Put_Byte DataCode.Server_SetCharDamage
        ConBuf.Put_Integer UserList(VictimIndex).Char.CharIndex
        ConBuf.Put_Integer -1
        Data_Send ToPCArea, VictimIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    
    'Update aggressive-face
    If UserList(AttackerIndex).Counters.AggressiveCounter <= 0 Then
        Log "User_AttackNPC: Making the user aggressive-faced", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_AggressiveFace
        ConBuf.Put_Integer UserList(AttackerIndex).Char.CharIndex
        ConBuf.Put_Byte 1
        Data_Send ToMap, AttackerIndex, ConBuf.Get_Buffer, UserList(AttackerIndex).Pos.Map
    End If
    UserList(AttackerIndex).Counters.AggressiveCounter = timeGetTime + AGGRESSIVEFACETIME
    
    'Calculate hit
    Hit = Server_RandomNumber(UserList(AttackerIndex).Stats.ModStat(SID.MinHIT), UserList(AttackerIndex).Stats.ModStat(SID.MaxHIT))
    Hit = Hit - (UserList(VictimIndex).Stats.ModStat(SID.DEF) \ 2)
    If Hit < 1 Then Hit = 1
    
    'Hit User
    UserList(VictimIndex).Stats.BaseStat(SID.MinHP) = UserList(VictimIndex).Stats.BaseStat(SID.MinHP) - Hit

    'User Die
    If UserList(VictimIndex).Stats.BaseStat(SID.MinHP) <= 0 Then
        Log "User_AttackUser: Killed the user", CodeTracker '//\\LOGLINE//\\

        'Kill user
        ConBuf.PreAllocate 3 + Len(UserList(VictimIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 22
        ConBuf.Put_String UserList(VictimIndex).Name
        Data_Send ToIndex, AttackerIndex, ConBuf.Get_Buffer

        ConBuf.PreAllocate 3 + Len(UserList(AttackerIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 23
        ConBuf.Put_String UserList(AttackerIndex).Name
        Data_Send ToIndex, VictimIndex, ConBuf.Get_Buffer

        User_Kill VictimIndex

    End If

End Sub

Private Sub User_ChangeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal UserIndex As Integer, Optional ByVal Body As Integer = -1, Optional ByVal Head As Integer = -1, Optional ByVal Heading As Byte = 0, Optional ByVal Weapon As Integer = -1, Optional ByVal Hair As Integer = -1, Optional ByVal Wings As Integer = -1)

'*****************************************************************
'Changes a user char's head,body and heading
'*****************************************************************
Dim ChangeFlags As Byte
Dim FlagSizes As Byte

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
            FlagSizes = FlagSizes + 2
        End If
        If Head > -1 Then
            If .Head <> Head Then .Head = Head
            ChangeFlags = ChangeFlags Or 2
            FlagSizes = FlagSizes + 2
        End If
        If Heading > 0 Then
            If .Heading <> Heading Then .Heading = Heading
            ChangeFlags = ChangeFlags Or 4
            FlagSizes = FlagSizes + 1
        End If
        If Weapon > -1 Then
            If .Weapon <> Weapon Then .Weapon = Weapon
            ChangeFlags = ChangeFlags Or 8
            FlagSizes = FlagSizes + 2
        End If
        If Hair > -1 Then
            If .Hair <> Hair Then .Hair = Hair
            ChangeFlags = ChangeFlags Or 16
            FlagSizes = FlagSizes + 2
        End If
        If Wings > -1 Then
            If .Wings <> Wings Then .Wings = Wings
            ChangeFlags = ChangeFlags Or 32
            FlagSizes = FlagSizes + 2
        End If
    End With

    'Make sure there is a packet to send
    If ChangeFlags = 0 Then Exit Sub

    'Create the packet
    ConBuf.PreAllocate 4 + FlagSizes
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

Private Sub User_ChangeInv(ByVal UserIndex As Integer, ByVal Slot As Byte, Object As UserOBJ)

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

    'If the object has an index, then send the related information of the object
    'If index = 0, then we assume we are deleting it
    If Object.ObjIndex Then
        ConBuf.PreAllocate 16 + Len(ObjData.Name(Object.ObjIndex))
        ConBuf.Put_Byte DataCode.User_SetInventorySlot
        ConBuf.Put_Byte Slot
        ConBuf.Put_Long Object.ObjIndex
        ConBuf.Put_String ObjData.Name(Object.ObjIndex)
        ConBuf.Put_Long Object.Amount
        ConBuf.Put_Byte Object.Equipped
        ConBuf.Put_Long ObjData.GrhIndex(Object.ObjIndex)
    Else
        ConBuf.PreAllocate 6
        ConBuf.Put_Byte DataCode.User_SetInventorySlot
        ConBuf.Put_Byte Slot
        ConBuf.Put_Long Object.ObjIndex
    End If
    
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Public Sub User_DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Drops a object from a User's slot
'*****************************************************************
Dim Obj As Obj
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
    If UserList(UserIndex).Pos.X < 1 Then
        Log "User_DropObj: User X < 1 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.X > MapInfo(UserList(UserIndex).Pos.Map).Width Then
        Log "User_DropObj: User X > XMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y < 1 Then
        Log "User_DropObj: User Y < 1 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y > MapInfo(UserList(UserIndex).Pos.Map).Height Then
        Log "User_DropObj: User Y > YMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
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
        Data_Send ToIndex, UserIndex, cMessage(24).Data
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

Public Sub User_EraseChar(ByVal UserIndex As Integer)

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
    
    'Confirm that the map is even loaded (if not, theres obviously no users on it)
    If MapInfo(UserList(UserIndex).Pos.Map).DataLoaded = 0 Then Exit Sub
    
    'Send erase command to clients
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.Server_EraseChar
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    
    'Remove from list
    CharList(UserList(UserIndex).Char.CharIndex).Index = 0
    CharList(UserList(UserIndex).Char.CharIndex).CharType = 0
    
    'Update userlist
    UserList(UserIndex).Char.CharIndex = 0
    
    If UserList(UserIndex).Pos.X < 1 Then
        Log "User_EraseChar: User X < 1 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.X > MapInfo(UserList(UserIndex).Pos.Map).Width Then
        Log "User_EraseChar: User X > XMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y < 1 Then
        Log "User_EraseChar: User Y < 1 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Y > MapInfo(UserList(UserIndex).Pos.Map).Height Then
        Log "User_EraseChar: User Y > YMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Remove from map
    MapInfo(UserList(UserIndex).Pos.Map).Data(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    
    Exit Sub '//\\LOGLINE//\\
    
ErrOut:

    Log "User_EraseChar: Unexpected error in User_EraseChar - GoTo ErrOut called!", CriticalError '//\\LOGLINE//\\

End Sub

Private Sub Obj_CleanMapTile(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

'*****************************************************************
'Removes all the unused obj slots on a map tile
'Make sure you call this every time you remove an object from a tile!
'*****************************************************************
Dim NumObjs As Byte
Dim i As Long
Dim j As Long

    Log "Call Obj_CleanMapTile(" & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the map is in memory
    If MapInfo(Map).DataLoaded = 0 Then Exit Sub

    'Make sure we wern't given an empty map tile
    If MapInfo(Map).ObjTile(X, Y).NumObjs = 0 Then
        Log "Obj_CleanMapTile: NumObjs = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check through all the object slots
    For i = 1 To MapInfo(Map).ObjTile(X, Y).NumObjs
        If MapInfo(Map).ObjTile(X, Y).ObjInfo(i).ObjIndex > 0 Then
            If MapInfo(Map).ObjTile(X, Y).ObjInfo(i).Amount > 0 Then
                
                'Object found, so raise the count
                NumObjs = NumObjs + 1
                
                'Move down in the array if possible
                If i > 1 Then   'We can't sort any lower then 1, so don't even try it
                    
                    'Loop through all the previous object slots
                    For j = 1 To (i - 1)    '(i - 1) is since we don't to check the slot it is already on!
                    
                        'If the object slot is unused then
                        If MapInfo(Map).ObjTile(X, Y).ObjInfo(j).ObjIndex = 0 Or MapInfo(Map).ObjTile(X, Y).ObjInfo(j).Amount = 0 Then
                
                            'Scoot the item's keester down to that slot (swap the used into the unused)
                            MapInfo(Map).ObjTile(X, Y).ObjInfo(j) = MapInfo(Map).ObjTile(X, Y).ObjInfo(i)
                            MapInfo(Map).ObjTile(X, Y).ObjLife(j) = MapInfo(Map).ObjTile(X, Y).ObjLife(i)
                            
                            'Clear the old object
                            ZeroMemory MapInfo(Map).ObjTile(X, Y).ObjInfo(i), Len(MapInfo(Map).ObjTile(X, Y).ObjInfo(i))
                            MapInfo(Map).ObjTile(X, Y).ObjLife(i) = 0
                        
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
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjInfo(1 To NumObjs)
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjLife(1 To NumObjs)
    Else
        'We have no slots at all used, so kill the whole damn thing
        Log "Obj_CleanMapTile: Erasing ObjInfo() array", CodeTracker '//\\LOGLINE//\\
        Erase MapInfo(Map).ObjTile(X, Y).ObjInfo
        Erase MapInfo(Map).ObjTile(X, Y).ObjLife
    End If
    
    'Assign the value to the map array for later usage
    MapInfo(Map).ObjTile(X, Y).NumObjs = NumObjs

End Sub

Public Sub User_GetObj(ByVal UserIndex As Integer)

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
    If Map <= 0 Then
        Log "User_GetObj: User map <= 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Map > NumMaps Then
        Log "User_GetObj: User map > NumMaps - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If X < 1 Then
        Log "User_GetObj: User X < 1 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If X > MapInfo(Map).Width Then
        Log "User_GetObj: User X > XMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Y < 1 Then
        Log "User_GetObj: User Y < 1 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Y > MapInfo(Map).Height Then
        Log "User_GetObj: User Y > YMaxMapSize - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    
    'Make sure the map is in memory
    If MapInfo(Map).DataLoaded = 0 Then Exit Sub
    
    'No objects exist on the tile
    If MapInfo(Map).ObjTile(X, Y).NumObjs = 0 Then
        Log "User_GetObj: NumObjs on tile (" & Map & "," & X & "," & Y & ") = 0", CodeTracker '//\\LOGLINE//\\
        Data_Send ToIndex, UserIndex, cMessage(25).Data
        Exit Sub
    End If

    'Check for object on ground
    For i = 1 To MapInfo(Map).ObjTile(X, Y).NumObjs
        If MapInfo(Map).ObjTile(X, Y).ObjInfo(i).ObjIndex > 0 Then
            ObjSlot = i
            Exit For
        End If
    Next i
    Log "User_GetObj: ObjSlot = " & ObjSlot, CodeTracker '//\\LOGLINE//\\
    
    'For some reason, the NumObjs value is > 0 but there are no objects - no objs found
    If ObjSlot = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(25).Data
        
        'Also request a cleaning of the map object array since it is obviously messy
        Obj_CleanMapTile Map, X, Y
        
        Exit Sub
    End If
    
    'Check to see if User already has the object type
    Slot = 1
    If ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex) > 1 Then
        Do Until UserList(UserIndex).Object(Slot).ObjIndex = MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex
            Slot = Slot + 1
            If Slot > MAX_INVENTORY_SLOTS Then
                Log "User_GetObj: Slot > MAX_INVENTORY_SLOTS", CodeTracker '//\\LOGLINE//\\
                Exit Do
            End If
        Loop
    Else
        Slot = MAX_INVENTORY_SLOTS + 1  'Override to force to check the next slot
    End If

    'Else check if there is a empty slot
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Data_Send ToIndex, UserIndex, cMessage(26).Data
                Exit Sub
            End If
        Loop
    End If

    'Fill object slot
    If UserList(UserIndex).Object(Slot).Amount + MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount <= ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex) Then

        'Tell the user they recieved the items
        ConBuf.PreAllocate 5 + Len(ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex))
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 27
        ConBuf.Put_Integer MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount
        ConBuf.Put_String ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex)
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'User takes all the items
        UserList(UserIndex).Object(Slot).ObjIndex = MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount
        Obj_Erase MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount, ObjSlot, Map, X, Y

    Else
        'Over MAX_INV_OBJS
        If MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount < UserList(UserIndex).Object(Slot).Amount Then
            'Tell the user they recieved the items
            ConBuf.PreAllocate 5 + Len(ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 27
            ConBuf.Put_Integer Abs(ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex) - (UserList(UserIndex).Object(Slot).Amount + MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).Amount))
            ConBuf.Put_String ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).ObjIndex)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).Amount = Abs(ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex) - (UserList(UserIndex).Object(Slot).Amount + MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).Amount))
        Else
            'Tell the user they recieved the items
            ConBuf.PreAllocate 5 + Len(ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).ObjIndex))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 27
            ConBuf.Put_Integer Abs((ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex) + UserList(UserIndex).Object(Slot).Amount) - MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).Amount)
            ConBuf.Put_String ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).ObjIndex)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).Amount = Abs((ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex) + UserList(UserIndex).Object(Slot).Amount) - MapInfo(Map).ObjTile(X, Y).ObjInfo(Slot).Amount)
        End If
        UserList(UserIndex).Object(Slot).Amount = ObjData.Stacking(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex)
    End If

    'Update the user's inventory
    User_UpdateInv False, UserIndex, Slot

ErrOut:

End Sub

Public Function User_CorrectServer(ByVal UserName As String, ByVal UserIndex As Integer, Optional ByVal UserMap As Integer = 0) As Byte

'*****************************************************************
'Checks if the user is on the right server - if not, moves them to the right one
'*****************************************************************

    'Get the user's map
    If UserList(UserIndex).Pos.Map = 0 Then
        DB_RS.Open "SELECT pos_map FROM users WHERE `name`='" & UserName & "'", DB_Conn, adOpenStatic, adLockOptimistic
        UserMap = DB_RS(0)
        DB_RS.Close
    Else
        UserMap = UserList(UserIndex).Pos.Map
    End If
    
    'Check if this is the right server for the map
    If ServerID <> ServerMap(UserMap) Then
    
        'Incorrect server, tell the user to change after saving their character
        Save_User UserList(UserIndex), UserIndex
        UserList(UserIndex).flags.DoNotSave = 1
        
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_ChangeServer
        ConBuf.Put_Integer ServerInfo(ServerMap(UserMap)).Port
        ConBuf.Put_String ServerInfo(ServerMap(UserMap)).EIP
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        
    Else
    
        'Everything was good, return a true
        User_CorrectServer = 1

    End If

End Function

Public Sub User_GiveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer)

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
                Data_Send ToIndex, UserIndex, cMessage(26).Data
                Exit Sub
            End If
        Loop
    End If

    'Fill object slot
    If UserList(UserIndex).Object(Slot).Amount + Amount <= ObjData.Stacking(UserList(UserIndex).Object(Slot).ObjIndex) Then

        'Tell the user they recieved the items
        ConBuf.PreAllocate 5 + Len(ObjData.Name(ObjIndex))
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 28
        ConBuf.Put_Integer Amount
        ConBuf.Put_String ObjData.Name(ObjIndex)
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'User takes all the items
        UserList(UserIndex).Object(Slot).ObjIndex = ObjIndex
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + Amount

    Else
        'Over MAX_INV_OBJS
        If Amount < UserList(UserIndex).Object(Slot).Amount Then
            'Tell the user they recieved the items
            ConBuf.PreAllocate 5 + Len(ObjData.Name(ObjIndex))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 28
            ConBuf.Put_Integer Abs(ObjData.Stacking(UserList(UserIndex).Object(Slot).ObjIndex) - (UserList(UserIndex).Object(Slot).Amount + Amount))
            ConBuf.Put_String ObjData.Name(ObjIndex)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Else
            'Tell the user they recieved the items
            ConBuf.PreAllocate 5 + Len(ObjData.Name(ObjIndex))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 28
            ConBuf.Put_Integer Abs((ObjData.Stacking(UserList(UserIndex).Object(Slot).ObjIndex) + UserList(UserIndex).Object(Slot).Amount) - Amount)
            ConBuf.Put_String ObjData.Name(ObjIndex)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        End If
        UserList(UserIndex).Object(Slot).Amount = ObjData.Stacking(UserList(UserIndex).Object(Slot).ObjIndex)
    End If

    'Update the user's inventory
    User_UpdateInv False, UserIndex, Slot

End Sub

Public Sub User_Kill(ByVal UserIndex As Integer)

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
    Server_ClosestLegalPos ResPos, TempPos
    If Server_LegalPos(TempPos.Map, TempPos.X, TempPos.Y, 0) = False Then
        Data_Send ToIndex, UserIndex, cMessage(83).Data
        User_Close UserIndex
        Exit Sub
    End If
    
    'Remove the targeted NPC
    UserList(UserIndex).flags.Target = 0
    UserList(UserIndex).flags.TargetIndex = 0
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.User_Target
    ConBuf.Put_Integer 0
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Warp him there
    User_WarpChar UserIndex, TempPos.Map, TempPos.X, TempPos.Y

End Sub

Public Sub User_LookAtTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Button As Byte)

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
    If Not Server_InMapBounds(Map, X, Y) Then
        Log "User_LookAtTile: Invalid tile looked at (X:" & X & " Y:" & Y & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
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

    'Make sure the map is in memory
    If MapInfo(Map).DataLoaded = 0 Then Exit Sub
    
    'Make sure the clicked position is in range of the screen
    If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y, MaxServerDistanceX, MaxServerDistanceY) = 0 Then Exit Sub

    '***** Right Click *****
    If Button = vbRightButton Then

        '*** Check for mailbox ***
        If MapInfo(Map).Data(X, Y).Mailbox = 1 Then

            'Only check mail if right next to the mailbox
            If UserList(UserIndex).Pos.Map = Map Then
                If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y, 1, 1) Then

                    'Store the position of the mailbox for later reference in case user tries to use items away from mailbox
                    UserList(UserIndex).MailboxPos.Map = Map
                    UserList(UserIndex).MailboxPos.X = X
                    UserList(UserIndex).MailboxPos.Y = Y

                    'Resend all the mail
                    ConBuf.PreAllocate 2    'One for header, one for end byte
                    ConBuf.Put_Byte DataCode.Server_MailBox
                    For LoopC = 1 To MaxMailPerUser
                        If UserList(UserIndex).MailID(LoopC) > 0 Then
                            MsgData = Load_Mail(UserList(UserIndex).MailID(LoopC))
                            ConBuf.Allocate 4 + Len(MsgData.WriterName) + Len(CStr(MsgData.RecieveDate)) + Len(MsgData.Subject)
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
            Data_Send ToIndex, UserIndex, cMessage(29).Data
            Exit Sub

        End If

        '*** Check for Characters ***
        If Y + 1 <= MapInfo(Map).Height Then
            If MapInfo(Map).Data(X, Y + 1).UserIndex > 0 Then
                TempIndex = MapInfo(Map).Data(X, Y + 1).UserIndex
                FoundChar = 1
            End If
            If MapInfo(Map).Data(X, Y + 1).NPCIndex > 0 Then
                TempIndex = MapInfo(Map).Data(X, Y + 1).NPCIndex
                FoundChar = 2
            End If
        End If
        'Check for Character
        If FoundChar = 0 Then
            If MapInfo(Map).Data(X, Y).UserIndex > 0 Then
                TempIndex = MapInfo(Map).Data(X, Y).UserIndex
                FoundChar = 1
            End If
            If MapInfo(Map).Data(X, Y).NPCIndex > 0 Then
                TempIndex = MapInfo(Map).Data(X, Y).NPCIndex
                FoundChar = 2
            End If
        End If
        'React to character
        If FoundChar = 1 Then
            If Len(UserList(TempIndex).Desc) > 1 Then
                ConBuf.PreAllocate 4 + Len(UserList(TempIndex).Name) + Len(UserList(TempIndex).Desc)
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 30
                ConBuf.Put_String UserList(TempIndex).Name
                ConBuf.Put_String UserList(TempIndex).Desc
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                ConBuf.PreAllocate 3 + Len(UserList(TempIndex).Name)
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 31
                ConBuf.Put_String UserList(TempIndex).Name
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            FoundSomething = 1
        End If
        If FoundChar = 2 Then
            FoundSomething = 1
            '*** Check for NPC banker ***
            If NPCList(TempIndex).AI = 6 Then
                UserList(UserIndex).flags.TradeWithNPC = TempIndex
                ConBuf.Put_Byte DataCode.User_Bank_Open
                For LoopC = 1 To MAX_INVENTORY_SLOTS
                    If UserList(UserIndex).Bank(LoopC).ObjIndex > 0 Then
                        ConBuf.Put_Byte LoopC
                        ConBuf.Put_Long ObjData.GrhIndex(UserList(UserIndex).Bank(LoopC).ObjIndex)
                        ConBuf.Put_Integer UserList(UserIndex).Bank(LoopC).Amount
                    End If
                Next LoopC
                ConBuf.Put_Byte 255 'Terminator byte - tells the client the list has ended
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                '*** Check for NPC vendor ***
                If NPCList(TempIndex).NumVendItems > 0 Then
                    User_TradeWithNPC UserIndex, TempIndex
                    FoundSomething = 1
                Else
                    '*** NPC not a vendor, give description ***
                    If Len(NPCList(TempIndex).Name) > 1 Then
                        ConBuf.PreAllocate 4 + Len(NPCList(TempIndex).Name) + Len(NPCList(TempIndex).Desc)
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 30
                        ConBuf.Put_String NPCList(TempIndex).Name
                        ConBuf.Put_String NPCList(TempIndex).Desc
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    Else
                        ConBuf.PreAllocate 3 + Len(NPCList(TempIndex).Name)
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 31
                        ConBuf.Put_String NPCList(TempIndex).Name
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    End If
                    '*** Quest NPC ***
                    If NPCList(TempIndex).Quest > 0 Then Quest_General UserIndex, TempIndex
                End If
            End If
        End If

        '*** Check for object ***
        If MapInfo(Map).ObjTile(X, Y).NumObjs > 0 Then
            For LoopC = 1 To MapInfo(Map).ObjTile(X, Y).NumObjs
                If MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).ObjIndex > 0 Then
                    'Check whether to use the singular or plural message
                    If MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).Amount = 1 Then
                        ConBuf.PreAllocate 3 + Len(ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).ObjIndex))
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 32
                        ConBuf.Put_String ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).ObjIndex)
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    Else
                        ConBuf.PreAllocate 5 + Len(ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).ObjIndex))
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 86
                        ConBuf.Put_String ObjData.Name(MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).ObjIndex)
                        ConBuf.Put_Integer MapInfo(Map).ObjTile(X, Y).ObjInfo(LoopC).Amount
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
            Data_Send ToIndex, UserIndex, cMessage(33).Data
        End If

        '***** Left Click *****
    ElseIf Button = vbLeftButton Then

        '*** Look for NPC/Player to target ***
        If Y + 1 <= MapInfo(Map).Height Then
            If MapInfo(Map).Data(X, Y + 1).UserIndex > 0 Then
                TempCharIndex = UserList(MapInfo(Map).Data(X, Y + 1).UserIndex).Char.CharIndex
                TempIndex = MapInfo(Map).Data(X, Y + 1).UserIndex
                FoundChar = 1
            End If
            If MapInfo(Map).Data(X, Y + 1).NPCIndex > 0 Then
                TempCharIndex = NPCList(MapInfo(Map).Data(X, Y + 1).NPCIndex).Char.CharIndex
                TempIndex = MapInfo(Map).Data(X, Y + 1).NPCIndex
                FoundChar = 2
            End If
        End If
        If FoundChar = 0 Then
            If MapInfo(Map).Data(X, Y).UserIndex > 0 Then
                TempCharIndex = UserList(MapInfo(Map).Data(X, Y).UserIndex).Char.CharIndex
                TempIndex = MapInfo(Map).Data(X, Y).UserIndex
                FoundChar = 1
            End If
            If MapInfo(Map).Data(X, Y).NPCIndex > 0 Then
                TempCharIndex = NPCList(MapInfo(Map).Data(X, Y).NPCIndex).Char.CharIndex
                TempIndex = MapInfo(Map).Data(X, Y).NPCIndex
                FoundChar = 2
            End If
        End If

        'Validate distance
        If FoundChar = 0 Then
            If UserList(UserIndex).flags.Target Then
                UserList(UserIndex).flags.Target = 0
                UserList(UserIndex).flags.TargetIndex = 0
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer 0
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            Exit Sub
        ElseIf FoundChar = 1 Then
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(TempIndex).Pos.X, UserList(TempIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                UserList(UserIndex).flags.Target = 1
                UserList(UserIndex).flags.TargetIndex = TempCharIndex
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer UserList(TempIndex).Char.CharIndex
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                If UserList(UserIndex).flags.Target Then
                    UserList(UserIndex).flags.Target = 0
                    UserList(UserIndex).flags.TargetIndex = 0
                    ConBuf.PreAllocate 3
                    ConBuf.Put_Byte DataCode.User_Target
                    ConBuf.Put_Integer 0
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                End If
            End If
        ElseIf FoundChar = 2 Then
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(TempIndex).Pos.X, NPCList(TempIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                UserList(UserIndex).flags.Target = 2
                UserList(UserIndex).flags.TargetIndex = TempCharIndex
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer NPCList(TempIndex).Char.CharIndex
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                If UserList(UserIndex).flags.Target Then
                    UserList(UserIndex).flags.Target = 0
                    UserList(UserIndex).flags.TargetIndex = 0
                    ConBuf.PreAllocate 3
                    ConBuf.Put_Byte DataCode.User_Target
                    ConBuf.Put_Integer 0
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                End If
            End If
        End If

    End If
    
ErrOut:

End Sub

Public Sub User_MakeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Makes and places a user's character
'*****************************************************************

Dim CharIndex As Integer

    Log "Call User_MakeChar(" & sndRoute & "," & sndIndex & "," & UserIndex & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

'Place character on map

    MapInfo(Map).Data(X, Y).UserIndex = UserIndex

    'Give it a char if needed
    If UserList(UserIndex).Char.CharIndex = 0 Then
        CharIndex = Server_NextOpenCharIndex
        UserList(UserIndex).Char.CharIndex = CharIndex
        CharList(CharIndex).Index = UserIndex
        CharList(CharIndex).CharType = CharType_PC
    End If

    'Send make character command to clients
    ConBuf.PreAllocate 22 + Len(UserList(UserIndex).Name)
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
    ConBuf.Put_Byte 0
    If UserList(UserIndex).GroupIndex > 0 And UserList(sndIndex).GroupIndex = UserList(UserIndex).GroupIndex Then
        ConBuf.Put_Byte ClientCharType_Grouped
    Else
        ConBuf.Put_Byte ClientCharType_PC
    End If
    
    '*** Status icons ***
    If UserList(UserIndex).Skills.Bless > 0 Then
        ConBuf.Allocate 4
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.Protect > 0 Then
        ConBuf.Allocate 4
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.IronSkin > 0 Then
        ConBuf.Allocate 4
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.Strengthen > 0 Then
        ConBuf.Allocate 4
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If
    If UserList(UserIndex).Skills.WarCurse > 0 Then
        ConBuf.Allocate 4
        ConBuf.Put_Byte DataCode.Server_IconWarCursed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    End If

    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, Map

End Sub

Public Sub User_MoveChar(ByVal UserIndex As Integer, ByVal nHeading As Byte, ByVal Running As Byte)

'*****************************************************************
'Moves a User from one tile to another
'*****************************************************************
Dim TempIndex As Integer
Dim nPos As WorldPos
Dim i As Long

    Log "Call User_MoveChar(" & UserIndex & "," & nHeading & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    If nHeading = 0 Then
        Log "User_MoveChar: nHeading = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If nHeading > 8 Then
        Log "User_MoveChar: nHeading > 8 - aborting", CodeTracker '//\\LOGLINE//\\
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
    
    'Clear the pending quest NPC number and trading NPC
    UserList(UserIndex).flags.QuestNPC = 0
    UserList(UserIndex).flags.TradeWithNPC = 0

    'Get the new time
    UserList(UserIndex).Counters.MoveCounter = timeGetTime + Server_WalkTimePerTile(UserList(UserIndex).Stats.ModStat(SID.Speed) + Running * RunningSpeed)

    'Get the new position
    nPos = UserList(UserIndex).Pos
    Server_HeadToPos nHeading, nPos

    'Move if legal pos
    If Server_LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, nHeading) Then
    
        'Send the movement
        ConBuf.PreAllocate 6
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
        MapInfo(UserList(UserIndex).Pos.Map).Data(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.Heading = nHeading
        UserList(UserIndex).Char.HeadHeading = nHeading
        MapInfo(UserList(UserIndex).Pos.Map).Data(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex

        'Do tile events
        Server_DoTileEvents UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y

    Else

        'Make sure the user's position is correct
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
    End If

End Sub

Function User_NameToIndex(ByVal strName As String) As Integer

'*****************************************************************
'Searches userlist for a name and return userindex
'*****************************************************************

Dim UserIndex As Integer

    Log "Call User_NameToIndex(" & strName & ")", CodeTracker '//\\LOGLINE//\\

    'Check for bad name
    If LenB(strName) = 0 Then
        User_NameToIndex = 0
        Log "Rtrn User_NameToIndex = " & User_NameToIndex, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Force the name to UCase
    strName = UCase$(strName)
    
    'Find the user
    UserIndex = 1
    Do Until UCase$(UserList(UserIndex).Name) = strName
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

Public Sub User_RaiseExp(ByVal UserIndex As Integer, ByVal lngEXP As Long)

'*****************************************************************
'Raise the user's experience - this should be the only way exp is raised!
'*****************************************************************

Dim Levels As Integer

    Log "Call User_RaiseExp(" & UserIndex & "," & lngEXP & ")", CodeTracker '//\\LOGLINE//\\

    'Update the user's experience
    UserList(UserIndex).Stats.BaseStat(SID.EXP) = UserList(UserIndex).Stats.BaseStat(SID.EXP) + lngEXP

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
        Data_Send ToIndex, UserIndex, cMessage(34).Data

    ElseIf Levels > 1 Then
        Log "User_RaiseExp: User gained multiple levels (" & Levels & ")", CodeTracker '//\\LOGLINE//\\

        'Say the user's level raised
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 35
        ConBuf.Put_Byte Levels
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    End If

End Sub

Public Sub User_RemoveInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal UpdateInv As Byte = 1)

'*****************************************************************
'Unequip a inventory item
'*****************************************************************

    Log "Call User_RemoveInvItem(" & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

    'Get the object type
    Select Case ObjData.ObjType(UserList(UserIndex).Object(Slot).ObjIndex)
    
            'Check for weapon
        Case OBJTYPE_WEAPON
            Log "User_RemoveInvItem: Object type OBJTYPE_WEAPON", CodeTracker '//\\LOGLINE//\\
            
            'Update the weapon distance on the client
            ConBuf.PreAllocate 2
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
            Log "User_RemoveInvItem: Unknown object type! Object type: " & ObjData.ObjType(UserList(UserIndex).Object(Slot).ObjIndex), CriticalError '//\\LOGLINE//\\

    End Select
    
    'Force update of the modstats
    UserList(UserIndex).Stats.Update = 1

    'Update the user's stats
    If UpdateInv = 1 Then User_UpdateInv False, UserIndex, Slot

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
    ReDim KnowSkillList(1 To NumBytesForSkills) As Byte

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
    ConBuf.PreAllocate 1 + NumBytesForSkills
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
        If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, 6, 6) = 0 Then
            Log "User_TradeWithNPC: Can not trade - user too far away", CodeTracker '//\\LOGLINE//\\
            Data_Send ToIndex, UserIndex, cMessage(36).Data
            Exit Sub
        End If

        Log "User_TradeWithNPC: Building vending items list", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 4 + Len(NPCList(NPCIndex).Name) + (NPCList(NPCIndex).NumVendItems * 9)   'We can't allocate the Obj names, so just allocate the byte for the string : /
        ConBuf.Put_Byte DataCode.User_Trade_StartNPCTrade
        ConBuf.Put_String NPCList(NPCIndex).Name
        ConBuf.Put_Integer NPCList(NPCIndex).NumVendItems
        For LoopC = 1 To NPCList(NPCIndex).NumVendItems
            ConBuf.Put_Long ObjData.GrhIndex(NPCList(NPCIndex).VendItems(LoopC).ObjIndex)
            ConBuf.Put_String ObjData.Name(NPCList(NPCIndex).VendItems(LoopC).ObjIndex)
            ConBuf.Put_Long ObjData.Price(NPCList(NPCIndex).VendItems(LoopC).ObjIndex)
        Next LoopC
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        UserList(UserIndex).flags.TradeWithNPC = NPCIndex
    End If

End Sub

Public Sub User_UpdateBank(ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Updates a user's bank slot
'*****************************************************************
    
    Log "Call User_UpdateBank(" & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

    'Sending empty object
    If UserList(UserIndex).Bank(Slot).ObjIndex = 0 Then
        ConBuf.PreAllocate 6
        ConBuf.Put_Byte DataCode.User_Bank_UpdateSlot
        ConBuf.Put_Byte Slot
        ConBuf.Put_Long 0
        
    'Sending object
    Else
        ConBuf.PreAllocate 8
        ConBuf.Put_Byte DataCode.User_Bank_UpdateSlot
        ConBuf.Put_Byte Slot
        ConBuf.Put_Long ObjData.GrhIndex(UserList(UserIndex).Bank(Slot).ObjIndex)
        ConBuf.Put_Integer UserList(UserIndex).Bank(Slot).Amount
    End If
    
    'Send the data
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Public Sub User_UpdateInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Updates a user's inventory slot
'*****************************************************************

Dim NullObj As UserOBJ
Dim LoopC As Long

    Log "Call User_UpdateInv(" & UpdateAll & "," & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

'Update one slot

    If Not UpdateAll Then
        'Update User inventory
        If UserList(UserIndex).Object(Slot).ObjIndex > 0 Then
            User_ChangeInv UserIndex, Slot, UserList(UserIndex).Object(Slot)
        Else
            User_ChangeInv UserIndex, Slot, NullObj
        End If
    Else
        'Update every slot
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            'Update User inventory
            If UserList(UserIndex).Object(LoopC).ObjIndex Then
                User_ChangeInv UserIndex, LoopC, UserList(UserIndex).Object(LoopC)
            End If
        Next LoopC
    End If

End Sub

Public Sub User_UpdateMap(ByVal UserIndex As Integer)

'*****************************************************************
'Updates a user with the place of all chars in the Map
'*****************************************************************

Dim Map As Integer
Dim X As Byte
Dim Y As Byte
Dim i As Long

    Log "Call User_UpdateMap(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Map = UserList(UserIndex).Pos.Map
    
    'Make sure the map is in memory
    If MapInfo(Map).DataLoaded = 0 Then Exit Sub

    'Send user char's pos
    Log "User_UpdateMap: For X = 1 to " & UBound(MapUsers(Map).Index()), CodeTracker '//\\LOGLINE//\\
    For X = 1 To UBound(MapUsers(Map).Index())
        User_MakeChar ToIndex, UserIndex, MapUsers(Map).Index(X), Map, UserList(MapUsers(Map).Index(X)).Pos.X, UserList(MapUsers(Map).Index(X)).Pos.Y
    Next X

    'Place chars and objects
    For Y = 1 To MapInfo(Map).Height
        For X = 1 To MapInfo(Map).Width
            
            'NPC update
            If MapInfo(Map).Data(X, Y).NPCIndex Then NPC_MakeChar ToIndex, UserIndex, MapInfo(Map).Data(X, Y).NPCIndex, Map, X, Y
            
            'Object update
            If MapInfo(Map).ObjTile(X, Y).NumObjs > 0 Then
                For i = 1 To MapInfo(Map).ObjTile(X, Y).NumObjs
                    If MapInfo(Map).ObjTile(X, Y).ObjInfo(i).ObjIndex Then
                        ConBuf.PreAllocate 7
                        ConBuf.Put_Byte DataCode.Server_MakeObject
                        ConBuf.Put_Long ObjData.GrhIndex(MapInfo(Map).ObjTile(X, Y).ObjInfo(i).ObjIndex)
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

Dim WeaponObj As Integer
Dim ArmorObj As Integer
Dim WingsObj As Integer
Dim i As Integer

    Log "Call User_UpdateModStats(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    If UserList(UserIndex).flags.UserLogged = 0 Then
        Log "User_UpdateModStats: UserLogged = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Set the equipted items
    If UserList(UserIndex).WeaponEqpObjIndex > 0 Then WeaponObj = UserList(UserIndex).WeaponEqpObjIndex
    If UserList(UserIndex).ArmorEqpObjIndex > 0 Then ArmorObj = UserList(UserIndex).ArmorEqpObjIndex
    If UserList(UserIndex).WingsEqpObjIndex > 0 Then WingsObj = UserList(UserIndex).WingsEqpObjIndex

    With UserList(UserIndex).Stats

        'Equipted items
        For i = FirstModStat To NumStats
            Log "User_UpdateModStats: Updating ModStat ID " & i, CodeTracker '//\\LOGLINE//\\
            .ModStat(i) = .BaseStat(i) + ObjData.AddStat(WeaponObj, i) + ObjData.AddStat(ArmorObj, i) + ObjData.AddStat(WingsObj, i)
        Next i
        
        'War curse
        If UserList(UserIndex).Skills.WarCurse > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell WarCurse", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.Agi) = .ModStat(SID.Agi) - (UserList(UserIndex).Skills.WarCurse \ 4)
            .ModStat(SID.DEF) = .ModStat(SID.DEF) - (UserList(UserIndex).Skills.WarCurse \ 4)
            .ModStat(SID.Str) = .ModStat(SID.Str) - (UserList(UserIndex).Skills.WarCurse \ 4)
            .ModStat(SID.Mag) = .ModStat(SID.Mag) - (UserList(UserIndex).Skills.WarCurse \ 4)
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - (UserList(UserIndex).Skills.WarCurse \ 4)
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - (UserList(UserIndex).Skills.WarCurse \ 4)
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
            .ModStat(SID.Agi) = .ModStat(SID.Agi) + UserList(UserIndex).Skills.Bless \ 2
            .ModStat(SID.Mag) = .ModStat(SID.Mag) + UserList(UserIndex).Skills.Bless \ 2
            .ModStat(SID.Str) = .ModStat(SID.Str) + UserList(UserIndex).Skills.Bless \ 2
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + UserList(UserIndex).Skills.Bless \ 4
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + UserList(UserIndex).Skills.Bless \ 4
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + UserList(UserIndex).Skills.Bless \ 4
        End If
        
        'Iron skin
        If UserList(UserIndex).Skills.IronSkin > 0 Then
            Log "User_UpdateModStats: Updating effects of skill/spell Iron Skin", CodeTracker '//\\LOGLINE//\\
            .ModStat(SID.DEF) = .ModStat(SID.DEF) + UserList(UserIndex).Skills.IronSkin * 2
            .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) - UserList(UserIndex).Skills.IronSkin * 1.5
            .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) - UserList(UserIndex).Skills.IronSkin * 1.5
        End If
        
        'Min/max hit (damage) modification from strength and agility
        .ModStat(SID.MinHIT) = .ModStat(SID.MinHIT) + .ModStat(SID.Str) + (.ModStat(SID.Agi) \ 4)
        .ModStat(SID.MaxHIT) = .ModStat(SID.MaxHIT) + .ModStat(SID.Str) + (.ModStat(SID.Agi) \ 4)

    End With
    
End Sub

Public Function User_NearBankNPC(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Checks if the user is near enough to a banking NPC to use the banking
'*****************************************************************
Dim MinX As Integer
Dim MinY As Integer
Dim MaxX As Integer
Dim MaxY As Integer
Dim X As Byte
Dim Y As Byte

    'Set the tiles we will search through
    MinX = (UserList(UserIndex).Pos.X - MaxServerDistanceX)
    MaxX = (UserList(UserIndex).Pos.X + MaxServerDistanceX)
    MinY = (UserList(UserIndex).Pos.Y - MaxServerDistanceY)
    MaxY = (UserList(UserIndex).Pos.Y + MaxServerDistanceY)
    If MinX < 1 Then MinX = 1
    If MinY < 1 Then MinY = 1
    If MaxX > MapInfo(UserList(UserIndex).Pos.Map).Width Then MaxX = MapInfo(UserList(UserIndex).Pos.Map).Width
    If MaxY > MapInfo(UserList(UserIndex).Pos.Map).Height Then MaxY = MapInfo(UserList(UserIndex).Pos.Map).Height

    'Loop through the tiles near the user
    For X = MinX To MaxX
        For Y = MinY To MaxY
            
            'Check for a NPC index
            If MapInfo(UserList(UserIndex).Pos.Map).Data(X, Y).NPCIndex > 0 Then
                
                'NPC was found, check if it is a banker NPC AI
                If NPCList(MapInfo(UserList(UserIndex).Pos.Map).Data(X, Y).NPCIndex).AI = 6 Then
                
                    'Banker NPC found! Screw the rest of the loops, we got what we wanted
                    User_NearBankNPC = 1
                    Exit Function
                    
                End If
                
            End If
        
        Next Y
    Next X
    
    'If we got this far, no banker was found :(
            
End Function

Public Sub User_UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Use/Equip a inventory item
'*****************************************************************
Dim ObjIndex As Integer

    Log "Call User_UseInvItem(" & UserIndex & "," & Slot & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    On Error GoTo ErrOut
    If UserList(UserIndex).flags.UserLogged = 0 Then
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
    If UserList(UserIndex).Object(Slot).ObjIndex < 0 Then
        Log "User_UseInvItem: ObjIndex < 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Object(Slot).ObjIndex > NumObjDatas Then
        Log "User_UseInvItem: ObjIndex > NumObjDatas - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    On Error GoTo 0
    
    ObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
    
    'Check if the user can use the item due to class restrictions
    If Obj_ValidObjForClass(UserList(UserIndex).Class, ObjIndex) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(125).Data()
        Exit Sub
    End If

    With UserList(UserIndex).Stats
    
        'Check if the user can use the item due to stat restrictions
        If .BaseStat(SID.Agi) < ObjData.ReqAgi(ObjIndex) Then
            Data_Send ToIndex, UserIndex, cMessage(125).Data()
            Exit Sub
        End If
        If .BaseStat(SID.Mag) < ObjData.ReqMag(ObjIndex) Then
            Data_Send ToIndex, UserIndex, cMessage(125).Data()
            Exit Sub
        End If
        If .BaseStat(SID.Str) < ObjData.ReqStr(ObjIndex) Then
            Data_Send ToIndex, UserIndex, cMessage(125).Data()
            Exit Sub
        End If
        If .BaseStat(SID.ELV) < ObjData.ReqLvl(ObjIndex) Then
            Data_Send ToIndex, UserIndex, cMessage(125).Data()
            Exit Sub
        End If
    
        'Apply the replenish values
        .BaseStat(SID.MinHP) = .BaseStat(SID.MinHP) + (.ModStat(SID.MaxHP) * ObjData.RepHPP(ObjIndex)) + ObjData.RepHP(ObjIndex)
        .BaseStat(SID.MinMAN) = .BaseStat(SID.MinMAN) + (.ModStat(SID.MaxMAN) * ObjData.RepMPP(ObjIndex)) + ObjData.RepMP(ObjIndex)
        .BaseStat(SID.MinSTA) = .BaseStat(SID.MinSTA) + (.ModStat(SID.MaxSTA) * ObjData.RepSPP(ObjIndex)) + ObjData.RepSP(ObjIndex)
    
    End With

    Select Case ObjData.ObjType(ObjIndex)
        
        Case OBJTYPE_USEONCE
            Log "User_UseInvItem: ObjType = OBJTYPE_USEONCE", CodeTracker '//\\LOGLINE//\\
    
            'Remove from inventory
            UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - 1
            If UserList(UserIndex).Object(Slot).Amount <= 0 Then UserList(UserIndex).Object(Slot).ObjIndex = 0
            
            'Set the paper-doll
            If ObjData.SpriteHair(ObjIndex) <> -1 Then UserList(UserIndex).Char.Hair = ObjData.SpriteHair(ObjIndex)
            If ObjData.SpriteBody(ObjIndex) <> -1 Then UserList(UserIndex).Char.Body = ObjData.SpriteBody(ObjIndex)
            If ObjData.SpriteHead(ObjIndex) <> -1 Then UserList(UserIndex).Char.Head = ObjData.SpriteHead(ObjIndex)
            If ObjData.SpriteWeapon(ObjIndex) <> -1 Then UserList(UserIndex).Char.Weapon = ObjData.SpriteWeapon(ObjIndex)
            User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings
            
            'Create the graphic effect
            If ObjData.UseGrh(ObjIndex) > 0 Then
                ConBuf.PreAllocate 7
                ConBuf.Put_Byte DataCode.Server_MakeEffect
                ConBuf.Put_Byte UserList(UserIndex).Pos.X
                ConBuf.Put_Byte UserList(UserIndex).Pos.Y
                ConBuf.Put_Long ObjData.UseGrh(ObjIndex)
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
            End If
            
            'Create the sound effect
            If ObjData.UseSfx(ObjIndex) > 0 Then
                ConBuf.PreAllocate 4
                ConBuf.Put_Byte DataCode.Server_PlaySound3D
                ConBuf.Put_Byte UserList(UserIndex).Pos.X
                ConBuf.Put_Byte UserList(UserIndex).Pos.Y
                ConBuf.Put_Byte ObjData.UseSfx(ObjIndex)
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer
            End If
            
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
            UserList(UserIndex).WeaponType = ObjData.WeaponType(ObjIndex)
            
            'Update the weapon distance on the client
            ConBuf.PreAllocate 2
            ConBuf.Put_Byte DataCode.User_SetWeaponRange
            ConBuf.Put_Byte ObjData.WeaponRange(UserList(UserIndex).WeaponEqpObjIndex)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
            'Set the paper-doll
            If ObjData.SpriteHair(ObjIndex) <> -1 Then UserList(UserIndex).Char.Hair = ObjData.SpriteHair(ObjIndex)
            If ObjData.SpriteBody(ObjIndex) <> -1 Then UserList(UserIndex).Char.Body = ObjData.SpriteBody(ObjIndex)
            If ObjData.SpriteHead(ObjIndex) <> -1 Then UserList(UserIndex).Char.Head = ObjData.SpriteHead(ObjIndex)
            If ObjData.SpriteWeapon(ObjIndex) <> -1 Then UserList(UserIndex).Char.Weapon = ObjData.SpriteWeapon(ObjIndex)
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
            If ObjData.SpriteHair(ObjIndex) <> -1 Then UserList(UserIndex).Char.Hair = ObjData.SpriteHair(ObjIndex)
            If ObjData.SpriteBody(ObjIndex) <> -1 Then UserList(UserIndex).Char.Body = ObjData.SpriteBody(ObjIndex)
            If ObjData.SpriteHead(ObjIndex) <> -1 Then UserList(UserIndex).Char.Head = ObjData.SpriteHead(ObjIndex)
            If ObjData.SpriteWeapon(ObjIndex) <> -1 Then UserList(UserIndex).Char.Weapon = ObjData.SpriteWeapon(ObjIndex)
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
            If ObjData.SpriteWings(ObjIndex) <> -1 Then UserList(UserIndex).Char.Wings = ObjData.SpriteWings(ObjIndex)
            User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair, UserList(UserIndex).Char.Wings
    
        Case Else
        
            'We have no idea what type of object it is! OMG!!!
            Log "User_UseInvItem: Unknown object type used! Object type: " & ObjData.ObjType(ObjIndex), CriticalError '//\\LOGLINE//\\

    End Select
    
    'Force update of the modstats
    UserList(UserIndex).Stats.Update = 1

    'Update user's stats and inventory
    User_UpdateInv False, UserIndex, Slot
    
ErrOut:

End Sub

Public Sub User_WarpChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal ForceSwitch As Boolean = False)

'*****************************************************************
'Warps user to another spot
'*****************************************************************
Dim CorrectServer As Byte
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
    
    'Clear the pending quest NPC number and trading NPC
    UserList(UserIndex).flags.QuestNPC = 0
    UserList(UserIndex).flags.TradeWithNPC = 0

    If (OldMap <> Map) Or ForceSwitch = True Then
        Log "User_WarpChar: Switching maps", CodeTracker '//\\LOGLINE//\\
        
        'Check if the user is on the correct server, or needs to be switched
        CorrectServer = User_CorrectServer(UserList(UserIndex).Name, UserIndex, Map)
        
        'Check if it's the first user on the map and is the correct server
        If CorrectServer = 1 Then
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
            If MapInfo(Map).NumUsers = 1 Then
                Load_Maps_Temp Map
                ReDim MapUsers(Map).Index(1 To 1)
            Else
                ReDim Preserve MapUsers(Map).Index(1 To MapInfo(Map).NumUsers)
            End If
            MapUsers(Map).Index(MapInfo(Map).NumUsers) = UserIndex
        End If
        
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
            Unload_Map OldMap
            Erase MapUsers(OldMap).Index()
        End If
            
        'Set the new position
        User_EraseChar UserIndex
        UserList(UserIndex).Pos.X = X
        UserList(UserIndex).Pos.Y = Y
        UserList(UserIndex).Pos.Map = Map
        
        'Check if the user is on the correct server
        If CorrectServer = 0 Then

            'Disconnect the user
            UserList(UserIndex).flags.Disconnecting = 1
        
        'User is already on the correct server
        Else
        
            'Check if it is a new map - if so, load the new map if needed
            If OldMap <> Map Then Load_Maps_Temp Map

            'Tell client to try switching maps
            ConBuf.PreAllocate 6
            ConBuf.Put_Byte DataCode.Map_LoadMap
            ConBuf.Put_Integer Map
            ConBuf.Put_Integer MapInfo(Map).MapVersion
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
            'Show Character to others
            User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y
            
            'Move the user's slaves
            For LoopC = 1 To UserList(UserIndex).NumSlaves
            
                With NPCList(UserList(UserIndex).SlaveNPCIndex(LoopC))

                    'Erase the NPC from the old map
                    MapInfo(.Pos.Map).Data(.Pos.X, .Pos.Y).NPCIndex = 0
                
                    'Send erase command to clients
                    ConBuf.PreAllocate 3
                    ConBuf.Put_Byte DataCode.Server_EraseChar
                    ConBuf.Put_Integer .Char.CharIndex
                    Data_Send ToMap, 0, ConBuf.Get_Buffer, .Pos.Map
                    
                    'Set the new position
                    Server_ClosestLegalPos UserList(UserIndex).Pos, .Pos
                    If Not Server_LegalPos(.Pos.Map, .Pos.X, .Pos.Y, 0) Then
                        NPC_Close UserList(UserIndex).SlaveNPCIndex(LoopC)
                    Else
                        NPC_MakeChar ToMap, UserIndex, UserList(UserIndex).SlaveNPCIndex(LoopC), .Pos.Map, .Pos.X, .Pos.Y
                    End If
                    
                End With
                
            Next LoopC
                
            'Check to update the database
            If MySQLUpdate_UserMap Then
                Log "User_WarpChar: Updating database with new map", CodeTracker '//\\LOGLINE//\\
                DB_RS.Open "SELECT * FROM users WHERE `name`='" & UserList(UserIndex).Name & "'", DB_Conn, adOpenStatic, adLockOptimistic
                DB_RS!pos_map = Map
                DB_RS.Update
                DB_RS.Close
            End If
            
        End If
            
    Else
        
        'User didn't change maps, just move their position
        Log "User_WarpChar: Moving user, map is not changing", CodeTracker '//\\LOGLINE//\\
        
        'Remove the user from the tile
        MapInfo(Map).Data(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        
        'Update their position
        UserList(UserIndex).Pos.X = X
        UserList(UserIndex).Pos.Y = Y
        
        'Set them on the new tile
        MapInfo(Map).Data(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        
        'Send the update packet to everyone on the map
        User_MakeChar ToMap, UserIndex, UserIndex, Map, X, Y
        
        'Update the user's char index
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_UserCharIndex
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
        'Move the user's slaves
        For LoopC = 1 To UserList(UserIndex).NumSlaves
        
            With NPCList(UserList(UserIndex).SlaveNPCIndex(LoopC))

                'Erase the NPC from the old map
                MapInfo(.Pos.Map).Data(.Pos.X, .Pos.Y).NPCIndex = 0
            
                'Send erase command to clients
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.Server_EraseChar
                ConBuf.Put_Integer .Char.CharIndex
                Data_Send ToMap, 0, ConBuf.Get_Buffer, .Pos.Map
                
                'Set the new position
                Server_ClosestLegalPos UserList(UserIndex).Pos, .Pos
                If Not Server_LegalPos(.Pos.Map, .Pos.X, .Pos.Y, 0) Then
                    NPC_Close UserList(UserIndex).SlaveNPCIndex(LoopC)
                Else
                    NPC_MakeChar ToMap, UserIndex, UserList(UserIndex).SlaveNPCIndex(LoopC), .Pos.Map, .Pos.X, .Pos.Y
                End If
                
            End With
            
        Next LoopC
        
    End If

End Sub

Public Sub Group_AddUser(ByVal UserIndex As Integer, ByVal GroupIndex As Byte)

'*****************************************************************
'Adds a user to an existing group
'*****************************************************************

    'Check for valid group information
    If GroupIndex > NumGroups Then Exit Sub
    If GroupIndex < 1 Then Exit Sub
    If GroupData(GroupIndex).NumUsers = 0 Then Exit Sub
    If GroupData(GroupIndex).NumUsers >= Group_MaxUsers Then
        ConBuf.PreAllocate 3 + Len(UserList(GroupData(GroupIndex).Users(1)).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 106
        ConBuf.Put_String UserList(GroupData(GroupIndex).Users(1)).Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    
    'Add the user to the group
    GroupData(GroupIndex).NumUsers = GroupData(GroupIndex).NumUsers + 1
    ReDim Preserve GroupData(GroupIndex).Users(1 To GroupData(GroupIndex).NumUsers)
    GroupData(GroupIndex).Users(GroupData(GroupIndex).NumUsers) = UserIndex
    UserList(UserIndex).GroupIndex = GroupIndex
    
    'Join group message
    ConBuf.PreAllocate 3 + Len(UserList(GroupData(GroupIndex).Users(1)).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 107
    ConBuf.Put_String UserList(GroupData(GroupIndex).Users(1)).Name
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
    'Tell the group members the user joined
    ConBuf.PreAllocate 3 + Len(UserList(UserIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 108
    ConBuf.Put_String UserList(UserIndex).Name
    Data_Send ToGroupButIndex, UserIndex, ConBuf.Get_Buffer
    
End Sub

Public Sub Group_EXPandGold(ByVal UserIndex As Integer, ByVal GroupIndex As Integer, ByVal EXP As Long, ByVal Gold As Long)

'*****************************************************************
'Splits up the EXP and Gold gained among group users in range
'*****************************************************************
Dim GiveUsers() As Integer
Dim NumUsersInRange As Byte
Dim tIndex As Integer
Dim i As Byte

    'Check for a valid group
    If GroupIndex = 0 Then Exit Sub

    'Default to give to all users
    ReDim GiveUsers(1 To GroupData(GroupIndex).NumUsers)

    'Loop through and find the users to give the exp to
    For i = 1 To GroupData(GroupIndex).NumUsers
    
        'Hold the index in a smaller variable
        tIndex = GroupData(GroupIndex).Users(i)
    
        'Confirm that it is a valid index
        If tIndex > 0 Then
        
            'Check if the user is on the same map
            If UserList(UserIndex).Pos.Map = UserList(tIndex).Pos.Map Then
            
                'Check the distance
                If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, Group_MaxDistanceX, Group_MaxDistanceY) Then
                    
                    'The user is in range, set them in the array to get the exp
                    NumUsersInRange = NumUsersInRange + 1
                    GiveUsers(NumUsersInRange) = tIndex
                    
                End If
                
            End If
        
        End If
        
    Next i
    
    'Divide the gold and exp among the users
    EXP = (EXP \ NumUsersInRange) + 1
    Gold = (Gold \ NumUsersInRange) + 1
    
    'Give the exp to all the users
    For i = 1 To NumUsersInRange
        User_RaiseExp GiveUsers(i), EXP
        UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + Gold
    Next i
    
    'Clear the GiveUsers array
    Erase GiveUsers()

End Sub

Public Sub Group_RemoveUser(ByVal UserIndex As Integer, ByVal GroupIndex As Integer)

'*****************************************************************
'Removes a user from an existing group
'*****************************************************************
Dim i As Byte
Dim j As Byte

    'Check for valid group information
    If GroupIndex > NumGroups Then Exit Sub
    If GroupIndex < 1 Then Exit Sub
    If GroupData(GroupIndex).NumUsers = 0 Then Exit Sub 'Group deleted

    'Tell the user they have left the group
    Data_Send ToIndex, UserIndex, cMessage(109).Data()
    
    'Check how many people are left in the group
    If GroupData(GroupIndex).NumUsers > 1 Then
    
        'Tell everyone else they have left
        ConBuf.PreAllocate 3 + Len(UserList(UserIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 110
        ConBuf.Put_String UserList(UserIndex).Name
        Data_Send ToGroupButIndex, UserIndex, ConBuf.Get_Buffer
        
        'Clear the user's group flag
        UserList(UserIndex).GroupIndex = 0
    
    Else
    
        'This is the last person, so just empty the group out
        GroupData(GroupIndex).NumUsers = 0
        Erase GroupData(GroupIndex).Users
        
        'Raise the empty group count
        NumEmptyGroups = NumEmptyGroups + 1
        
        'Clear the user's group flag
        UserList(UserIndex).GroupIndex = 0
        
        Exit Sub
        
    End If
    
    'Update everyone in the map that is part of the gorup that the user has changed out of the group
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_ChangeCharType
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte ClientCharType_PC
    Data_Send ToMapGroupButIndex, UserIndex, ConBuf.Get_Buffer
    
    'Find the slot the user has in the group
    For i = 1 To GroupData(GroupIndex).NumUsers
    
        'Index found
        If GroupData(GroupIndex).Users(i) = UserIndex Then

            'If the user is the last one in the group, just resize the array
            If GroupData(GroupIndex).NumUsers = i Then
                GroupData(GroupIndex).NumUsers = GroupData(GroupIndex).NumUsers - 1
                ReDim Preserve GroupData(GroupIndex).Users(1 To GroupData(GroupIndex).NumUsers)
                Exit Sub
            End If
            
            'The user is not at the end of the array, and theres more then one user
            For j = i To GroupData(GroupIndex).NumUsers - 1
                
                'Move all the users in the group down in the array to fill the now empty slow
                GroupData(GroupIndex).Users(j) = GroupData(GroupIndex).Users(j + 1)
                
            Next j
            
            'Remove the left-over slot at the end
            GroupData(GroupIndex).NumUsers = GroupData(GroupIndex).NumUsers - 1
            ReDim Preserve GroupData(GroupIndex).Users(1 To GroupData(GroupIndex).NumUsers)
            Exit Sub
            
        End If
        
    Next i
    
End Sub

Public Function Group_Create(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Find the next free group index and return it
'*****************************************************************
Dim i As Byte

    'Check if theres any groups yet
    If NumGroups = 0 Then
        NumGroups = 1
        Group_Create = 1
        ReDim GroupData(1 To 1) As GroupData
        Exit Function
    End If
    
    'Check if there are any empty groups - if so, find out which index it is
    If NumEmptyGroups > 0 Then
        For i = 1 To NumGroups
            If GroupData(i).NumUsers = 0 Then
                
                'Found a group not in use, use it
                NumEmptyGroups = NumEmptyGroups - 1 'We took one of the empty groups
                Group_Create = i
                Exit Function
                
            End If
        Next i
    End If
    
    'No groups found, check if we can make a new one
    If NumGroups + 1 >= Group_MaxGroups Then
        Data_Send ToIndex, UserIndex, cMessage(111).Data()
        Exit Function
    End If
    
    'Add the new group
    NumGroups = NumGroups + 1
    ReDim Preserve GroupData(1 To NumGroups)
    Group_Create = NumGroups

End Function

Public Sub Quest_SendText(ByVal UserIndex As Integer, Optional ByVal QuestIndex As Byte = 0)

'*****************************************************************
'Sends the active quest information to the user
'*****************************************************************
Dim i As Byte

    'No index specified, update them all
    If QuestIndex = 0 Then
        ConBuf.Clear
        For i = 1 To MaxQuests
            If UserList(UserIndex).Quest(i) > 0 Then
                DB_RS.Open "SELECT text_info FROM quests WHERE `id`='" & UserList(UserIndex).Quest(i) & "'"
                ConBuf.Allocate 3
                ConBuf.Put_Byte DataCode.Server_SendQuestInfo
                ConBuf.Put_Byte i
                ConBuf.Put_String QuestData(UserList(UserIndex).Quest(i)).Name
                ConBuf.Put_StringEX DB_RS(0)
                DB_RS.Close
            Else
                ConBuf.Allocate 3
                ConBuf.Put_Byte DataCode.Server_SendQuestInfo
                ConBuf.Put_Byte i
                ConBuf.Put_String vbNullString
            End If
        Next i
    
    'Index specified, update only that index
    Else
        If UserList(UserIndex).Quest(QuestIndex) > 0 Then
            DB_RS.Open "SELECT text_info FROM quests WHERE `id`='" & QuestIndex & "'"
            ConBuf.PreAllocate 3
            ConBuf.Put_Byte DataCode.Server_SendQuestInfo
            ConBuf.Put_Byte QuestIndex
            ConBuf.Put_String QuestData(QuestIndex).Name
            ConBuf.Put_StringEX DB_RS(0)
            DB_RS.Close
        Else
            ConBuf.PreAllocate 3
            ConBuf.Put_Byte DataCode.Server_SendQuestInfo
            ConBuf.Put_Byte QuestIndex
            ConBuf.Put_String vbNullString
        End If
    End If
    
    'If we put information in the buffer, send it
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
End Sub
