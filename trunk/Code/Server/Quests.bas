Attribute VB_Name = "Quests"
Option Explicit

Private Sub Quest_CheckIfComplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal UserQuestSlot As Byte)

'*****************************************************************
'Checks if a quest is ready to be completed
'*****************************************************************
Dim ReqObjSlot As Byte
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
                
                'Cache the required object slot (we will be using it later if the quest is complete)
                ReqObjSlot = Slot

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

    '*** Quest was completed! ***

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
        UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold
        ConBuf.PreAllocate 6
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 4
        ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
    'Remove the quest object
    If ReqObjSlot > 0 Then
        User_RemoveObj UserIndex, ReqObjSlot, QuestData(NPCList(NPCIndex).Quest).FinishReqObjAmount
    End If
    
    'Give the object reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObj > 0 Then
        User_GiveObj UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObj, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObjAmount
    End If
    
    'Learn skill reward
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill > 0 Then
        User_GiveSkill UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill
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

Public Sub Quest_Cancel(ByVal UserIndex As Integer, ByVal QuestSlot As Byte)

'*****************************************************************
'Cancel a user's quest by removing the flags
'*****************************************************************

    If QuestSlot = 0 Then Exit Sub
    If QuestSlot > MaxQuests Then Exit Sub

    UserList(UserIndex).QuestStatus(QuestSlot).NPCKills = 0
    UserList(UserIndex).Quest(QuestSlot) = 0
    
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 129
    ConBuf.Put_Byte QuestSlot
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Public Sub Quest_General(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

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
            DB_RS.Open "SELECT text_info FROM quests WHERE `id`='" & UserList(UserIndex).Quest(QuestIndex) & "'"
            ConBuf.PreAllocate 3
            ConBuf.Put_Byte DataCode.Server_SendQuestInfo
            ConBuf.Put_Byte QuestIndex
            ConBuf.Put_String QuestData(UserList(UserIndex).Quest(QuestIndex)).Name
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

