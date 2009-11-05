Attribute VB_Name = "GameLogic"
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

Sub Obj_Erase(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal Num As Integer, ByVal Map As Byte, ByVal x As Integer, ByVal Y As Integer)

'*****************************************************************
'Erase a object
'*****************************************************************

    If Num = -1 Then Num = MapData(Map, x, Y).ObjInfo.Amount

    MapData(Map, x, Y).ObjInfo.Amount = MapData(Map, x, Y).ObjInfo.Amount - Num

    If MapData(Map, x, Y).ObjInfo.Amount <= 0 Then
        MapData(Map, x, Y).ObjInfo.ObjIndex = 0
        MapData(Map, x, Y).ObjInfo.Amount = 0
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Obj_Eraseect
        ConBuf.Put_Byte CByte(x)
        ConBuf.Put_Byte CByte(Y)
        Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, Map
    End If

End Sub

Sub Obj_Make(ByVal sndRoute As Byte, ByVal sndIndex As Integer, Obj As Obj, ByVal x As Integer, ByVal Y As Integer)

'*****************************************************************
'Erase a object
'*****************************************************************

    MapData(UserList(sndIndex).Pos.Map, x, Y).ObjInfo = Obj
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Obj_Makeect
    ConBuf.Put_Integer ObjData(Obj.ObjIndex).GrhIndex
    ConBuf.Put_Byte CByte(x)
    ConBuf.Put_Byte CByte(Y)
    Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, UserList(sndIndex).Pos.Map

End Sub

Function Quest_BuildReqString(ByVal QuestID As Integer) As String

'*****************************************************************
'Builds the string that says what is required for the quest
'*****************************************************************

Dim FileNum As Byte
Dim TempNPC As NPC
Dim S As String

'Get the target NPC's name if there is one - to do this, we have to open up the NPC file since we dont store the "defaults" like we do with objects/quests/etc

    If QuestData(QuestID).FinishReqNPC Then
        FileNum = FreeFile
        Open App.Path & "\NPCs\" & QuestData(QuestID).FinishReqNPC & ".npc" For Binary As FileNum
        Get #FileNum, , TempNPC
        Close #FileNum
    End If

    'We must put a must, or else no must will be given, and it IS A MUST!!!
    S = "You must "

    'See if we need to pop some caps in any homies
    If QuestData(QuestID).FinishReqNPC Then S = S & "kill " & QuestData(QuestID).FinishReqNPCAmount & " " & TempNPC.Name & "s"

    'See if we need to acquire any bling-blings
    If QuestData(QuestID).FinishReqObj Then

        'Make sure we use proper grammar since we are civilized townfolk
        If QuestData(QuestID).FinishReqNPC Then S = S & " and "

        'Put the object requirement string
        S = S & " get " & QuestData(QuestID).FinishReqObjAmount & " " & ObjData(QuestData(QuestID).FinishReqObj).Name & "s"

    End If

    'Ain't nuttin like a period at da end of a statement.
    S = S & "."

    'Return the string
    Quest_BuildReqString = S

End Function

Sub Quest_CheckIfComplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal UserQuestSlot As Byte)

'*****************************************************************
'Checks if a quest is ready to be completed
'*****************************************************************

Dim Slot As Byte

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
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp > 0 Then
        User_RaiseExp UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You got " & QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewExp & " experience!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold > 0 Then
        UserList(UserIndex).Gold = UserList(UserIndex).Gold + QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You got " & QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewGold & " gold!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObj > 0 Then
        User_GiveObj UserIndex, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObj, QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishRewObjAmount
    End If
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill > 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        If UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) = 1 Then
            ConBuf.Put_String "You already know " & Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) & "."
        Else
            ConBuf.Put_String "You have learned " & Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) & "!"
        End If
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(UserQuestSlot)).FinishLearnSkill) = 1
    End If

    'Add the quest to the user's finished quest list
    If QuestData(UserList(UserIndex).Quest(UserQuestSlot)).Redoable Then

        'Only add a redoable quest in the list once
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

'Check for valid values

    If UserIndex <= 0 Then Exit Sub
    If UserIndex > LastUser Then Exit Sub
    If NPCIndex <= 0 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If NPCList(NPCIndex).Quest <= 0 Then Exit Sub

    'Check if the user is currently involved in the quest
    For i = 1 To MaxQuests

        'If they are involved in a quest, then we will send it off to another sub
        If UserList(UserIndex).Quest(i) = NPCList(NPCIndex).Quest Then
            Quest_CheckIfComplete UserIndex, NPCIndex, i
            Exit Sub
        End If

    Next i

    'The user is not involved in this quest currently - check if they have already completed it
    If InStr(1, UserList(UserIndex).CompletedQuests & "-", "-" & NPCList(NPCIndex).Quest & "-") Then

        'The user has completed this quest before, so check if it is redoable
        If QuestData(NPCList(NPCIndex).Quest).Redoable = 0 Then

            'The quest is not redoable, so sorry dude, no quest fo' j00
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You have already completed this quest!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Quest
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
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Type /accept to accept the quest."
    ConBuf.Put_Byte DataCode.Comm_FontType_Quest
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Set the pending quest to the selected quest
    UserList(UserIndex).Flags.QuestNPC = NPCIndex

End Sub

Sub Quest_SayIncomplete(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Make the targeted NPC say the "incomplete quest" text
'*****************************************************************

'Incomplete text

    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String NPCList(NPCIndex).Name & ": " & QuestData(NPCList(NPCIndex).Quest).IncompleteTxt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
    Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer

    'Requirements text
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String Quest_BuildReqString(NPCList(NPCIndex).Quest)
    ConBuf.Put_Byte DataCode.Comm_FontType_Quest
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Function Server_CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean

'*****************************************************************
'Checks for a user with the same IP
'*****************************************************************

Dim LoopC As Long

    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged = 1 Then
            If UserList(LoopC).IP = UserIP And UserIndex <> LoopC Then
                Server_CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC

    Server_CheckForSameIP = False

End Function

Function Server_CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean

'*****************************************************************
'Checks for a user with the same Name
'*****************************************************************

Dim LoopC As Long

    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged = 1 Then
            If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserIndex <> LoopC Then
                Server_CheckForSameName = True
                Exit Function
            End If
        End If
    Next LoopC

    Server_CheckForSameName = False

End Function

Function Server_CheckTargetedDistance(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Checks if a user is targeting a character in range
'*****************************************************************

Dim TargetID As Integer

    Select Case UserList(UserIndex).Flags.Target

        'User
    Case 1
        TargetID = UserList(UserIndex).Flags.TargetIndex
        If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, UserList(TargetID).Pos.x, UserList(TargetID).Pos.Y) <= Max_Server_Distance Then
            Server_CheckTargetedDistance = 1
            Exit Function
        End If

        'NPC
    Case 2
        TargetID = UserList(UserIndex).Flags.TargetIndex
        If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, NPCList(TargetID).Pos.x, NPCList(TargetID).Pos.Y) <= Max_Server_Distance Then
            Server_CheckTargetedDistance = 1
            Exit Function
        End If

    End Select

    'Not in distance or nothing targeted
    If TargetID Or UserList(UserIndex).Flags.TargetIndex Then
        UserList(UserIndex).Flags.Target = 0
        UserList(UserIndex).Flags.TargetIndex = 0
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Target
        ConBuf.Put_Integer 0
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If

End Function

Sub Server_ClosestLegalPos(Pos As WorldPos, nPos As WorldPos)

'*****************************************************************
'Finds the closest legal tile to Pos and stores it in nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

'Set the new map

    nPos.Map = Pos.Map

    'Keep looping while the position is not legal
    Do While Not Server_LegalPos(Pos.Map, nPos.x, nPos.Y, 0)

        'If we have checked too much, then just leave
        If LoopC > 3 Then   'How many tiles in all directions to search
            Notfound = True
            Exit Do
        End If

        'Loop through the tiles
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.x - LoopC To Pos.x + LoopC

                'Check if the position is legal
                If Server_LegalPos(nPos.Map, tX, tY, 0) = True Then
                    nPos.x = tX
                    nPos.Y = tY
                    tX = Pos.x + LoopC
                    tY = Pos.Y + LoopC
                End If

            Next tX
        Next tY

        'Check the next set of tiles
        LoopC = LoopC + 1

    Loop

    'If no position was found, return empty positions
    If Notfound Then
        nPos.x = 0
        nPos.Y = 0
    End If

End Sub

Sub Server_DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)

'*****************************************************************
'Do any events on a tile
'*****************************************************************

Dim TempPos As WorldPos
Dim NewPos As WorldPos

'Check for tile exit

    If MapData(Map, x, Y).TileExit.Map Then

        'Set the position values
        TempPos.x = MapData(Map, x, Y).TileExit.x
        TempPos.Y = MapData(Map, x, Y).TileExit.Y
        TempPos.Map = MapData(Map, x, Y).TileExit.Map

        'Get the closest legal position
        Server_ClosestLegalPos TempPos, NewPos

        'If the position is legal, then warp the user there
        If Server_LegalPos(NewPos.Map, NewPos.x, NewPos.Y, 0) Then User_WarpChar UserIndex, MapData(Map, x, Y).TileExit.Map, MapData(Map, x, Y).TileExit.x, MapData(Map, x, Y).TileExit.Y

    End If

End Sub

Function Server_FindDirection(Pos As WorldPos, Target As WorldPos) As Byte

'*****************************************************************
'Returns the direction in which the Target is from the Pos, 0 if equal
'*****************************************************************

Dim x As Integer
Dim Y As Integer

    x = Pos.x - Target.x
    Y = Pos.Y - Target.Y

    'NE
    If x <= -1 Then
        If Y >= 1 Then
            Server_FindDirection = NORTHEAST
            Exit Function
        End If
    End If

    'NW
    If x >= 1 Then
        If Y >= 1 Then
            Server_FindDirection = NORTHWEST
            Exit Function
        End If
    End If

    'SW
    If x >= 1 And Y <= -1 Then
        Server_FindDirection = SOUTHWEST
        Exit Function
    End If

    'SE
    If x <= -1 Then
        If Y <= -1 Then
            Server_FindDirection = SOUTHEAST
            Exit Function
        End If
    End If

    'South
    If Y <= -1 Then
        Server_FindDirection = SOUTH
        Exit Function
    End If

    'north
    If Y >= 1 Then
        Server_FindDirection = NORTH
        Exit Function
    End If

    'West
    If x >= 1 Then
        Server_FindDirection = WEST
        Exit Function
    End If

    'East
    If x <= -1 Then
        Server_FindDirection = EAST
        Exit Function
    End If

End Function

Sub Server_HeadToPos(ByVal Head As Byte, ByRef Pos As WorldPos)

'*****************************************************************
'Takes Pos and moves it in heading direction
'*****************************************************************

Dim x As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

    x = Pos.x
    Y = Pos.Y

    If Head = NORTH Then
        nX = x
        nY = Y - 1
    End If

    If Head = SOUTH Then
        nX = x
        nY = Y + 1
    End If

    If Head = EAST Then
        nX = x + 1
        nY = Y
    End If

    If Head = WEST Then
        nX = x - 1
        nY = Y
    End If

    If Head = NORTHEAST Then
        nX = x + 1
        nY = Y - 1
    End If

    If Head = SOUTHEAST Then
        nX = x + 1
        nY = Y + 1
    End If

    If Head = SOUTHWEST Then
        nX = x - 1
        nY = Y + 1
    End If

    If Head = NORTHWEST Then
        nX = x - 1
        nY = Y - 1
    End If

    'return values
    Pos.x = nX
    Pos.Y = nY

End Sub

Function Server_InMapBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean

'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

    If x > MinXBorder Then
        If x < MaxXBorder Then
            If Y > MinYBorder Then
                If Y < MaxYBorder Then Server_InMapBounds = True
            End If
        End If
    End If
    
End Function

Function Server_LegalPos(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal Heading As Byte) As Boolean

'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Make sure it's a legal map

    If Map <= 0 Then Exit Function
    If Map > NumMaps Then Exit Function

    'Check to see if its out of bounds
    If x < MinXBorder Then Exit Function
    If x > MaxXBorder Then Exit Function
    If Y < MinYBorder Then Exit Function
    If Y > MaxYBorder Then Exit Function

    'Check if a character (User or NPC) is already at the tile
    If MapData(Map, x, Y).UserIndex > 0 Then Exit Function
    If MapData(Map, x, Y).NPCIndex > 0 Then Exit Function

    'Check to see if its blocked
    If MapData(Map, x, Y).Blocked = BlockedAll Then Exit Function

    'Check the heading for directional blocking
    If Heading > 0 Then
        If MapData(Map, x, Y).Blocked And BlockedNorth Then
            If Heading = NORTH Then Exit Function
            If Heading = NORTHEAST Then Exit Function
            If Heading = NORTHWEST Then Exit Function
        End If
        If MapData(Map, x, Y).Blocked And BlockedEast Then
            If Heading = EAST Then Exit Function
            If Heading = NORTHEAST Then Exit Function
            If Heading = SOUTHEAST Then Exit Function
        End If
        If MapData(Map, x, Y).Blocked And BlockedSouth Then
            If Heading = SOUTH Then Exit Function
            If Heading = SOUTHEAST Then Exit Function
            If Heading = SOUTHWEST Then Exit Function
        End If
        If MapData(Map, x, Y).Blocked And BlockedWest Then
            If Heading = WEST Then Exit Function
            If Heading = NORTHWEST Then Exit Function
            If Heading = SOUTHWEST Then Exit Function
        End If
    End If

    'If we are still in this routine, then it must be a legal position
    Server_LegalPos = True

End Function


Function Server_NextOpenCharIndex() As Integer

'*****************************************************************
'Finds the next open CharIndex in Charlist
'*****************************************************************

Dim LoopC As Long

'Check for the first char creation

    If LastChar = 0 Then
        ReDim CharList(1 To 1)
        LastChar = 1
        Server_NextOpenCharIndex = 1
        Exit Function
    End If

    'Loop through the character slots
    For LoopC = 1 To LastChar + 1

        'We need to create a new slot
        If LoopC > LastChar Then
            LastChar = LoopC
            Server_NextOpenCharIndex = LoopC
            ReDim Preserve CharList(1 To LastChar)
            Exit Function
        End If

        'Re-use an old slot that is not being used
        If CharList(LoopC).Index = 0 Then
            Server_NextOpenCharIndex = LoopC
            Exit Function
        End If

    Next LoopC

End Function

Public Function Server_SkillIDtoSkillName(ByVal SkillID As Byte) As String

'***************************************************
'Takes in a SkillID and returns the name of that skill
'***************************************************

    Select Case SkillID
    Case SkID.Bless: Server_SkillIDtoSkillName = "Bless"
    Case SkID.IronSkin: Server_SkillIDtoSkillName = "Iron Skin"
    Case SkID.Strengthen: Server_SkillIDtoSkillName = "Strengthen"
    Case SkID.Warcry: Server_SkillIDtoSkillName = "War Cry"
    Case SkID.Protection: Server_SkillIDtoSkillName = "Protection"
    Case SkID.Curse: Server_SkillIDtoSkillName = "Curse"
    Case SkID.SpikeField: Server_SkillIDtoSkillName = "Spike Field"
    Case SkID.Heal: Server_SkillIDtoSkillName = "Heal"
    Case Else: Server_SkillIDtoSkillName = "Unknown Skill"
    End Select

End Function

Public Sub Server_WriteMail(WriterIndex As Integer, RecieverName As String, Subject As String, Message As String, ObjIndexString As String, ObjAmountString As String)

Dim MailIndex As Integer
Dim MailData As MailData
Dim TempSplit() As String
Dim TempSplit2() As String
Dim TempUser As User
Dim LoopC As Byte
Dim LoopX As Byte

'Check for a valid reciever name

    If Server_FileExist(CharPath & UCase$(RecieverName) & ".chr", vbNormal) = False Then
        If WriterIndex <> -1 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "User" & RecieverName & " does not exist!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
        End If
        Exit Sub
    End If

    'Get the next open mail slot
    Do
        MailIndex = MailIndex + 1
        If MailIndex > MaxMail Then Exit Sub
    Loop While Server_FileExist(App.Path & "\Mail\" & MailIndex & ".mail", vbNormal)

    'Set up the mail type
    MailData.New = 1
    MailData.Message = Message
    MailData.RecieveDate = Date
    MailData.Subject = Subject
    If WriterIndex <> -1 Then
        MailData.WriterName = UserList(WriterIndex).Name
    Else
        MailData.WriterName = "Game Admin"
    End If

    'Split up the object index string
    TempSplit = Split(ObjIndexString, ",")
    For LoopC = 0 To UBound(TempSplit())
        MailData.Obj(LoopC + 1).ObjIndex = TempSplit(LoopC)
    Next LoopC

    'Split up the object amount string
    TempSplit2 = Split(ObjAmountString, ",")
    For LoopC = 0 To UBound(TempSplit2())
        MailData.Obj(LoopC + 1).Amount = TempSplit2(LoopC)
    Next LoopC

    'Check if the reciever is on
    For LoopC = 1 To LastUser
        If UserList(LoopC).Flags.UserLogged Then
            If UCase$(UserList(LoopC).Name) = UCase$(RecieverName) Then

                'Get the user's next open MailID slot
                LoopX = 0
                Do
                    LoopX = LoopX + 1
                    If LoopX > MaxMailPerUser Then
                        If WriterIndex <> -1 Then
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Comm_Talk
                            ConBuf.Put_String UserList(WriterIndex).Name & " tried to send you a message, but you could not recieve it because your mailbox is full!"
                            ConBuf.Put_Byte DataCode.Comm_FontType_Info
                            Data_Send ToIndex, LoopC, ConBuf.Get_Buffer
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Comm_Talk
                            ConBuf.Put_String RecieverName & " can not recieve any more mail because their mailbox is full!"
                            ConBuf.Put_Byte DataCode.Comm_FontType_Info
                            Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                            Exit Sub
                        Else
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.Comm_Talk
                            ConBuf.Put_String "The Game Admin tried to send you a message, but you could not recieve it because your mailbox is full!"
                            ConBuf.Put_Byte DataCode.Comm_FontType_Info
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
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "You message has been sent to " & RecieverName & " successfully!"
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "You have recieved a new message from " & UserList(WriterIndex).Name & "!"
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, LoopC, ConBuf.Get_Buffer
                Else
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "You have recieved a new message from The Game Admin!"
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, LoopC, ConBuf.Get_Buffer
                End If
                Exit Sub

            End If
        End If
    Next LoopC

    'The user is not on, so load up his character data and impliment it into the character
    Set TempUser.Stats = New UserStats
    Load_User TempUser, CharPath & UCase$(RecieverName) & ".chr"
    LoopC = 0
    Do
        LoopC = LoopC + 1
        If LoopC > MaxMailPerUser Then
            If WriterIndex <> -1 Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "Could not send message because " & RecieverName & "'s mailbox is full!"
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
            End If
            Exit Sub
        End If
    Loop While TempUser.MailID(LoopC) > 0

    'Load the mail data into the temp character
    TempUser.MailID(LoopC) = MailIndex

    'Save the temp user
    Save_User TempUser, CharPath & UCase$(RecieverName) & ".chr"

    'Send the message of success
    If WriterIndex <> -1 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "Message was successfully sent to " & RecieverName & "!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, WriterIndex, ConBuf.Get_Buffer
    End If

    'Save the mail
    Save_Mail MailIndex, MailData

End Sub

Public Sub User_AddObjToInv(ByVal UserIndex As Integer, ByRef Object As Obj)

Dim LoopC As Long

'Look for a slot

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
    Obj_Make ToMap, UserIndex, Object, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y

End Sub

Sub User_Attack(ByVal UserIndex As Integer)

'*****************************************************************
'Begin a user attack sequence
'*****************************************************************

Dim AttackPos As WorldPos

'Check for invalid values

    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.DownloadingMap Then Exit Sub
    If UserList(UserIndex).Stats.ModStat(SID.MinSTA) <= 0 Then Exit Sub
    If UserList(UserIndex).Counters.AttackCounter > timeGetTime - STAT_ATTACKWAIT Then Exit Sub

    'Update counters
    UserList(UserIndex).Counters.AttackCounter = timeGetTime

    'Get tile user is attacking
    AttackPos = UserList(UserIndex).Pos
    Server_HeadToPos UserList(UserIndex).Char.Heading, AttackPos

    'Exit if not legal
    If AttackPos.x < XMinMapSize Or AttackPos.x > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then Exit Sub

    'Look for user
    If MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).UserIndex > 0 Then

        'Play attack sound
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_PlaySound
        ConBuf.Put_Byte SOUND_SWING
        Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

        'Go to the user attacking user sub
        'User_AttackUser UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "PVP is currently disabled."
        ConBuf.Put_Byte DataCode.Comm_FontType_Fight
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub

    End If

    'Look for NPC
    If MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NPCIndex > 0 Then
        If NPCList(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NPCIndex).Attackable Then

            'Play attack sound
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_PlaySound
            ConBuf.Put_Byte SOUND_SWING
            Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

            'Go to user attacking npc sub
            User_AttackNPC UserIndex, MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NPCIndex

        Else

            'Can not attack the selected NPC, NPC is not attackable
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "A mysterious force prevents you from attacking..."
            ConBuf.Put_Byte DataCode.Comm_FontType_Fight
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        End If
        Exit Sub
    End If

End Sub

Sub User_AttackNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Have a User attack a NPC
'*****************************************************************

Dim HitSkill As Long    'User hit skill
Dim Hit As Integer      'Hit damage

'Get the user hit skill

    If UserList(UserIndex).WeaponType = Hand Then
        HitSkill = UserList(UserIndex).Stats.ModStat(SID.Fist)
    ElseIf UserList(UserIndex).WeaponType = Dagger Then
        HitSkill = UserList(UserIndex).Stats.ModStat(SID.Dagger)
    ElseIf UserList(UserIndex).WeaponType = Staff Then
        HitSkill = UserList(UserIndex).Stats.ModStat(SID.Staff)
    ElseIf UserList(UserIndex).WeaponType = Sword Then
        HitSkill = UserList(UserIndex).Stats.ModStat(SID.Staff)
    End If

    'Display the attack
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Attack
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer

    'Check if the user has a 100% chance to miss
    If HitSkill + 50 < NPCList(NPCIndex).ModStat(SID.Parry) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetCharDamage
        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
        ConBuf.Put_Integer -1
        Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'If user weapon skill is at least 50 points greater, 100% chance to hit
    If HitSkill - 50 <= NPCList(NPCIndex).ModStat(SID.Parry) Then

        'Since the user doesn't have 100% chance to hit, calculate if they hit
        If Server_RandomNumber(1, 100) >= ((HitSkill + 50) - NPCList(NPCIndex).ModStat(SID.Parry)) Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_SetCharDamage
            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
            ConBuf.Put_Integer -1
            Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
            Exit Sub
        End If

    End If

    'Update aggressive-face
    If UserList(UserIndex).Counters.AggressiveCounter <= 0 Then
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

    'Hurt the NPC
    NPC_Damage NPCIndex, UserIndex, Hit

End Sub

Sub User_AttackUser(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

'*****************************************************************
'Have a user attack a user
'*****************************************************************

Dim Hit As Integer

'Don't allow if switchingmaps maps

    If UserList(VictimIndex).Flags.SwitchingMaps Then
        Exit Sub
    End If

    'Calculate hit
    Hit = Server_RandomNumber(UserList(AttackerIndex).Stats.ModStat(SID.MinHIT), UserList(AttackerIndex).Stats.ModStat(SID.MaxHIT))
    Hit = Hit - (UserList(VictimIndex).Stats.ModStat(SID.DEF) / 2)
    If Hit < 1 Then Hit = 1

    'Hit User
    UserList(VictimIndex).Stats.ModStat(SID.MinHP) = UserList(VictimIndex).Stats.ModStat(SID.MinHP) - Hit

    'Play the attack animation
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Attack
    ConBuf.Put_Integer UserList(AttackerIndex).Char.CharIndex
    Data_Send ToPCArea, AttackerIndex, ConBuf.Get_Buffer

    'User Die
    If UserList(VictimIndex).Stats.ModStat(SID.MinHP) <= 0 Then

        'Kill user
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You kill " & UserList(VictimIndex).Name & "!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Fight
        Data_Send ToIndex, AttackerIndex, ConBuf.Get_Buffer

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String UserList(AttackerIndex).Name & " kills you!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Fight
        Data_Send ToIndex, VictimIndex, ConBuf.Get_Buffer

        User_Kill VictimIndex

    End If

End Sub

Sub User_ChangeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal Weapon As Integer, ByVal Hair As Integer)

'*****************************************************************
'Changes a user char's head,body and heading
'*****************************************************************

Dim i As Byte

'Check for invalid values

    If UserIndex > MaxUsers Then Exit Sub
    If UserIndex <= 0 Then Exit Sub
    If Body < 0 Then Exit Sub
    If Head < 0 Then Exit Sub
    If Heading < 0 Then Exit Sub
    If Weapon < 0 Then Exit Sub
    If Hair < 0 Then Exit Sub

    'Apply the values
    If UserList(UserIndex).Char.Body <> Body Then
        UserList(UserIndex).Char.Body = Body
        i = 1
    End If
    If UserList(UserIndex).Char.Head <> Head Then
        UserList(UserIndex).Char.Head = Head
        i = 1
    End If
    If UserList(UserIndex).Char.Heading <> Heading Then
        UserList(UserIndex).Char.Heading = Heading
        i = 1
    End If
    If UserList(UserIndex).Char.HeadHeading <> Heading Then
        UserList(UserIndex).Char.HeadHeading = Heading
        i = 1
    End If
    If UserList(UserIndex).Char.Weapon <> Weapon Then
        UserList(UserIndex).Char.Weapon = Weapon
        i = 1
    End If
    If UserList(UserIndex).Char.Hair <> Hair Then
        UserList(UserIndex).Char.Hair = Hair
        i = 1
    End If

    'Send the update
    If i = 1 Then   'Make sure we only send update when it is needed
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_ChangeChar
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Integer Body
        ConBuf.Put_Integer Head
        ConBuf.Put_Byte Heading
        ConBuf.Put_Integer Weapon
        ConBuf.Put_Integer Hair
        Data_Send sndRoute, sndIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If

End Sub

Sub User_ChangeInv(ByVal UserIndex As Integer, ByVal Slot As Byte, Object As UserOBJ)

'*****************************************************************
'Changes a user's inventory
'*****************************************************************

    UserList(UserIndex).Object(Slot) = Object

    If Object.ObjIndex Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_SetInventorySlot
        ConBuf.Put_Byte Slot
        ConBuf.Put_Long Object.ObjIndex
        ConBuf.Put_String ObjData(Object.ObjIndex).Name
        ConBuf.Put_Long Object.Amount
        ConBuf.Put_Byte Object.Equipped
        ConBuf.Put_Integer ObjData(Object.ObjIndex).GrhIndex
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    Else
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_SetInventorySlot
        ConBuf.Put_Byte Slot
        ConBuf.Put_Long 0
        ConBuf.Put_String "(None)"
        ConBuf.Put_Long 0
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer 0
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If

End Sub

Sub User_DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal x As Integer, ByVal Y As Integer)

'*****************************************************************
'Drops a object from a User's slot
'*****************************************************************

Dim Obj As Obj

'Check for invalid values

    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.DownloadingMap Then Exit Sub
    If Num > UserList(UserIndex).Object(Slot).Amount Then Num = UserList(UserIndex).Object(Slot).Amount
    If Num <= 0 Then Exit Sub

    'Check for object on gorund
    If MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex <> 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "No room on ground."
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    Obj.ObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
    Obj.Amount = Num
    Obj_Make ToMap, UserIndex, Obj, x, Y

    'Remove object
    UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - Num
    If UserList(UserIndex).Object(Slot).Amount <= 0 Then
        'Unequip is the object is currently equipped
        If UserList(UserIndex).Object(Slot).Equipped = 1 Then User_RemoveInvItem UserIndex, Slot

        UserList(UserIndex).Object(Slot).ObjIndex = 0
        UserList(UserIndex).Object(Slot).Amount = 0
        UserList(UserIndex).Object(Slot).Equipped = 0
    End If

    User_UpdateInv False, UserIndex, Slot

End Sub

Sub User_EraseChar(ByVal UserIndex As Integer)

'*****************************************************************
'Erase a character
'*****************************************************************
'Remove from list

    CharList(UserList(UserIndex).Char.CharIndex).Index = 0
    CharList(UserList(UserIndex).Char.CharIndex).CharType = 0

    'Remove from map
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).UserIndex = 0

    'Send erase command to clients
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_EraseChar
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map

    'Update userlist
    UserList(UserIndex).Char.CharIndex = 0

End Sub

Sub User_GetObj(ByVal UserIndex As Integer)

'*****************************************************************
'Puts a object in a User's slot from the current User's position
'*****************************************************************

Dim x As Integer
Dim Y As Integer
Dim Slot As Byte

'Check for invalud values

    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.DownloadingMap Then Exit Sub

    x = UserList(UserIndex).Pos.x
    Y = UserList(UserIndex).Pos.Y

    'Check for object on ground
    If MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "Nothing there."
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check to see if User already has object type
    Slot = 1
    Do Until UserList(UserIndex).Object(Slot).ObjIndex = MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex
        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then
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
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "Inventory full."
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        Loop
    End If

    'Fill object slot
    If UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount <= MAX_INVENTORY_OBJS Then

        'Tell the user they recieved the items
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You pick up " & MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount & " " & ObjData(MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex).Name
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'User takes all the items
        UserList(UserIndex).Object(Slot).ObjIndex = MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount
        Obj_Erase ToMap, UserIndex, MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y

    Else
        'Over MAX_INV_OBJS
        If MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount < UserList(UserIndex).Object(Slot).Amount Then
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You pick up " & Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount)) & " " & ObjData(MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex).Name & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount = Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount))
        Else
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You pick up " & Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount) & " " & ObjData(MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex).Name & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount = Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount)
        End If
        UserList(UserIndex).Object(Slot).Amount = MAX_INVENTORY_OBJS
    End If

    'Update the user's inventory
    Call User_UpdateInv(False, UserIndex, Slot)

End Sub

Sub User_GiveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer)

'*****************************************************************
'Give the user an object
'*****************************************************************

Dim Slot As Byte

'Check for invalid values

    If UserIndex <= 0 Then Exit Sub
    If ObjIndex <= 0 Then Exit Sub
    If UserIndex > LastUser Then Exit Sub
    If ObjIndex > NumObjDatas Then Exit Sub

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
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "Inventory full."
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        Loop
    End If

    'Fill object slot
    If UserList(UserIndex).Object(Slot).Amount + Amount <= MAX_INVENTORY_OBJS Then

        'Tell the user they recieved the items
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You pick up " & Amount & " " & ObjData(ObjIndex).Name
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        'User takes all the items
        UserList(UserIndex).Object(Slot).ObjIndex = ObjIndex
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount + Amount
        Obj_Erase ToMap, UserIndex, Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y

    Else
        'Over MAX_INV_OBJS
        If Amount < UserList(UserIndex).Object(Slot).Amount Then
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You get " & Abs(MAX_INVENTORY_OBJS - (UserList(UserIndex).Object(Slot).Amount + Amount)) & " " & ObjData(ObjIndex).Name & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Else
            'Tell the user they recieved the items
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You get " & Abs((MAX_INVENTORY_OBJS + UserList(UserIndex).Object(Slot).Amount) - Amount) & " " & ObjData(ObjIndex).Name & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
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

    Do
        i = i + 1
        If i > NumUsers + 10 Then
            User_IndexFromSox = -1
            Exit Function
        End If
    Loop While UserList(i).ConnID <> inSox
    User_IndexFromSox = i

End Function

Sub User_Kill(ByVal UserIndex As Integer)

'*****************************************************************
'Kill a user
'*****************************************************************

Dim TempPos As WorldPos

'Set user health back to full

    UserList(UserIndex).Stats.ModStat(SID.MinHP) = UserList(UserIndex).Stats.ModStat(SID.MaxHP)

    'Find a place to put user
    Call Server_ClosestLegalPos(ResPos, TempPos)
    If Server_LegalPos(TempPos.Map, TempPos.x, TempPos.Y, 0) = False Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_UMsgbox
        ConBuf.Put_String "No legal position found. Please try again."
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        User_Close UserIndex
        Exit Sub
    End If

    'Warp him there
    User_WarpChar UserIndex, TempPos.Map, TempPos.x, TempPos.Y

End Sub

Sub User_LookAtTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal Button As Byte)

'*****************************************************************
'Responds to the user clicking on a square
'*****************************************************************

Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim LoopC As Byte
Dim TempCharIndex As Integer
Dim MsgData As MailData

'Check for invalid values

    If Not Server_InMapBounds(x, Y) Then Exit Sub
    If UserIndex <= 0 Then Exit Sub
    If UserIndex > MaxUsers Then Exit Sub
    If Map <= 0 Then Exit Sub
    If Map > NumMaps Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.DownloadingMap Then Exit Sub

    '***** Right Click *****
    If Button = vbRightButton Then

        '*** Check for mailbox ***
        If MapData(Map, x, Y).Mailbox = 1 Then

            'Only check mail if right next to the mailbox
            If UserList(UserIndex).Pos.Map = Map Then
                If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, CInt(x), CInt(Y)) < 2 Then

                    'Store the position of the mailbox for later reference in case user tries to use items away from mailbox
                    UserList(UserIndex).MailboxPos.Map = Map
                    UserList(UserIndex).MailboxPos.x = x
                    UserList(UserIndex).MailboxPos.Y = Y

                    'Resend all the mail
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_MailBox
                    For LoopC = 1 To MaxMailPerUser
                        If UserList(UserIndex).MailID(LoopC) > 0 Then
                            Load_Mail UserList(UserIndex).MailID(LoopC), MsgData
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
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You must be next to a mailbox to read your messages."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub

        End If

        '*** Check for Characters ***
        If Y + 1 <= YMaxMapSize Then
            If MapData(Map, x, Y + 1).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y + 1).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, x, Y + 1).NPCIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y + 1).NPCIndex
                FoundChar = 2
            End If
        End If
        'Check for Character
        If FoundChar = 0 Then
            If MapData(Map, x, Y).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, x, Y).NPCIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y).NPCIndex
                FoundChar = 2
            End If
        End If
        'React to character
        If FoundChar = 1 Then
            If Len(UserList(TempCharIndex).Desc) > 1 Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "You see " & UserList(TempCharIndex).Name & ". " & UserList(TempCharIndex).Desc
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "You see " & UserList(TempCharIndex).Name & "."
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            FoundSomething = 1
        End If
        If FoundChar = 2 Then
            FoundSomething = 1
            '*** Check for NPC vendor ***
            If NPCList(TempCharIndex).NumVendItems > 0 Then
                User_TradeWithNPC UserIndex, TempCharIndex
                FoundSomething = 1
            End If
            '*** NPC not a vendor, give description ***
            If Len(NPCList(TempCharIndex).Name) > 1 Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "You see " & NPCList(TempCharIndex).Name & ". " & NPCList(TempCharIndex).Desc
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Else
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "You see " & NPCList(TempCharIndex).Name & "."
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            '*** Quest NPC ***
            If NPCList(TempCharIndex).Quest > 0 Then Quest_General UserIndex, TempCharIndex
        End If

        '*** Check for object ***
        If MapData(Map, x, Y).ObjInfo.ObjIndex > 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You see a " & ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).Name
            ConBuf.Put_Byte DataCode.Comm_FontType_Talk
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            FoundSomething = 1
        End If

        '*** Didn't find anything ***
        If FoundSomething = 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You see nothing of interest."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        End If

        '***** Left Click *****
    ElseIf Button = vbLeftButton Then

        '*** Look for NPC/Player to target ***
        If Y + 1 <= YMaxMapSize Then
            If MapData(Map, x, Y + 1).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y + 1).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, x, Y + 1).NPCIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y + 1).NPCIndex
                FoundChar = 2
            End If
        End If
        If FoundChar = 0 Then
            If MapData(Map, x, Y).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y).UserIndex
                FoundChar = 1
            End If
            If MapData(Map, x, Y).NPCIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y).NPCIndex
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
            If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, UserList(TempCharIndex).Pos.x, UserList(TempCharIndex).Pos.Y) <= Max_Server_Distance Then
                UserList(UserIndex).Flags.Target = 1
                UserList(UserIndex).Flags.TargetIndex = TempCharIndex
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer UserList(TempCharIndex).Char.CharIndex
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
            If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, NPCList(TempCharIndex).Pos.x, NPCList(TempCharIndex).Pos.Y) <= Max_Server_Distance Then
                UserList(UserIndex).Flags.Target = 2
                UserList(UserIndex).Flags.TargetIndex = TempCharIndex
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer NPCList(TempCharIndex).Char.CharIndex
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

End Sub

Sub User_MakeChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)

'*****************************************************************
'Makes and places a user's character
'*****************************************************************

Dim CharIndex As Integer

'Place character on map

    MapData(Map, x, Y).UserIndex = UserIndex

    'Give it a char if needed
    If UserList(UserIndex).Char.CharIndex = 0 Then
        CharIndex = Server_NextOpenCharIndex
        UserList(UserIndex).Char.CharIndex = CharIndex
        CharList(CharIndex).Index = UserIndex
        CharList(CharIndex).CharType = CharType_NPC
    End If

    'Send make character command to clients
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_MakeChar
    ConBuf.Put_Integer UserList(UserIndex).Char.Body
    ConBuf.Put_Integer UserList(UserIndex).Char.Head
    ConBuf.Put_Byte UserList(UserIndex).Char.Heading
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte x
    ConBuf.Put_Byte Y
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_Integer UserList(UserIndex).Char.Weapon
    ConBuf.Put_Integer UserList(UserIndex).Char.Hair
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

Sub User_MoveChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

'*****************************************************************
'Moves a User from one tile to another
'*****************************************************************

Dim nPos As WorldPos

'Check for invalid values

    If UserIndex <= 0 Then Exit Sub
    If UserIndex > MaxUsers Then Exit Sub
    If nHeading <= 0 Then Exit Sub
    If nHeading > 8 Then Exit Sub

    'Update the move counter
    If timeGetTime < UserList(UserIndex).Counters.MoveCounter + 50 Then

        'If the user is moving too fast, then put them back
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte CByte(UserList(UserIndex).Pos.x)
        ConBuf.Put_Byte CByte(UserList(UserIndex).Pos.Y)
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub

    End If

    'Clear the pending quest NPC number
    UserList(UserIndex).Flags.QuestNPC = 0

    'Get the new time
    UserList(UserIndex).Counters.MoveCounter = timeGetTime

    'Get the new position
    nPos = UserList(UserIndex).Pos
    Server_HeadToPos nHeading, nPos

    'Move if legal pos
    If Server_LegalPos(UserList(UserIndex).Pos.Map, nPos.x, nPos.Y, nHeading) = True Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MoveChar
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte nPos.x
        ConBuf.Put_Byte nPos.Y
        Data_Send ToMapButIndex, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map

        'Update map and user pos
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.Heading = nHeading
        UserList(UserIndex).Char.HeadHeading = nHeading
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).UserIndex = UserIndex

        'Do tile events
        Server_DoTileEvents UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y

    Else

        'Else correct user's pos
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte CByte(UserList(UserIndex).Pos.x)
        ConBuf.Put_Byte CByte(UserList(UserIndex).Pos.Y)
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    End If

End Sub

Function User_NameToIndex(ByVal Name As String) As Integer

'*****************************************************************
'Searches userlist for a name and return userindex
'*****************************************************************

Dim UserIndex As Integer

'check for bad name

    If LenB(Name) = 0 Then
        User_NameToIndex = 0
        Exit Function
    End If

    UserIndex = 1
    Do Until UCase$(Left$(UserList(UserIndex).Name, Len(Name))) = UCase$(Name)
        UserIndex = UserIndex + 1

        If UserIndex > LastUser Then
            UserIndex = 0
            Exit Do
        End If
    Loop

    User_NameToIndex = UserIndex

End Function

Function User_NextOpen() As Integer

'*****************************************************************
'Finds the next open UserIndex in UserList
'*****************************************************************

Dim LoopC As Long

    LoopC = 1

    Do Until UserList(LoopC).Flags.UserLogged = 0
        LoopC = LoopC + 1
        If LoopC > MaxUsers Then Exit Do
    Loop

    User_NextOpen = LoopC

End Function

Public Sub User_RaiseExp(ByVal UserIndex As Integer, ByVal EXP As Long)

'*****************************************************************
'Raise the user's experience - this should be the only way exp is raised!
'*****************************************************************

Dim Levels As Integer

    On Error GoTo ErrOut

    'Raise the user's points
    UserList(UserIndex).Stats.BaseStat(SID.Points) = UserList(UserIndex).Stats.BaseStat(SID.Points) + EXP

    'Update the user's experience
    UserList(UserIndex).Stats.BaseStat(SID.EXP) = UserList(UserIndex).Stats.BaseStat(SID.EXP) + EXP

    'Loop as many times as needed to get every level gained in
    Do While UserList(UserIndex).Stats.BaseStat(SID.EXP) >= UserList(UserIndex).Stats.BaseStat(SID.ELU)

        'Once the user reaches level 1000, the exp to level is the same
        If UserList(UserIndex).Stats.BaseStat(SID.ELU) >= 1000 Then

            'Reset the level requirements
            UserList(UserIndex).Stats.BaseStat(SID.ELU) = UserList(UserIndex).Stats.BaseStat(SID.ELU) + 1
            UserList(UserIndex).Stats.BaseStat(SID.EXP) = UserList(UserIndex).Stats.BaseStat(SID.EXP) - UserList(UserIndex).Stats.BaseStat(SID.ELU)
            UserList(UserIndex).Stats.BaseStat(SID.ELU) = LARGESTLONG

            'Update client-side
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You gained a level!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub

        End If

        'Set the number of levels gained
        Levels = Levels + 1

        'Raise stats
        With UserList(UserIndex).Stats
            .BaseStat(SID.Agil) = .BaseStat(SID.Agil) + 1
            .BaseStat(SID.Clairovoyance) = .BaseStat(SID.Clairovoyance) + 1
            .BaseStat(SID.Dagger) = .BaseStat(SID.Dagger) + 1
            .BaseStat(SID.DefensiveMag) = .BaseStat(SID.DefensiveMag) + 1
            .BaseStat(SID.Fist) = .BaseStat(SID.Fist) + 1
            .BaseStat(SID.Immunity) = .BaseStat(SID.Immunity) + 1
            .BaseStat(SID.Mag) = .BaseStat(SID.Mag) + 1
            .BaseStat(SID.MaxHIT) = .BaseStat(SID.MaxHIT) + 1
            .BaseStat(SID.MaxHP) = .BaseStat(SID.MaxHP) + 10
            .BaseStat(SID.MaxMAN) = .BaseStat(SID.MaxMAN) + 10
            .BaseStat(SID.MaxSTA) = .BaseStat(SID.MaxSTA) + 10
            .BaseStat(SID.Meditate) = .BaseStat(SID.Meditate) + 1
            .BaseStat(SID.MinHIT) = .BaseStat(SID.MinHIT) + 1
            .BaseStat(SID.OffensiveMag) = .BaseStat(SID.OffensiveMag) + 1
            .BaseStat(SID.Parry) = .BaseStat(SID.Parry) + 1
            .BaseStat(SID.Regen) = .BaseStat(SID.Regen) + 1
            .BaseStat(SID.Rest) = .BaseStat(SID.Rest) + 1
            .BaseStat(SID.Staff) = .BaseStat(SID.Staff) + 1
            .BaseStat(SID.Str) = .BaseStat(SID.Str) + 1
            .BaseStat(SID.SummoningMag) = .BaseStat(SID.SummoningMag) + 1
            .BaseStat(SID.Sword) = .BaseStat(SID.Sword) + 1
        End With
        
        'Reset the level requirements
        UserList(UserIndex).Stats.BaseStat(SID.ELU) = UserList(UserIndex).Stats.BaseStat(SID.ELU) + 1
        UserList(UserIndex).Stats.BaseStat(SID.EXP) = UserList(UserIndex).Stats.BaseStat(SID.EXP) - UserList(UserIndex).Stats.BaseStat(SID.ELU)
        UserList(UserIndex).Stats.BaseStat(SID.ELU) = Int(8 * (1.2 ^ (UserList(UserIndex).Stats.BaseStat(SID.ELU) - 1)))

    Loop

    'Check if needing to update from leveling
    If Levels = 1 Then

        'Say the user's level raised
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You gained a level!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    ElseIf Levels > 1 Then

        'Say the user's level raised
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You gained " & Levels & " levels!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    End If

Exit Sub

ErrOut:

    'Tell user they have too many points
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "You have too many points! Points must be used before you gain more experience!"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

Exit Sub

End Sub

Sub User_RemoveInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Unequip a inventory item
'*****************************************************************

Dim Obj As ObjData

'Set the object

    Obj = ObjData(UserList(UserIndex).Object(Slot).ObjIndex)

    'Get the object type
    Select Case Obj.ObjType

        'Check for weapon
    Case OBJTYPE_WEAPON

        'Set the equipted variables
        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).WeaponEqpObjIndex = 0
        UserList(UserIndex).WeaponEqpSlot = 0
        UserList(UserIndex).Char.Weapon = 0
        UserList(UserIndex).WeaponType = Hand
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair

        'Check for armor
    Case OBJTYPE_ARMOR

        'Set the equipted variables
        UserList(UserIndex).Object(Slot).Equipped = 0
        UserList(UserIndex).ArmorEqpObjIndex = 0
        UserList(UserIndex).ArmorEqpSlot = 0
        UserList(UserIndex).Char.Body = 1
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair

    End Select

    'Update the user's stats
    User_UpdateInv False, UserIndex, Slot

End Sub

Public Sub User_SendKnownSkills(ByVal UserIndex As Integer)

Dim KnowSkillList As Long
Dim i As Byte

'Check for a valid userindex

    If UserIndex <= 0 Then Exit Sub
    If UserIndex > MaxUsers Then Exit Sub

    'Compile all the known skills into a long
    'Once you start getting overflows, you will have to program in another long
    For i = 1 To NumSkills
        If UserList(UserIndex).KnownSkills(i) Then KnowSkillList = KnowSkillList Or (1 * (2 ^ (i - 1)))
    Next i

    'Send the long to the user
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_KnownSkills
    ConBuf.Put_Long KnowSkillList
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Public Sub User_TradeWithNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

'*****************************************************************
'Start a trade with a NPC
'*****************************************************************

Dim LoopC As Integer

'Trade with a NPC

    If NPCList(NPCIndex).NumVendItems > 0 Then

        'Check if close enough to trade with
        If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.x, NPCList(NPCIndex).Pos.Y) > 5 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You are too far away to trade."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Trade_StartNPCTrade
        ConBuf.Put_String NPCList(NPCIndex).Name
        ConBuf.Put_Integer NPCList(NPCIndex).NumVendItems
        For LoopC = 1 To NPCList(NPCIndex).NumVendItems
            ConBuf.Put_Integer ObjData(NPCList(NPCIndex).VendItems(LoopC).ObjIndex).GrhIndex
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
Dim x As Long
Dim Y As Long

    Map = UserList(UserIndex).Pos.Map

    'Send user char's pos
    For x = 1 To UBound(ConnectionGroups(Map).UserIndex())
        Call User_MakeChar(ToIndex, UserIndex, ConnectionGroups(Map).UserIndex(x), Map, UserList(ConnectionGroups(Map).UserIndex(x)).Pos.x, UserList(ConnectionGroups(Map).UserIndex(x)).Pos.Y)
    Next x

    'Place chars and objects
    For Y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            If MapData(Map, x, Y).NPCIndex Then NPC_MakeChar ToIndex, UserIndex, MapData(Map, x, Y).NPCIndex, Map, x, Y
            If MapData(Map, x, Y).ObjInfo.ObjIndex Then Obj_Make ToIndex, UserIndex, MapData(Map, x, Y).ObjInfo, x, Y
        Next x
    Next Y

End Sub

Public Sub User_UpdateModStats(ByVal UserIndex As Integer)

'*****************************************************************
'Set the user's mod stats based on base stats and equipted items
'*****************************************************************

Dim WeaponObj As ObjData
Dim ArmorObj As ObjData
Dim TempVal As Single
Dim i As Integer

    If UserList(UserIndex).Flags.UserLogged = 0 Then Exit Sub

    'Set the equipted items
    If UserList(UserIndex).WeaponEqpObjIndex > 0 Then WeaponObj = ObjData(UserList(UserIndex).WeaponEqpObjIndex)
    If UserList(UserIndex).ArmorEqpObjIndex > 0 Then ArmorObj = ObjData(UserList(UserIndex).ArmorEqpObjIndex)

    With UserList(UserIndex).Stats

        'Equipted items
        For i = 1 To NumStats
            .ModStat(i) = .BaseStat(i) + WeaponObj.AddStat(i)
        Next i

    End With
    
End Sub

Sub User_UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

'*****************************************************************
'Use/Equip a inventory item
'*****************************************************************

Dim Obj As ObjData

'Check for invalid values

    If UserIndex > MaxUsers Then Exit Sub
    If UserIndex <= 0 Then Exit Sub
    If Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    If Slot <= 0 Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.DownloadingMap Then Exit Sub

    Obj = ObjData(UserList(UserIndex).Object(Slot).ObjIndex)

    'Apply the replenish values
    With UserList(UserIndex).Stats
        .ModStat(SID.MinHP) = .ModStat(SID.MinHP) + (.ModStat(SID.MaxHP) * Obj.RepHPP) + Obj.RepHP
        .ModStat(SID.MinMAN) = .ModStat(SID.MinMAN) + (.ModStat(SID.MaxMAN) * Obj.RepMPP) + Obj.RepMP
        .ModStat(SID.MinSTA) = .ModStat(SID.MinSTA) + (.ModStat(SID.MaxSTA) * Obj.RepSPP) + Obj.RepSP
    End With

    Select Case Obj.ObjType
    Case OBJTYPE_USEONCE

        'Remove from inventory
        UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - 1
        If UserList(UserIndex).Object(Slot).Amount <= 0 Then UserList(UserIndex).Object(Slot).ObjIndex = 0

    Case OBJTYPE_WEAPON

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

        'Set the paper-doll
        If Obj.SpriteHair <> -1 Then UserList(UserIndex).Char.Hair = Obj.SpriteHair
        If Obj.SpriteBody <> -1 Then UserList(UserIndex).Char.Body = Obj.SpriteBody
        If Obj.SpriteHead <> -1 Then UserList(UserIndex).Char.Head = Obj.SpriteHead
        If Obj.SpriteWeapon <> -1 Then UserList(UserIndex).Char.Weapon = Obj.SpriteWeapon
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair

    Case OBJTYPE_ARMOR

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
        User_ChangeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.Weapon, UserList(UserIndex).Char.Hair

    End Select

    'Update user's stats and inventory
    User_UpdateInv False, UserIndex, Slot

End Sub

Sub User_WarpChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal ForceSwitch As Boolean = False)

'*****************************************************************
'Warps user to another spot
'*****************************************************************

Dim OldMap As Integer
Dim LoopC As Long

    OldMap = UserList(UserIndex).Pos.Map

    User_EraseChar UserIndex

    UserList(UserIndex).Pos.x = x
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.Map = Map

    If (OldMap <> Map) Or ForceSwitch = True Then
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
        'Check if it's the first user on the map
        If MapInfo(Map).NumUsers = 1 Then
            ReDim ConnectionGroups(Map).UserIndex(1 To 1)
        Else
            ReDim Preserve ConnectionGroups(Map).UserIndex(1 To MapInfo(Map).NumUsers)
        End If
        ConnectionGroups(Map).UserIndex(MapInfo(Map).NumUsers) = UserIndex

        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers Then
            'Find current pos within connection group
            For LoopC = 1 To MapInfo(OldMap).NumUsers + 1
                If ConnectionGroups(OldMap).UserIndex(LoopC) = UserIndex Then Exit For
            Next LoopC
            'Move the rest of the list backwards
            For LoopC = LoopC To MapInfo(OldMap).NumUsers
                ConnectionGroups(OldMap).UserIndex(LoopC) = ConnectionGroups(OldMap).UserIndex(LoopC + 1)
            Next LoopC
            'Resize the list
            ReDim Preserve ConnectionGroups(OldMap).UserIndex(1 To MapInfo(OldMap).NumUsers)
        Else
            ReDim ConnectionGroups(OldMap).UserIndex(0)
        End If

        'Show Character to others
        User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y
    Else
        User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_UserCharIndex
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:48)  Decl: 1  Code: 2408  Total: 2409 Lines
':) CommentOnly: 361 (15%)  Commented: 5 (0.2%)  Empty: 499 (20.7%)  Max Logic Depth: 7
