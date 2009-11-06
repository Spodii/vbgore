Attribute VB_Name = "TCP"
Option Explicit

'Packet priorities constants (influncing when packets are sent)
Public Const PP_None As Byte = 0        'Does not force sending
Public Const PP_Low As Byte = 1         'Enables a count-down timer
Public Const PP_High As Byte = 2        'Packet sent as soon as possible
Private Const PP_Time As Long = 1000    'How much time must pass for a PP_Low packet to be sent

'Constants used in the Data_Send sub
Public Const ToIndex As Byte = 0            'Send data to a single User index
Public Const ToAll As Byte = 1              'Send it to all User index
Public Const ToMap As Byte = 2              'Send it to all users in a map
Public Const ToPCArea As Byte = 3           'Send to all users in a user's area
Public Const ToMapButIndex As Byte = 4      'Send to all users on the map but the userindex
Public Const ToNPCArea As Byte = 5          'Send to all users in a NPC's area
Public Const ToUserMove As Byte = 6         'Send one priority to local map users, another priority to distance map users
Public Const ToNPCMove As Byte = 7          'Same as ToUserMove, but if a NPC moves

'*************************
'*** PACKET PRIORITIES ***
'*************************

'*** Packet_WaitTime ***
'How many miliseconds must go by for a packet to be sent (even if PP_High, the time must be waited)
' Keep in mind the highest frame rate is 60 (~17ms), so it is not recommended to go below that value
' or else you are just wasting bandwidth.
'0 is used in the demo so people don't get the wrong idea on the ping times.
'Recommended about 200~300 at highest, about 50 at lowest

'Packet priorities (placed here for easy changing - for advanced users -
' recommended to go packet-by-packet and decide what you want)
'If you have more bandwidth then you could possibly use, then just set
' all the packet priorities to HIGH. Packet priorities is aimed for those
' who want and need an extra booster from their bandwidth.
'vbGORE, by default, aims for the lowest packet priority values that will
' affect the gameplay the least (most stuff not listed as HIGH you wont notice the lag time)

'*** Option 1 - Optimized for minimal bandwidth usage ***
'Public Const Packet_WaitTime As Long = 100      'Recommended about 200~300 at highest, about 50 at lowest

'Communication
'Public Const PP_GlobalChat As Byte = PP_Low     'Talking to everyone online (shout)
'Public Const PP_LocalChat As Byte = PP_Low      'Talking to people locally (normal chat, emoting)
'Public Const PP_PrivateChat As Byte = PP_Low    'Talking to someone privately (private messages, NOT mailing)
'Public Const PP_GMMessages As Byte = PP_Low     'Messages sent back from using a GM command (only GMs can see)
'Public Const PP_GlobalMessage As Byte = PP_None 'Messages sent by the server globally not specified above (kicking a user, server messages, etc)

'Misc
'Public Const PP_ChangeChar As Byte = PP_None    'Updating character paperdoll values
'Public Const PP_NewMail As Byte = PP_None       '"You just got new mail" message
'Public Const PP_Connect As Byte = PP_Low        '"User xxx connected" message
'Public Const PP_Stat As Byte = PP_Low           'Base/mod stat updating
'Public Const PP_StatPercent As Byte = PP_Low    'Stat percentage (MP and SP bars) updating
'Public Const PP_Blink As Byte = PP_None         'Sending character eye blinks
'Public Const PP_Look As Byte = PP_Low           'Change looking direction (lookleft, lookright)
'Public Const PP_Rotate As Byte = PP_Low         'Change body direction (rotating)
'Public Const PP_CloseCharMove As Byte = PP_High 'Character movement of near-by characters (in screen)
'Public Const PP_FarCharMove As Byte = PP_None   'Character movement of characters out of the screen

'*** Option 2 - Optimized for speed and performance ***
Public Const Packet_WaitTime As Long = 0        'Recommended about 200~300 at highest, about 50 at lowest

'Communication
Public Const PP_GlobalChat As Byte = PP_High    'Talking to everyone online (shout)
Public Const PP_LocalChat As Byte = PP_High     'Talking to people locally (normal chat, emoting)
Public Const PP_PrivateChat As Byte = PP_High   'Talking to someone privately (private messages, NOT mailing)
Public Const PP_GMMessages As Byte = PP_High    'Messages sent back from using a GM command (only GMs can see)
Public Const PP_GlobalMessage As Byte = PP_High 'Messages sent by the server globally not specified above (kicking a user, server messages, etc)

'Misc
Public Const PP_ChangeChar As Byte = PP_High    'Updating character paperdoll values
Public Const PP_NewMail As Byte = PP_High       '"You just got new mail" message
Public Const PP_Connect As Byte = PP_High       '"User xxx connected" message
Public Const PP_Stat As Byte = PP_High          'Base/mod stat updating
Public Const PP_StatPercent As Byte = PP_High   'Stat percentage (MP and SP bars) updating
Public Const PP_Blink As Byte = PP_High         'Sending character eye blinks
Public Const PP_Look As Byte = PP_High          'Change looking direction (lookleft, lookright)
Public Const PP_Rotate As Byte = PP_High        'Change body direction (rotating)
Public Const PP_CloseCharMove As Byte = PP_High 'Character movement of near-by characters (in screen)
Public Const PP_FarCharMove As Byte = PP_None   'Character movement of characters out of the screen

Sub Data_User_Bank_PutItem(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Client requests to put an item in bank
'<Slot(B)><Amount(I)>
'*****************************************************************
Dim Amount As Integer
Dim Slot As Byte
Dim PutSlot As Byte

    'Get the values
    Slot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer
    
    'Check for invalid values
    If Slot <= 0 Then Exit Sub
    If Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    If UserList(UserIndex).Object(Slot).ObjIndex = 0 Then Exit Sub
    If Amount > UserList(UserIndex).Object(Slot).Amount Then Amount = UserList(UserIndex).Object(Slot).Amount
    If Amount <= 0 Then Exit Sub
    
    'Check for a valid distance from the banker NPC
    If UserList(UserIndex).flags.TradeWithNPC > 0 Then
        With NPCList(UserList(UserIndex).flags.TradeWithNPC)
            If NPCList(UserList(UserIndex).flags.TradeWithNPC).AI <> 6 Then Exit Sub    'Not a banker
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Pos.X, .Pos.Y, MaxServerDistanceX, MaxServerDistanceY) = 0 Then Exit Sub  'Out of range
        End With
    Else
        Exit Sub    'No clicked NPC
    End If
    
    'Check for item of the same type already in there
    Do
        PutSlot = PutSlot + 1
        If PutSlot > MAX_INVENTORY_SLOTS Then
            PutSlot = 0
            Exit Do
        End If
    Loop While UserList(UserIndex).Bank(PutSlot).ObjIndex <> UserList(UserIndex).Object(Slot).ObjIndex
 
    'If PutSlot = 0, no duplicate item was found, so use the next free slot
    If PutSlot = 0 Then
        Do
            PutSlot = PutSlot + 1
            If PutSlot > MAX_INVENTORY_SLOTS Then
                
                'Bank is full
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "Your bank depot is full!"
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
                
            End If
        Loop While UserList(UserIndex).Bank(PutSlot).ObjIndex > 0
        
        'Just as a pre-caution, we empty the amount value since we are going to be adding on to it and it should be empty
        UserList(UserIndex).Bank(PutSlot).Amount = 0
        
    End If
    
    'Check if theres room for the item
    If UserList(UserIndex).Bank(PutSlot).Amount + Amount > MAX_INVENTORY_OBJS Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "Your bank depot is full!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    
    'Put the items
    UserList(UserIndex).Bank(PutSlot).ObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
    UserList(UserIndex).Bank(PutSlot).Amount = UserList(UserIndex).Bank(PutSlot).Amount + Amount
    
    'Remove the items from the user
    UserList(UserIndex).Object(Slot).Amount = UserList(UserIndex).Object(Slot).Amount - Amount

    'Check if the user ran out of items
    If UserList(UserIndex).Object(Slot).Amount <= 0 Then
        
        'User depotted all the items, so remove it from the inventory
        If UserList(UserIndex).Object(Slot).Equipped Then User_RemoveInvItem UserIndex, Slot, 0
        
        'Remove the item
        UserList(UserIndex).Object(Slot).ObjIndex = 0
    
    End If
    
    'Update the inventory
    User_UpdateInv False, UserIndex, Slot
    
    'Update the bank slot
    User_UpdateBank UserIndex, PutSlot

End Sub

Sub Data_User_Bank_TakeItem(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Client requests to take an item from their bank
'<Slot(B)><Amount(I)>
'*****************************************************************
Dim Amount As Integer
Dim Slot As Byte
Dim PutSlot As Byte

    'Get the values
    Slot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer
    
    'Check for invalid values
    If Slot <= 0 Then Exit Sub
    If Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    If UserList(UserIndex).Bank(Slot).ObjIndex = 0 Then Exit Sub
    If Amount > UserList(UserIndex).Bank(Slot).Amount Then Amount = UserList(UserIndex).Bank(Slot).Amount
    If Amount <= 0 Then Exit Sub
    
    'Check for a valid distance from the banker NPC
    If UserList(UserIndex).flags.TradeWithNPC > 0 Then
        With NPCList(UserList(UserIndex).flags.TradeWithNPC)
            If NPCList(UserList(UserIndex).flags.TradeWithNPC).AI <> 6 Then Exit Sub    'Not a banker
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Pos.X, .Pos.Y, MaxServerDistanceX, MaxServerDistanceY) = 0 Then Exit Sub  'Out of range
        End With
    Else
        Exit Sub    'No clicked NPC
    End If
    
    'Check for item of the same type already in there
    Do
        PutSlot = PutSlot + 1
        If PutSlot > MAX_INVENTORY_SLOTS Then
            PutSlot = 0
            Exit Do
        End If
    Loop While UserList(UserIndex).Object(PutSlot).ObjIndex <> UserList(UserIndex).Bank(Slot).ObjIndex
 
    'If PutSlot = 0, no duplicate item was found, so use the next free slot
    If PutSlot = 0 Then
        Do
            PutSlot = PutSlot + 1
            If PutSlot > MAX_INVENTORY_SLOTS Then
                
                'Bank is full
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 26
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
                
            End If
        Loop While UserList(UserIndex).Object(PutSlot).ObjIndex > 0
        
        'Just as a pre-caution, we empty the amount value since we are going to be adding on to it and it should be empty
        UserList(UserIndex).Object(PutSlot).Amount = 0
        
    End If
    
    'Check if theres room for the item
    If UserList(UserIndex).Object(PutSlot).Amount + Amount > MAX_INVENTORY_OBJS Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 26
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    
    'Put the items in the inventory
    UserList(UserIndex).Object(PutSlot).ObjIndex = UserList(UserIndex).Bank(Slot).ObjIndex
    UserList(UserIndex).Object(PutSlot).Amount = UserList(UserIndex).Object(PutSlot).Amount + Amount
    
    'Remove the items from the bank
    UserList(UserIndex).Bank(Slot).Amount = UserList(UserIndex).Bank(Slot).Amount - Amount

    'Check if the user ran out of items
    If UserList(UserIndex).Bank(Slot).Amount <= 0 Then UserList(UserIndex).Bank(Slot).ObjIndex = 0
    
    'Update the inventory
    User_UpdateInv False, UserIndex, PutSlot
    
    'Update the bank slot
    User_UpdateBank UserIndex, Slot

End Sub

Sub Data_User_RequestUserCharIndex(ByVal UserIndex As Integer)

'*****************************************************************
'Client requests their user char index
'<>
'*****************************************************************

    'Send the user char index
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_UserCharIndex
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_RequestMakeChar(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Client requests for the information on a character index
'<CharIndex(I)>
'*****************************************************************
Dim CharIndex As Integer
Dim Index As Integer

    CharIndex = rBuf.Get_Integer
    If CharIndex <= LastChar Then
        If CharIndex > 0 Then
        
            'Sort by the char type
            Index = CharList(CharIndex).Index
            Select Case CharList(CharIndex).CharType
                Case CharType_PC
                    User_MakeChar ToIndex, UserIndex, Index, UserList(Index).Pos.Map, UserList(Index).Pos.X, UserList(Index).Pos.Y
                Case CharType_NPC
                    NPC_MakeChar ToIndex, UserIndex, Index, NPCList(Index).Pos.Map, NPCList(Index).Pos.X, NPCList(Index).Pos.Y
            End Select
            
        End If
    End If
            
End Sub

Sub Data_Comm_Emote(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Emote
'<txt(S-EX)>
'*****************************************************************

Dim txt As String

    Log "Call Data_Comm_Emote([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    txt = rBuf.Get_String

    'Check for invalid conditions
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then
        Log "Data_Comm_Emote: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > CurrentTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = CurrentTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    If Server_ValidString(txt) = False Then Exit Sub
    
    'Write the message
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).name & " " & txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk Or DataCode.Comm_UseBubble
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_LocalChat

End Sub

Sub Data_Comm_Shout(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Shout
'<txt(S-EX)>
'*****************************************************************

Dim txt As String

    Log "Call Data_Comm_Shout([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    txt = rBuf.Get_String

    'Check for invalid conditions
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then
        Log "Data_Comm_Shout: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > CurrentTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = CurrentTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    If Not Server_ValidString(txt) Then Exit Sub
    
    'Write the message
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 76
    ConBuf.Put_String UserList(UserIndex).name
    ConBuf.Put_String txt
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToAll, UserIndex, ConBuf.Get_Buffer, , PP_GlobalChat

End Sub

Sub Data_Comm_Talk(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Talk
'<txt(S-EX)>
'*****************************************************************

Dim txt As String

    Log "Call Data_Comm_Talk([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    txt = rBuf.Get_String

    'Check for invalid conditions
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then
        Log "Data_Comm_Talk: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > CurrentTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = CurrentTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    If Not Server_ValidString(txt) Then Exit Sub
    
    'Write the message
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).name & ": " & txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk Or DataCode.Comm_UseBubble
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_LocalChat

End Sub

Sub Data_Comm_Whisper(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Whisper
'<txt(S)>
'*****************************************************************

Dim txt As String
Dim tName As String
Dim tIndex As Integer
Dim tMessage As String
Dim TempStr() As String

    Log "Call Data_Comm_Whisper([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    txt = rBuf.Get_String
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > CurrentTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = CurrentTime + DelayTimeTalk
    End If

    'Split the name from the rest of the string
    TempStr = Split(txt, " ")
    
    'This will make sure we actually have a string (dont crash if there is no space)
    On Error GoTo ErrOut
    tName = TempStr(0)
    On Error GoTo 0 'Removes the previous On Error

    'Check for valid tName
    If Not Server_LegalString(tName) Then Exit Sub

    'Retrieve the index of the player being whispered to
    tIndex = User_NameToIndex(tName)

    'Check for invalid conditions
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then
        Log "Data_Comm_Whisper: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'User is not online
    If tIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    If Len(txt) <> Len(tName) Then
        tMessage = Right$(txt, Len(txt) - (1 + Len(tName)))
    Else
        tMessage = " "
    End If
    
    'Check if the string is valid
    If Not Server_ValidString(txt) Then Exit Sub

    'Tell the target what they got whispered
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 52
    ConBuf.Put_String UserList(UserIndex).name
    ConBuf.Put_String tMessage
    Data_Send ToIndex, tIndex, ConBuf.Get_Buffer, , PP_PrivateChat

    'Tell the whisperer what they whispered
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 53
    ConBuf.Put_String UserList(tIndex).name
    ConBuf.Put_String tMessage
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_PrivateChat

    Exit Sub

ErrOut:

    Log "Data_Comm_Whisper: TempStr() not sized (delimeter did not exist)", InvalidPacketData '//\\LOGLINE//\\

End Sub

Sub Data_GM_SetGMLevel(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM sets another player as GM
'<TargetName(S)><GMLevel(B)>
'*****************************************************************
Dim TargetName As String
Dim TargetIndex As String
Dim GMLevel As Byte

    Log "Call Data_GM_SetGMLevel([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    TargetName = rBuf.Get_String
    GMLevel = rBuf.Get_Byte
    
    'Make sure the one using the command is admin
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub
    
    'Get the target user's ID
    TargetIndex = User_NameToIndex(TargetName)
    
    'Set the GM level (if online)
    If TargetIndex > 0 Then
        UserList(TargetIndex).flags.GMLevel = GMLevel
    Else
        'Make sure the user exists
        If Server_UserExist(TargetName) Then
        
            'Set GM level (if offline)
            DB_RS.Open "SELECT gm FROM users WHERE `name`='" & TargetName & "'"
            DB_RS!gm = GMLevel
            DB_RS.Update
            DB_RS.Close
            
        Else
            
            'Send the "unknown user" message
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "Can not set user " & TargetName & " to GM - user does not exist!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, PP_GMMessages
            Exit Sub
        
        End If
            
    End If
    
    'Say the change
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 54
    ConBuf.Put_String TargetName
    ConBuf.Put_Byte GMLevel
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, PP_GMMessages

End Sub

Sub Data_GM_Thrall(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM creates NPCs
'<NPCIndex(I)><Amount(I)>
'*****************************************************************
Dim CharIndex As Integer
Dim tIndex As Integer
Dim NPCIndex As Integer
Dim Amount As Integer
Dim i As Long

    Log "Call Data_GM_Thrall([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Get the values
    NPCIndex = rBuf.Get_Integer
    Amount = rBuf.Get_Integer
    
    'Check for invalid values
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If NPCIndex <= 0 Then
        Log "Data_GM_Thrall: Invalid NPCIndex (" & NPCIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If Amount <= 0 Then
        Log "Data_GM_Thrall: Invalid Amount (" & Amount & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If Amount > 50 Then
        Log "Data_GM_Thrall: Invalid Amount (" & Amount & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Thrall the NPCs
    For i = 1 To Amount
    
        'Load up the NPC
        tIndex = Load_NPC(NPCIndex, 1)
        
        'Error in the thrall, so obviously the next ones wont work
        If tIndex < 1 Then
            If i = 1 Then   'Only show the fail if this is the first NPC being thralled
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Comm_Talk
                ConBuf.Put_String "Failed to thrall NPC id " & NPCIndex & " - NPC does not exist!"
                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        End If
        
        'Find the closest legal position from the user to place the NPC
        Server_ClosestLegalPos UserList(UserIndex).Pos, NPCList(tIndex).Pos
        
        'If we ran out of room, remove the NPC and skip making the rest
        If Server_LegalPos(NPCList(tIndex).Pos.Map, NPCList(tIndex).Pos.X, NPCList(tIndex).Pos.Y, NPCList(tIndex).Char.Heading) = False Then
            NPC_Close tIndex
            Exit Sub
        End If
        
        'Set up the NPC on the map / char array
        MapInfo(NPCList(tIndex).Pos.Map).Data(NPCList(tIndex).Pos.X, NPCList(tIndex).Pos.Y).NPCIndex = tIndex
        CharIndex = Server_NextOpenCharIndex
        NPCList(NPCIndex).Char.CharIndex = CharIndex
        CharList(CharIndex).Index = NPCIndex
        CharList(CharIndex).CharType = CharType_NPC
        
        'Display the NPC
        NPC_Spawn tIndex
        NPC_MakeChar ToMap, UserIndex, tIndex, NPCList(tIndex).Pos.Map, NPCList(tIndex).Pos.X, NPCList(tIndex).Pos.Y
        
    Next i
        
End Sub

Sub Data_GM_DeThrall(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Removes all thralled creatures in range
'<>
'*****************************************************************
Dim NPCIndex As Integer
Dim X As Long
Dim Y As Long

    Log "Call Data_GM_DeThrall([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub
    
    'Loop through the nearby tiles
    For X = UserList(UserIndex).Pos.X - 4 To UserList(UserIndex).Pos.X + 4
        For Y = UserList(UserIndex).Pos.Y - 4 To UserList(UserIndex).Pos.Y + 4
            
            'Check for a thralled NPC
            If MapInfo(UserList(UserIndex).Pos.Map).Data(X, Y).NPCIndex Then
                NPCIndex = MapInfo(UserList(UserIndex).Pos.Map).Data(X, Y).NPCIndex
                If NPCList(NPCIndex).flags.Thralled Then
                    NPC_Kill NPCIndex
                End If
            End If
        
        Next Y
    Next X
    
End Sub

Sub Data_GM_Approach(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM approaches user
'<TargetName(S)>
'*****************************************************************

Dim TargetName As String
Dim TargetIndex As Integer
Dim nPos As WorldPos

    Log "Call Data_GM_Approach([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_GMMessages
        Exit Sub
    End If

    'Don't allow if switching maps
    If UserList(TargetIndex).flags.SwitchingMaps Then Exit Sub

    'Find closest legal position and warp there
    Server_ClosestLegalPos UserList(TargetIndex).Pos, nPos
    User_WarpChar UserIndex, nPos.Map, nPos.X, nPos.Y

End Sub

Sub Data_GM_Kick(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM kicked the user
'<TargetName(S)>
'*****************************************************************

Dim TargetName As String
Dim TargetIndex As Integer

    Log "Call Data_GM_Kick([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub

    'Check if the user is online
    If TargetIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_GMMessages
        Exit Sub
    End If

    'Tell everyone that the user was booted
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 77
    ConBuf.Put_String UserList(UserIndex).name
    ConBuf.Put_String UserList(TargetIndex).name
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_GlobalMessage

    'Close the user
    Server_CloseSocket TargetIndex

End Sub

Sub Data_GM_Raise(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM raised a user's experience
'<TargetName(S)><RaiseValue(L)>
'*****************************************************************
Dim TargetName As String
Dim TargetIndex As Integer
Dim RaiseValue As Long

    Log "Call Data_GM_Raise([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

On Error GoTo RaiseErrOut

    TargetName = rBuf.Get_String
    RaiseValue = rBuf.Get_Long
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_GMMessages
        Exit Sub
    End If

    'Raise user's experience
    User_RaiseExp TargetIndex, RaiseValue

    'Show raise message
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 55
    ConBuf.Put_Long RaiseValue
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

Exit Sub

RaiseErrOut:

    Log "Data_GM_Raise: Error raising experience (unknown section of routine) - possible variable overflow", InvalidPacketData '//\\LOGLINE//\\

    'Error raising the char's exp
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "There was an error raising user's experience by " & RaiseValue & " points."
    ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_GM_Summon(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM summons user
'<TargetName(S)>
'*****************************************************************

Dim TargetName As String
Dim TargetIndex As Integer
Dim nPos As WorldPos

    Log "Call Data_GM_Summon([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check for invalid values
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub

    'Check the user admin ID
    If UserList(UserIndex).flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_GMMessages
        Exit Sub
    End If

    'Don't allow if switching maps
    If UserList(TargetIndex).flags.SwitchingMaps Then Exit Sub

    'Find closest legal position and warp there
    Server_ClosestLegalPos UserList(UserIndex).Pos, nPos

    'If the position is legal, warp the user then tell them they have warped
    If Server_LegalPos(nPos.Map, nPos.X, nPos.Y, 0) Then
        User_WarpChar TargetIndex, nPos.Map, nPos.X, nPos.Y
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 56
        ConBuf.Put_String UserList(UserIndex).name
        Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    End If

End Sub

Sub Data_Map_DoneLoadingMap(ByVal UserIndex As Integer)

'*****************************************************************
'User is done loading the map
'<>
'*****************************************************************

    Log "Call Data_Map_DoneLoadingMap(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    If UserList(UserIndex).Pos.Map <= 0 Then
        Log "Data_Map_DoneLoadingMap: Invalid user map (" & UserList(UserIndex).Pos.Map & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).Pos.Map > NumMaps Then
        Log "Data_Map_DoneLoadingMap: Invalid user map (" & UserList(UserIndex).Pos.Map & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If

    UserList(UserIndex).flags.SwitchingMaps = 0

    'Send the map name
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Map_SendName
    ConBuf.Put_String MapInfo(UserList(UserIndex).Pos.Map).name
    ConBuf.Put_Byte MapInfo(UserList(UserIndex).Pos.Map).Weather
    ConBuf.Put_Byte MapInfo(UserList(UserIndex).Pos.Map).Music
    ConBuf.Put_Byte DataCode.Map_DoneSwitching
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Fill in all the characters and objects
    User_UpdateMap UserIndex

    'Send the user's char index
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_UserCharIndex
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Private Sub Data_Send_Update(ByVal UserIndex As Integer, ByVal Priority As Byte)

    Log "Call Data_Send_Update(" & UserIndex & "," & Priority & ")", CodeTracker '//\\LOGLINE//\\

    'The user has a buffer
    UserList(UserIndex).HasBuffer = 1

    'Update the wait time
    If UserList(UserIndex).PacketWait <= 0 Then UserList(UserIndex).PacketWait = CurrentTime + Packet_WaitTime

    'Check if moving from PP_None to PP_Low
    If UserList(UserIndex).PPValue = PP_None Then
        If Priority = PP_Low Then UserList(UserIndex).PPCount = CurrentTime + PP_Time
    End If
    
    'Update the PP value
    If UserList(UserIndex).PPValue < Priority Then UserList(UserIndex).PPValue = Priority

End Sub

Sub Data_Send(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByRef sndData() As Byte, Optional ByVal sndMap As Integer, Optional ByVal Priority As Byte = PP_High)

'*****************************************************************
'Sends data to sendRoute
'*****************************************************************
Dim CopySize As Long
Dim CopyPos As Long
Dim LoopC As Long

    Log "Call Data_Send(" & sndRoute & "," & sndIndex & ",[" & ByteArrayToStr(sndData) & "]," & sndMap & "," & Priority & ")", CodeTracker '//\\LOGLINE//\\

    On Error Resume Next
    
    'Set the copy size
    CopySize = UBound(sndData()) + 1

    'Check the send routes
    Select Case sndRoute
    
            'Send to the sndIndex
        Case ToIndex
            Log "Send ToIndex(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            If sndIndex > LastUser Then Exit Sub
            If UserList(sndIndex).flags.UserLogged = 0 Then Exit Sub
            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
            CopyPos = UserList(sndIndex).BufferSize + 1
            ReDim Preserve UserList(sndIndex).SendBuffer(UserList(sndIndex).BufferSize + CopySize)
            CopyMemory UserList(sndIndex).SendBuffer(CopyPos), sndData(0), CopySize
            UserList(sndIndex).BufferSize = UserList(sndIndex).BufferSize + CopySize
            Data_Send_Update sndIndex, Priority
            
            'Send to All
        Case ToAll
            Log "Send ToAll: " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            For LoopC = 1 To LastUser
                If UserList(LoopC).flags.UserLogged Then
                    If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                    CopyPos = UserList(LoopC).BufferSize + 1
                    ReDim Preserve UserList(LoopC).SendBuffer(UserList(LoopC).BufferSize + CopySize)
                    CopyMemory UserList(LoopC).SendBuffer(CopyPos), sndData(0), CopySize
                    UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                    Data_Send_Update LoopC, Priority
                End If
            Next LoopC
    
            'Send to Map
        Case ToMap
            Log "Send ToMap(" & sndMap & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            If sndMap = 0 Then Exit Sub
            For LoopC = 1 To MapInfo(sndMap).NumUsers
                If UserList(LoopC).flags.UserLogged Then
                    If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                    CopyPos = UserList(LoopC).BufferSize + 1
                    ReDim Preserve UserList(LoopC).SendBuffer(UserList(MapUsers(sndMap).Index(LoopC)).BufferSize + CopySize)
                    CopyMemory UserList(MapUsers(sndMap).Index(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                    UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                    Data_Send_Update LoopC, Priority
                End If
            Next LoopC
    
            'Send to PC Area
        Case ToPCArea
            Log "Send ToPCArea(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = UserList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If UserList(MapUsers(sndMap).Index(LoopC)).flags.UserLogged Then
                        If Server_RectDistance(UserList(sndIndex).Pos.X, UserList(sndIndex).Pos.Y, UserList(MapUsers(sndMap).Index(LoopC)).Pos.X, UserList(MapUsers(sndMap).Index(LoopC)).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(LoopC).BufferSize + 1
                            ReDim Preserve UserList(LoopC).SendBuffer(UserList(MapUsers(sndMap).Index(LoopC)).BufferSize + CopySize)
                            CopyMemory UserList(MapUsers(sndMap).Index(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                            Data_Send_Update LoopC, Priority
                        End If
                    End If
                Next LoopC
            End If
            
            'Send to map (except the index)
        Case ToMapButIndex
            Log "Send ToMapButIndex(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = UserList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If LoopC <> sndIndex Then
                        If UserList(MapUsers(sndMap).Index(LoopC)).flags.UserLogged Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(LoopC).BufferSize + 1
                            ReDim Preserve UserList(LoopC).SendBuffer(UserList(MapUsers(sndMap).Index(LoopC)).BufferSize + CopySize)
                            CopyMemory UserList(MapUsers(sndMap).Index(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                            Data_Send_Update LoopC, Priority
                        End If
                    End If
                Next LoopC
            End If
    
            'Send to NPC Area
        Case ToNPCArea
            Log "Send ToNPCArea(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = NPCList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If UserList(MapUsers(sndMap).Index(LoopC)).flags.UserLogged Then
                        If Server_RectDistance(NPCList(sndIndex).Pos.X, NPCList(sndIndex).Pos.Y, UserList(MapUsers(sndMap).Index(LoopC)).Pos.X, UserList(MapUsers(sndMap).Index(LoopC)).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(LoopC).BufferSize + 1
                            ReDim Preserve UserList(LoopC).SendBuffer(UserList(MapUsers(sndMap).Index(LoopC)).BufferSize + CopySize)
                            CopyMemory UserList(MapUsers(sndMap).Index(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                            Data_Send_Update LoopC, Priority
                        End If
                    End If
                Next LoopC
            End If
            
            'A user moved
        Case ToUserMove
            Log "Send ToUserMove(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = UserList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If MapUsers(sndMap).Index(LoopC) <> sndIndex Then   'Dont send to the user who moved
                        If UserList(MapUsers(sndMap).Index(LoopC)).flags.UserLogged Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(LoopC).BufferSize + 1
                            ReDim Preserve UserList(LoopC).SendBuffer(UserList(MapUsers(sndMap).Index(LoopC)).BufferSize + CopySize)
                            CopyMemory UserList(MapUsers(sndMap).Index(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                            If Server_RectDistance(UserList(sndIndex).Pos.X, UserList(sndIndex).Pos.Y, UserList(LoopC).Pos.X, UserList(LoopC).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                                Data_Send_Update LoopC, PP_CloseCharMove
                            Else
                                Data_Send_Update LoopC, PP_FarCharMove
                            End If
                            UserList(LoopC).HasMovePacket = 1
                        End If
                    End If
                Next LoopC
            End If
            
            'A NPC moved
        Case ToNPCMove
            Log "Send ToNPCMove(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = NPCList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If UserList(MapUsers(sndMap).Index(LoopC)).flags.UserLogged Then
                        If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                        CopyPos = UserList(LoopC).BufferSize + 1
                        ReDim Preserve UserList(LoopC).SendBuffer(UserList(MapUsers(sndMap).Index(LoopC)).BufferSize + CopySize)
                        CopyMemory UserList(MapUsers(sndMap).Index(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                        UserList(LoopC).BufferSize = UserList(LoopC).BufferSize + CopySize
                        If Server_RectDistance(NPCList(sndIndex).Pos.X, NPCList(sndIndex).Pos.Y, UserList(LoopC).Pos.X, UserList(LoopC).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                            Data_Send_Update LoopC, PP_CloseCharMove
                        Else
                            Data_Send_Update LoopC, PP_FarCharMove
                        End If
                        UserList(LoopC).HasMovePacket = 1
                    End If
                Next LoopC
            End If
                    
    End Select
    
ErrOut:

End Sub

Sub Data_Send_Buffer(ByVal UserIndex As Integer)

'*****************************************************************
'Sends the data buffer to the user index
'*****************************************************************

    Log "Call Data_Send_Buffer(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    If GOREsock_State(UserList(UserIndex).ConnID) = soxIdle Or soxListening Then
        If UBound(UserList(UserIndex).SendBuffer) > 0 Then
            
            Log "Send Buffer: " & ByteArrayToStr(UserList(UserIndex).SendBuffer), PacketOut '//\\LOGLINE//\\
            
            'Encrypt the packet
            Select Case PacketEncType
                Case PacketEncTypeXOR
                    Encryption_XOR_EncryptByte UserList(UserIndex).SendBuffer(), PacketEncKey
                Case PacketEncTypeRC4
                    Encryption_RC4_EncryptByte UserList(UserIndex).SendBuffer(), PacketEncKey
            End Select

            'Send the packet
            GOREsock_SendData UserList(UserIndex).ConnID, UserList(UserIndex).SendBuffer()

            'Calculate bandwidth usage
            'TCP header = 20 bytes, IPv4 header = 20 bytes, socket header = 4 bytes
            If CalcTraffic Then DataOut = DataOut + UserList(UserIndex).BufferSize + 44

            'Clear the buffer
            Erase UserList(UserIndex).SendBuffer()
            UserList(UserIndex).HasBuffer = 0
            UserList(UserIndex).BufferSize = -1
            UserList(UserIndex).HasMovePacket = 0

        End If
    End If

    'Check if the user was disconnected
    If GOREsock_State(UserList(UserIndex).ConnID) = 1 Then User_Close UserIndex

    'Check if we are disconnecting
    If UserList(UserIndex).flags.Disconnecting = 1 Then Server_CloseSocket UserIndex

End Sub

Sub Data_Server_Help(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to retrieve help buffer from server
'<>
'*****************************************************************
Dim LoopC As Byte

    Log "Call Data_Server_Help(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Loop through the lines and send them
    ConBuf.Clear
    For LoopC = 1 To NumHelpLines
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String Var_Get(ServerDataPath & "Help.dat", "INIT", "Line" & LoopC)
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Next LoopC
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_Server_MailCompose(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Compose a new message
'<RecieverName(S)><Subject(S)><Message(S-EX)><NumObjs(B)>(<Slot(B)><Amount(I)>)
'*****************************************************************
Dim NumObjs As Byte
Dim RecieverName As String
Dim Subject As String
Dim Message As String
Dim Objs() As Obj
Dim Slot As Byte
Dim i As Long
Dim j As Long
Dim Slots() As Byte

    Log "Call Data_Server_MailCompose([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for enough time to elapse since last mail packet (since this is an easy one to spam to ruin the server CPU)
    If UserList(UserIndex).Counters.DelayTimeMail > CurrentTime Then
        Log "Data_Server_MailCompose: Not enough time elapsed since last mail write for user " & UserIndex & " (" & UserList(UserIndex).name & ")", CodeTracker '//\\LOGLINE//\\
        rBuf.Overflow   'We risk breaking something if we break the packet, so remove the whole thing
        Exit Sub        ' If it bypassed the client-side delay, it was most likely a hacked packet, anyways
    End If
    
    'Get the data
    RecieverName = rBuf.Get_String
    Subject = rBuf.Get_String
    Message = rBuf.Get_StringEX
    NumObjs = rBuf.Get_Byte
    
    'Get the objects (plus lots of valid data checking... -.-)
    If NumObjs > 0 Then
        If NumObjs > MaxMailObjs Then   'Too many objects passed
            rBuf.Overflow
            Exit Sub
        End If
        ReDim Objs(1 To NumObjs) As Obj
        ReDim Slots(1 To NumObjs) As Byte
        For i = 1 To NumObjs
            Slot = rBuf.Get_Byte
            Slots(i) = Slot
            If Slot = 0 Then    'Invalid slot
                rBuf.Overflow
                Exit Sub
            End If
            If Slot > MAX_INVENTORY_SLOTS Then  'Invalid slot
                rBuf.Overflow
                Exit Sub
            End If
            If UserList(UserIndex).Object(Slot).ObjIndex = 0 Then   'User has no object in the slot
                rBuf.Overflow
                Exit Sub
            End If
            Objs(i).ObjIndex = UserList(UserIndex).Object(Slot).ObjIndex
            Objs(i).Amount = rBuf.Get_Integer
            If Objs(i).Amount <= 0 Then 'Invalid amount
                rBuf.Overflow
                Exit Sub
            End If
            If Objs(i).Amount > UserList(UserIndex).Object(Slot).Amount Then    'User doesn't have enough objects
                rBuf.Overflow
                Exit Sub
            End If
        Next i
    End If
    
    'Check for duplicate slots used
    For i = 1 To NumObjs - 1
        For j = i + 1 To NumObjs
            If Slots(i) = Slots(j) Then Exit Sub    'Duplicate slots sent, uh oh!
        Next j
    Next i
    
    'Remove the items from the user
    For i = 1 To NumObjs
        With UserList(UserIndex).Object(Slots(i))
            .Amount = .Amount - Objs(i).Amount
            If .Amount <= 0 Then
        
                'Unequip if the object is currently equipped
                If .Equipped = 1 Then User_RemoveInvItem UserIndex, Slots(i)
                
                'Clear the values
                .ObjIndex = 0
                .Amount = 0
                .Equipped = 0
                
            End If
            
            'Update the inventory slot
            User_UpdateInv False, UserIndex, Slot
            
        End With
    Next i

    'Check for invalid values (most is handled in Server_WriteMail sub)
    If UserList(UserIndex).flags.SwitchingMaps = 0 Then
        If Server_ValidString(RecieverName) Then
            If Server_ValidString(Subject) Then
                If Server_ValidString(Message) Then
            
                    'Write the mail to the RecieverName
                    Server_WriteMail UserIndex, RecieverName, Subject, Message, Objs
                    
                End If
            End If
        End If
    End If

End Sub

Sub Data_Server_MailDelete(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Delete mail
'<MailIndex(B)>
'*****************************************************************

Dim MailIndex As Byte
Dim LoopC As Byte
Dim MsgData As MailData

    Log "Call Data_Server_MailDelete([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    MailIndex = rBuf.Get_Byte

    'Check for invalid values
    If MailIndex < 1 Then
        Log "Data_Server_MailDelete: Invalid mail index (" & MailIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If MailIndex > MaxMailPerUser Then
        Log "Data_Server_MailDelete: Invalid mail index (" & MailIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).MailID(MailIndex) <= 0 Then
        Log "Data_Server_MailDelete: Invalid mail ID (" & UserList(UserIndex).MailID(MailIndex) & ") on mail index (" & MailIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Erase the mail from the file and the user
    DB_Conn.Execute "DELETE FROM mail WHERE id=" & UserList(UserIndex).MailID(MailIndex)
    UserList(UserIndex).MailID(MailIndex) = 0

    'Shift all mail down
    For LoopC = 2 To MaxMailPerUser
        If UserList(UserIndex).MailID(LoopC) > 0 Then
            If UserList(UserIndex).MailID(LoopC - 1) = 0 Then
                UserList(UserIndex).MailID(LoopC - 1) = UserList(UserIndex).MailID(LoopC)
                UserList(UserIndex).MailID(LoopC) = 0
            End If
        End If
    Next LoopC

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

End Sub

Sub Data_Server_MailItemTake(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Take an item out of the mail
'<ObjIndex(B)>
'*****************************************************************

Dim ObjIndex As Byte
Dim LoopC As Byte
Dim MsgData As MailData
Dim X As Long
Dim Y As Long

    Log "Call Data_Server_MailItemTake([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    ObjIndex = rBuf.Get_Byte

    'Check for invalid values
    If UserList(UserIndex).flags.LastViewedMail <= 0 Then Exit Sub
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If ObjIndex > MaxMailObjs Then
        Log "Data_Server_MailItemTake: Invalid object index (" & ObjIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If ObjIndex <= 0 Then
        Log "Data_Server_MailItemTake: Invalid object index (" & ObjIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If

    'Check to make sure the user is still next to the same mailbox
    If UserList(UserIndex).Pos.Map = UserList(UserIndex).MailboxPos.Map Then
        If Not Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).MailboxPos.X, UserList(UserIndex).MailboxPos.Y, 1, 1) Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 57
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    'Get the item information and check if theres an item
    MsgData = Load_Mail(UserList(UserIndex).MailID(UserList(UserIndex).flags.LastViewedMail))
    If MsgData.Obj(ObjIndex).ObjIndex <= 0 Then
        Log "Data_Server_MailItemTake: Invalid object index (" & MsgData.Obj(ObjIndex).ObjIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If MsgData.Obj(ObjIndex).Amount <= 0 Then
        Log "Data_Server_MailItemTake: Invalid object amount (" & MsgData.Obj(ObjIndex).Amount & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If

    'Check if the user already has an item of the type
    LoopC = 1
    Do Until UserList(UserIndex).Object(LoopC).ObjIndex = MsgData.Obj(ObjIndex).ObjIndex
        LoopC = LoopC + 1
        If LoopC > MAX_INVENTORY_SLOTS Then Exit Do
    Loop

    'Check that the user has room in their inventory
    If LoopC > MAX_INVENTORY_SLOTS Then
        LoopC = 1
        Do Until UserList(UserIndex).Object(LoopC).ObjIndex = 0
            LoopC = LoopC + 1
            If LoopC > MAX_INVENTORY_SLOTS Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 26
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        Loop
    Else
        'Check if user has less then MAX_INVENTORY_OBJS objects
        If UserList(UserIndex).Object(LoopC).Amount >= MAX_INVENTORY_OBJS Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 26
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    End If

    'Take all the items
    If UserList(UserIndex).Object(LoopC).Amount + MsgData.Obj(ObjIndex).Amount <= MAX_INVENTORY_OBJS Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 58
        ConBuf.Put_String ObjData(MsgData.Obj(ObjIndex).ObjIndex).name
        ConBuf.Put_Integer MsgData.Obj(ObjIndex).Amount
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
        UserList(UserIndex).Object(LoopC).ObjIndex = MsgData.Obj(ObjIndex).ObjIndex
        UserList(UserIndex).Object(LoopC).Amount = UserList(UserIndex).Object(LoopC).Amount + MsgData.Obj(ObjIndex).Amount
        MsgData.Obj(ObjIndex).ObjIndex = 0
        MsgData.Obj(ObjIndex).Amount = 0

        'Get the number of objects
        Y = 0
        For X = 1 To MaxMailObjs
            If MsgData.Obj(X).ObjIndex > 0 Then Y = Y + 1
        Next X
        
        'Update the mail objects
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MailObjUpdate
        ConBuf.Put_Byte Y   'Number of objects
        For X = 1 To MaxMailObjs
            If MsgData.Obj(X).ObjIndex > 0 Then
                ConBuf.Put_Long ObjData(MsgData.Obj(X).ObjIndex).GrhIndex
                ConBuf.Put_String ObjData(MsgData.Obj(X).ObjIndex).name
                ConBuf.Put_Integer MsgData.Obj(X).Amount
            End If
        Next X
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    Else

        'Take as many items as possible
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 59
        ConBuf.Put_String ObjData(MsgData.Obj(ObjIndex).ObjIndex).name
        ConBuf.Put_Integer (MAX_INVENTORY_OBJS - UserList(UserIndex).Object(LoopC).Amount)
        ConBuf.Put_Integer (MsgData.Obj(ObjIndex).Amount - (MAX_INVENTORY_OBJS - UserList(UserIndex).Object(LoopC).Amount))
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        MsgData.Obj(ObjIndex).Amount = (MsgData.Obj(ObjIndex).Amount + UserList(UserIndex).Object(LoopC).Amount) - MAX_INVENTORY_OBJS
        UserList(UserIndex).Object(LoopC).Amount = MAX_INVENTORY_OBJS

    End If

    'Update the changed inventory slot
    User_UpdateInv False, UserIndex, LoopC

    'Save the changes done to the mail
    Save_Mail UserList(UserIndex).MailID(UserList(UserIndex).flags.LastViewedMail), MsgData

End Sub

Sub Data_Server_MailMessage(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Request to read a message
'<MessageID(B)>
'*****************************************************************

Dim MessageID As Byte
Dim LoopC As Byte
Dim MsgData As MailData
Dim NumObjs As Byte

    Log "Call Data_Server_MailMessage([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    MessageID = rBuf.Get_Byte

    'Check for invalid values
    If MessageID > MaxMailPerUser Then Exit Sub
    If MessageID <= 0 Then Exit Sub
    If UserList(UserIndex).MailID(MessageID) <= 0 Then
        Log "Data_Server_MailMessage: Invalid mail ID (" & UserList(UserIndex).MailID(MessageID) & ") on message ID (" & MessageID & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub

    'Load the mail into a temporary mail buffer
    MsgData = Load_Mail(UserList(UserIndex).MailID(MessageID))

    'Send the message information
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_MailMessage
    ConBuf.Put_StringEX MsgData.Message
    ConBuf.Put_String MsgData.Subject
    ConBuf.Put_String MsgData.WriterName
    NumObjs = 0
    For LoopC = 1 To MaxMailObjs
        If MsgData.Obj(LoopC).ObjIndex > 0 Then NumObjs = NumObjs + 1
    Next LoopC
    ConBuf.Put_Byte NumObjs
    For LoopC = 1 To MaxMailObjs
        If MsgData.Obj(LoopC).ObjIndex > 0 Then
            ConBuf.Put_Long ObjData(MsgData.Obj(LoopC).ObjIndex).GrhIndex
            ConBuf.Put_String ObjData(MsgData.Obj(LoopC).ObjIndex).name
            ConBuf.Put_Integer MsgData.Obj(LoopC).Amount
        End If
    Next LoopC
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Remember the last message the user viewed
    UserList(UserIndex).flags.LastViewedMail = MessageID

    'If the mail was new, then make it old (and save it)
    If MsgData.New = 1 Then
        MsgData.New = 0
        Save_Mail UserList(UserIndex).MailID(MessageID), MsgData
    End If

End Sub

Sub Data_Server_Ping(ByVal UserIndex As Integer)

'*****************************************************************
'We got pinged!
'<>
'*****************************************************************

    Log "Call Data_Server_Ping(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Pong the ping so we can be pinged from our pong

    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Ping
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer()

End Sub

Sub Data_Server_Who(ByVal UserIndex As Integer)

'*****************************************************************
'Send list of who is online
'<>
'*****************************************************************

Dim UserNames() As String
Dim intNumUsers As Integer
Dim LoopC As Long
Dim tStr As String
Dim i As Long

    Log "Call Data_Server_Who(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Get the list of users

    For LoopC = 1 To LastUser
        If LenB(UserList(LoopC).name) Then
            If UserList(LoopC).flags.UserLogged Then
                intNumUsers = intNumUsers + 1
                ReDim Preserve UserNames(1 To intNumUsers)
                UserNames(intNumUsers) = UserList(LoopC).name
            End If
        End If
    Next LoopC

    'Set the first line (total users)
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 78
    ConBuf.Put_Integer intNumUsers

    'Send the user names
    For i = 1 To (intNumUsers \ 10) + 1 'Replace the 10's with the number of users per line
        tStr = vbNullString             'Clear the string buffer
        For LoopC = 1 To 10
            If LoopC + ((i - 1) * 10) > intNumUsers Then Exit For
            tStr = UserNames(LoopC + ((i - 1) * 10))
            tStr = tStr & ", "
        Next LoopC
        tStr = Left$(tStr, Len(tStr) - 2)   'Crop off the last comma
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String tStr
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Next i

    'Send all the lines as a whole
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_Attack(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User attacked
'<Heading(B)>
'*****************************************************************
Dim Heading As Byte

    Log "Call Data_User_Attack(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Invalid values checked by User_Attack routine
    Heading = rBuf.Get_Byte
    User_Attack UserIndex, Heading

End Sub

Sub Data_User_BaseStat(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User wants to raise a stat
'<StatID(B)>
'*****************************************************************
Dim StatID As Byte

    Log "Call Data_User_BaseStat([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    StatID = rBuf.Get_Byte

    'Check for invalid values
    If UserList(UserIndex).flags.UserLogged = 0 Then Exit Sub
    If StatID <= 0 Then
        Log "Data_User_BaseStat: Invalid stat ID (" & StatID & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If StatID > NumStats Then
        Log "Data_User_BaseStat: Invalid stat ID (" & StatID & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If

    'Raise the stat if possible
    If UserList(UserIndex).Stats.BaseStat(SID.Points) >= 1 Then
        UserList(UserIndex).Stats.BaseStat(SID.Points) = UserList(UserIndex).Stats.BaseStat(SID.Points) - 1
        UserList(UserIndex).Stats.BaseStat(StatID) = UserList(UserIndex).Stats.BaseStat(StatID) + 1
    End If

End Sub

Sub Data_User_Blink(ByVal UserIndex As Integer)

'*****************************************************************
'User blinked
'<>
'*****************************************************************

    Log "Call Data_User_Blink(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Force a blink to everyone on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Blink
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_Blink

End Sub

Sub Data_User_CastSkill(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User casts a skill
'<SkillID(B)><TargetIndex(I)>
'*****************************************************************

Dim TargetCharIndex As Integer
Dim TargetIndex As Integer
Dim TargetType As Byte
Dim SkillID As Byte
Dim i As Integer

    Log "Call Data_User_CastSkill([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    SkillID = rBuf.Get_Byte

    'Easiest to always pass a target, even if a target can not be selected for the skill
    TargetCharIndex = rBuf.Get_Integer
    
    'Check for a valid char index
    If TargetCharIndex < 0 Then TargetCharIndex = 0
    If TargetCharIndex > LastChar Then TargetCharIndex = 0
    If CharList(TargetCharIndex).Index <= 0 Then TargetCharIndex = 0
    
    'Clear the pending quest NPC
    UserList(UserIndex).flags.QuestNPC = 0

    'Turn TargetCharIndex into the TargetIndex - get the array index, not char index
    If TargetCharIndex = 0 Then
        TargetIndex = UserIndex
        TargetType = CharType_PC
        UserList(UserIndex).flags.TargetIndex = 0
        UserList(UserIndex).flags.Target = 0
    Else
        TargetIndex = CharList(TargetCharIndex).Index
        'Since PCs and NPCs are in different arrays, we have to pass which one the TargetIndex belongs to, too
        TargetType = CharList(TargetCharIndex).CharType
    End If

    'Call the skill by the SkillID
    Select Case SkillID
    
        Case SkID.Bless
            If TargetType = CharType_PC Then
                Skill_Bless_PCtoPC UserIndex, TargetIndex
            Else
                Skill_Bless_PCtoNPC UserIndex, TargetIndex
            End If
            
        Case SkID.Protection
            If TargetType = CharType_PC Then
                Skill_Protection_PCtoPC UserIndex, TargetIndex
            Else
                Skill_Protection_PCtoNPC UserIndex, TargetIndex
            End If
            
        Case SkID.Heal
            If TargetType = CharType_PC Then
                Skill_Heal_PCtoPC UserIndex, TargetIndex
            ElseIf TargetType = CharType_NPC Then
                Skill_Heal_PCtoNPC UserIndex, TargetIndex
            End If
            
        Case SkID.IronSkin
            Skill_IronSkin_PC UserIndex
            
        Case SkID.SpikeField
            Skill_SpikeField UserIndex
            
        Case SkID.Strengthen
            If TargetType = CharType_PC Then
                Skill_Strengthen_PCtoPC UserIndex, TargetIndex
            Else
                Skill_Strengthen_PCtoNPC UserIndex, TargetIndex
            End If
            
        Case SkID.Warcry
            Skill_Warcry_PC UserIndex
            
    End Select

End Sub

Sub Data_User_ChangeInvSlot(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User changes an item in their inventory from one slot to another
'<SourceSlot(B)><DestinationSlot(B)>
'*****************************************************************

Dim DestObj As UserOBJ
Dim SrcObj As UserOBJ
Dim DestSlot As Byte
Dim SrcSlot As Byte

    Log "Call Data_User_ChangeInvSlot([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    SrcSlot = rBuf.Get_Byte
    DestSlot = rBuf.Get_Byte

    'Check for valid numbers
    If SrcSlot <= 0 Then
        Log "Data_User_ChangeInvSlot: Invalid source slot (" & SrcSlot & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If SrcSlot > MAX_INVENTORY_SLOTS Then
        Log "Data_User_ChangeInvSlot: Invalid source slot (" & SrcSlot & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If DestSlot <= 0 Then
        Log "Data_User_ChangeInvSlot: Invalid destination slot (" & SrcSlot & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If DestSlot > MAX_INVENTORY_SLOTS Then
        Log "Data_User_ChangeInvSlot: Invalid destination slot (" & SrcSlot & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Re-arrange the slots
    DestObj = UserList(UserIndex).Object(SrcSlot)
    SrcObj = UserList(UserIndex).Object(DestSlot)

    'Apply the values and send the update
    UserList(UserIndex).Object(SrcSlot) = SrcObj
    UserList(UserIndex).Object(DestSlot) = DestObj
    User_UpdateInv False, UserIndex, SrcSlot
    User_UpdateInv False, UserIndex, DestSlot

End Sub

Sub Data_User_Desc(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Change user's description
'<Desc(S)>
'*****************************************************************

Dim Desc As String

    Log "Call Data_User_Desc([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Desc = rBuf.Get_String

    'Set the description
    UserList(UserIndex).Desc = Desc

    'Tell the user their description has changed
    If Len(Desc) > 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 60
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        'Tell the user their description has been removed
    Else
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 61
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If

End Sub

Sub Data_User_Drop(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User dropping an object to the ground
'<ObjSlot(B)><Amount(I)>
'*****************************************************************

Dim ObjSlot As Byte
Dim Amount As Integer

    Log "Call Data_User_Drop([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    ObjSlot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer

    'Invalid values handled by sub
    User_DropObj UserIndex, ObjSlot, Amount, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y

End Sub

Sub Data_User_Emote(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Brings up emoticon to everyone on the screen (PCArea)
'<>
'*****************************************************************

Dim EmoteIndex As Byte

    Log "Call Data_User_Emote([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    EmoteIndex = rBuf.Get_Byte

    'Check for invalid values
    If EmoteIndex <= 0 Then
        Log "EmoteIndex: Invalid emote index (" & EmoteIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If EmoteIndex > NumEmotes Then
        Log "EmoteIndex: Invalid emote index (" & EmoteIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If

    'Send the emoticon to everyone nearby
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Emote
    ConBuf.Put_Byte EmoteIndex
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map

End Sub

Sub Data_User_Get(ByVal UserIndex As Integer)

'*****************************************************************
'User grabbing object off of the ground
'<>
'*****************************************************************

    Log "Call Data_User_Get(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Invalid values handled by sub

    User_GetObj UserIndex

End Sub

Sub Data_User_KnownSkills(ByVal UserIndex As Integer)

'*****************************************************************
'Requested to have list of known spells/skills
'<>
'*****************************************************************

    Log "Call Data_User_KnownSkills(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Invalid values handled by sub

    User_SendKnownSkills UserIndex

End Sub

Sub Data_User_LeftClick(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User left-clicked a tile
'<X(B)><Y(B)>
'*****************************************************************

Dim X As Byte
Dim Y As Byte

    Log "Call Data_User_LeftClick([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Invalid values handled by sub
    User_LookAtTile UserIndex, UserList(UserIndex).Pos.Map, X, Y, vbLeftButton

End Sub

Sub Data_User_Login(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Log on existing character
'<Name(S)><Password(S)>
'*****************************************************************

Dim name As String
Dim Pass As String

    Log "Call Data_User_Login([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    name = rBuf.Get_String
    Pass = rBuf.Get_String

    'Invalid values handled by sub
    User_Connect UserIndex, name, Pass

End Sub

Sub Data_User_LookLeft(ByVal UserIndex As Integer)

'*****************************************************************
'User looked left
'<>
'*****************************************************************

    Log "Call Data_User_LookLeft(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Check for invalid values

    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Char.HeadHeading = 0 Then Exit Sub
    If UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.Heading - 1 Then Exit Sub

    'Turn the head
    UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.HeadHeading - 1
    If UserList(UserIndex).Char.HeadHeading = 0 Then UserList(UserIndex).Char.HeadHeading = 4

    'Send the update to the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_LookLeft
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(UserIndex).Char.HeadHeading
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_Look

End Sub

Sub Data_User_LookRight(ByVal UserIndex As Integer)

'*****************************************************************
'User looked right
'<>
'*****************************************************************

    Log "Call Data_User_LookRight(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Check for invalid values

    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.Heading + 1 Then Exit Sub

    'Turn the head
    UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.HeadHeading + 1
    If UserList(UserIndex).Char.HeadHeading = 5 Then UserList(UserIndex).Char.HeadHeading = 1

    'Send the update to the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_LookLeft
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(UserIndex).Char.HeadHeading
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_Look

End Sub

Sub Data_User_Move(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Move the user
'<Direction(B)><X(B)><Y(B)>
'*****************************************************************
Dim uX As Byte
Dim uY As Byte
Dim Dir As Byte
Dim Running As Byte

    Log "Call Data_User_Move([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Dir = rBuf.Get_Byte
    uX = rBuf.Get_Byte
    uY = rBuf.Get_Byte
    
    'Check if running
    If Dir > 128 Then
        Running = 1
        Dir = Dir Xor 128
    End If
    
    'Check if the user's X and Y co-ordinates are correct - if not, send them the right position they should be at
    If uX <> UserList(UserIndex).Pos.X Or uY <> UserList(UserIndex).Pos.Y Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If

    'Invalid values handled by sub
    User_MoveChar UserIndex, Dir, Running

End Sub

Sub Data_User_NewLogin(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Log on a new character
'<Name(S)><Password(S)>
'*****************************************************************

Dim name As String
Dim Pass As String

    Log "Call Data_User_NewLogin([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    name = rBuf.Get_String
    Pass = rBuf.Get_String

    'Invalid values handled by sub
    User_ConnectNew UserIndex, name, Pass, 1, 1

End Sub

Sub Data_User_RightClick(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User right-clicked a tile
'<X(B)><Y(B)>
'*****************************************************************

Dim X As Byte
Dim Y As Byte

    Log "Call Data_User_RightClick([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Invalid values handled by sub
    User_LookAtTile UserIndex, UserList(UserIndex).Pos.Map, CInt(X), CInt(Y), vbRightButton

End Sub

Sub Data_User_Rotate(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Rotate the user
'<Direction(B)>
'*****************************************************************

Dim Dir As Byte

    Log "Call Data_User_Rotate([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Dir = rBuf.Get_Byte

    'Check for invalid direction
    If Dir < 1 Then Exit Sub
    If Dir > 8 Then Exit Sub
    If UserList(UserIndex).Char.Heading = Dir Then Exit Sub 'Dont rotate if we are already facing that direction

    'Rotate the user
    UserList(UserIndex).Char.Heading = Dir

    'Send the rotation update
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte Dir
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_Rotate

End Sub

Sub Data_User_StartQuest(ByVal UserIndex As Integer)

'*****************************************************************
'Start a quest
'<>
'*****************************************************************

Dim QuestNPC As Integer
Dim Slot As Byte
Dim i As Integer

    Log "Call Data_User_StartQuest(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Make sure the user has a quests pending

    If UserList(UserIndex).flags.QuestNPC = 0 Then Exit Sub

    'Check the distance of the user to the NPC
    QuestNPC = UserList(UserIndex).flags.QuestNPC
    If NPCList(QuestNPC).Pos.Map = UserList(UserIndex).Pos.Map Then
        If Server_RectDistance(NPCList(QuestNPC).Pos.X, NPCList(QuestNPC).Pos.Y, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, MaxServerDistanceX + 2, MaxServerDistanceY + 2) Then 'Add a little bit extra to distance, in case NPC moves

            'Check the requirements
            If QuestData(NPCList(QuestNPC).Quest).AcceptReqLvl > UserList(UserIndex).Stats.BaseStat(SID.ELV) Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 62
                ConBuf.Put_Long QuestData(NPCList(QuestNPC).Quest).AcceptReqLvl
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                UserList(UserIndex).flags.QuestNPC = 0
                Exit Sub
            End If
            If QuestData(NPCList(QuestNPC).Quest).AcceptReqObj > 0 Then
                For Slot = 1 To MAX_INVENTORY_SLOTS
                    If UserList(UserIndex).Object(Slot).ObjIndex = QuestData(NPCList(QuestNPC).Quest).AcceptReqObj Then
                        If UserList(UserIndex).Object(Slot).Amount = QuestData(NPCList(QuestNPC).Quest).AcceptReqObjAmount Then
                            Slot = 0
                            Exit For
                        End If
                    End If
                Next Slot
                If Slot <> 0 Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 63
                    ConBuf.Put_Integer QuestData(NPCList(QuestNPC).Quest).AcceptReqObjAmount
                    ConBuf.Put_String ObjData(QuestData(NPCList(QuestNPC).Quest).AcceptReqObj).name
                    UserList(UserIndex).flags.QuestNPC = 0
                    Exit Sub
                End If
            End If

            'Make sure the user has room in the quest queue
            i = 0
            Do
                i = i + 1

                'Uhoh, the user has no more quest room!
                If i > MaxQuests Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 12
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    UserList(UserIndex).flags.QuestNPC = 0
                    Exit Sub
                End If

            Loop While UserList(UserIndex).Quest(i) <> 0

            'We made it out safely, so that means we got an open slot, horray! (>^_^)>
            UserList(UserIndex).Quest(i) = NPCList(QuestNPC).Quest
            
            'Send the requirements
            Quest_SendReqString UserIndex, UserList(UserIndex).Quest(i)

            'Give the user the quest starting stuff
            ConBuf.Clear
            If QuestData(UserList(UserIndex).Quest(i)).AcceptRewExp > 0 Then
                User_RaiseExp UserIndex, QuestData(UserList(UserIndex).Quest(i)).AcceptRewExp
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 3
                ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(i)).AcceptRewExp
            End If
            If QuestData(UserList(UserIndex).Quest(i)).AcceptRewGold > 0 Then
                UserList(UserIndex).Gold = UserList(UserIndex).Gold + QuestData(UserList(UserIndex).Quest(i)).AcceptRewGold
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 4
                ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(i)).AcceptRewGold
            End If
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            
            'Starting reward skills
            If QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill > 0 Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                If UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill) = 1 Then
                    'User already knew the skill
                    ConBuf.Put_Byte 5
                    ConBuf.Put_String Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill)
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Else
                    'User learns the new skill
                    ConBuf.Put_Byte 6
                    ConBuf.Put_String Server_SkillIDtoSkillName(QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill)
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    'Give the user the skill
                    UserList(UserIndex).KnownSkills(QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill) = 1
                    User_SendKnownSkills UserIndex
                End If
            End If

            'Check for an object reward
            If QuestData(UserList(UserIndex).Quest(i)).AcceptRewObj > 0 Then
                User_GiveObj UserIndex, QuestData(UserList(UserIndex).Quest(i)).AcceptRewObj, QuestData(UserList(UserIndex).Quest(i)).AcceptRewObjAmount
            End If

            Exit Sub

        End If
    End If

    'The user is not close enough to the NPC
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 64
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_Trade_SellToNPC(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Sell an item to a NPC
'<Slot(B)><Amount(I)>
'*****************************************************************
Dim NPCIndex As Integer
Dim AddMoney As Long
Dim Amount As Integer
Dim Slot As Byte

    Log "Call Data_User_Trade_SellToNPC([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Get the values
    Slot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer
    
    'Set the NPC index to trade with
    NPCIndex = UserList(UserIndex).flags.TradeWithNPC

    'Check for invalid values
    If NPCIndex < 1 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If NPCList(NPCIndex).NumVendItems < 1 Then Exit Sub
    If Slot <= 0 Then Exit Sub
    If Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).flags.TradeWithNPC <= 0 Then Exit Sub
    If Amount < 0 Then Exit Sub
    If Amount > UserList(UserIndex).Object(Slot).Amount Then Exit Sub
    If UserList(UserIndex).Object(Slot).ObjIndex < 1 Then Exit Sub
    
    'Check for valid locations
    If UserList(UserIndex).Pos.Map <> NPCList(NPCIndex).Pos.Map Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 36
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    If Not Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, 6, 6) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 36
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    
    With UserList(UserIndex).Object(Slot)
    
        'Give the money
        AddMoney = ObjData(.ObjIndex).Price * 0.5 * Amount
        UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + AddMoney
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You sold " & Amount & " " & ObjData(.ObjIndex).name & " for " & AddMoney & " gold."
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
        'Remove the user's objects
        .Amount = .Amount - Amount
    
        'Check if the user lost all their objects
        If .Amount = 0 Then
            User_RemoveInvItem UserIndex, Slot
            .ObjIndex = 0
        End If
        
    End With
    
    'Update the inventory slot
    User_UpdateInv False, UserIndex, Slot

End Sub

Sub Data_User_Trade_BuyFromNPC(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Buy an item from NPC
'<Slot(B)><Amount(I)>
'*****************************************************************

Dim PurchaseObj As ObjData
Dim NPCIndex As Integer
Dim UserSlot As Integer
Dim Amount As Integer
Dim Slot As Byte

    Log "Call Data_User_Trade_BuyFromNPC([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Slot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer

    'Set the NPC index to trade with
    NPCIndex = UserList(UserIndex).flags.TradeWithNPC

    'Check for invalid values
    If NPCIndex <= 0 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If Slot > NPCList(NPCIndex).NumVendItems Then Exit Sub
    If Slot <= 0 Then Exit Sub
    If UserList(UserIndex).flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).flags.TradeWithNPC <= 0 Then Exit Sub
    If NPCList(NPCIndex).VendItems(Slot).ObjIndex <= 0 Then Exit Sub
    If NPCList(NPCIndex).VendItems(Slot).Amount = 0 Then Exit Sub
    If Amount < 0 Then Exit Sub

    'Check for valid locations
    If UserList(UserIndex).Pos.Map <> NPCList(NPCIndex).Pos.Map Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 36
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    If Not Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, 6, 6) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 36
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check to see if User already has object type
    UserSlot = 1
    Do Until UserList(UserIndex).Object(UserSlot).ObjIndex = NPCList(NPCIndex).VendItems(Slot).ObjIndex
        UserSlot = UserSlot + 1
        If UserSlot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop

    'If not, check if there is a empty UserSlot
    If UserSlot > MAX_INVENTORY_SLOTS Then
        UserSlot = 1
        Do Until UserList(UserIndex).Object(UserSlot).ObjIndex = 0
            UserSlot = UserSlot + 1
            If UserSlot > MAX_INVENTORY_SLOTS Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 26
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            End If
        Loop
    End If

    'See how much the user can hold at max
    If UserList(UserIndex).Object(UserSlot).Amount + Amount > MAX_INVENTORY_OBJS Then Amount = MAX_INVENTORY_OBJS - UserList(UserIndex).Object(UserSlot).Amount

    'Set the ObjData of the item to be purchased to the PurchaseObj variable
    PurchaseObj = ObjData(NPCList(NPCIndex).VendItems(Slot).ObjIndex)

    'Check that the user has enough money
    If UserList(UserIndex).Stats.BaseStat(SID.Gold) < PurchaseObj.Price * Amount Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 65
        ConBuf.Put_Integer Amount
        ConBuf.Put_String PurchaseObj.name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Reduce the amount of items the NPC has
    If NPCList(NPCIndex).VendItems(Slot).Amount <> -1 Then

        'Check if there is enough
        If NPCList(NPCIndex).VendItems(Slot).Amount - Amount < 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 66
            ConBuf.Put_Integer Amount
            ConBuf.Put_String PurchaseObj.name
            ConBuf.Put_Integer NPCList(NPCIndex).VendItems(Slot).Amount
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Exit Sub
        End If

        'Reduce the amount
        NPCList(NPCIndex).VendItems(Slot).Amount = NPCList(NPCIndex).VendItems(Slot).Amount - Amount

        'Check if the NPC has hit 0
        If NPCList(NPCIndex).VendItems(Slot).Amount = 0 Then
            NPCList(NPCIndex).VendItems(Slot).ObjIndex = 0
            User_TradeWithNPC UserIndex, NPCIndex    'Update the NPC trade page user-side
        End If

    End If

    'Change the user's inventory item
    If UserList(UserIndex).Object(UserSlot).ObjIndex = 0 Then UserList(UserIndex).Object(UserSlot).ObjIndex = NPCList(NPCIndex).VendItems(Slot).ObjIndex

    'Take the user's money
    UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) - (PurchaseObj.Price * Amount)

    'Give the user the items
    UserList(UserIndex).Object(UserSlot).Amount = UserList(UserIndex).Object(UserSlot).Amount + Amount

    'Send the purchase message
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 67
    ConBuf.Put_Integer Amount
    ConBuf.Put_String PurchaseObj.name
    ConBuf.Put_Long (PurchaseObj.Price * Amount)
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Update the user's inventory
    User_UpdateInv False, UserIndex, UserSlot    'Give the item to the user

End Sub

Sub Data_User_Use(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User uses an object
'<ObjSlot(B)>
'*****************************************************************

Dim ObjSlot As Byte

    Log "Call Data_User_Use([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    ObjSlot = rBuf.Get_Byte

    'Invalid values handled by sub
    User_UseInvItem UserIndex, ObjSlot

End Sub

Sub Server_CloseSocket(ByVal UserIndex As Integer)

'*****************************************************************
'Close the users socket
'*****************************************************************

    Log "Call Server_CloseSocket(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the user is valid
    If UserIndex > 0 Then
        If UserList(UserIndex).ConnID <> -1 Then
            UserList(UserIndex).flags.Disconnecting = 0

            'Send the buffer
            Data_Send_Buffer UserIndex

            'Shut down the socket
            GOREsock_Shut UserList(UserIndex).ConnID

            'If the user hasn't been closed, close it
            If UserList(UserIndex).flags.UserLogged = 1 Then User_Close UserIndex

            'If User_Close was called, the array may have been resized, so make sure we are still in bounds
            If UserIndex > LastUser Then Exit Sub

            'Clear the record of an old inSox, in case it ever did get transfered to another char
            UserList(UserIndex).ConnID = -1
            UserList(UserIndex).flags.UserLogged = 0    'Set to 0 again, just in case

        End If
    End If

End Sub

Sub User_Close(ByVal UserIndex As Integer)

'*****************************************************************
'save user then reset user's slot
'*****************************************************************
Dim LoopC As Long
Dim Map As Integer
Dim name As String

    Log "Call User_Close(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    UserList(UserIndex).flags.Disconnecting = 0

    'Save temps
    Map = (UserList(UserIndex).Pos.Map)
    name = UserList(UserIndex).name

    'Set logged to false
    UserList(UserIndex).flags.UserLogged = 0

    'Save user
    Save_User UserList(UserIndex)
    
    'Erase the user from the map on the other clients
    If UserList(UserIndex).Char.CharIndex > 0 Then User_EraseChar UserIndex

    'Remove user from connection groups and map
    If Map > 0 Then
        If Map <= NumMaps Then
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
            If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0
            If MapInfo(Map).NumUsers Then
                For LoopC = 1 To MapInfo(Map).NumUsers + 1
                    If MapUsers(Map).Index(LoopC) = UserIndex Then Exit For
                Next LoopC
                For LoopC = LoopC To MapInfo(Map).NumUsers
                    MapUsers(Map).Index(LoopC) = MapUsers(Map).Index(LoopC + 1)
                Next LoopC
                ReDim Preserve MapUsers(Map).Index(1 To MapInfo(Map).NumUsers)
            Else
                Unload_Map Map
                Erase MapUsers(Map).Index()
            End If
        End If
    End If

    'Clear the user from memory
    ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))

    'Update the last user value
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged = 1
            LastUser = LastUser - 1
            If LastUser = 0 Then Exit Do
        Loop
    End If
    
    'Resize the userlist according to the LastUser value
    If LastUser = 0 Then
        Erase UserList
    Else
        ReDim Preserve UserList(1 To LastUser)
    End If

    'Update number of users
    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    TrayModify ToolTip, Server_BuildToolTipString

    'Send log off phrase
    If LenB(name) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 68
        ConBuf.Put_String name
        Data_Send ToAll, 0, ConBuf.Get_Buffer
    End If

End Sub

Sub User_Connect(ByVal UserIndex As Integer, ByVal strName As String, ByVal Password As String, Optional ByVal NewUser As Boolean = False)

'*****************************************************************
'Reads the users .chr file and loads into Userlist array
'*****************************************************************
Dim TempPass As String
Dim CharIndex As Integer
Dim TempConnID As Long
Dim LoopC As Byte
Dim Count As Byte

    Log "Call User_Connect(" & UserIndex & "," & strName & "," & Password & "," & NewUser & ")", CodeTracker '//\\LOGLINE//\\
    
    'Make sure the user is not in use
    If UserList(UserIndex).flags.UserLogged Then
        Log "User_Connect: User already logged in", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for valid name and password
    If Not Server_LegalString(strName) Then
        Log "User_Connect: User name (" & strName & ") not legal", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Not Server_LegalString(Password) Then
        Log "User_Connect: Password (" & Password & ") not legal", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(strName) > 10 Then
        Log "User_Connect: Name too long", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Password) > 10 Then
        Log "User_Connect: Password too long", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(strName) < 3 Then
        Log "User_Connect: Name too short", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Password) < 3 Then
        Log "User_Connect: Password too short", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If

    'Set up the variables - we only have to do this if not a new char
    If NewUser = False Then
        TempConnID = UserList(UserIndex).ConnID                     'Transfer over the ConnID or else it'd be 0
        ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))    'Empty the character variables
        UserList(UserIndex).ConnID = TempConnID                     'Set the ConnID back
        UserList(UserIndex).BufferSize = -1                         'Set the buffer start position (-1 = no buffer)
        Set UserList(UserIndex).Stats = New UserStats               'Create the stats class
        UserList(UserIndex).Stats.UserIndex = UserIndex             'Set the user index
    End If
    
    'Set the user as logged in
    UserList(UserIndex).flags.UserLogged = 1

    'Check to see is user already logged with name
    If Server_CheckForSameName(UserIndex, strName) Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 79
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If

    'Check for character file
    If Not Server_UserExist(strName) Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 80
        ConBuf.Put_String strName
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If
    
    'Get the password
    DB_RS.Open "SELECT password FROM users WHERE `name`='" & strName & "'", DB_Conn, adOpenStatic, adLockOptimistic
    TempPass = DB_RS!Password
    DB_RS.Close
    If LenB(TempPass) = 0 Then
    
        'Error getting the password
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 81
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If
    
    'Check password
    If MD5_String(Password) <> TempPass Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 82
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If

    'Load character information from file
    Load_User UserList(UserIndex), strName
    
    'Clear the idle and last packet counter
    UserList(UserIndex).Counters.IdleCount = CurrentTime
    UserList(UserIndex).Counters.LastPacket = CurrentTime

    'Set switching map flag
    UserList(UserIndex).flags.SwitchingMaps = 1

    'Update inventory
    User_UpdateInv True, UserIndex, 0

    'Update number of users
    NumUsers = NumUsers + 1

    'Update map and connection groups data
    MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

    'Check if it's the first user on the map
    If MapInfo(UserList(UserIndex).Pos.Map).NumUsers = 1 Then
        Load_Maps_Temp UserList(UserIndex).Pos.Map
        ReDim MapUsers(UserList(UserIndex).Pos.Map).Index(1 To 1)
    Else
        ReDim Preserve MapUsers(UserList(UserIndex).Pos.Map).Index(1 To MapInfo(UserList(UserIndex).Pos.Map).NumUsers)
    End If
    MapUsers(UserList(UserIndex).Pos.Map).Index(MapInfo(UserList(UserIndex).Pos.Map).NumUsers) = UserIndex
    
    'Get closest legal pos
    Server_ClosestLegalPos UserList(UserIndex).Pos, UserList(UserIndex).Pos
    If Server_LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0) = False Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 83
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        User_Close UserIndex
        Exit Sub
        
    End If
    
    'Tell the user they have successfully connected
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Connect

    'Tell client to try switching maps
    ConBuf.Put_Byte DataCode.Map_LoadMap
    ConBuf.Put_Integer UserList(UserIndex).Pos.Map
    ConBuf.Put_Integer MapInfo(UserList(UserIndex).Pos.Map).MapVersion
    ConBuf.Put_Byte MapInfo(UserList(UserIndex).Pos.Map).Weather
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer()

    'Send data
    Data_Send_Buffer UserIndex

    'Give user a charindex
    CharIndex = Server_NextOpenCharIndex
    UserList(UserIndex).Char.CharIndex = CharIndex
    CharList(CharIndex).Index = UserIndex
    CharList(CharIndex).CharType = CharType_PC

    'Show Character to others
    User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y

    'Refresh tooltip
    TrayModify ToolTip, Server_BuildToolTipString

    'If the user already has a weapon equiped with a range, tell what that range is
    ConBuf.Clear
    If UserList(UserIndex).WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).WeaponEqpObjIndex).WeaponRange > 0 Then
            ConBuf.Put_Byte DataCode.User_SetWeaponRange
            ConBuf.Put_Byte ObjData(UserList(UserIndex).WeaponEqpObjIndex).WeaponRange
        End If
    End If

    'Send welcome message
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Welcome to vbGORE Version 0.3.1!"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "For help, please type |/help|"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "|Ctrl+W| for inventory, |Ctrl+S| for stats."
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Please visit our site at www.vbgore.com"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Questions, code, help, etc can be found there. :)"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info
    
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Remember to go under the |waterfall| to fight some enemies!"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Tell the user if they have new mail
    For LoopC = 1 To MaxMailPerUser
        If UserList(UserIndex).MailID(LoopC) <> 0 Then
            DB_RS.Open "SELECT new FROM mail WHERE id=" & UserList(UserIndex).MailID(LoopC), DB_Conn, adOpenStatic, adLockOptimistic
            If Val(DB_RS!New) = 1 Then Count = Count + 1
            DB_RS.Close
        End If
    Next LoopC
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    If Count > 1 Then
        ConBuf.Put_Byte 71
        ConBuf.Put_Byte Count
    ElseIf Count = 1 Then
        ConBuf.Put_Byte 70
    Else
        ConBuf.Put_Byte 69
    End If
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Send list of known skills
    User_SendKnownSkills UserIndex

    'Send login in sound and login phrase
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte SOUND_WARP
    ConBuf.Put_Byte UserList(UserIndex).Pos.X
    ConBuf.Put_Byte UserList(UserIndex).Pos.Y
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
    'Connect message
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 72
    ConBuf.Put_String UserList(UserIndex).name
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_Connect

End Sub

Sub User_ConnectNew(ByVal UserIndex As Integer, ByVal name As String, ByVal Password As String, ByVal Body As Integer, ByVal Head As Integer)

'*****************************************************************
'Opens a new user. Loads default vars, saves then calls connectuser
'*****************************************************************
Dim TempConnID As Long
Dim TempObjs(1 To 5) As Obj
Dim i As Byte

    Log "Call User_Connect(" & UserIndex & "," & name & "," & Password & "," & Body & "," & Head & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the user is not in use
    If UserList(UserIndex).flags.UserLogged Then Exit Sub

    'Check for a valid body and head
    If Head <> 1 Then Exit Sub
    If Body <> 1 Then Exit Sub

    'Check for valid name and password
    If Len(name) > 10 Then
        Log "User_ConnectNew: Name too long", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Password) > 10 Then
        Log "User_ConnectNew: Password too long", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(name) < 3 Then
        Log "User_ConnectNew: Name too short", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Password) < 3 Then
        Log "User_ConnectNew: Password too short", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Server_LegalString(name) = False Then
        Log "User_ConnectNew: Invalid name string", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Server_LegalString(Password) = False Then
        Log "User_ConnectNew: Invalid password string", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Clear user character
    TempConnID = UserList(UserIndex).ConnID 'Transfer over the ConnID or else it'd be 0
    ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))
    UserList(UserIndex).ConnID = TempConnID
    UserList(UserIndex).BufferSize = -1
    
    'Set the user as logged in
    UserList(UserIndex).flags.UserLogged = 1

    'Check for character file
    If Server_UserExist(name) Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 84
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If

    'Set the stats class
    Set UserList(UserIndex).Stats = New UserStats
    UserList(UserIndex).Stats.UserIndex = UserIndex

    'This is for testing purposes only - remove this in any public release!
    UserList(UserIndex).flags.GMLevel = 1

    'Set the user's variables
    UserList(UserIndex).name = name
    UserList(UserIndex).Char.Heading = SOUTH
    UserList(UserIndex).Char.HeadHeading = SOUTH
    UserList(UserIndex).Char.Head = Head
    UserList(UserIndex).Char.Hair = 1
    UserList(UserIndex).Char.Body = Body
    UserList(UserIndex).Counters.IdleCount = CurrentTime
    UserList(UserIndex).Stats.UserIndex = UserIndex
    UserList(UserIndex).Pos.Map = StartPos.Map
    UserList(UserIndex).Pos.X = StartPos.X
    UserList(UserIndex).Pos.Y = StartPos.Y

    'Set the user's starting stats
    UserList(UserIndex).Stats.BaseStat(SID.ELU) = 10
    UserList(UserIndex).Stats.BaseStat(SID.ELV) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Str) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Agi) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Mag) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Speed) = 5
    UserList(UserIndex).Stats.BaseStat(SID.Gold) = 100
    UserList(UserIndex).Stats.BaseStat(SID.DEF) = 1
    UserList(UserIndex).Stats.BaseStat(SID.MinHIT) = 1
    UserList(UserIndex).Stats.BaseStat(SID.MaxHIT) = 1
    UserList(UserIndex).Stats.BaseStat(SID.MaxHP) = 50
    UserList(UserIndex).Stats.BaseStat(SID.MaxMAN) = 50
    UserList(UserIndex).Stats.BaseStat(SID.MaxSTA) = 50
    UserList(UserIndex).Stats.ModStat(SID.MaxHP) = UserList(UserIndex).Stats.BaseStat(SID.MaxHP)
    UserList(UserIndex).Stats.ModStat(SID.MaxMAN) = UserList(UserIndex).Stats.BaseStat(SID.MaxMAN)
    UserList(UserIndex).Stats.ModStat(SID.MaxSTA) = UserList(UserIndex).Stats.BaseStat(SID.MaxSTA)
    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = 50
    UserList(UserIndex).Stats.BaseStat(SID.MinMAN) = 50
    UserList(UserIndex).Stats.BaseStat(SID.MinSTA) = 50
    
    'Give the user newbie items
    UserList(UserIndex).Object(1).ObjIndex = 1
    UserList(UserIndex).Object(1).Amount = 5
    UserList(UserIndex).Object(2).ObjIndex = 2
    UserList(UserIndex).Object(2).Amount = 1
    UserList(UserIndex).Object(3).ObjIndex = 3
    UserList(UserIndex).Object(3).Amount = 1
    UserList(UserIndex).Object(7).ObjIndex = 8
    UserList(UserIndex).Object(7).Amount = 1
    UserList(UserIndex).Object(8).ObjIndex = 9
    UserList(UserIndex).Object(8).Amount = 50
    
    'Equipt the armor
    UserList(UserIndex).Object(4).ObjIndex = 5
    UserList(UserIndex).Object(4).Amount = 1
    UserList(UserIndex).Object(4).Equipped = 0
    User_UseInvItem UserIndex, 4
    
    'Equipt the weapon
    UserList(UserIndex).Object(5).ObjIndex = 6
    UserList(UserIndex).Object(5).Amount = 1
    UserList(UserIndex).Object(5).Equipped = 0
    User_UseInvItem UserIndex, 5
    
    'Equipt the wings
    UserList(UserIndex).Object(6).ObjIndex = 7
    UserList(UserIndex).Object(6).Amount = 1
    UserList(UserIndex).Object(6).Equipped = 0
    User_UseInvItem UserIndex, 6

    'For testing only
    UserList(UserIndex).KnownSkills(SkID.Bless) = 1
    UserList(UserIndex).KnownSkills(SkID.Heal) = 1
    UserList(UserIndex).KnownSkills(SkID.IronSkin) = 1
    UserList(UserIndex).KnownSkills(SkID.Protection) = 1
    UserList(UserIndex).KnownSkills(SkID.Strengthen) = 1
    UserList(UserIndex).KnownSkills(SkID.Warcry) = 1
    UserList(UserIndex).KnownSkills(SkID.SpikeField) = 1

    'Save the user
    Save_User UserList(UserIndex), MD5_String(Password), 1

    'Write a test message to the user
    For i = 1 To 5
        TempObjs(i).ObjIndex = Int(Rnd * 7) + 1
        TempObjs(i).Amount = Int(Rnd * 10) + 1
    Next i
    For i = 1 To 5
        Server_WriteMail -1, UserList(UserIndex).name, "Test Message", "This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! " & Rnd * 100, TempObjs
    Next i
    
    'Disconnect the user so we can re-connect them through the connect sub
    UserList(UserIndex).flags.UserLogged = 0

    'Go through regular connecting routine
    User_Connect UserIndex, name, Password, True
    
ErrOut:

End Sub

Sub GOREsock_Close(inSox As Long)

'*********************************************
'Socket was closed - make sure the user is logged off and reset the ConnID
'*********************************************

Dim UserIndex As Integer

    Log "Call Socket_OnClose(" & inSox & ")", CodeTracker '//\\LOGLINE//\\

    UserIndex = User_IndexFromSox(inSox)
    
    'Make sure the user is in a valid range
    If UserIndex < 0 Then Exit Sub
    If UserIndex > LastUser Then Exit Sub
    
    'If the user is logged in still, close them down so they can be removed properly
    If UserList(UserIndex).flags.UserLogged = 1 Then User_Close UserIndex

End Sub

Sub GOREsock_Connecting(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Sub GOREsock_Connection(inSox As Long)

'*********************************************
'Accepts new user and assigns an open Index
'*********************************************

Dim Index As Integer

    Log "Call Socket_OnConnection(" & inSox & ")", CodeTracker '//\\LOGLINE//\\

    Index = User_NextOpen

    'Check for max users
    If Index > MaxUsers Then Exit Sub
    UserList(Index).ConnID = inSox
    GOREsock_SetOption inSox, soxSO_TCP_NODELAY, True

End Sub

Sub GOREsock_DataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************

Dim Index As Integer
Dim rBuf As DataBuffer
Dim BufUBound As Long
Dim CommandID As Byte

    Log "Call Socket_OnDataArrival(" & inSox & "," & ByteArrayToStr(inData) & ")", CodeTracker '//\\LOGLINE//\\

    'Get the UserIndex
    Index = User_IndexFromSox(inSox)
    If Index = -1 Then Exit Sub

    'If it is a character disconnecting, do not check their packets since they're doodie heads
    If UserList(Index).flags.Disconnecting Then Exit Sub
    
    'Reset the user's packet counter
    UserList(Index).Counters.LastPacket = CurrentTime
    
    'Calculate data transfer rate
    'TCP header = 20 bytes, IPv4 header = 20 bytes, socket header = 4 bytes
    BufUBound = UBound(inData)
    If CalcTraffic = True Then DataIn = DataIn + BufUBound + 45 '+ 1 because we have to count inData(0)
    
    'Check if to reset the packet flood timer
    If UserList(Index).Counters.PacketsInTime + 1000 < CurrentTime Then
        UserList(Index).Counters.PacketsInTime = CurrentTime
        UserList(Index).Counters.PacketsInCount = 0
    End If
    
    'Decrypt the packet
    Select Case PacketEncType
        Case PacketEncTypeXOR
            Encryption_XOR_DecryptByte inData(), PacketEncKey
        Case PacketEncTypeRC4
            Encryption_RC4_DecryptByte inData(), PacketEncKey
    End Select
    
    'Create the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    
    'Uncomment this to see packets going into the client
    'Dim i As Long
    'Dim s As String
    'For i = LBound(inData) To UBound(inData)
    '    If inData(i) >= 100 Then
    '        s = s & inData(i) & " "
    '    ElseIf inData(i) >= 10 Then
    '        s = s & "0" & inData(i) & " "
    '    Else
    '        s = s & "00" & inData(i) & " "
    '    End If
    'Next i
    'Debug.Print s
    
    Log "Receive: " & ByteArrayToStr(rBuf.Get_Buffer), PacketIn '//\\LOGLINE//\\

    'Loop through the data buffer until it's empty
    'If done right, we should use up exactly every byte in the buffer
    Do

        'Raise the packets in count and check if the user has been flooding packets
        UserList(Index).Counters.PacketsInCount = UserList(Index).Counters.PacketsInCount + 1
        If UserList(Index).Counters.PacketsInCount > 100 Then Exit Do   '100 is our flood limit

        'Get the CommandID
        CommandID = rBuf.Get_Byte
    
        If CommandID >= 100 Then
            Log " * ID: " & CommandID & "  Data Left: " & ByteArrayToStr(rBuf.Get_Buffer_Remainder), PacketIn '//\\LOGLINE//\\
        ElseIf CommandID >= 10 Then
            Log " * ID: 0" & CommandID & "  Data Left: " & ByteArrayToStr(rBuf.Get_Buffer_Remainder), PacketIn '//\\LOGLINE//\\
        Else
            Log " * ID: 00" & CommandID & "  Data Left: " & ByteArrayToStr(rBuf.Get_Buffer_Remainder), PacketIn '//\\LOGLINE//\\
        End If
        
        'Make the appropriate call based on the CommandID
        With DataCode
            
            'Reset idle counter
            If CommandID <> .Server_Ping Then UserList(Index).Counters.IdleCount = CurrentTime
        
            Select Case CommandID
            
            Case 0
                Exit Do
       
            Case .Comm_Emote: Data_Comm_Emote rBuf, Index
            Case .Comm_Shout: Data_Comm_Shout rBuf, Index
            Case .Comm_Talk: Data_Comm_Talk rBuf, Index
            Case .Comm_Whisper: Data_Comm_Whisper rBuf, Index

            Case .GM_Approach: Data_GM_Approach rBuf, Index
            Case .GM_Kick: Data_GM_Kick rBuf, Index
            Case .GM_Raise: Data_GM_Raise rBuf, Index
            Case .GM_SetGMLevel: Data_GM_SetGMLevel rBuf, Index
            Case .GM_Summon: Data_GM_Summon rBuf, Index
            Case .GM_Thrall: Data_GM_Thrall rBuf, Index
            Case .GM_DeThrall: Data_GM_DeThrall rBuf, Index
            
            Case .Map_DoneLoadingMap: Data_Map_DoneLoadingMap Index

            Case .Server_Help: Data_Server_Help Index
            Case .Server_MailCompose: Data_Server_MailCompose rBuf, Index
            Case .Server_MailDelete: Data_Server_MailDelete rBuf, Index
            Case .Server_MailItemTake: Data_Server_MailItemTake rBuf, Index
            Case .Server_MailMessage: Data_Server_MailMessage rBuf, Index
            Case .Server_Ping: Data_Server_Ping Index
            Case .Server_Who: Data_Server_Who Index

            Case .User_Attack: Data_User_Attack rBuf, Index
            Case .User_Bank_PutItem: Data_User_Bank_PutItem rBuf, Index
            Case .User_Bank_TakeItem: Data_User_Bank_TakeItem rBuf, Index
            Case .User_BaseStat: Data_User_BaseStat rBuf, Index
            Case .User_Blink: Data_User_Blink Index
            Case .User_CastSkill: Data_User_CastSkill rBuf, Index
            Case .User_ChangeInvSlot: Data_User_ChangeInvSlot rBuf, Index
            Case .User_Desc: Data_User_Desc rBuf, Index
            Case .User_Drop: Data_User_Drop rBuf, Index
            Case .User_Emote: Data_User_Emote rBuf, Index
            Case .User_Get: Data_User_Get Index
            Case .User_KnownSkills: Data_User_KnownSkills Index
            Case .User_LeftClick: Data_User_LeftClick rBuf, Index
            Case .User_Login: Data_User_Login rBuf, Index
            Case .User_LookLeft: Data_User_LookLeft Index
            Case .User_LookRight: Data_User_LookRight Index
            Case .User_Move: Data_User_Move rBuf, Index
            Case .User_NewLogin: Data_User_NewLogin rBuf, Index
            Case .User_RequestMakeChar: Data_User_RequestMakeChar rBuf, Index
            Case .User_RequestUserCharIndex: Data_User_RequestUserCharIndex Index
            Case .User_RightClick: Data_User_RightClick rBuf, Index
            Case .User_Rotate: Data_User_Rotate rBuf, Index
            Case .User_StartQuest: Data_User_StartQuest Index
            Case .User_Trade_BuyFromNPC: Data_User_Trade_BuyFromNPC rBuf, Index
            Case .User_Trade_SellToNPC: Data_User_Trade_SellToNPC rBuf, Index
            Case .User_Use: Data_User_Use rBuf, Index
            
            Case Else
                Log "OnDataArrival: Command ID " & CommandID & " caused a premature packet handling abortion!", CriticalError '//\\LOGLINE//\\
                Exit Do 'Something went wrong or we hit the end, either way, RUN!!!!
                
            End Select

        End With

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do

    Loop

End Sub
