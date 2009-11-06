Attribute VB_Name = "TCP"
Option Explicit

'Packet priorities constants (influncing when packets are sent)
Public Const PP_None As Long = 2 ^ 28       'Does not force sending
Public Const PP_Low As Long = 2 ^ 29        'Enables a count-down timer
Public Const PP_High As Long = 2 ^ 30       'Packet sent as soon as possible
Public Const PP_DefaultTime As Long = 1000  'How much time must pass for a PP_Low packet to be sent when no time is specified

'Constants used in the Data_Send sub
Public Const ToIndex As Byte = 0                'Send data to a single User index
Public Const ToAll As Byte = 1                  'Send it to all User index
Public Const ToMap As Byte = 2                  'Send it to all users in a map
Public Const ToPCArea As Byte = 3               'Send to all users in a user's area
Public Const ToMapButIndex As Byte = 4          'Send to all users on the map but the userindex
Public Const ToNPCArea As Byte = 5              'Send to all users in a NPC's area
Public Const ToUserMove As Byte = 6             'Send one priority to local map users, another priority to distance map users
Public Const ToNPCMove As Byte = 7              'Same as ToUserMove, but if a NPC moves
Public Const ToGroupButIndex As Byte = 8        'Send to all users in a group but the userindex
Public Const ToGroup As Byte = 9                'Send to all users in a group
Public Const ToMapGroupButIndex As Byte = 10    'Send to all users in a group but the userindex that are on the same map
Public Const ToPCAreaButIndex As Byte = 11      'Send to all the users in the user's area, but the user themself

'*************************
'*** PACKET PRIORITIES ***
'*************************

'*** Packet_WaitTime ***
'How many miliseconds must go by for a packet to be sent (even if PP_High, the time must be waited)
' Keep in mind the highest frame rate is 60 (~17ms), so it is not recommended to go below that value
' or else you are just wasting bandwidth.
'Recommended about 200~300 at highest, about 50 at lowest

'Packet priorities (placed here for easy changing - for advanced users -
' recommended to go packet-by-packet and decide what you want)
'If you have more bandwidth then you could possibly use, then just set
' all the packet priorities to HIGH. Packet priorities is aimed for those
' who want and need an extra booster from their bandwidth.
'vbGORE, by default, aims for the lowest packet priority values that will
' affect the gameplay the least (most stuff not listed as HIGH you wont notice the lag time)

'*** Option 1 - Optimized for speed and performance ***
Public Const Packet_WaitTime As Long = 0        'Recommended about 200~300 at highest - lower is faster, but more packets sent (more bandwidth used)

'Communication
Public Const PP_GlobalChat As Long = PP_High    'Talking to everyone online (shout)
Public Const PP_LocalChat As Long = PP_High     'Talking to people locally (normal chat, emoting)
Public Const PP_PrivateChat As Long = PP_High   'Talking to someone privately (private messages, NOT mailing)
Public Const PP_GMMessages As Long = PP_High    'Messages sent back from using a GM command (only GMs can see)
Public Const PP_GlobalMessage As Long = PP_High 'Messages sent by the server globally not specified above (kicking a user, server messages, etc)

'Visuals information
Public Const PP_StatusIcons As Long = PP_High   'Displaying and removing status icons
Public Const PP_DisplaySpell As Long = PP_High  'Displaying spell effects
Public Const PP_ChangeChar As Long = PP_High    'Updating character paperdoll values
Public Const PP_Blink As Long = PP_High         'Sending character eye blinks
Public Const PP_Look As Long = PP_High          'Change looking direction (lookleft, lookright)
Public Const PP_Rotate As Long = PP_High        'Change body direction (rotating)
Public Const PP_CloseCharMove As Long = PP_High 'Character movement of near-by characters (in screen)
Public Const PP_FarCharMove As Long = PP_None   'Character movement of characters out of the screen
Public Const PP_GroupCharMove As Long = PP_High 'Character in the user's group moves
Public Const PP_GroundObjects As Long = PP_High 'Updating the objects on the ground, adding and removing them

'Misc
Public Const PP_Banking As Long = PP_High       'All information related to banking
Public Const PP_Trading As Long = PP_High       'All information on peer-to-peer trading and npc-to-pc trading
Public Const PP_Mail As Long = PP_High          'Mailing information
Public Const PP_Sound As Long = PP_High         'Playing sound effects
Public Const PP_NewMail As Long = PP_High       '"You just got new mail" message
Public Const PP_Connect As Long = PP_High       '"User xxx connected" message
Public Const PP_Stat As Long = PP_High          'Base/mod stat updating
Public Const PP_StatPercent As Long = PP_High   'Stat percentage (MP and SP bars) updating

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
    If UserList(UserIndex).Flags.TradeWithNPC > 0 Then
        With NPCList(UserList(UserIndex).Flags.TradeWithNPC)
            If NPCList(UserList(UserIndex).Flags.TradeWithNPC).AI <> 6 Then Exit Sub    'Not a banker
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Pos.X, .Pos.Y, MaxServerDistanceX, MaxServerDistanceY) = 0 Then Exit Sub  'Out of range
        End With
    Else
        Exit Sub    'No clicked NPC
    End If
    
    'Check for item of the same type already in there
    PutSlot = 1
    If ObjData.Stacking(UserList(UserIndex).Bank(PutSlot).ObjIndex) > 1 Then
        PutSlot = 0
        Do
            PutSlot = PutSlot + 1
            If PutSlot > MAX_INVENTORY_SLOTS Then
                PutSlot = 0
                Exit Do
            End If
        Loop While UserList(UserIndex).Bank(PutSlot).ObjIndex <> UserList(UserIndex).Object(Slot).ObjIndex
    Else
        PutSlot = 0  'Force to check the next free slot
    End If
    
    'If PutSlot = 0, no duplicate item was found, so use the next free slot
    If PutSlot = 0 Then
        Do
            PutSlot = PutSlot + 1
            If PutSlot > MAX_INVENTORY_SLOTS Then
                
                'Bank is full
                Data_Send ToIndex, UserIndex, cMessage(97).Data
                Exit Sub
                
            End If
        Loop While UserList(UserIndex).Bank(PutSlot).ObjIndex > 0
        
        'Just as a pre-caution, we empty the amount value since we are going to be adding on to it and it should be empty
        UserList(UserIndex).Bank(PutSlot).Amount = 0
        
    End If
    
    'Check if theres room for the item
    If UserList(UserIndex).Bank(PutSlot).Amount + Amount > ObjData.Stacking(UserList(UserIndex).Object(Slot).ObjIndex) Then
        Data_Send ToIndex, UserIndex, cMessage(97).Data
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
        If UserList(UserIndex).Object(Slot).Equipped Then User_RemoveInvItem UserIndex, Slot
        
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
Dim AmountTaken As Integer
Dim Amount As Integer
Dim Slot As Byte

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
    If UserList(UserIndex).Flags.TradeWithNPC > 0 Then
        With NPCList(UserList(UserIndex).Flags.TradeWithNPC)
            If NPCList(UserList(UserIndex).Flags.TradeWithNPC).AI <> 6 Then Exit Sub    'Not a banker
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Pos.X, .Pos.Y, MaxServerDistanceX, MaxServerDistanceY) = 0 Then Exit Sub  'Out of range
        End With
    Else
        Exit Sub    'No clicked NPC
    End If
    
    'Grab the objects from the bank account
    AmountTaken = User_GiveObj(UserIndex, UserList(UserIndex).Bank(Slot).ObjIndex, UserList(UserIndex).Bank(Slot).Amount)
    
    'Take away the objects from the bank
    UserList(UserIndex).Bank(Slot).Amount = UserList(UserIndex).Bank(Slot).Amount - AmountTaken
    If UserList(UserIndex).Bank(Slot).Amount <= 0 Then
        UserList(UserIndex).Bank(Slot).Amount = 0
        UserList(UserIndex).Bank(Slot).ObjIndex = 0
    End If
    
    'Update the bank slot
    User_UpdateBank UserIndex, Slot

End Sub

Sub Data_User_RequestUserCharIndex(ByVal UserIndex As Integer)

'*****************************************************************
'Client requests their user char index
'<>
'*****************************************************************

    'Send the user char index
    ConBuf.PreAllocate 3
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

Dim Txt As String

    Log "Call Data_Comm_Emote([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Txt = rBuf.Get_String

    'Check for invalid conditions
    If Txt = vbNullString Then
        Log "Data_Comm_Emote: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > timeGetTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = timeGetTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    'Uncomment this if filtering issues arise and the filtering needs to be done on the server
    'If Server_ValidString(Txt) = False Then Exit Sub
    
    'Write the message
    ConBuf.PreAllocate 6 + Len(UserList(UserIndex).Name) + Len(Txt)
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).Name & " " & Txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk Or DataCode.Comm_UseBubble
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_LocalChat

End Sub

Sub Data_Comm_Shout_ServerToServer(ByRef rBuf As DataBuffer)

'*****************************************************************
'Another server told us someone shouted
'<Name(S)><txt(S)>
'*****************************************************************
Dim Name As String
Dim Txt As String

    Log "Call Data_Comm_Shout_ServerToServer([" & ByteArrayToStr(rBuf.Get_Buffer) & "])", CodeTracker '//\\LOGLINE//\\

    'Get the values out of the buffer (we assume they are valid because it is from another server)
    Name = rBuf.Get_String
    Txt = rBuf.Get_String

    'Create the packet and send it to everyone
    ConBuf.PreAllocate 6 + Len(Name) + Len(Txt)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 76
    ConBuf.Put_String Name
    ConBuf.Put_String Txt
    ConBuf.Put_Integer 0    'We know the user isn't in the same screen (or even server) so don't pass an index
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_GlobalChat
    
End Sub

Sub Data_Comm_Shout(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Shout
'<txt(S)>
'*****************************************************************

Dim Txt As String
Dim i As Byte

    Log "Call Data_Comm_Shout([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Txt = rBuf.Get_String

    'Check for invalid conditions
    If Txt = vbNullString Then
        Log "Data_Comm_Shout: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > timeGetTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = timeGetTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    'Uncomment this if filtering issues arise and the filtering needs to be done on the server
    'If Not Server_ValidString(Txt) Then Exit Sub
    
    'Write the message to all players on this server
    ConBuf.PreAllocate 6 + Len(UserList(UserIndex).Name) + Len(Txt)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 76
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_String Txt
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToAll, UserIndex, ConBuf.Get_Buffer, , PP_GlobalChat
    
    'Check for more servers
    If NumServers > 1 Then
    
        'Build the packet now so we only have to do it once
        ConBuf.PreAllocate 3 + Len(Txt) + Len(UserList(UserIndex).Name)
        ConBuf.Put_Byte DataCode.Comm_Shout
        ConBuf.Put_String UserList(UserIndex).Name
        ConBuf.Put_String Txt
        
        'Write the message to every other server
        For i = 1 To NumServers
        
            'Make sure we're not going to THIS server
            If ServerID <> i Then
            
                'Confirm the connection to the server
                Server_ConnectToServer i
                
                'Send the data
                If frmMain.ServerSocket(i).State = sckConnected Then frmMain.ServerSocket(i).SendData ConBuf.Get_Buffer
                
            End If
        
        Next i
        
    End If

End Sub

Sub Data_Comm_Talk(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Talk
'<txt(S-EX)>
'*****************************************************************

Dim Txt As String

    Log "Call Data_Comm_Talk([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Txt = rBuf.Get_String

    'Check for invalid conditions
    If Txt = vbNullString Then
        Log "Data_Comm_Talk: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > timeGetTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = timeGetTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    'Uncomment this if filtering issues arise and the filtering needs to be done on the server
    'If Not Server_ValidString(Txt) Then Exit Sub
    
    'Write the message
    ConBuf.PreAllocate 7 + Len(UserList(UserIndex).Name) + Len(Txt)
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).Name & ": " & Txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk Or DataCode.Comm_UseBubble
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_LocalChat

End Sub

Sub Data_Comm_GroupTalk(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Talk
'<txt(S-EX)>
'*****************************************************************
Dim Txt As String

    Log "Call Data_Comm_GroupTalk([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Txt = rBuf.Get_String
    
    'Check if the user is in a group
    If UserList(UserIndex).GroupIndex = 0 Then Exit Sub

    'Check for invalid conditions
    If Txt = vbNullString Then
        Log "Data_Comm_Talk: txt = vbNullString", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > timeGetTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = timeGetTime + DelayTimeTalk
    End If
    
    'Check if the string is valid
    'Uncomment this if filtering issues arise and the filtering needs to be done on the server
    'If Not Server_ValidString(Txt) Then Exit Sub
    
    'Write the message
    ConBuf.PreAllocate 7 + Len(UserList(UserIndex).Name) + Len(Txt)
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).Name & ": " & Txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Group
    Data_Send ToGroup, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_Comm_Whisper_ServerToServer(ByRef rBuf As DataBuffer)

'*****************************************************************
'Whisper (message from another server)
'<Name(S)><Txt(S)>
'*****************************************************************
Dim Name As String
Dim Txt As String
Dim tIndex As Integer

    Name = rBuf.Get_String
    Txt = rBuf.Get_String
    
    'Get the user's ID
    tIndex = User_NameToIndex(Name)
    If tIndex <= 0 Then Exit Sub
    
    'Send the user the message (since it is from another server, we assume it is valid)
    ConBuf.PreAllocate 4 + Len(Name) + Len(Txt)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 52
    ConBuf.Put_String Name
    ConBuf.Put_String Txt
    Data_Send ToIndex, tIndex, ConBuf.Get_Buffer, , PP_PrivateChat

End Sub

Sub Data_Comm_Whisper(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Whisper
'<Name(S)><Message(S)>
'*****************************************************************
Dim ServerIndex As Byte
Dim tName As String
Dim tIndex As Integer
Dim tMessage As String

    Log "Call Data_Comm_Whisper([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    tName = rBuf.Get_String
    tMessage = rBuf.Get_String
    
    'Check for enough elapsed time
    If UserList(UserIndex).Counters.DelayTimeTalk > timeGetTime Then
        Exit Sub
    Else
        UserList(UserIndex).Counters.DelayTimeTalk = timeGetTime + DelayTimeTalk
    End If

    'Check for valid tName
    If Not Server_LegalString(tName) Then Exit Sub
    
    'Check if the string is valid
    'Uncomment this if filtering issues arise and the filtering needs to be done on the server
    'If Not Server_ValidString(tMessage) Then Exit Sub

    'Retrieve the index of the player being whispered to
    tIndex = User_NameToIndex(tName)
    
    'User is not on this server
    If tIndex <= 0 Then
    
        'Make sure there is more then one server before checking if the user is on another server
        If NumServers = 1 Then Exit Sub
    
        'Check what server the user is on
        DB_RS.Open "SELECT server FROM users WHERE `name`='" & tName & "'", DB_Conn, adOpenStatic, adLockOptimistic
        ServerIndex = Val(DB_RS!server)
        If DB_RS.EOF Or ServerIndex = 0 Then
            DB_RS.Close
            
            'User doesn't exist or is offline
            Data_Send ToIndex, UserIndex, cMessage(51).Data
        
        Else
            DB_RS.Close
            
            'User has a valid server, message the server
            Server_ConnectToServer ServerIndex 'Confirm the connection
            If frmMain.ServerSocket(ServerIndex).State = sckConnected Then
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.Comm_Whisper
                ConBuf.Put_String tName
                ConBuf.Put_String tMessage
                frmMain.ServerSocket(ServerIndex).SendData ConBuf.Get_Buffer
            End If
        
        End If
        
    Else
    
        'Tell the target what they got whispered
        ConBuf.PreAllocate 4 + Len(UserList(UserIndex).Name) + Len(tMessage)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 52
        ConBuf.Put_String UserList(UserIndex).Name
        ConBuf.Put_String tMessage
        Data_Send ToIndex, tIndex, ConBuf.Get_Buffer, , PP_PrivateChat
    
        'Tell the whisperer what they whispered
        ConBuf.PreAllocate 4 + Len(UserList(tIndex).Name) + Len(tMessage)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 53
        ConBuf.Put_String UserList(tIndex).Name
        ConBuf.Put_String tMessage
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_PrivateChat

    End If

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
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
    
    'Get the target user's ID
    TargetIndex = User_NameToIndex(TargetName)
    
    'Set the GM level (if online)
    If TargetIndex > 0 Then
        UserList(TargetIndex).Flags.GMLevel = GMLevel
    Else
        'Make sure the user exists
        If Server_UserExist(TargetName) Then
        
            'Set GM level (if offline)
            DB_RS.Open "SELECT gm FROM users WHERE `name`='" & TargetName & "'", DB_Conn, adOpenStatic, adLockOptimistic
            DB_RS!gm = GMLevel
            DB_RS.Update
            DB_RS.Close
            
        Else
            
            'Send the "unknown user" message
            ConBuf.PreAllocate 3 + Len(TargetName)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 94
            ConBuf.Put_String TargetName
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, PP_GMMessages
            Exit Sub
        
        End If
            
    End If
    
    'Say the change
    ConBuf.PreAllocate 4 + Len(TargetName)
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
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
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
                ConBuf.PreAllocate 4
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 95
                ConBuf.Put_Integer NPCIndex
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
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
    
    'Loop through the nearby tiles
    For X = UserList(UserIndex).Pos.X - 4 To UserList(UserIndex).Pos.X + 4
        For Y = UserList(UserIndex).Pos.Y - 4 To UserList(UserIndex).Pos.Y + 4
            
            'Check for a thralled NPC
            If X > 0 Then
                If Y > 0 Then
                    If X <= MapInfo(UserList(UserIndex).Pos.Map).Width Then
                        If Y <= MapInfo(UserList(UserIndex).Pos.Map).Height Then
                            If MapInfo(UserList(UserIndex).Pos.Map).Data(X, Y).NPCIndex Then
                                NPCIndex = MapInfo(UserList(UserIndex).Pos.Map).Data(X, Y).NPCIndex
                                If NPCList(NPCIndex).Flags.Thralled Then
                                    If NPCList(NPCIndex).OwnerIndex = 0 Then
                                        NPC_Kill NPCIndex
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        
        Next Y
    Next X
    
End Sub

Sub Data_GM_UnBanIP(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM removes an IP ban
'<IP(S)>
'*****************************************************************
Dim TempS() As String
Dim IP As String
Dim i As Long

    Log "Call Data_GM_UnBanIP([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Get the IP
    IP = rBuf.Get_String
    
    'Check that the user is an admin
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
    
    'Check for a valid IP
    If LenB(IP) = 0 Then Exit Sub
    
    'Split up the IP values
    TempS = Split(IP, ".")
    If UBound(TempS) <> 3 Then Exit Sub
    
    'Check for a valid value on each slot
    For i = 0 To 3
        If Val(TempS(i)) < 0 Or Val(TempS(i)) > 255 Then Exit Sub
    Next i
    
    'Check if the IP is in the database
    DB_RS.Open "SELECT * FROM banned_ips WHERE `ip`='" & IP & "'", DB_Conn, adOpenStatic, adLockOptimistic
    If Not DB_RS.EOF Then
        'IP found, remove it
        DB_RS.Delete
        Data_Send ToIndex, UserIndex, cMessage(101).Data
    Else
        'IP not found, send error message
        Data_Send ToIndex, UserIndex, cMessage(102).Data
    End If
    DB_RS.Close
    
End Sub

Sub Data_GM_BanIP(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM bans an IP
'<IP(S)><Reason(S)>
'*****************************************************************
Dim TempS() As String
Dim IP As String
Dim Reason As String
Dim i As Long

    Log "Call Data_GM_BanIP([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\
    
    'Get the IP and reason
    IP = rBuf.Get_String
    Reason = rBuf.Get_String
    
    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
    
    'Check for a reason and IP
    If LenB(Reason) = 0 Then Exit Sub
    If LenB(IP) = 0 Then Exit Sub
    
    'Check if the IP is in the database
    DB_RS.Open "SELECT ip FROM banned_ips WHERE `ip`='" & IP & "'", DB_Conn, adOpenStatic, adLockOptimistic
    If Not DB_RS.EOF Then
        'IP already exists
        DB_RS.Close
        Data_Send ToIndex, UserIndex, cMessage(98).Data
        Exit Sub
    End If
    DB_RS.Close
    
    'Split up the IP values
    TempS = Split(IP, ".")
    If UBound(TempS) <> 3 Then Exit Sub
    
    'Check for a valid value on each slot
    For i = 0 To 3
        If Val(TempS(i)) < 0 Or Val(TempS(i)) > 255 Then Exit Sub
    Next i
    
    'Add the IP and reason to the database
    DB_RS.Open "SELECT * FROM banned_ips WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.AddNew
    DB_RS!IP = IP
    DB_RS!Reason = Reason
    DB_RS.Update
    DB_RS.Close
    
    'Send the success message
    Data_Send ToIndex, UserIndex, cMessage(99).Data
    
    'Check if anyone online is affected by the ban
    For i = 1 To LastUser
        If UserList(i).Flags.UserLogged Then
            If Server_IPisBanned(frmMain.GOREsock.Address(i), Reason) Then
                ConBuf.PreAllocate 3 + Len(Reason)
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 100
                ConBuf.Put_String Reason
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Data_Send_Buffer UserIndex
                Server_CloseSocket UserIndex
            End If
        End If
    Next i

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
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        Data_Send ToIndex, UserIndex, cMessage(51).Data, , PP_GMMessages
        Exit Sub
    End If

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
Dim i As Byte

    Log "Call Data_GM_Kick([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

    'Check if the user is online
    If TargetIndex <= 0 Then
        Data_Send ToIndex, UserIndex, cMessage(51).Data, , PP_GMMessages
        Exit Sub
    End If

    'Tell everyone that the user was booted
    ConBuf.PreAllocate 4 + Len(UserList(UserIndex).Name) + Len(UserList(TargetIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 77
    ConBuf.Put_String UserList(TargetIndex).Name
    ConBuf.Put_String UserList(UserIndex).Name
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_GlobalMessage

    'For multiple servers, send to all the servers
    If NumServers > 1 Then
        For i = 1 To NumServers
            If ServerID <> i Then
                Server_ConnectToServer i
                If frmMain.ServerSocket(i).State = sckConnected Then frmMain.ServerSocket(i).SendData ConBuf.Get_Buffer
            End If
        Next i
    End If

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
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        Data_Send ToIndex, UserIndex, cMessage(51).Data, , PP_GMMessages
        Exit Sub
    End If

    'Raise user's experience
    User_RaiseExp TargetIndex, RaiseValue
    
    On Error GoTo 0

    'Show raise message
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 55
    ConBuf.Put_Long RaiseValue
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

Exit Sub

RaiseErrOut:

    Log "Data_GM_Raise: Error raising experience (unknown section of routine) - possible variable overflow", InvalidPacketData '//\\LOGLINE//\\

    'Error raising the char's exp
    ConBuf.PreAllocate 59 + Len(CStr(RaiseValue))
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "There was an error raising user's experience by " & RaiseValue & " points."
    ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_Trade_Trade(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User requests to trade with another user
'<TargetName(S)>
'*****************************************************************
Dim TargetName As String
Dim TargetIndex As Integer

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)
    
    'Don't allow trading with yourself
    If TargetIndex = UserIndex Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        Data_Send ToIndex, UserIndex, cMessage(51).Data
        Exit Sub
    End If
    
    'Check the distance
    If UserList(UserIndex).Pos.Map = UserList(TargetIndex).Pos.Map Then
        If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(TargetIndex).Pos.X, UserList(TargetIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) = 1 Then
                    
            'Put the user in their trade request flag
            UserList(UserIndex).Flags.TradeWith = TargetIndex
            
            'Check if the user has them in their trade flag
            If UserList(TargetIndex).Flags.TradeWith = UserIndex Then
            
                'Woot, they want to trade with eachother - isn't that CUTE!!! ^_^
                TradeTable_Create UserIndex, TargetIndex
                
                'Clear the tradewith flags
                UserList(UserIndex).Flags.TradeWith = 0
                UserList(TargetIndex).Flags.TradeWith = 0
                
            Else
            
                'Send the trade request message
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 128
                ConBuf.Put_String UserList(UserIndex).Name
                Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
                
            End If
            
            Exit Sub
            
        End If
    End If
    
    'Invalid distance
    Data_Send ToIndex, UserIndex, cMessage(36).Data()

End Sub

Sub Data_User_Trade_RemoveItem(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User removes an item from their trade table
'<TableSlot(B)>
'*****************************************************************
Dim TableSlot As Byte

    TableSlot = rBuf.Get_Byte
    
    'Invalid values handled by the sub
    TradeTable_RemoveItem UserIndex, TableSlot

End Sub

Sub Data_User_Trade_UpdateTrade(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Update a trade table
'<InvSlot(B)><Amount(L)>
'*****************************************************************
Dim InvSlot As Byte
Dim Amount As Long

    'Get the values
    InvSlot = rBuf.Get_Byte
    Amount = rBuf.Get_Long
    
    'Invalid values handled by the sub
    TradeTable_UpdateSlot UserIndex, InvSlot, Amount

End Sub

Sub Data_User_Trade_Cancel(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to cancel a trade
'<>
'*****************************************************************

    TradeTable_Close UserList(UserIndex).Flags.TradeTable

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

    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        Data_Send ToIndex, UserIndex, cMessage(51).Data, , PP_GMMessages
        Exit Sub
    End If

    'Find closest legal position and warp there
    Server_ClosestLegalPos UserList(UserIndex).Pos, nPos

    'If the position is legal, warp the user then tell them they have warped
    If Server_LegalPos(nPos.Map, nPos.X, nPos.Y, 0) Then
        User_WarpChar TargetIndex, nPos.Map, nPos.X, nPos.Y
        ConBuf.PreAllocate 3 + Len(UserList(UserIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 56
        ConBuf.Put_String UserList(UserIndex).Name
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

    'Send the map name
    ConBuf.PreAllocate 4 + Len(MapInfo(UserList(UserIndex).Pos.Map).Name)
    ConBuf.Put_Byte DataCode.Map_SendName
    ConBuf.Put_String MapInfo(UserList(UserIndex).Pos.Map).Name
    ConBuf.Put_Byte MapInfo(UserList(UserIndex).Pos.Map).Weather
    ConBuf.Put_Byte MapInfo(UserList(UserIndex).Pos.Map).Music
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Fill in all the characters and objects
    User_UpdateMap UserIndex

    'Send the user's char index
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.Server_UserCharIndex
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Private Sub Data_Send_Update(ByVal UserIndex As Integer, ByVal Priority As Long)
Dim PPTime As Long

    Log "Call Data_Send_Update(" & UserIndex & "," & Priority & ")", CodeTracker '//\\LOGLINE//\\

    'The user has a buffer
    UserList(UserIndex).HasBuffer = 1

    'Update the overall wait time (only happens on the first packet added to the buffer)
    If UserList(UserIndex).PacketWait = 0 Then UserList(UserIndex).PacketWait = timeGetTime + Packet_WaitTime

    'If the priority is None, we don't have to update anything
    If Priority And PP_None Then Exit Sub

    'If the priority PP_Low, then remove the PP_Low flag and get the value
    If Priority And PP_Low Then
    
        'Get the time to wait by removing the priority
        PPTime = Priority Xor PP_Low

        'If no priority was specified (time = 0), we stick on the default time, since if you want time = 0, use PP_High
        If PPTime = 0 Then PPTime = PP_DefaultTime
        
        'Check if the PPCount has even been set yet
        If UserList(UserIndex).PPCount = 0 Then
            
            'No time has been set yet, so set it
            UserList(UserIndex).PPCount = PPTime
            
        Else
        
            'Check if the time specified is sooner then the time we're already waiting
            If UserList(UserIndex).PPCount - timeGetTime > PPTime Then
            
                'The time we just specified in PPTime is lower than the time we were waiting, so change to the lower time
                UserList(UserIndex).PPCount = timeGetTime + PPTime
                
            End If
            
        End If
        
        Exit Sub
    End If

    'The priority must be high (or invalid, so treat it as high) so send it instantly
    UserList(UserIndex).PPCount = timeGetTime

End Sub

Sub Data_Send(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByRef sndData() As Byte, Optional ByVal sndMap As Integer, Optional ByVal Priority As Long = PP_High)

'*****************************************************************
'Sends data to sendRoute
'*****************************************************************
Dim CopySize As Long
Dim CopyPos As Long
Dim LoopC As Long
Dim tIndex As Integer

    Log "Call Data_Send(" & sndRoute & "," & sndIndex & ",[" & ByteArrayToStr(sndData) & "]," & sndMap & "," & Priority & ")", CodeTracker '//\\LOGLINE//\\

    'Set the copy size
    CopySize = UBound(sndData()) + 1

    'Check the send routes
    Select Case sndRoute
    
            'Send to the sndIndex
        Case ToIndex
            Log "Send ToIndex(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            If sndIndex < 1 Then Exit Sub
            If sndIndex > LastUser Then Exit Sub
            If UserList(sndIndex).Flags.UserLogged = 0 Then Exit Sub
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
                If UserList(LoopC).Flags.UserLogged Then
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
                tIndex = MapUsers(sndMap).Index(LoopC)
                If UserList(tIndex).Flags.UserLogged Then
                    If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                    CopyPos = UserList(tIndex).BufferSize + 1
                    ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                    CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                    UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                    Data_Send_Update tIndex, Priority
                End If
            Next LoopC
    
            'Send to PC Area
        Case ToPCArea
            Log "Send ToPCArea(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = UserList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    tIndex = MapUsers(sndMap).Index(LoopC)
                    If UserList(tIndex).Flags.UserLogged Then
                        If Server_RectDistance(UserList(sndIndex).Pos.X, UserList(sndIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(tIndex).BufferSize + 1
                            ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                            CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                            Data_Send_Update tIndex, Priority
                        End If
                    End If
                Next LoopC
            End If
            
            'Send to PC Area but the sndIndex
        Case ToPCAreaButIndex
            Log "Send ToPCAreaButIndex(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = UserList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    tIndex = MapUsers(sndMap).Index(LoopC)
                    If tIndex <> sndIndex Then
                        If UserList(tIndex).Flags.UserLogged Then
                            If Server_RectDistance(UserList(sndIndex).Pos.X, UserList(sndIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                                If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                                CopyPos = UserList(tIndex).BufferSize + 1
                                ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                                CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                                UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                                Data_Send_Update tIndex, Priority
                            End If
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
                    tIndex = MapUsers(sndMap).Index(LoopC)
                    If tIndex <> sndIndex Then
                        If UserList(tIndex).Flags.UserLogged Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(tIndex).BufferSize + 1
                            ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                            CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                            Data_Send_Update tIndex, Priority
                        End If
                    End If
                Next LoopC
            End If
            
            'Send to all users in a group on the specified map (except the index)
        Case ToMapGroupButIndex
            Log "Send ToMapGroupButIndex(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            If UserList(sndIndex).GroupIndex > 0 Then
                For LoopC = 1 To GroupData(UserList(sndIndex).GroupIndex).NumUsers
                    tIndex = GroupData(UserList(sndIndex).GroupIndex).Users(LoopC)
                    If tIndex <> LoopC Then
                        If UserList(tIndex).Pos.Map = UserList(sndIndex).Pos.Map Then
                            If UserList(tIndex).Flags.UserLogged Then
                                If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                                CopyPos = UserList(tIndex).BufferSize + 1
                                ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                                CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                                UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                                Data_Send_Update tIndex, Priority
                            End If
                        End If
                    End If
                Next LoopC
            End If
            
            'Send to all users in a group (except the index)
        Case ToGroupButIndex
            Log "Send ToGroupButIndex(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            If UserList(sndIndex).GroupIndex > 0 Then
                For LoopC = 1 To GroupData(UserList(sndIndex).GroupIndex).NumUsers
                    tIndex = GroupData(UserList(sndIndex).GroupIndex).Users(LoopC)
                    If tIndex <> LoopC Then
                        If UserList(tIndex).Flags.UserLogged Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(tIndex).BufferSize + 1
                            ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                            CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                            Data_Send_Update tIndex, Priority
                        End If
                    End If
                Next LoopC
            End If
            
            'Send to all users in a group
        Case ToGroup
            Log "Send ToGroup(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            If UserList(sndIndex).GroupIndex > 0 Then
                For LoopC = 1 To GroupData(UserList(sndIndex).GroupIndex).NumUsers
                    tIndex = GroupData(UserList(sndIndex).GroupIndex).Users(LoopC)
                    If UserList(tIndex).Flags.UserLogged Then
                        If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                        CopyPos = UserList(tIndex).BufferSize + 1
                        ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                        CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                        UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                        Data_Send_Update tIndex, Priority
                    End If
                Next LoopC
            End If
    
            'Send to NPC Area
        Case ToNPCArea
            Log "Send ToNPCArea(" & sndIndex & "): " & ByteArrayToStr(sndData), PacketOut '//\\LOGLINE//\\
            sndMap = NPCList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    tIndex = MapUsers(sndMap).Index(LoopC)
                    If UserList(tIndex).Flags.UserLogged Then
                        If Server_RectDistance(NPCList(sndIndex).Pos.X, NPCList(sndIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(tIndex).BufferSize + 1
                            ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                            CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                            Data_Send_Update tIndex, Priority
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
                    tIndex = MapUsers(sndMap).Index(LoopC)
                    If tIndex <> sndIndex Then   'Dont send to the user who moved
                        If UserList(tIndex).Flags.UserLogged Then
                            If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                            CopyPos = UserList(tIndex).BufferSize + 1
                            ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                            CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                            UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                            If UserList(tIndex).GroupIndex = UserList(sndIndex).GroupIndex Then
                                Data_Send_Update tIndex, PP_GroupCharMove
                            Else
                                If Server_RectDistance(UserList(sndIndex).Pos.X, UserList(sndIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                                    Data_Send_Update tIndex, PP_CloseCharMove
                                Else
                                    Data_Send_Update tIndex, PP_FarCharMove
                                End If
                            End If
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
                    tIndex = MapUsers(sndMap).Index(LoopC)
                    If UserList(tIndex).Flags.UserLogged Then
                        If DEBUG_RecordPacketsOut Then DebugPacketsOut(sndData(0)) = DebugPacketsOut(sndData(0)) + 1
                        CopyPos = UserList(tIndex).BufferSize + 1
                        ReDim Preserve UserList(tIndex).SendBuffer(UserList(tIndex).BufferSize + CopySize)
                        CopyMemory UserList(tIndex).SendBuffer(CopyPos), sndData(0), CopySize
                        UserList(tIndex).BufferSize = UserList(tIndex).BufferSize + CopySize
                        If Server_RectDistance(NPCList(sndIndex).Pos.X, NPCList(sndIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
                            Data_Send_Update tIndex, PP_CloseCharMove
                        Else
                            Data_Send_Update tIndex, PP_FarCharMove
                        End If
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

    If frmMain.GOREsock.State(UserIndex) = soxIdle Or soxListening Then
        If UserList(UserIndex).BufferSize > -1 Then
            
            Log "Send Buffer: " & ByteArrayToStr(UserList(UserIndex).SendBuffer), PacketOut '//\\LOGLINE//\\

            'Send the packet
            DoEvents
            frmMain.GOREsock.SendData UserIndex, UserList(UserIndex).SendBuffer()

            'Calculate bandwidth usage
            'TCP header = 20 bytes, IPv4 header = 20 bytes
            If CalcTraffic Then DataOut = DataOut + UserList(UserIndex).BufferSize + 41

            'Clear the buffer
            Erase UserList(UserIndex).SendBuffer()
            UserList(UserIndex).HasBuffer = 0
            UserList(UserIndex).BufferSize = -1
            UserList(UserIndex).PPCount = 0
            UserList(UserIndex).PacketWait = 0

        End If
    End If

    'Check if the user was disconnected
    If frmMain.GOREsock.State(UserIndex) = 1 Then
        Server_CloseSocket UserIndex
    End If

    'Check if we are disconnecting
    If UserList(UserIndex).Flags.Disconnecting = 1 Then
        Server_CloseSocket UserIndex
    End If

End Sub

Sub Data_Server_Help(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to retrieve help buffer from server
'<>
'*****************************************************************

    Log "Call Data_Server_Help(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Send the help buffer
    Data_Send ToIndex, UserIndex, HelpBuffer

End Sub

Sub Data_Server_MailCompose_ServerToServer(ByRef rBuf As DataBuffer)

'*****************************************************************
'Compose a new message (from another server)
'<ReceiverName(S)><WriterName(S)><Subject(S)><Message(B)> (<ObjIndex(I)><Amount(I)>)
'*****************************************************************
Dim ReceiverName As String
Dim NumObjs As Byte
Dim tMail As MailData
Dim i As Byte

    'Get the variables
    ReceiverName = rBuf.Get_String
    tMail.WriterName = rBuf.Get_String
    tMail.Subject = rBuf.Get_String
    tMail.Message = rBuf.Get_StringEX
    NumObjs = rBuf.Get_Byte
    For i = 1 To NumObjs
        tMail.Obj(i).ObjIndex = rBuf.Get_Integer
        tMail.Obj(i).Amount = rBuf.Get_Integer
    Next i
    
    'Send the message to the mail handler
    Server_WriteMail 0, ReceiverName, tMail.Subject, tMail.Message, tMail.Obj

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
    If UserList(UserIndex).Counters.DelayTimeMail > timeGetTime Then
        Log "Data_Server_MailCompose: Not enough time elapsed since last mail write for user " & UserIndex & " (" & UserList(UserIndex).Name & ")", CodeTracker '//\\LOGLINE//\\
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

    'Check for invalid values (most is handled in Server_WriteMail sub)
    If Server_ValidString(RecieverName) Then
        If Server_ValidString(Subject) Then
            If Server_ValidString(Message) Then
        
                'Write the mail to the RecieverName
                If Server_WriteMail(UserIndex, RecieverName, Subject, Message, Objs) = 0 Then Exit Sub
                
            End If
        End If
    End If
    
    'Remove the items from the user
    For i = 1 To NumObjs
        With UserList(UserIndex).Object(Slots(i))
            .Amount = .Amount - Objs(i).Amount
            If .Amount <= 0 Then
        
                'Unequip if the object is currently equipped
                If .Equipped = 1 Then User_RemoveInvItem UserIndex, Slots(i), False
                
                'Clear the values
                .ObjIndex = 0
                .Amount = 0
                .Equipped = 0
                
            End If
            
            'Update the inventory slot
            User_UpdateInv False, UserIndex, Slots(i)
            
        End With
    Next i

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
    ConBuf.PreAllocate 2
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
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_Mail

End Sub

Sub Data_Server_MailItemTake(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Take an item out of the mail
'<ObjIndex(B)>
'*****************************************************************
Dim AmountTaken As Integer
Dim ObjIndex As Byte
Dim MsgData As MailData
Dim X As Long
Dim Y As Long

    Log "Call Data_Server_MailItemTake([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    ObjIndex = rBuf.Get_Byte

    'Check for invalid values
    If UserList(UserIndex).Flags.LastViewedMail <= 0 Then Exit Sub
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
        If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).MailboxPos.X, UserList(UserIndex).MailboxPos.Y, 1, 1) = 0 Then
            Data_Send ToIndex, UserIndex, cMessage(57).Data
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    'Get the item information and check if theres an item
    MsgData = Load_Mail(UserList(UserIndex).MailID(UserList(UserIndex).Flags.LastViewedMail))
    If MsgData.Obj(ObjIndex).ObjIndex <= 0 Then
        Log "Data_Server_MailItemTake: Invalid object index (" & MsgData.Obj(ObjIndex).ObjIndex & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    If MsgData.Obj(ObjIndex).Amount <= 0 Then
        Log "Data_Server_MailItemTake: Invalid object amount (" & MsgData.Obj(ObjIndex).Amount & ")", InvalidPacketData '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Give the user the objects
    AmountTaken = User_GiveObj(UserIndex, MsgData.Obj(ObjIndex).ObjIndex, MsgData.Obj(ObjIndex).Amount)
    If AmountTaken = 0 Then Exit Sub
    MsgData.Obj(ObjIndex).Amount = MsgData.Obj(ObjIndex).Amount - AmountTaken
    
    'Get the number of objects
    Y = 0
    For X = 1 To MaxMailObjs
        If MsgData.Obj(X).Amount > 0 Then Y = Y + 1
    Next X
        
    'Update the mail objects
    ConBuf.PreAllocate 2
    ConBuf.Put_Byte DataCode.Server_MailObjUpdate
    ConBuf.Put_Byte Y   'Number of objects
    For X = 1 To MaxMailObjs
        If MsgData.Obj(X).Amount > 0 Then
            ConBuf.Allocate 7 + Len(ObjData.Name(MsgData.Obj(X).ObjIndex))
            ConBuf.Put_Long ObjData.GrhIndex(MsgData.Obj(X).ObjIndex)
            ConBuf.Put_String ObjData.Name(MsgData.Obj(X).ObjIndex)
            ConBuf.Put_Integer MsgData.Obj(X).Amount
        End If
    Next X
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_Mail
    
    'Save the changes done to the mail
    Save_Mail UserList(UserIndex).MailID(UserList(UserIndex).Flags.LastViewedMail), MsgData

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

    'Load the mail into a temporary mail buffer
    MsgData = Load_Mail(UserList(UserIndex).MailID(MessageID))

    'Get the number of objects
    NumObjs = 0
    For LoopC = 1 To MaxMailObjs
        If MsgData.Obj(LoopC).ObjIndex > 0 Then NumObjs = NumObjs + 1
    Next LoopC

    'Send the message information
    ConBuf.PreAllocate 6 + Len(MsgData.Message) + Len(MsgData.Subject) + Len(MsgData.WriterName)
    ConBuf.Put_Byte DataCode.Server_MailMessage
    ConBuf.Put_StringEX MsgData.Message
    ConBuf.Put_String MsgData.Subject
    ConBuf.Put_String MsgData.WriterName
    ConBuf.Put_Byte NumObjs
    For LoopC = 1 To MaxMailObjs
        If MsgData.Obj(LoopC).ObjIndex > 0 Then
            ConBuf.Allocate 7 + Len(ObjData.Name(MsgData.Obj(LoopC).ObjIndex))
            ConBuf.Put_Long ObjData.GrhIndex(MsgData.Obj(LoopC).ObjIndex)
            ConBuf.Put_String ObjData.Name(MsgData.Obj(LoopC).ObjIndex)
            ConBuf.Put_Integer MsgData.Obj(LoopC).Amount
        End If
    Next LoopC
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_Mail

    'Remember the last message the user viewed
    UserList(UserIndex).Flags.LastViewedMail = MessageID

    'If the mail was new, then make it old (and save it)
    If MsgData.New = 1 Then
        MsgData.New = 0
        Save_Mail UserList(UserIndex).MailID(MessageID), MsgData
    End If

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
        If LenB(UserList(LoopC).Name) <> 0 Then
            If UserList(LoopC).Flags.UserLogged Then
                intNumUsers = intNumUsers + 1
                ReDim Preserve UserNames(1 To intNumUsers)
                UserNames(intNumUsers) = UserList(LoopC).Name
            End If
        End If
    Next LoopC

    'Set the first line (total users)
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 78
    ConBuf.Put_Integer intNumUsers

    'Send the user names
    For i = 1 To (intNumUsers \ 10) + 1 'Replace the 10's with the number of users per line
        tStr = vbNullString             'Clear the string buffer
        For LoopC = 1 To 10
            If LoopC + ((i - 1) * 10) > intNumUsers Then Exit For
            tStr = tStr & UserNames(LoopC + ((i - 1) * 10)) & ", "
        Next LoopC
        tStr = Left$(tStr, Len(tStr) - 2)   'Crop off the last comma
        ConBuf.Allocate 3 + Len(tStr)
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String tStr
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Next i

    'Send all the lines as a whole
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_GM_Warp(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM warps to a new location
'<X(B)><Y(B)>
'*****************************************************************
Dim nPos As WorldPos
Dim tPos As WorldPos

    'Get the location
    nPos.Map = UserList(UserIndex).Pos.Map
    nPos.X = rBuf.Get_Byte
    nPos.Y = rBuf.Get_Byte
    
    'Confirm that the user is a GM
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
    
    'Check if the location is valid - if not, get the closest legal position
    Server_ClosestLegalPos nPos, tPos
    
    'Warp the user to the new position if it is valid
    If Server_LegalPos(tPos.Map, tPos.X, tPos.Y, 0) Then
        User_WarpChar UserIndex, tPos.Map, tPos.X, tPos.Y
    End If

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
    If UserList(UserIndex).Flags.UserLogged = 0 Then Exit Sub
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
    ConBuf.PreAllocate 3
    ConBuf.Put_Byte DataCode.User_Blink
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_Blink

End Sub

Sub Data_User_CastSkill(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User casts a skill
'<SkillID(B)><TargetIndex(I)><Heading(B)>
'*****************************************************************
Dim TargetCharIndex As Integer
Dim TargetIndex As Integer
Dim TargetType As Byte
Dim Heading As Byte
Dim SkillID As Byte

    Log "Call Data_User_CastSkill([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    SkillID = rBuf.Get_Byte

    'Easiest to always pass a target, even if a target can not be selected for the skill
    TargetCharIndex = rBuf.Get_Integer
    
    'Same with the heading
    Heading = rBuf.Get_Byte
    
    'Check for a valid char index
    If Heading > 8 Then Heading = 8
    If Heading < 1 Then Heading = 1
    If TargetCharIndex < 0 Then TargetCharIndex = 0
    If TargetCharIndex > LastChar Then TargetCharIndex = 0
    If CharList(TargetCharIndex).Index <= 0 Then TargetCharIndex = 0
    
    'Check if the skill can be used by the user's class
    If Not Skill_ValidSkillForClass(UserList(UserIndex).Class, SkillID) Then Exit Sub
    
    'Clear the pending quest NPC
    UserList(UserIndex).Flags.QuestNPC = 0
    
    'Turn TargetCharIndex into the TargetIndex - get the array index, not char index
    If TargetCharIndex = 0 Then
        TargetIndex = UserIndex
        TargetType = CharType_PC
        UserList(UserIndex).Flags.TargetIndex = 0
        UserList(UserIndex).Flags.Target = 0
    Else
        TargetIndex = CharList(TargetCharIndex).Index
        'Since PCs and NPCs are in different arrays, we have to pass which one the TargetIndex belongs to, too
        TargetType = CharList(TargetCharIndex).CharType
    End If
    
    'Set the heading
    If UserList(UserIndex).Char.Heading <> Heading Then
        UserList(UserIndex).Char.Heading = Heading
        UserList(UserIndex).Char.HeadHeading = Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte Heading
        Data_Send ToPCAreaButIndex, UserIndex, ConBuf.Get_Buffer
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
            Else
                Skill_Heal_PCtoPC UserIndex, UserIndex
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
            
        Case SkID.SummonBandit
            Skill_SummonBandit_PC UserIndex
            
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
        Data_Send ToIndex, UserIndex, cMessage(60).Data
        'Tell the user their description has been removed
    Else
        Data_Send ToIndex, UserIndex, cMessage(61).Data
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
    ConBuf.PreAllocate 4
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

Dim Name As String
Dim Pass As String

    Log "Call Data_User_Login([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Name = rBuf.Get_String
    Pass = rBuf.Get_String

    'Invalid values handled by sub
    User_Connect UserIndex, Name, Pass

End Sub

Sub Data_User_LookLeft(ByVal UserIndex As Integer)

'*****************************************************************
'User looked left
'<>
'*****************************************************************

    Log "Call Data_User_LookLeft(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    'Check for invalid values
    If UserList(UserIndex).Char.HeadHeading = 0 Then Exit Sub
    If UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.Heading - 1 Then Exit Sub

    'Turn the head
    UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.HeadHeading - 1
    If UserList(UserIndex).Char.HeadHeading = 0 Then UserList(UserIndex).Char.HeadHeading = 4

    'Send the update to the map
    ConBuf.PreAllocate 4
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
    If UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.Heading + 1 Then Exit Sub

    'Turn the head
    UserList(UserIndex).Char.HeadHeading = UserList(UserIndex).Char.HeadHeading + 1
    If UserList(UserIndex).Char.HeadHeading = 5 Then UserList(UserIndex).Char.HeadHeading = 1

    'Send the update to the map
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_LookLeft
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(UserIndex).Char.HeadHeading
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_Look

End Sub

Sub Data_Server_SetUserPosition(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User requests to make sure their position is correct
'<X(B)<Y(B)>
'*****************************************************************
Dim nX As Byte
Dim nY As Byte

    'Get the client's position
    nX = rBuf.Get_Byte
    nY = rBuf.Get_Byte
    
    'Check the values against the values on the server
    If nX <> UserList(UserIndex).Pos.X Or nY <> UserList(UserIndex).Pos.Y Then
    
        'The position the client holds is incorrect, send the correction
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
    End If

End Sub

Sub Data_User_Bank_Withdraw(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User wants to take gold out of their bank account
'<Gold(L)>
'*****************************************************************
Dim Gold As Long

    'Get the value
    Gold = rBuf.Get_Long
    
    'Check for a valid value
    If Gold <= 0 Then Exit Sub
    
    'Confirm that the user is near a banking NPC
    If User_NearBankNPC(UserIndex) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(121).Data()
        Exit Sub
    End If
    
    'Make sure the user can withdraw this much
    If UserList(UserIndex).BankGold - Gold < 0 Then
    
        'User is trying to take out too much gold, give error message
        Data_Send ToIndex, UserIndex, cMessage(119).Data()
        Exit Sub
        
    End If
    
    'Withdraw the money
    UserList(UserIndex).BankGold = UserList(UserIndex).BankGold - Gold
    UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + Gold
    
    'Give the success message
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 120
    ConBuf.Put_Long Gold
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_Bank_Deposit(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User wants to put gold in their bank account
'<Gold(L)>
'*****************************************************************
Dim Gold As Long

    'Get the value
    Gold = rBuf.Get_Long
    
    'Check for a valid value
    If Gold <= 0 Then Exit Sub
    
    'Confirm the user has enough gold
    If UserList(UserIndex).Stats.BaseStat(SID.Gold) < Gold Then Exit Sub
    
    'Confirm that the user is near a banking NPC
    If User_NearBankNPC(UserIndex) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(121).Data()
        Exit Sub
    End If
    
    'Put the money in the bank
    UserList(UserIndex).BankGold = UserList(UserIndex).BankGold + Gold
    UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) - Gold

End Sub

Sub Data_User_Bank_Balance(ByVal UserIndex As Integer)

'*****************************************************************
'User requests to know their bank balance
'<>
'*****************************************************************

    'Confirm that the user is near a banking NPC
    If User_NearBankNPC(UserIndex) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(121).Data()
        Exit Sub
    End If

    'Send the balance to the user
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 117
    ConBuf.Put_Long UserList(UserIndex).BankGold
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_Move(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Move the user
'<Direction(B)>
'*****************************************************************
Dim Dir As Byte
Dim Running As Byte

    Log "Call Data_User_Move([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Dir = rBuf.Get_Byte
    
    'Check the move counter
    If 1 = 2 Then   '//TEMP
    If timeGetTime < UserList(UserIndex).Counters.MoveCounter Then
        Log "User_MoveChar: Not enough time has elapsed for movement.", CodeTracker '//\\LOGLINE//\\

        'If the user is moving too fast, then put them back
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_SetUserPosition
        ConBuf.Put_Byte UserList(UserIndex).Pos.X
        ConBuf.Put_Byte UserList(UserIndex).Pos.Y
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub

    End If
    End If
    
    'Check if running
    If Dir > 128 Then
        Running = 1
        Dir = Dir Xor 128
    End If

    'Invalid values handled by sub
    User_MoveChar UserIndex, Dir, Running

End Sub

Sub Data_User_Group_Info(ByVal UserIndex As Integer)

'*****************************************************************
'User requests the information of their group
'<>
'*****************************************************************
Dim s As String
Dim i As Byte

    'Check if the user is in a group
    If UserList(UserIndex).GroupIndex = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(123).Data()
        Exit Sub
    End If
    
    'Create the list to send to the user
    For i = 1 To GroupData(UserList(UserIndex).GroupIndex).NumUsers
        If LenB(s) = 0 Then
            s = s & UserList(GroupData(UserList(UserIndex).GroupIndex).Users(i)).Name & "(" & UserList(GroupData(UserList(UserIndex).GroupIndex).Users(i)).Stats.BaseStat(SID.ELV) & ")"
        Else
            s = s & vbNewLine & UserList(GroupData(UserList(UserIndex).GroupIndex).Users(i)).Name & "(" & UserList(GroupData(UserList(UserIndex).GroupIndex).Users(i)).Stats.BaseStat(SID.ELV) & ")"
        End If
    Next i
    
    'Create the group text packet
    If LenB(s) Then
        ConBuf.PreAllocate Len(s) + 3
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String s
        ConBuf.Put_Byte DataCode.Comm_FontType_Group
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    End If
    
End Sub

Sub Data_User_Group_Join(ByVal UserIndex As Integer)

'*****************************************************************
'User joins a group they were invited to
'<>
'*****************************************************************

    'Check for invalid values
    If UserList(UserIndex).GroupIndex > 0 Then Exit Sub
    If UserList(UserIndex).Counters.GroupCounter < timeGetTime Then
        UserList(UserIndex).Counters.GroupCounter = 0
        UserList(UserIndex).Flags.InviteGroup = 0
        Exit Sub
    End If
    If UserList(UserIndex).Flags.InviteGroup = 0 Then
        UserList(UserIndex).Counters.GroupCounter = 0
        UserList(UserIndex).Flags.InviteGroup = 0
        Exit Sub
    End If
    
    'Join the group
    Group_AddUser UserIndex, UserList(UserIndex).Flags.InviteGroup

End Sub

Sub Data_User_NewLogin(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Log on a new character
'<Name(S)><Password(S)><Head(I)><Body(I)><Class(B)>
'*****************************************************************

Dim Name As String
Dim Pass As String
Dim Head As Integer
Dim Body As Integer
Dim Class As Byte

    Log "Call Data_User_NewLogin([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Name = rBuf.Get_String
    Pass = rBuf.Get_String
    Head = rBuf.Get_Integer
    Body = rBuf.Get_Integer
    Class = rBuf.Get_Byte

    'Invalid values handled by sub
    User_ConnectNew UserIndex, Name, Pass, Body, Head, Class

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
    ConBuf.PreAllocate 4
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
Dim i As Long
Dim s As String

    Log "Call Data_User_StartQuest(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

'Make sure the user has a quests pending

    If UserList(UserIndex).Flags.QuestNPC = 0 Then Exit Sub

    'Check the distance of the user to the NPC
    QuestNPC = UserList(UserIndex).Flags.QuestNPC
    If NPCList(QuestNPC).Pos.Map = UserList(UserIndex).Pos.Map Then
        If Server_RectDistance(NPCList(QuestNPC).Pos.X, NPCList(QuestNPC).Pos.Y, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, MaxServerDistanceX + 2, MaxServerDistanceY + 2) Then 'Add a little bit extra to distance, in case NPC moves

            'Check the requirements
            If QuestData(NPCList(QuestNPC).Quest).AcceptReqLvl > UserList(UserIndex).Stats.BaseStat(SID.ELV) Then
                ConBuf.PreAllocate 6
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 62
                ConBuf.Put_Long QuestData(NPCList(QuestNPC).Quest).AcceptReqLvl
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                UserList(UserIndex).Flags.QuestNPC = 0
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
                    ConBuf.PreAllocate 5 + Len(ObjData.Name(QuestData(NPCList(QuestNPC).Quest).AcceptReqObj))
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 63
                    ConBuf.Put_Integer QuestData(NPCList(QuestNPC).Quest).AcceptReqObjAmount
                    ConBuf.Put_String ObjData.Name(QuestData(NPCList(QuestNPC).Quest).AcceptReqObj)
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer()
                    UserList(UserIndex).Flags.QuestNPC = 0
                    Exit Sub
                End If
            End If
            If QuestData(NPCList(QuestNPC).Quest).AcceptReqFinishQuest > 0 Then
                For i = 1 To UserList(UserIndex).NumCompletedQuests
                    If UserList(UserIndex).CompletedQuests(i) = QuestData(NPCList(QuestNPC).Quest).AcceptReqFinishQuest Then
                        i = 0   'Lets us know we found the required quest
                        Exit For
                    End If
                Next i
                If i > 0 Then
                    ConBuf.PreAllocate 3 + Len(QuestData(QuestData(NPCList(QuestNPC).Quest).AcceptReqFinishQuest).Name)
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 134
                    ConBuf.Put_String QuestData(QuestData(NPCList(QuestNPC).Quest).AcceptReqFinishQuest).Name
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer()
                    UserList(UserIndex).Flags.QuestNPC = 0
                    Exit Sub
                End If
            End If

            'Make sure the user has room in the quest queue
            i = 0
            Do
                i = i + 1

                'Uhoh, the user has no more quest room!
                If i > MaxQuests Then
                    Data_Send ToIndex, UserIndex, cMessage(12).Data
                    UserList(UserIndex).Flags.QuestNPC = 0
                    Exit Sub
                End If

            Loop While UserList(UserIndex).Quest(i) <> 0

            'We made it out safely, so that means we got an open slot, horray! (>^_^)>
            UserList(UserIndex).Quest(i) = NPCList(QuestNPC).Quest
            
            'Send the requirements
            Quest_SendReqString UserIndex, QuestNPC

            'Give the user the quest starting stuff
            If QuestData(UserList(UserIndex).Quest(i)).AcceptRewExp > 0 Then
                User_RaiseExp UserIndex, QuestData(UserList(UserIndex).Quest(i)).AcceptRewExp
                ConBuf.PreAllocate 6
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 3
                ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(i)).AcceptRewExp
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            If QuestData(UserList(UserIndex).Quest(i)).AcceptRewGold > 0 Then
                UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + QuestData(UserList(UserIndex).Quest(i)).AcceptRewGold
                ConBuf.PreAllocate 6
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 4
                ConBuf.Put_Long QuestData(UserList(UserIndex).Quest(i)).AcceptRewGold
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            End If
            
            'Send the quest text to the client
            Quest_SendText UserIndex, i

            'Starting reward skills
            If QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill > 0 Then
                User_GiveSkill UserIndex, QuestData(UserList(UserIndex).Quest(i)).AcceptLearnSkill
            End If

            'Check for an object reward
            If QuestData(UserList(UserIndex).Quest(i)).AcceptRewObj > 0 Then
                User_GiveObj UserIndex, QuestData(UserList(UserIndex).Quest(i)).AcceptRewObj, QuestData(UserList(UserIndex).Quest(i)).AcceptRewObjAmount
            End If

            Exit Sub

        End If
    End If

    'The user is not close enough to the NPC
    Data_Send ToIndex, UserIndex, cMessage(64).Data

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
    NPCIndex = UserList(UserIndex).Flags.TradeWithNPC

    'Check for invalid values
    If NPCIndex < 1 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If NPCList(NPCIndex).NumVendItems < 1 Then Exit Sub
    If Slot <= 0 Then Exit Sub
    If Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    If UserList(UserIndex).Flags.TradeWithNPC <= 0 Then Exit Sub
    If Amount < 0 Then Exit Sub
    If Amount > UserList(UserIndex).Object(Slot).Amount Then Exit Sub
    If UserList(UserIndex).Object(Slot).ObjIndex < 1 Then Exit Sub
    
    'Check for valid locations
    If UserList(UserIndex).Pos.Map <> NPCList(NPCIndex).Pos.Map Then
        Data_Send ToIndex, UserIndex, cMessage(36).Data
        Exit Sub
    End If
    If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, 6, 6) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(36).Data
        Exit Sub
    End If
    
    With UserList(UserIndex).Object(Slot)
    
        'Give the money
        AddMoney = ObjData.Value(.ObjIndex) * 0.5 * Amount
        UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) + AddMoney
        ConBuf.PreAllocate 9 + Len(ObjData.Name(.ObjIndex))
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 96
        ConBuf.Put_Integer Amount
        ConBuf.Put_String ObjData.Name(.ObjIndex)
        ConBuf.Put_Long AddMoney
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

Sub Data_User_Trade_Accept(ByVal UserIndex As Integer)

'*****************************************************************
'User accepts the trade
'<>
'*****************************************************************

    TradeTable_Accept UserIndex

End Sub

Sub Data_User_Trade_Finish(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to finish the trade
'<>
'*****************************************************************

    TradeTable_RequestFinish UserIndex

End Sub

Sub Data_User_Trade_BuyFromNPC(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Buy an item from NPC
'<Slot(B)><Amount(I)>
'*****************************************************************
Dim AmountBought As Integer
Dim PurchaseObj As Integer
Dim NPCIndex As Integer
Dim UserSlot As Integer
Dim Amount As Integer
Dim Slot As Byte

    Log "Call Data_User_Trade_BuyFromNPC([" & ByteArrayToStr(rBuf.Get_Buffer) & "]," & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    Slot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer

    'Set the NPC index to trade with
    NPCIndex = UserList(UserIndex).Flags.TradeWithNPC

    'Check for invalid values
    If NPCIndex <= 0 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If Slot > NPCList(NPCIndex).NumVendItems Then Exit Sub
    If Slot <= 0 Then Exit Sub
    If UserList(UserIndex).Flags.TradeWithNPC <= 0 Then Exit Sub
    If NPCList(NPCIndex).VendItems(Slot).ObjIndex <= 0 Then Exit Sub
    If NPCList(NPCIndex).VendItems(Slot).Amount = 0 Then Exit Sub
    If Amount < 0 Then Exit Sub

    'Check for valid locations
    If UserList(UserIndex).Pos.Map <> NPCList(NPCIndex).Pos.Map Then
        Data_Send ToIndex, UserIndex, cMessage(36).Data
        Exit Sub
    End If
    If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, 6, 6) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(36).Data
        Exit Sub
    End If

    'Set the ObjData of the item to be purchased to the PurchaseObj variable
    PurchaseObj = NPCList(NPCIndex).VendItems(Slot).ObjIndex
    
    'Check that the user has enough money
    If UserList(UserIndex).Stats.BaseStat(SID.Gold) < ObjData.Value(PurchaseObj) * Amount Then
        ConBuf.PreAllocate 5 + Len(ObjData.Name(PurchaseObj))
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 65
        ConBuf.Put_Integer Amount
        ConBuf.Put_String ObjData.Name(PurchaseObj)
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Reduce the amount of items the NPC has
    If NPCList(NPCIndex).VendItems(Slot).Amount <> -1 Then

        'Check if there is enough
        If NPCList(NPCIndex).VendItems(Slot).Amount - Amount < 0 Then
            ConBuf.PreAllocate 7 + Len(ObjData.Name(PurchaseObj))
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 66
            ConBuf.Put_Integer Amount
            ConBuf.Put_String ObjData.Name(PurchaseObj)
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
    
    'Give the user the objects
    AmountBought = User_GiveObj(UserIndex, PurchaseObj, Amount, True)

    'Take the user's money
    UserList(UserIndex).Stats.BaseStat(SID.Gold) = UserList(UserIndex).Stats.BaseStat(SID.Gold) - (ObjData.Value(PurchaseObj) * AmountBought)

    'Send the purchase message
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 67
    ConBuf.Put_Integer Amount
    ConBuf.Put_String ObjData.Name(PurchaseObj)
    ConBuf.Put_Long (ObjData.Value(PurchaseObj) * Amount)
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

End Sub

Sub Data_User_Group_Make(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to create a new group
'<>
'*****************************************************************
Dim i As Byte

    'Make sure the user isn't already in a group
    If UserList(UserIndex).GroupIndex > 0 Then Exit Sub
    
    'Create the group
    i = Group_Create(UserIndex)
    If i > 0 Then
        Data_Send ToIndex, UserIndex, cMessage(116).Data()
        GroupData(i).NumUsers = 1
        ReDim GroupData(i).Users(1 To 1)
        GroupData(i).Users(1) = UserIndex
        UserList(UserIndex).GroupIndex = i
    End If

End Sub

Sub Data_User_CancelQuest(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User cancels a quest
'<QuestSlot(B)>
'*****************************************************************
Dim QuestSlot As Byte

    QuestSlot = rBuf.Get_Byte
    
    'Check for a valid slot
    Quest_Cancel UserIndex, QuestSlot

End Sub

Sub Data_User_Group_Leave(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to leave their group
'<>
'*****************************************************************

    'Leave the group
    Group_RemoveUser UserIndex, UserList(UserIndex).GroupIndex
    
End Sub

Sub Data_User_Target(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Client wants to target a character
'<CharIndex(I)>
'*****************************************************************
Dim CharIndex As Integer
Dim Index As Integer

    'Get the character index
    CharIndex = rBuf.Get_Integer
    
    'Check for a valid character index
    If CharIndex < 1 Then GoTo Deselect
    If CharIndex > LastChar Then GoTo Deselect
    
    'Get the index and make sure it is valid
    Index = CharList(CharIndex).Index
    If Index = 0 Then Exit Sub
    
    'Check if the characer is a NPC or PC
    If CharList(CharIndex).CharType = CharType_PC Then
    
        'Make sure the index is a valid PC
        If Index > LastUser Then Exit Sub
        If UserList(Index).Flags.Disconnecting = 1 Then Exit Sub
        If UserList(Index).Flags.UserLogged = 0 Then Exit Sub

        'Check the distance from the user
        If UserList(UserIndex).Pos.Map = UserList(Index).Pos.Map Then
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(Index).Pos.X, UserList(Index).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
    
                'Valid distance, set as the new target
                UserList(UserIndex).Flags.TargetIndex = UserList(Index).Char.CharIndex
                UserList(UserIndex).Flags.Target = CharType_PC
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer UserList(Index).Char.CharIndex
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            
            End If
        End If
        
    'The character is a NPC
    ElseIf CharList(CharIndex).CharType = CharType_NPC Then
    
        'Make sure the index is a valid NPC
        If Index > LastNPC Then Exit Sub
        If NPCList(Index).Flags.NPCAlive = 0 Then Exit Sub
        
        'Check the distance from the user
        If UserList(UserIndex).Pos.Map = NPCList(Index).Pos.Map Then
            If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, NPCList(Index).Pos.X, NPCList(Index).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
    
                'Valid distance, set as the new target
                UserList(UserIndex).Flags.TargetIndex = NPCList(Index).Char.CharIndex
                UserList(UserIndex).Flags.Target = CharType_NPC
                ConBuf.PreAllocate 3
                ConBuf.Put_Byte DataCode.User_Target
                ConBuf.Put_Integer NPCList(Index).Char.CharIndex
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Exit Sub
            
            End If
        End If
        
    End If
    
Deselect:
    
    'The target wasn't valid, check if the user has a target already
    If UserList(UserIndex).Flags.Target > 0 Then
        
        'Clear the user's target
        UserList(UserIndex).Flags.Target = 0
        UserList(UserIndex).Flags.TargetIndex = 0
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.User_Target
        ConBuf.Put_Integer 0
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
    End If

End Sub

Sub Data_User_Group_Invite(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Invite a user to part of a group
'<Name(S)>
'*****************************************************************
Dim tName As String
Dim tIndex As Integer

    'Get the name
    tName = rBuf.Get_String
    
    'Make sure the user is in a group
    If UserList(UserIndex).GroupIndex = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(113).Data()
        Exit Sub
    End If
        
    'Make sure the user is the leader of the group
    If GroupData(UserList(UserIndex).GroupIndex).Users(1) <> UserIndex Then
        Data_Send ToIndex, UserIndex, cMessage(113).Data()
        Exit Sub
    End If
    
    'Make sure the group isn't full
    If GroupData(UserList(UserIndex).GroupIndex).NumUsers >= Group_MaxUsers Then
        Data_Send ToIndex, UserIndex, cMessage(112).Data()
        Exit Sub
    End If
    
    'Get the user's index
    tIndex = User_NameToIndex(tName)
    
    'Make sure the user is online
    If tIndex = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(51).Data()
        Exit Sub
    End If
    
    'Make sure the user isn't already part of a group
    If UserList(tIndex).GroupIndex > 0 Then
        Data_Send ToIndex, UserIndex, cMessage(114).Data()
        Exit Sub
    End If
        
    'Tell the user the invite message
    ConBuf.PreAllocate 4 + Len(UserList(UserIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 115
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_Integer (Group_InviteTime \ 1000)
    Data_Send ToIndex, tIndex, ConBuf.Get_Buffer
    UserList(tIndex).Flags.InviteGroup = UserList(UserIndex).GroupIndex
    UserList(tIndex).Counters.GroupCounter = timeGetTime + Group_InviteTime
    
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
        If UserIndex <= LastUser Then
            
            UserList(UserIndex).Flags.Disconnecting = 0
    
            'Send the buffer
            Data_Send_Buffer UserIndex
    
            'Shut down the socket
            frmMain.GOREsock.Shut UserIndex
    
            'If the user hasn't been closed, close it
            If UserList(UserIndex).Flags.UserLogged = 1 Then User_Close UserIndex
    
            'If User_Close was called, the array may have been resized, so make sure we are still in bounds
            If UserIndex > LastUser Then Exit Sub
    
            'Set to 0 again, just in case
            UserList(UserIndex).Flags.UserLogged = 0
        
        End If
    End If

End Sub

Sub User_Close(ByVal UserIndex As Integer)

'*****************************************************************
'Save user then reset the user's slot
'*****************************************************************
Dim NPCEraseArray() As Integer
Dim LoopC As Long
Dim Map As Integer
Dim Name As String
Dim i As Byte

    Log "Call User_Close(" & UserIndex & ")", CodeTracker '//\\LOGLINE//\\

    UserList(UserIndex).Flags.Disconnecting = 0

    'Save temps
    Map = (UserList(UserIndex).Pos.Map)
    Name = UserList(UserIndex).Name
    
    'Remove the user from their group
    If UserList(UserIndex).GroupIndex > 0 Then Group_RemoveUser UserIndex, UserList(UserIndex).GroupIndex

    'Remove the user's slave NPCs
    If UserList(UserIndex).NumSlaves > 0 Then
        
        'Store in a temp array or else things will get very, very dirty (and crashy)
        ReDim NPCEraseArray(1 To UserList(UserIndex).NumSlaves)
        CopyMemory NPCEraseArray(1), UserList(UserIndex).SlaveNPCIndex(1), 2 * UserList(UserIndex).NumSlaves
        
        'Loop through the NPCs - by the time this loop is done, the user's SlaveNPCIndex array should be erased
        For LoopC = 1 To UserList(UserIndex).NumSlaves
            If NPCEraseArray(LoopC) > 0 Then NPC_Kill NPCEraseArray(LoopC)
        Next LoopC
        
    End If
    
    'Set logged to false
    UserList(UserIndex).Flags.UserLogged = 0

    'Save user
    Save_User UserList(UserIndex), UserIndex
    
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

    'Update number of users
    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    TrayModify ToolTip, Server_BuildToolTipString

    'Send log off phrase
    If LenB(Name) Then
        ConBuf.PreAllocate 3 + Len(Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 68
        ConBuf.Put_String Name
        Data_Send ToAll, 0, ConBuf.Get_Buffer
        
        'For multiple servers, send to all the servers
        If NumServers > 1 Then
            For i = 1 To NumServers
                If ServerID <> i Then
                    Server_ConnectToServer i
                    If frmMain.ServerSocket(i).State = sckConnected Then frmMain.ServerSocket(i).SendData ConBuf.Get_Buffer
                End If
            Next i
        End If
        
    End If
    
    'Request the users array to be cleaned up
    CallUserCleanArray = 1

End Sub

Sub User_CleanArray()

'*****************************************************************
'Cleans up the user array by sizing it as small as possible to fit in all active users
'*****************************************************************

    'Check for a valid LastUser
    If LastUser = 0 Then Exit Sub

    'Update the last user value
    Do Until UserList(LastUser).Flags.UserLogged = 1
        LastUser = LastUser - 1
        If LastUser = 0 Then Exit Do
    Loop
    
    'Resize the userlist according to the LastUser value
    If LastUser = 0 Then
        Erase UserList
    Else
        ReDim Preserve UserList(1 To LastUser)
    End If
    
End Sub

Sub Data_Server_Message_ServerToServer(ByRef rBuf As DataBuffer)

'*****************************************************************
'Sends a Server Message between servers
'This has to be a little tricky since each message is handled differently
'<MessageID(B)> (<?(?)>)
'*****************************************************************
Dim MessageID As Byte
Dim Str1 As String
Dim Str2 As String

    MessageID = rBuf.Get_Byte
    
    'Handle based on the message ID
    Select Case MessageID
    
        'Disconnect message
        Case 68
            Str1 = rBuf.Get_String
            ConBuf.PreAllocate 3 + Len(Str1)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte MessageID
            ConBuf.Put_String Str1
            Data_Send ToAll, 0, ConBuf.Get_Buffer
            
        'Connect message
        Case 72
            Str1 = rBuf.Get_String
            ConBuf.PreAllocate 3 + Len(Str1)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte MessageID
            ConBuf.Put_String Str1
            Data_Send ToAll, 0, ConBuf.Get_Buffer
            
        'Kick message
        Case 77
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            ConBuf.PreAllocate 3 + Len(Str1)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte MessageID
            ConBuf.Put_String Str1
            ConBuf.Put_String Str2
            Data_Send ToAll, 0, ConBuf.Get_Buffer
            
    End Select

End Sub

Sub User_Connect(ByVal UserIndex As Integer, ByVal strName As String, ByVal Password As String, Optional ByVal NewUser As Boolean = False)

'*****************************************************************
'Reads the users .chr file and loads into Userlist array
'*****************************************************************
Dim TempPass As String
Dim CharIndex As Integer
Dim TempPos As WorldPos
Dim LoopC As Byte
Dim Count As Byte
Dim IPs() As String
Dim CurrIP As String
Dim s As String
Dim i As Byte
Dim j As Long

    Log "Call User_Connect(" & UserIndex & "," & strName & "," & Password & "," & NewUser & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the user is not in use
    If UserList(UserIndex).Flags.UserLogged Then
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
    
    If Not NewUser Then
    
        'Check if the character exists (we know that those marked NewUser = True exist since we just made them)
        If Not Server_UserExist(strName) Then
    
            'Tell the client we're disconnecting
            UserList(UserIndex).Flags.UserLogged = 1    'This is required to make the packet go through
            ConBuf.PreAllocate 3 + Len(strName)
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 80
            ConBuf.Put_String strName
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
            Data_Send_Buffer UserIndex
            Exit Sub
            
        End If
    
        'Set up the variables - we only have to do this if not a new char
        ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))    'Empty the character variables
        UserList(UserIndex).BufferSize = -1                         'Set the buffer start position (-1 = no buffer)
        Set UserList(UserIndex).Stats = New UserStats               'Create the stats class
        UserList(UserIndex).Flags.CreatedStats = 1
        UserList(UserIndex).Stats.UserIndex = UserIndex             'Set the user index
        
    End If
    
    'Set the user as logged in
    UserList(UserIndex).Flags.UserLogged = 1
    
    'Check if the user is on the correct server (pass the map just in case its a new char since they have the map already)
    If User_CorrectServer(strName, UserIndex, UserList(UserIndex).Pos.Map) = 0 Then
        Exit Sub
    End If
    
    'Store the user's IP in a temp variable
    CurrIP = frmMain.GOREsock.Address(UserIndex)
    
    'Check if the user is IP banned
    If Server_IPisBanned(CurrIP, s) Then
        ConBuf.PreAllocate 3 + Len(s)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 100
        ConBuf.Put_String s
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_Send_Buffer UserIndex
        Exit Sub
    End If

    'Check to see is user already logged with name
    j = Server_CheckForSameName(UserIndex, strName)
    If j > 0 Then

        'Tell the client we're disconnecting
        Data_Send ToIndex, UserIndex, cMessage(79).Data
        Data_Send_Buffer UserIndex
        
        'Disconnect the user currently on
        UserList(j).Flags.Disconnecting = 1
        
        Exit Sub
        
    End If
    
    'Get the password
    DB_RS.Open "SELECT password FROM users WHERE `name`='" & strName & "'", DB_Conn, adOpenStatic, adLockOptimistic
    TempPass = DB_RS!Password
    DB_RS.Close
    If LenB(TempPass) = 0 Then
    
        'Error getting the password
        Data_Send ToIndex, UserIndex, cMessage(81).Data
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If
    
    'Check password
    If MD5_String(Password) <> TempPass Then

        'Tell the client we're disconnecting
        Data_Send ToIndex, UserIndex, cMessage(82).Data
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If
    
    'Update the IP list
    DB_RS.Open "SELECT ip FROM users WHERE `name`='" & strName & "'", DB_Conn, adOpenStatic, adLockOptimistic

    'First IP entered into the list
    If LenB(DB_RS!IP) = 0 Then
        DB_RS!IP = CurrIP
        DB_RS.Update
    
    'Enter in a new IP into the list if its not already in there
    Else
        IPs() = Split(DB_RS!IP, vbNewLine)  'Create the IP list
        Count = UBound(IPs)
        For LoopC = 0 To Count
            If IPs(LoopC) = CurrIP Then
                LoopC = 250
                Exit For    'IP already in list, abort!
            End If
        Next LoopC
        If LoopC < 250 Then 'Check if we have a unique value
            If Count >= 9 Then  'If we have too many IPs already, just add to the bottom of the list
                For LoopC = 1 To Count
                    IPs(LoopC - 1) = IPs(LoopC)
                Next LoopC
            Else
                Count = Count + 1
                ReDim Preserve IPs(0 To Count)
            End If
            IPs(Count) = CurrIP 'Add the new IP to the end of the list
            s = vbNullString    'Clear out the string
            For LoopC = 0 To Count
                s = s & IPs(LoopC)  'Add the IP onto the string
                If LoopC < Count Then s = s & vbNewLine 'If not the last entry, add the line break
            Next LoopC
            DB_RS!IP = s
            DB_RS.Update
        End If
    End If
    DB_RS.Close

    'Load character information from file
    Load_User UserIndex, strName
    
    'Clear the idle and last packet counter
    UserList(UserIndex).Counters.IdleCount = timeGetTime
    UserList(UserIndex).Counters.LastPacket = timeGetTime

    'Update inventory
    User_UpdateInv True, UserIndex, 0

    'Update number of users
    NumUsers = NumUsers + 1

    'Update map and connection groups data
    MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

    'Check if it's the first user on the map
    If MapInfo(UserList(UserIndex).Pos.Map).DataLoaded = 0 Or MapInfo(UserList(UserIndex).Pos.Map).NumUsers = 1 Then
        Load_Maps_Temp UserList(UserIndex).Pos.Map
        ReDim MapUsers(UserList(UserIndex).Pos.Map).Index(1 To 1)
    Else
        ReDim Preserve MapUsers(UserList(UserIndex).Pos.Map).Index(1 To MapInfo(UserList(UserIndex).Pos.Map).NumUsers)
    End If
    MapUsers(UserList(UserIndex).Pos.Map).Index(MapInfo(UserList(UserIndex).Pos.Map).NumUsers) = UserIndex

    'Get closest legal pos
    Server_ClosestLegalPos UserList(UserIndex).Pos, TempPos
    If Not Server_LegalPos(TempPos.Map, TempPos.X, TempPos.Y, 0) Then

        'Tell the client we're disconnecting
        Data_Send ToIndex, UserIndex, cMessage(83).Data
        Data_Send_Buffer UserIndex
        User_Close UserIndex
        Exit Sub
        
    End If
    UserList(UserIndex).Pos = TempPos
    
    'Tell the user they have successfully connected
    ConBuf.PreAllocate 7
    ConBuf.Put_Byte DataCode.Server_Connect

    'Tell client to try switching maps
    ConBuf.Put_Byte DataCode.Map_LoadMap
    ConBuf.Put_Integer UserList(UserIndex).Pos.Map
    ConBuf.Put_Integer MapInfo(UserList(UserIndex).Pos.Map).MapVersion
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer()

    'Give user a charindex
    CharIndex = Server_NextOpenCharIndex
    UserList(UserIndex).Char.CharIndex = CharIndex
    CharList(CharIndex).Index = UserIndex
    CharList(CharIndex).CharType = CharType_PC

    'Show Character to others
    User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y

    'Refresh tooltip
    TrayModify ToolTip, Server_BuildToolTipString

    'Send the MOTD
    Data_Send ToIndex, UserIndex, MOTDBuffer()
    
    'If the user already has a weapon equiped with a range, tell what that range is
    If UserList(UserIndex).WeaponEqpObjIndex > 0 Then
        If ObjData.WeaponRange(UserList(UserIndex).WeaponEqpObjIndex) > 0 Then
            ConBuf.PreAllocate 2
            ConBuf.Put_Byte DataCode.User_SetWeaponRange
            ConBuf.Put_Byte ObjData.WeaponRange(UserList(UserIndex).WeaponEqpObjIndex)
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        End If
    End If

    'Tell the user if they have new mail
    Count = 0
    For LoopC = 1 To MaxMailPerUser
        If UserList(UserIndex).MailID(LoopC) <> 0 Then
            DB_RS.Open "SELECT new FROM mail WHERE id=" & UserList(UserIndex).MailID(LoopC), DB_Conn, adOpenStatic, adLockOptimistic
            If Val(DB_RS!New) = 1 Then Count = Count + 1
            DB_RS.Close
        End If
    Next LoopC
    
    'Send the appropriate message according to how much mail they have
    If Count > 1 Then
        ConBuf.PreAllocate 3
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 71
        ConBuf.Put_Byte Count
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    ElseIf Count = 1 Then
        Data_Send ToIndex, UserIndex, cMessage(70).Data
    Else
        Data_Send ToIndex, UserIndex, cMessage(69).Data
    End If
    
    'Send the user their quest information for their active quests
    Quest_SendText UserIndex  'Default quest index (0) means send all

    'Send list of known skills
    User_SendKnownSkills UserIndex

    'Connect message
    ConBuf.PreAllocate 3 + Len(UserList(UserIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 72
    ConBuf.Put_String UserList(UserIndex).Name
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_Connect
    
    'For multiple servers, send to all the servers
    If NumServers > 1 Then
        For i = 1 To NumServers
            If ServerID <> i Then
                Server_ConnectToServer i
                If frmMain.ServerSocket(i).State = sckConnected Then frmMain.ServerSocket(i).SendData ConBuf.Get_Buffer
            End If
        Next i
    End If

End Sub

Sub User_ConnectNew(ByVal UserIndex As Integer, ByVal Name As String, ByVal Password As String, ByVal Body As Integer, ByVal Head As Integer, ByVal Class As Byte)

'*****************************************************************
'Opens a new user. Loads default vars, saves then calls connectuser
'*****************************************************************
Dim TempObjs(1 To 5) As Obj
Dim i As Byte

    Log "Call User_Connect(" & UserIndex & "," & Name & "," & Password & "," & Body & "," & Head & "," & Class & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the user is not in use
    If UserList(UserIndex).Flags.UserLogged Then Exit Sub

    'Check for a valid body and head
    If Head <> 1 Then Exit Sub
    If Body <> 1 Then Exit Sub
    
    'Check for valid starting classes
    If Class <> ClassID.Mage And Class <> ClassID.Warrior And Class <> ClassID.Rogue Then Exit Sub

    'Check for valid name and password
    If Len(Name) > 10 Then
        Log "User_ConnectNew: Name too long", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Password) > 10 Then
        Log "User_ConnectNew: Password too long", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Name) < 3 Then
        Log "User_ConnectNew: Name too short", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Len(Password) < 3 Then
        Log "User_ConnectNew: Password too short", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Server_LegalString(Name) = False Then
        Log "User_ConnectNew: Invalid name string", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    If Server_LegalString(Password) = False Then
        Log "User_ConnectNew: Invalid password string", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Clear user character
    ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))
    UserList(UserIndex).BufferSize = -1
    
    'Set the user as logged in
    UserList(UserIndex).Flags.UserLogged = 1

    'Check for character file
    If Server_UserExist(Name) Then

        'Tell the client we're disconnecting
        Data_Send ToIndex, UserIndex, cMessage(84).Data
        Data_Send_Buffer UserIndex
        Exit Sub
        
    End If

    'Set the stats class
    Set UserList(UserIndex).Stats = New UserStats
    UserList(UserIndex).Flags.CreatedStats = 1
    UserList(UserIndex).Stats.UserIndex = UserIndex

    'This is for testing purposes only - remove this in any public release!
    UserList(UserIndex).Flags.GMLevel = 1

    'Set the user's variables
    UserList(UserIndex).Name = Name
    UserList(UserIndex).Class = Class
    UserList(UserIndex).Char.Heading = SOUTH
    UserList(UserIndex).Char.HeadHeading = SOUTH
    UserList(UserIndex).Char.Head = Head
    UserList(UserIndex).Char.Hair = 1
    UserList(UserIndex).Char.Body = Body
    UserList(UserIndex).Counters.IdleCount = timeGetTime
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

    'For testing only (gives user all the skills
    For i = 1 To NumSkills
        UserList(UserIndex).KnownSkills(i) = 1
    Next i

    'Save the user
    Save_User UserList(UserIndex), UserIndex, MD5_String(Password), 1

    'Write a test message to the user
    For i = 1 To 5
        TempObjs(i).ObjIndex = Int(Rnd * 7) + 1
        TempObjs(i).Amount = Int(Rnd * 10) + 1
    Next i
    For i = 1 To 5
        Server_WriteMail -1, UserList(UserIndex).Name, "Test Message", "This is a test message that simply shows the pwnification of the mailing system. Here, have a random number! " & Rnd * 100, TempObjs
    Next i
    
    'Disconnect the user so we can re-connect them through the connect sub
    UserList(UserIndex).Flags.UserLogged = 0

    'Go through regular connecting routine
    User_Connect UserIndex, Name, Password, True
    
ErrOut:

End Sub
