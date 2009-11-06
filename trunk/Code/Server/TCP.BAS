Attribute VB_Name = "TCP"
Option Explicit

'Packet priorities constants (influncing when packets are sent)
Public Const PP_None As Byte = 0        'Does not force sending
Public Const PP_Low As Byte = 1         'Enables a count-down timer
Public Const PP_High As Byte = 2        'Packet sent as soon as possible
Public Const PP_Time As Long = 1000     'How much time must pass for a PP_Low packet to be sent

'Constants used in the Data_Send sub
Public Const ToIndex = 0            'Send data to a single User index
Public Const ToAll = 1              'Send it to all User index
Public Const ToMap = 2              'Send it to all users in a map
Public Const ToPCArea = 3           'Send to all users in a user's area
Public Const ToMapButIndex = 4      'Send to all on a map but the index
Public Const ToNPCArea = 5          'Send to all users in a NPC's area

'How many miliseconds must go by for a packet to be sent (even if PP_High, the time must be waited)
' Keep in mind the highest frame rate is 60 (~17ms), so it is not recommended to go below that value
' or else you are just wasting bandwidth.
Public Const Packet_WaitTime As Long = 50   'Recommended about 200~300 at highest, about 50 at lowest

'MTU size (how large the packet buffer can be)
Public Const MTU_Nondialup  As Integer = 1472
Public Const MTU_Dialup As Integer = 576
Public Const SOCKET_BUFFER_SIZE As Integer = MTU_Dialup 'Buffer in bytes for each socket
Public Const UserBufferSize As Integer = SOCKET_BUFFER_SIZE - 2


'*************************
'*** PACKET PRIORITIES ***
'*************************

'Packet priorities (placed here for easy changing - for advanced users -
' recommended to go packet-by-packet and decide what you want)
'If you have more bandwidth then you could possibly use, then just set
' all the packet priorities to HIGH. Packet priorities is aimed for those
' who want and need an extra booster from their bandwidth.
'vbGORE, by default, aims for the lowest packet priority values that will
' affect the gameplay the least (most stuff not listed as HIGH you wont notice the lag time)

'*** Option 1 - Optimized for minimal bandwidth usage ***

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

'*** Option 2 - Optimized for speed and performance ***

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


Sub Data_Comm_Emote(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Emote
'<txt(S-EX)>
'*****************************************************************

Dim txt As String

    txt = rBuf.Get_String

    'Check for invalid conditions
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then Exit Sub
    
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).Name & " " & txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
    Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_LocalChat

End Sub

Sub Data_Comm_Shout(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Shout
'<txt(S-EX)>
'*****************************************************************

Dim txt As String

    txt = rBuf.Get_String

    'Check for invalid conditions
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then Exit Sub
    
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 76
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_String txt
    Data_Send ToAll, UserIndex, ConBuf.Get_Buffer, , PP_GlobalChat

End Sub

Sub Data_Comm_Talk(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Talk
'<txt(S-EX)>
'*****************************************************************

Dim txt As String

    txt = rBuf.Get_String

    'Check for invalid conditions
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then Exit Sub
    
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String UserList(UserIndex).Name & ": " & txt
    ConBuf.Put_Byte DataCode.Comm_FontType_Talk
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

    txt = rBuf.Get_String

    'Split the name from the rest of the string
    TempStr = Split(txt, " ")
    
    'This will make sure we actually have a string (dont crash if there is no space)
    On Error GoTo ErrOut
    tName = TempStr(0)
    On Error GoTo 0 'Removes the previous On Error

    'Check for valid tName
    If Server_LegalString(tName) = False Then Exit Sub

    'Retrieve the index of the player being whispered to
    tIndex = User_NameToIndex(tName)

    'Check for invalid conditions
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If txt = vbNullString Then Exit Sub
    
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

    'Tell the target what they got whispered
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 52
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_String tMessage
    Data_Send ToIndex, tIndex, ConBuf.Get_Buffer, , PP_PrivateChat

    'Tell the whisperer what they whispered
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 53
    ConBuf.Put_String UserList(tIndex).Name
    ConBuf.Put_String tMessage
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_PrivateChat

ErrOut:

End Sub

Sub Data_GM_SetGMLevel(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM sets another player as GM
'<TargetName(S)><GMLevel(B)>
'*****************************************************************
Dim TargetName As String
Dim TargetIndex As String
Dim GMLevel As Byte

    TargetName = rBuf.Get_String
    GMLevel = rBuf.Get_Byte
    
    'Make sure the one using the command is admin
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub
    
    'Get the target user's ID
    TargetIndex = User_NameToIndex(TargetName)
    
    'Set the GM level (if online)
    If TargetIndex > 0 Then
        UserList(TargetIndex).Flags.GMLevel = GMLevel
    Else    'Set GM level (if offline)
        DB_RS.Open "SELECT * FROM users WHERE `name`='" & TargetName & "'"
        DB_RS!gm = GMLevel
        DB_RS.Update
        DB_RS.Close
    End If
    
    'Say the change
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 54
    ConBuf.Put_String TargetName
    ConBuf.Put_Byte GMLevel
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, PP_GMMessages

End Sub

Sub Data_GM_Approach(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM approaches user
'<TargetName(S)>
'*****************************************************************

Dim TargetName As String
Dim TargetIndex As Integer
Dim nPos As WorldPos

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_GMMessages
        Exit Sub
    End If

    'Don't allow if switching maps
    If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Sub

    'Find closest legal position and warp there
    Server_ClosestLegalPos UserList(TargetIndex).Pos, nPos
    If Server_LegalPos(nPos.Map, nPos.x, nPos.Y, 0) Then User_WarpChar UserIndex, nPos.Map, nPos.x, nPos.Y

End Sub

Sub Data_GM_Kick(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM kicked the user
'<TargetName(S)>
'*****************************************************************

Dim TargetName As String
Dim TargetIndex As Integer

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

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
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_String UserList(TargetIndex).Name
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_GlobalMessage

    'Close the user
    Server_CloseSocket TargetIndex

End Sub

Sub Data_GM_Raise(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'GM raised a user's experience
'<TargetName(S)><RaiseValue(L)>
'*****************************************************************

    On Error GoTo RaiseErrOut
Dim TargetName As String
Dim TargetIndex As Integer
Dim RaiseValue As Long

    TargetName = rBuf.Get_String
    RaiseValue = rBuf.Get_Long
    TargetIndex = User_NameToIndex(TargetName)

    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

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

    TargetName = rBuf.Get_String
    TargetIndex = User_NameToIndex(TargetName)

    'Check for invalid values
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub

    'Check the user admin ID
    If UserList(UserIndex).Flags.GMLevel = 0 Then Exit Sub

    'See if user online
    If TargetIndex <= 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 51
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_GMMessages
        Exit Sub
    End If

    'Don't allow if switching maps
    If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Sub

    'Find closest legal position and warp there
    Server_ClosestLegalPos UserList(UserIndex).Pos, nPos

    'If the position is legal, warp the user then tell them they have warped
    If Server_LegalPos(nPos.Map, nPos.x, nPos.Y, 0) Then
        User_WarpChar TargetIndex, nPos.Map, nPos.x, nPos.Y
        ConBuf.Clear
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

    If UserList(UserIndex).Pos.Map <= 0 Then Exit Sub
    If UserList(UserIndex).Pos.Map > NumMaps Then Exit Sub

    UserList(UserIndex).Flags.SwitchingMaps = 0

    'Send the map name
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Map_SendName
    ConBuf.Put_String MapInfo(UserList(UserIndex).Pos.Map).Name
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

Private Sub Data_Send_UpdatePP(ByVal UserIndex As Integer, ByVal Priority As Byte)

    'Update the wait time
    If UserList(UserIndex).PacketWait <= 0 Then UserList(UserIndex).PacketWait = Packet_WaitTime

    'Check if moving from PP_None to PP_Low
    If UserList(UserIndex).PPValue = PP_None And Priority = PP_Low Then
        UserList(UserIndex).PPCount = PP_Time
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
    
    'Set the copy size
    CopySize = UBound(sndData())
    
    'Check for invalid values
    If CopySize = 0 Then Exit Sub
    
    Select Case sndRoute
    
            'Send to the sndIndex
        Case ToIndex
            If UserList(sndIndex).Flags.UserLogged = 0 Then Exit Sub
            If sndIndex > UBound(UserList()) Then Exit Sub
            If UBound(UserList(sndIndex).SendBuffer) + 1 + CopySize >= UserBufferSize Then Data_SendBuffer sndIndex
            CopyPos = UBound(UserList(sndIndex).SendBuffer) + 1
            ReDim Preserve UserList(sndIndex).SendBuffer(UBound(UserList(sndIndex).SendBuffer) + CopySize)
            CopyMemory UserList(sndIndex).SendBuffer(CopyPos), sndData(0), CopySize
            Data_Send_UpdatePP sndIndex, Priority
    
            'Send to All
        Case ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).Flags.UserLogged Then
                    If UBound(UserList(LoopC).SendBuffer) + 1 + CopySize >= UserBufferSize Then Data_SendBuffer LoopC
                    CopyPos = UBound(UserList(LoopC).SendBuffer) + 1
                    ReDim Preserve UserList(LoopC).SendBuffer(UBound(UserList(LoopC).SendBuffer) + CopySize)
                    CopyMemory UserList(LoopC).SendBuffer(CopyPos), sndData(0), CopySize
                    Data_Send_UpdatePP LoopC, Priority
                End If
            Next LoopC
    
            'Send to Map
        Case ToMap
            If sndMap = 0 Then Exit Sub
            For LoopC = 1 To MapInfo(sndMap).NumUsers
                If UserList(LoopC).Flags.UserLogged = 1 Then
                    If UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1 + CopySize >= UserBufferSize Then Data_SendBuffer ConnectionGroups(sndMap).UserIndex(LoopC)
                    CopyPos = UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1
                    ReDim Preserve UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + CopySize)
                    CopyMemory UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                    Data_Send_UpdatePP LoopC, Priority
                End If
            Next LoopC
    
            'Send to everone on map but sndIndex
        Case ToMapButIndex
            If sndMap = 0 Then Exit Sub
            For LoopC = 1 To MapInfo(sndMap).NumUsers
                If UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Flags.UserLogged Then
                    If ConnectionGroups(sndMap).UserIndex(LoopC) <> sndIndex Then
                        If UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1 + CopySize >= UserBufferSize Then Data_SendBuffer ConnectionGroups(sndMap).UserIndex(LoopC)
                        CopyPos = UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1
                        ReDim Preserve UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + CopySize)
                        CopyMemory UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                        Data_Send_UpdatePP LoopC, Priority
                    End If
                End If
            Next LoopC
    
            'Send to PC Area
        Case ToPCArea
            sndMap = UserList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Flags.UserLogged Then
                        If Server_Distance(UserList(sndIndex).Pos.x, UserList(sndIndex).Pos.Y, UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Pos.x, UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Pos.Y) <= Max_Server_Distance Then
                            If UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1 + CopySize >= UserBufferSize Then Data_SendBuffer ConnectionGroups(sndMap).UserIndex(LoopC)
                            CopyPos = UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1
                            ReDim Preserve UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + CopySize)
                            CopyMemory UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                            Data_Send_UpdatePP LoopC, Priority
                        End If
                    End If
                Next LoopC
            End If
    
            'Send to NPC Area
        Case ToNPCArea
            sndMap = NPCList(sndIndex).Pos.Map
            If sndMap > 0 Then
                For LoopC = 1 To MapInfo(sndMap).NumUsers
                    If UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Flags.UserLogged Then
                        If Server_Distance(NPCList(sndIndex).Pos.x, NPCList(sndIndex).Pos.Y, UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Pos.x, UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).Pos.Y) <= Max_Server_Distance Then
                            If UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1 + CopySize >= UserBufferSize Then Data_SendBuffer ConnectionGroups(sndMap).UserIndex(LoopC)
                            CopyPos = UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + 1
                            ReDim Preserve UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(UBound(UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer) + CopySize)
                            CopyMemory UserList(ConnectionGroups(sndMap).UserIndex(LoopC)).SendBuffer(CopyPos), sndData(0), CopySize
                            Data_Send_UpdatePP LoopC, Priority
                        End If
                    End If
                Next LoopC
            End If

    End Select
    
ErrOut:

End Sub

Sub Data_SendBuffer(ByVal UserIndex As Integer)

'*****************************************************************
'Sends the data buffer to the user index
'*****************************************************************

Dim bArray() As Byte

    If frmMain.Sox.State(UserList(UserIndex).ConnID) = soxIdle Or soxListening Then
        If UBound(UserList(UserIndex).SendBuffer) > 0 Then

            'Because the server's send buffer's array starts at 0, the first array element (0) will be
            '0 always. To work around this little technicality, copy memory to a new buffer to move
            'all array elements down one. I know, excessive memory useage for one byte per packet, but
            'I think I want to keep my byte thank you very much!
            ReDim bArray(UBound(UserList(UserIndex).SendBuffer) - 1)
            CopyMemory bArray(0), UserList(UserIndex).SendBuffer(1), UBound(UserList(UserIndex).SendBuffer)
            
            'Uncomment this to see packets going out from the server
            'Dim i As Long
            'Dim S As String
            'For i = LBound(bArray) To UBound(bArray)
            '    S = S & bArray(i) & " "
            'Next i
            'Debug.Print S

            'Encrypt our packet
            Select Case EncryptionType
            Case EncryptionTypeBlowfish
                Encryption_Blowfish_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeCryptAPI
                Encryption_CryptAPI_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeDES
                Encryption_DES_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeGost
                Encryption_Gost_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeRC4
                Encryption_RC4_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeXOR
                Encryption_XOR_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeSkipjack
                Encryption_Skipjack_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeTEA
                Encryption_TEA_EncryptByte bArray, EncryptionKey
            Case EncryptionTypeTwofish
                Encryption_Twofish_EncryptByte bArray, EncryptionKey
            End Select

            'Send the packet
            frmMain.Sox.SendData UserList(UserIndex).ConnID, bArray()

            'Calculate bandwidth usage
            If CalcTraffic = True Then DataOut = DataOut + UBound(bArray) + 42  '1 for the array ofset, 20 for the TCP header, 20 for IPv4 header, 1 for Sox header

            'Clear the buffer
            ReDim UserList(UserIndex).SendBuffer(0)
            
            'Clear packet values
            UserList(UserIndex).PPCount = 0
            UserList(UserIndex).PPValue = 0
            UserList(UserIndex).PacketWait = 0
            
        End If
    End If

    'Check if the user was disconnected
    If frmMain.Sox.State(UserList(UserIndex).ConnID) = 1 Then User_Close UserIndex

    'Check if we are disconnecting
    If UserList(UserIndex).Flags.Disconnecting = 1 Then Server_CloseSocket UserIndex

End Sub

Sub Data_Server_Help(ByVal UserIndex As Integer)

'*****************************************************************
'User wants to retrieve help buffer from server
'<>
'*****************************************************************

Dim NumHelpLines As Integer
Dim LoopC As Long

'Get the number of lines

    NumHelpLines = Val(Var_Get(ServerDataPath & "Help.dat", "INIT", "NumLines"))

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
'<RecieverName(S)><Subject(S)><Message(S-EX)><ObjIndexString(S)><ObjAmountString(S)>
'*****************************************************************

Dim RecieverName As String
Dim Subject As String
Dim Message As String
Dim ObjIndexStr As String
Dim ObjAmountStr As String

    RecieverName = rBuf.Get_String
    Subject = rBuf.Get_String
    Message = rBuf.Get_StringEX
    ObjIndexStr = rBuf.Get_String
    ObjAmountStr = rBuf.Get_String

    'Check for invalid values (most is handled in Server_WriteMail sub)
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub

    'Write the mail to the RecieverName
    Server_WriteMail UserIndex, RecieverName, Subject, Message, ObjIndexStr, ObjAmountStr

End Sub

Sub Data_Server_MailDelete(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Delete mail
'<MailIndex(B)>
'*****************************************************************

Dim MailIndex As Byte
Dim LoopC As Byte
Dim MsgData As MailData

    MailIndex = rBuf.Get_Byte

    'Check for invalid values
    If MailIndex < 1 Then Exit Sub
    If MailIndex > MaxMailPerUser Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).MailID(MailIndex) <= 0 Then Exit Sub

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

Sub Data_Server_MailItemInfo(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Get info of an item in mail
'<ObjIndex(B)>
'*****************************************************************

Dim ObjIndex As Byte
Dim MsgData As MailData

    ObjIndex = rBuf.Get_Byte

    'Check for invalid values
    If UserList(UserIndex).Flags.LastViewedMail <= 0 Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If ObjIndex > MaxMailObjs Then Exit Sub
    If ObjIndex < 1 Then Exit Sub

    'Load the mail into a temporary mail buffer
    MsgData = Load_Mail(UserList(UserIndex).MailID(UserList(UserIndex).Flags.LastViewedMail))

    'Send the info
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_MailItemInfo
    ConBuf.Put_String ObjData(MsgData.Obj(ObjIndex).ObjIndex).Name
    ConBuf.Put_Integer MsgData.Obj(ObjIndex).Amount
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

    ObjIndex = rBuf.Get_Byte

    'Check for invalid values
    If UserList(UserIndex).Flags.LastViewedMail <= 0 Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If ObjIndex > MaxMailObjs Then Exit Sub
    If ObjIndex <= 0 Then Exit Sub

    'Check to make sure the user is still next to the same mailbox
    If UserList(UserIndex).Pos.Map = UserList(UserIndex).MailboxPos.Map Then
        If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, UserList(UserIndex).MailboxPos.x, UserList(UserIndex).MailboxPos.Y) > 2 Then
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
    MsgData = Load_Mail(UserList(UserIndex).MailID(UserList(UserIndex).Flags.LastViewedMail))
    If MsgData.Obj(ObjIndex).ObjIndex <= 0 Then Exit Sub
    If MsgData.Obj(ObjIndex).Amount <= 0 Then Exit Sub

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
        ConBuf.Put_String ObjData(MsgData.Obj(ObjIndex).ObjIndex).Name
        ConBuf.Put_Integer MsgData.Obj(ObjIndex).Amount
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        
        UserList(UserIndex).Object(LoopC).ObjIndex = MsgData.Obj(ObjIndex).ObjIndex
        UserList(UserIndex).Object(LoopC).Amount = UserList(UserIndex).Object(LoopC).Amount + MsgData.Obj(ObjIndex).Amount
        MsgData.Obj(ObjIndex).ObjIndex = 0
        MsgData.Obj(ObjIndex).Amount = 0

        'Update the mail object slot
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_MailItemRemove
        ConBuf.Put_Byte ObjIndex
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    Else

        'Take as many items as possible
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 59
        ConBuf.Put_String ObjData(MsgData.Obj(ObjIndex).ObjIndex).Name
        ConBuf.Put_Integer (MAX_INVENTORY_OBJS - UserList(UserIndex).Object(LoopC).Amount)
        ConBuf.Put_Integer (MsgData.Obj(ObjIndex).Amount - (MAX_INVENTORY_OBJS - UserList(UserIndex).Object(LoopC).Amount))
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        MsgData.Obj(ObjIndex).Amount = (MsgData.Obj(ObjIndex).Amount + UserList(UserIndex).Object(LoopC).Amount) - MAX_INVENTORY_OBJS
        UserList(UserIndex).Object(LoopC).Amount = MAX_INVENTORY_OBJS

    End If

    'Update the changed inventory slot
    User_UpdateInv False, UserIndex, LoopC

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

    MessageID = rBuf.Get_Byte

    'Check for invalid values
    If MessageID > MaxMailPerUser Then Exit Sub
    If MessageID <= 0 Then Exit Sub
    If UserList(UserIndex).MailID(MessageID) <= 0 Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub

    'Load the mail into a temporary mail buffer
    MsgData = Load_Mail(UserList(UserIndex).MailID(MessageID))

    'Send the message information
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_MailMessage
    ConBuf.Put_StringEX MsgData.Message
    ConBuf.Put_String MsgData.Subject
    ConBuf.Put_String MsgData.WriterName
    For LoopC = 1 To MaxMailObjs    'Get all the objects
        ConBuf.Put_Integer ObjData(MsgData.Obj(LoopC).ObjIndex).GrhIndex
    Next LoopC
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Remember the last message the user viewed
    UserList(UserIndex).Flags.LastViewedMail = MessageID

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

'Pong the ping so we can be pinged from out pong

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
Dim NumUsers As Integer
Dim LoopC As Long
Dim tStr As String
Dim i As Long

'Get the list of users

    For LoopC = 1 To LastUser
        If UserList(LoopC).Name <> "" Then
            If UserList(LoopC).Flags.UserLogged Then
                NumUsers = NumUsers + 1
                ReDim Preserve UserNames(1 To NumUsers)
                UserNames(NumUsers) = UserList(LoopC).Name
            End If
        End If
    Next LoopC

    'Set the first line (total users)
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 78
    ConBuf.Put_Integer NumUsers

    'Send the user names
    For i = 1 To (NumUsers \ 10) + 1    'Replace the 10's with the number of users per line
        tStr = vbNullString             'Clear the string buffer
        For LoopC = 1 To 10
            If LoopC + ((i - 1) * 10) > NumUsers Then Exit For
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

Sub Data_User_Attack(ByVal UserIndex As Integer)

'*****************************************************************
'User attacked
'<>
'*****************************************************************

'Invalid values checked by User_Attack routine

    User_Attack UserIndex

End Sub

Sub Data_User_BaseStat(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User wants to raise a stat
'<StatID(B)>
'*****************************************************************
Dim StatID As Byte

    StatID = rBuf.Get_Byte

    'Check for invalid values
    If UserList(UserIndex).Flags.UserLogged = 0 Then Exit Sub
    If StatID <= 0 Then Exit Sub
    If StatID > NumStats Then Exit Sub

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

    SkillID = rBuf.Get_Byte

    'Easiest to always pass a target, even if a target can not be selected for the skill
    TargetCharIndex = rBuf.Get_Integer
    
    'Check for a valid char index
    If TargetCharIndex < 0 Then TargetCharIndex = 0
    If TargetCharIndex > LastChar Then TargetCharIndex = 0
    If CharList(TargetCharIndex).Index <= 0 Then TargetCharIndex = 0
    
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

    'Call the skill by the SkillID
    Select Case SkillID
        Case SkID.Bless
            If Server_CheckTargetedDistance(UserIndex) = 0 Then Exit Sub
            Skill_Bless TargetIndex, UserIndex, TargetType, CharType_PC
        Case SkID.Protection
            If Server_CheckTargetedDistance(UserIndex) = 0 Then Exit Sub
            Skill_Protection TargetIndex, UserIndex, TargetType, CharType_PC
        Case SkID.Heal
            If Server_CheckTargetedDistance(UserIndex) = 0 Then Exit Sub
            Skill_Heal TargetIndex, UserIndex, TargetType, CharType_PC
        Case SkID.IronSkin
            Skill_IronSkin UserIndex
        Case SkID.SpikeField
            Skill_SpikeField UserIndex
        Case SkID.Strengthen
            If Server_CheckTargetedDistance(UserIndex) = 0 Then Exit Sub
            Skill_Strengthen TargetIndex, UserIndex, TargetType, CharType_PC
        Case SkID.Warcry
            Skill_Warcry UserIndex
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

    SrcSlot = rBuf.Get_Byte
    DestSlot = rBuf.Get_Byte

    'Check for valid numbers
    If SrcSlot <= 0 Then Exit Sub
    If SrcSlot > MAX_INVENTORY_SLOTS Then Exit Sub
    If DestSlot <= 0 Then Exit Sub
    If DestSlot > MAX_INVENTORY_SLOTS Then Exit Sub

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
    
    'Update the database
    If MySQLUpdate_Desc Then
        DB_RS.Open "SELECT * FROM users WHERE `name`='" & UserList(UserIndex).Name & "'", DB_Conn, adOpenStatic, adLockOptimistic
        DB_RS!Desc = Desc
        DB_RS.Update
        DB_RS.Close
    End If

End Sub

Sub Data_User_Drop(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User dropping an object to the ground
'<ObjSlot(B)><Amount(I)>
'*****************************************************************

Dim ObjSlot As Byte
Dim Amount As Integer

    ObjSlot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer

    'Invalid values handled by sub
    User_DropObj UserIndex, ObjSlot, Amount, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y

End Sub

Sub Data_User_Emote(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Brings up emoticon to everyone on the screen (PCArea)
'<>
'*****************************************************************

Dim EmoteIndex As Byte

    EmoteIndex = rBuf.Get_Byte

    'Check for invalid values
    If EmoteIndex <= 0 Then Exit Sub
    If EmoteIndex > NumEmotes Then Exit Sub

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

'Invalid values handled by sub

    User_GetObj UserIndex

End Sub

Sub Data_User_KnownSkills(ByVal UserIndex As Integer)

'*****************************************************************
'Requested to have list of known spells/skills
'<>
'*****************************************************************

'Invalid values handled by sub

    User_SendKnownSkills UserIndex

End Sub

Sub Data_User_LeftClick(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User left-clicked a tile
'<X(B)><Y(B)>
'*****************************************************************

Dim x As Byte
Dim Y As Byte

    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Invalid values handled by sub
    User_LookAtTile UserIndex, UserList(UserIndex).Pos.Map, x, Y, vbLeftButton

End Sub

Sub Data_User_Login(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Log on existing character
'<Name(S)><Password(S)>
'*****************************************************************

Dim Name As String
Dim Pass As String

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

'Check for invalid values

    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
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

'Check for invalid values

    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
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
'<Direction(B)>
'*****************************************************************

Dim Dir As Byte

    Dir = rBuf.Get_Byte

    'Invalid values handled by sub
    User_MoveChar UserIndex, Dir

End Sub

Sub Data_User_NewLogin(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Log on a new character
'<Name(S)><Password(S)>
'*****************************************************************

Dim Name As String
Dim Pass As String

    Name = rBuf.Get_String
    Pass = rBuf.Get_String

    'Invalid values handled by sub
    User_ConnectNew UserIndex, Name, Pass, 1, 1

End Sub

Sub Data_User_RightClick(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'User right-clicked a tile
'<X(B)><Y(B)>
'*****************************************************************

Dim x As Byte
Dim Y As Byte

    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Invalid values handled by sub
    User_LookAtTile UserIndex, UserList(UserIndex).Pos.Map, CInt(x), CInt(Y), vbRightButton

End Sub

Sub Data_User_Rotate(ByRef rBuf As DataBuffer, ByVal UserIndex As Integer)

'*****************************************************************
'Rotate the user
'<Direction(B)>
'*****************************************************************

Dim Dir As Byte

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

'Make sure the user has a quests pending

    If UserList(UserIndex).Flags.QuestNPC = 0 Then Exit Sub

    'Check the distance of the user to the NPC
    QuestNPC = UserList(UserIndex).Flags.QuestNPC
    If NPCList(QuestNPC).Pos.Map = UserList(UserIndex).Pos.Map Then
        If Server_Distance(NPCList(QuestNPC).Pos.x, NPCList(QuestNPC).Pos.Y, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) < Max_Server_Distance + 2 Then 'Add a little bit extra to distance, in case NPC moves

            'Check the requirements
            If QuestData(NPCList(QuestNPC).Quest).AcceptReqLvl > UserList(UserIndex).Stats.BaseStat(SID.ELV) Then
                ConBuf.Clear
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
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 63
                    ConBuf.Put_Integer QuestData(NPCList(QuestNPC).Quest).AcceptReqObjAmount
                    ConBuf.Put_String ObjData(QuestData(NPCList(QuestNPC).Quest).AcceptReqObj).Name
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
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_Message
                    ConBuf.Put_Byte 12
                    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                    UserList(UserIndex).Flags.QuestNPC = 0
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

            'Check for sending an object
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

    Slot = rBuf.Get_Byte
    Amount = rBuf.Get_Integer

    'Set the NPC index to trade with
    NPCIndex = UserList(UserIndex).Flags.TradeWithNPC

    'Check for invalid values
    If NPCIndex <= 0 Then Exit Sub
    If NPCIndex > LastNPC Then Exit Sub
    If Slot > NPCList(NPCIndex).NumVendItems Then Exit Sub
    If Slot <= 0 Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.TradeWithNPC <= 0 Then Exit Sub
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
    If Server_Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, NPCList(NPCIndex).Pos.x, NPCList(NPCIndex).Pos.Y) > 5 Then
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
        ConBuf.Put_String PurchaseObj.Name
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
            ConBuf.Put_String PurchaseObj.Name
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
    ConBuf.Put_String PurchaseObj.Name
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

    ObjSlot = rBuf.Get_Byte

    'Invalid values handled by sub
    User_UseInvItem UserIndex, ObjSlot

End Sub

Sub Server_CloseSocket(ByVal UserIndex As Integer)

'*****************************************************************
'Close the users socket
'*****************************************************************

    If DEBUG_PacketFlood = True Then Exit Sub

    'Make sure the user is valid
    If UserIndex > 0 Then
        If UserList(UserIndex).ConnID <> -1 Then
            UserList(UserIndex).Flags.Disconnecting = 0

            'Send the buffer
            Data_SendBuffer UserIndex

            'Shut down the socket
            frmMain.Sox.Shut UserList(UserIndex).ConnID

            'If the user hasn't been closed, close it
            If UserList(UserIndex).Flags.UserLogged = 1 Then User_Close UserIndex

            'Clear the record of an old inSox, in case it ever did get transfered to another char
            UserList(UserIndex).ConnID = -1
            UserList(UserIndex).Flags.UserLogged = 0    'Set to 0 again, just in case

        End If
    End If

End Sub

Sub User_Close(ByVal UserIndex As Integer)

'*****************************************************************
'save user then reset user's slot
'*****************************************************************
Dim LoopC As Long
Dim Map As Integer
Dim Name As String

    If DEBUG_PacketFlood = True Then Exit Sub

    UserList(UserIndex).Flags.Disconnecting = 0

    'Save temps
    Map = (UserList(UserIndex).Pos.Map)
    Name = UserList(UserIndex).Name

    'Set logged to false
    UserList(UserIndex).Flags.UserLogged = 0

    'Save user
    Save_User UserList(UserIndex)

    'Erase user's character
    UserList(UserIndex).Char.Body = 0
    UserList(UserIndex).Char.Head = 0
    UserList(UserIndex).Char.Wings = 0
    UserList(UserIndex).Char.Heading = 0
    UserList(UserIndex).Char.HeadHeading = 0

    If UserList(UserIndex).Char.CharIndex > 0 Then User_EraseChar UserIndex

    'Remove user from connection groups and map
    If Map > 0 Then
        If Map < NumMaps Then
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
            If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0
            If MapInfo(Map).NumUsers Then
                For LoopC = 1 To MapInfo(Map).NumUsers + 1
                    If ConnectionGroups(Map).UserIndex(LoopC) = UserIndex Then Exit For
                Next LoopC
                For LoopC = LoopC To MapInfo(Map).NumUsers
                    ConnectionGroups(Map).UserIndex(LoopC) = ConnectionGroups(Map).UserIndex(LoopC + 1)
                Next LoopC
                ReDim Preserve ConnectionGroups(Map).UserIndex(1 To MapInfo(Map).NumUsers)
            Else
                ReDim ConnectionGroups(Map).UserIndex(0)
            End If
        End If
    End If

    'Clear the user from memory
    ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))

    'update last user
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).Flags.UserLogged = 1
            LastUser = LastUser - 1
            If LastUser = 0 Then Exit Do
        Loop
    End If

    'Update number of users
    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    Server_RefreshUserListBox

    'Send log off phrase
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 68
    ConBuf.Put_String Name
    Data_Send ToAll, 0, ConBuf.Get_Buffer

End Sub

Sub User_Connect(ByVal UserIndex As Integer, ByVal Name As String, ByVal Password As String, Optional ByVal NewUser As Boolean = False)

'*****************************************************************
'Reads the users .chr file and loads into Userlist array
'*****************************************************************
Dim TempPass As String
Dim CharIndex As Integer
Dim TempConnID As Long
Dim LoopC As Byte
Dim Count As Byte
    
    'Make sure the user is not in use
    If UserList(UserIndex).Flags.UserLogged Then Exit Sub
    
    'Check for valid name and password
    If Server_LegalString(Name) = False Then Exit Sub
    If Server_LegalString(Password) = False Then Exit Sub
    If Len(Name) > 10 Then Exit Sub
    If Len(Password) > 10 Then Exit Sub

    'Set up the variables - we only have to do this if not a new char
    If NewUser = False Then
        TempConnID = UserList(UserIndex).ConnID 'Transfer over the ConnID or else it'd be 0
        ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))
        UserList(UserIndex).ConnID = TempConnID
        ReDim UserList(UserIndex).SendBuffer(0)
        Set UserList(UserIndex).Stats = New UserStats
        UserList(UserIndex).Stats.UserIndex = UserIndex
    End If
    
    'Set the user as logged in
    UserList(UserIndex).Flags.UserLogged = 1

    'Check to see is user already logged with name
    If Server_CheckForSameName(UserIndex, Name) = True Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 79
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_SendBuffer UserIndex
        Exit Sub
        
    End If

    'Check for character file
    If Server_UserExist(Name) = False Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 80
        ConBuf.Put_String Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_SendBuffer UserIndex
        Exit Sub
        
    End If
    
    'Get the password
    DB_RS.Open "SELECT password FROM users WHERE `name`='" & Name & "'", DB_Conn, adOpenStatic, adLockOptimistic
    If DB_PasswordKey <> "" Then
        TempPass = Encryption_RC4_DecryptString(Trim$(DB_RS(0)), DB_PasswordKey)
    Else
        
        'Error getting the password
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 81
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_SendBuffer UserIndex
        Exit Sub
        
    End If
    DB_RS.Close
    
    'Check password
    If Password <> TempPass Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 82
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_SendBuffer UserIndex
        Exit Sub
        
    End If

    'Load character information from file
    Load_User UserList(UserIndex), Name

    'Get closest legal pos
    Server_ClosestLegalPos UserList(UserIndex).Pos, UserList(UserIndex).Pos
    If Server_LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, 0) = False Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 83
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_SendBuffer UserIndex
        Exit Sub
        
    End If
    
    'Clear the idle and last packet counter
    UserList(UserIndex).Counters.IdleCount = timeGetTime
    UserList(UserIndex).Counters.LastPacket = timeGetTime

    'Set name
    UserList(UserIndex).Name = Name

    'Set the password
    UserList(UserIndex).Password = Password

    'Set switching map flag
    UserList(UserIndex).Flags.SwitchingMaps = 1
    
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
    Data_SendBuffer UserIndex

    'Update inventory
    User_UpdateInv True, UserIndex, 0

    'Update number of users
    If UserIndex > LastUser Then LastUser = UserIndex
    NumUsers = NumUsers + 1

    'Update map and connection groups data
    MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

    'Check if it's the first user on the map
    If MapInfo(UserList(UserIndex).Pos.Map).NumUsers = 1 Then
        ReDim ConnectionGroups(UserList(UserIndex).Pos.Map).UserIndex(1 To 1)
    Else
        ReDim Preserve ConnectionGroups(UserList(UserIndex).Pos.Map).UserIndex(1 To MapInfo(UserList(UserIndex).Pos.Map).NumUsers)
    End If
    ConnectionGroups(UserList(UserIndex).Pos.Map).UserIndex(MapInfo(UserList(UserIndex).Pos.Map).NumUsers) = UserIndex

    'Give user a charindex
    CharIndex = Server_NextOpenCharIndex
    UserList(UserIndex).Char.CharIndex = CharIndex
    CharList(CharIndex).Index = UserIndex
    CharList(CharIndex).CharType = CharType_PC

    'Show Character to others
    User_MakeChar ToMap, UserIndex, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y

    'Refresh list box and send log on string
    Server_RefreshUserListBox

    'Send welcome message
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Welcome to vbGORE Version 0.2.1!"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "For help, please type /help"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Ctrl+W for inventory, Ctrl+S for stats."
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Please visit our site at:"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "www.vbgore.com"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info

    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "Questions, code, help, etc can be found there. :)"
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
    ConBuf.Put_Byte DataCode.Server_PlaySound
    ConBuf.Put_Byte SOUND_WARP
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 72
    ConBuf.Put_String UserList(UserIndex).Name
    Data_Send ToAll, 0, ConBuf.Get_Buffer, , PP_Connect

End Sub

Sub User_ConnectNew(ByVal UserIndex As Integer, ByVal Name As String, ByVal Password As String, ByVal Body As Integer, ByVal Head As Integer)

'*****************************************************************
'Opens a new user. Loads default vars, saves then calls connectuser
'*****************************************************************
Dim TempConnID As Long
Dim i As Byte

    'Make sure the user is not in use
    If UserList(UserIndex).Flags.UserLogged Then Exit Sub
    
    'Check for valid name and password
    If Server_LegalString(Name) = False Then Exit Sub
    If Server_LegalString(Password) = False Then Exit Sub
    If Len(Name) > 10 Then Exit Sub
    If Len(Password) > 10 Then Exit Sub
    If Len(Name) < 3 Then Exit Sub
    If Len(Password) < 3 Then Exit Sub
    
    'Check to remove the Packetflooder char
    If DEBUG_PacketFlood Then
        On Error GoTo ErrOut    'Catch the character path being open already preventing errors
        Server_UserExist Name, True
        On Error GoTo 0
    End If

    'Clear user character
    TempConnID = UserList(UserIndex).ConnID 'Transfer over the ConnID or else it'd be 0
    ZeroMemory UserList(UserIndex), Len(UserList(UserIndex))
    UserList(UserIndex).ConnID = TempConnID
    ReDim UserList(UserIndex).SendBuffer(0)
    
    'Set the user as logged in
    UserList(UserIndex).Flags.UserLogged = 1

    'Check for character file
    If Server_UserExist(Name) Then

        'Tell the client we're disconnecting
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_String 84
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Data_SendBuffer UserIndex
        Exit Sub
        
    End If

    'Set the stats class
    Set UserList(UserIndex).Stats = New UserStats
    UserList(UserIndex).Stats.UserIndex = UserIndex

    'Set the user's variables
    UserList(UserIndex).Name = Name
    UserList(UserIndex).Password = Password
    UserList(UserIndex).Char.Heading = SOUTH
    UserList(UserIndex).Char.HeadHeading = SOUTH
    UserList(UserIndex).Char.Head = Head
    UserList(UserIndex).Char.Hair = 1
    UserList(UserIndex).Char.Body = Body
    UserList(UserIndex).Counters.IdleCount = timeGetTime
    UserList(UserIndex).Stats.UserIndex = UserIndex
    UserList(UserIndex).Pos.Map = StartPos.Map
    UserList(UserIndex).Pos.x = StartPos.x
    UserList(UserIndex).Pos.Y = StartPos.Y

    'Set the user's starting stats
    UserList(UserIndex).Stats.BaseStat(SID.ELV) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Str) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Agi) = 1
    UserList(UserIndex).Stats.BaseStat(SID.Mag) = 1
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
    Save_User UserList(UserIndex)

    'Write a test message to the user
    For i = 1 To 5
        Server_WriteMail -1, UserList(UserIndex).Name, "Test Message", "This is a test message that simply shows the pwnification of the mailing system.", "1,1,1,1,1,1,1,1,1,1", "1,2,3,4,94,95,96,97,98,99"
    Next i
    
    'Disconnect the user so we can re-connect them through the connect sub
    UserList(UserIndex).Flags.UserLogged = 0

    'Go through regular connecting routine
    User_Connect UserIndex, Name, Password, True
    
ErrOut:

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:47)  Decl: 15  Code: 2436  Total: 2451 Lines
':) CommentOnly: 461 (18.8%)  Commented: 17 (0.7%)  Empty: 541 (22.1%)  Max Logic Depth: 7
