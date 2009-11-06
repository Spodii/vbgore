Attribute VB_Name = "TCP"
Option Explicit

'Size of our buffer - the lower the value, the faster the send/recv, if more is sent or
'recieved then buffer size, then becomes slower (recommended to leave as is)
Public Const TCPBufferSize As Long = 512

Sub Data_Server_SetCharSpeed(ByRef rBuf As DataBuffer)
'*********************************************
'Update a char's speed so we can move them the right speed
'<CharIndex(I)><Speed(B)>
'*********************************************
Dim CharIndex As Integer
Dim Speed As Byte

    CharIndex = rBuf.Get_Integer
    Speed = rBuf.Get_Byte
    CharList(CharIndex).Speed = Speed

End Sub

Sub Data_Server_Message(ByRef rBuf As DataBuffer)
'*********************************************
'Server sending a common message to client (reccomended you send
' as many messages as possible via this method to save bandwidth)
'<MessageID(B)><...depends on the message>
'*********************************************
Dim MessageID As Byte
Dim TempStr As String
Dim Str1 As String
Dim Str2 As String
Dim Lng1 As Long
Dim Int1 As Integer
Dim Int2 As Integer
Dim Byt1 As Byte

    'Get the message ID
    MessageID = rBuf.Get_Byte
    
    'Check what to do depending on the message ID
    '*** Please refer to the language file for the description of the numbers ***
    Select Case MessageID
        Case 1
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(1), "<npcname>", Str1), FontColor_Info
        Case 2
            Engine_AddToChatTextBuffer Message(2), FontColor_Fight
        Case 3
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(3), "<exp>", Lng1), FontColor_Info
        Case 4
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(4), "<gold>", Lng1), FontColor_Info
        Case 5
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(5), "<skill>", Str1), FontColor_Info
        Case 6
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(6), "<skill>", Str1), FontColor_Info
        Case 7
            Engine_AddToChatTextBuffer Message(7), FontColor_Quest
        Case 8
            Engine_AddToChatTextBuffer Message(8), FontColor_Quest
        Case 9
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(9), "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<npcname>", Str1), FontColor_Quest
        Case 10
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(10), "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<objname>", Str1), FontColor_Quest
        Case 11
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            Int2 = rBuf.Get_Integer
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(11), "<npcamount>", Int1)
            TempStr = Replace$(TempStr, "<npcname>", Str1)
            TempStr = Replace$(TempStr, "<objamount>", Int2)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<objname>", Str2), FontColor_Quest
        Case 12
            Engine_AddToChatTextBuffer Message(12), FontColor_Quest
        Case 13
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(13), "<name>", Str1), FontColor_Info
        Case 14
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(14), "<cost>", Lng1), FontColor_Info
        Case 15
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(15), "<sender>", Str1), FontColor_Info
        Case 16
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(16), "<receiver>", Str1), FontColor_Info
        Case 17
            Engine_AddToChatTextBuffer Message(17), FontColor_Info
        Case 18
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(18), "<sender>", Str1), FontColor_Info
        Case 19
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(19), "<receiver>", Str1), FontColor_Info
        Case 20
            Engine_AddToChatTextBuffer Message(20), FontColor_Info
        Case 21
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(21), "<cost>", Lng1), FontColor_Info
        Case 22
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(22), "<name>", Str1), FontColor_Info
        Case 23
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(23), "<name>", Str1), FontColor_Info
        Case 24
            Engine_AddToChatTextBuffer Message(24), FontColor_Info
        Case 25
            Engine_AddToChatTextBuffer Message(25), FontColor_Info
        Case 26
            Engine_AddToChatTextBuffer Message(26), FontColor_Info
        Case 27
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(27), "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<name>", Str1), FontColor_Info
        Case 28
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(28), "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<name>", Str1), FontColor_Info
        Case 29
            Engine_AddToChatTextBuffer Message(29), FontColor_Info
        Case 30
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(30), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<desc>", Str2), FontColor_Info
        Case 31
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(31), "<name>", Str1), FontColor_Info
        Case 32
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(32), "<name>", Str1), FontColor_Info
        Case 33
            Engine_AddToChatTextBuffer Message(33), FontColor_Info
        Case 34
            Engine_AddToChatTextBuffer Message(34), FontColor_Info
        Case 35
            Byt1 = rBuf.Get_Byte
            Engine_AddToChatTextBuffer Replace$(Message(35), "<amount>", Byt1), FontColor_Info
        Case 36
            Engine_AddToChatTextBuffer Message(36), FontColor_Info
        Case 37
            Engine_AddToChatTextBuffer Message(37), FontColor_Info
        Case 38
            Engine_AddToChatTextBuffer Message(38), FontColor_Info
        Case 39
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(39), "<skill>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<name>", Str2), FontColor_Info
        Case 40
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(40), "<name>", Str1), FontColor_Info
        Case 41
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(40), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<power>", Int1), FontColor_Info
        Case 42
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(42), "<name>", Str1), FontColor_Info
        Case 43
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(43), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<power>", Int1), FontColor_Info
        Case 44
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(44), "<name>", Str1), FontColor_Info
        Case 45
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(45), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<power>", Int1), FontColor_Info
        Case 46
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(46), "<name>", Str1), FontColor_Info
        Case 47
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(47), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<power>", Int1), FontColor_Info
        Case 48
            Engine_AddToChatTextBuffer Message(48), FontColor_Info
        Case 49
            Engine_AddToChatTextBuffer Message(49), FontColor_Info
        Case 50
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(50), "<name>", Str1), FontColor_Info
        Case 51
            Engine_AddToChatTextBuffer Message(51), FontColor_Info
        Case 52
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(52), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<message>", Str2), FontColor_Talk
        Case 53
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(53), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<message>", Str2), FontColor_Talk
        Case 54
            Str1 = rBuf.Get_String
            Byt1 = rBuf.Get_Byte
            TempStr = Replace$(Message(54), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<value>", Byt1), FontColor_Info
        Case 55
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(55), "<value>", Lng1), FontColor_Info
        Case 56
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(56), "<name>", Str1), FontColor_Info
        Case 57
            Engine_AddToChatTextBuffer Message(57), FontColor_Info
        Case 58
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(58), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<amount>", Int1), FontColor_Info
        Case 59
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            Int2 = rBuf.Get_Integer
            TempStr = Replace$(Message(59), "<name>", Str1)
            TempStr = Replace$(TempStr, "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<leftover>", Int2), FontColor_Info
        Case 60
            Engine_AddToChatTextBuffer Message(60), FontColor_Info
        Case 61
            Engine_AddToChatTextBuffer Message(61), FontColor_Info
        Case 62
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(62), "<level>", Lng1), FontColor_Info
        Case 63
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(63), "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<name>", Str1), FontColor_Info
        Case 64
            Engine_AddToChatTextBuffer Message(64), FontColor_Info
        Case 65
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(65), "<amount>", Int1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<name>", Str1), FontColor_Info
        Case 66
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            Int2 = rBuf.Get_Integer
            TempStr = Replace$(Message(66), "<amount>", Int1)
            TempStr = Replace$(TempStr, "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<leftover>", Int2), FontColor_Info
        Case 67
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            Lng1 = rBuf.Get_Long
            TempStr = Replace$(Message(67), "<amount>", Int1)
            TempStr = Replace$(TempStr, "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<total>", Lng1), FontColor_Info
        Case 68
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(68), "<name>", Str1), FontColor_Info
        Case 69
            Engine_AddToChatTextBuffer Message(69), FontColor_Info
        Case 70
            Engine_AddToChatTextBuffer Message(70), FontColor_Info
        Case 71
            Byt1 = rBuf.Get_Byte
            Engine_AddToChatTextBuffer Replace$(Message(71), "<value>", Byt1), FontColor_Info
        Case 72
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(72), "<name>", Str1), FontColor_Info
        Case 73
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(73), "<name>", Str1), FontColor_Info
        Case 74
            Int1 = rBuf.Get_Integer
            Int2 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(74), "<amount>", Int1)
            TempStr = Replace$(TempStr, "<total>", Int2)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<name>", Str1), FontColor_Quest
        Case 75
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(75), "<name>", Str1), FontColor_Info
        Case 76
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(76), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<message>", Str2), FontColor_Info
        Case 77
            Str1 = rBuf.Get_String
            Str2 = rBuf.Get_String
            TempStr = Replace$(Message(77), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<gm>", Str2), FontColor_Info
        Case 78
            Int1 = rBuf.Get_Integer
            Engine_AddToChatTextBuffer Replace$(Message(78), "<value>", Int1), FontColor_Info
        Case 79
            MsgBox Message(79)
        Case 80
            Str1 = rBuf.Get_String
            MsgBox Replace$(Message(80), "<name>", Str1)
        Case 81
            MsgBox Message(81)
        Case 82
            MsgBox Message(82)
        Case 83
            MsgBox Message(83)
        Case 84
            MsgBox Message(84)
        Case 85
            MsgBox Message(85)
        Case 86
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(86), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<amount>", Int1), FontColor_Info
    End Select

End Sub

Sub Data_Server_Connect()
'*********************************************
'Server is telling the client they have successfully logged in
'<>
'*********************************************

    'Set the socket state
    SocketOpen = 1

    'Load user config
    Game_Config_Load

    'Unload the connect form
    Unload frmConnect

    'Load main form
    frmMain.Visible = True
    frmMain.Show
    frmMain.SetFocus

    'Load the engine
    Engine_Init_TileEngine

    'Get the device
    frmMain.Show
    frmMain.SetFocus
    DoEvents
    DIDevice.Acquire

    'Init the Ping timer
    frmMain.PingTmr.Enabled = True

    'Send the data
    Data_Send

End Sub

Sub Data_Server_Disconnect()
'*********************************************
'Forces the client to disconnect from the server
'<>
'*********************************************

    IsUnloading = 1

End Sub

Sub Data_Comm_Talk(ByRef rBuf As DataBuffer)

'*********************************************
'Send data to chat buffer
'<Text(S)><FontColorID(B)>
'*********************************************

Dim TempStr As String
Dim TempLng As Long
Dim TempByte As Byte

    TempStr = rBuf.Get_String
    TempByte = rBuf.Get_Byte
    
    Select Case TempByte
    Case DataCode.Comm_FontType_Fight
        TempLng = FontColor_Fight
    Case DataCode.Comm_FontType_Info
        TempLng = FontColor_Info
    Case DataCode.Comm_FontType_Quest
        TempLng = FontColor_Quest
    Case DataCode.Comm_FontType_Talk
        TempLng = FontColor_Talk
    Case Else
        TempLng = FontColor_Talk
    End Select
    Engine_AddToChatTextBuffer TempStr, TempLng

End Sub

Sub Data_Map_DoneSwitching()

'*********************************************
'Done switching maps, load engine back up
'<>
'*********************************************

    DownloadingMap = False
    EngineRun = True

End Sub

Sub Data_Map_LoadMap(ByRef rBuf As DataBuffer)

'*********************************************
'Load the map the server told us to load
'<MapNum(I)><ServerSideVersion(I)><Weather(B)>
'*********************************************

Dim MapNumInt As Integer
Dim SSV As Integer
Dim Weather As Byte
Dim TempInt As Integer

    EngineRun = False

    DownloadingMap = True
    MapNumInt = rBuf.Get_Integer
    SSV = rBuf.Get_Integer
    Weather = rBuf.Get_Byte
    If Engine_FileExist(MapPath & MapNumInt & ".map", vbNormal) Then  'Get Version Num
        Open MapPath & MapNumInt & ".map" For Binary As #1
        Seek #1, 1
        Get #1, , TempInt
        Close #1
        If TempInt = SSV Then   'Correct Version
            Game_Map_Switch MapNumInt
            sndBuf.Put_Byte DataCode.Map_DoneLoadingMap 'Tell the server we are done loading map
            MapInfo.Weather = Weather

        Else
            'Not correct version
            MsgBox "Error! Your map version is not up to date with the server's map! Please run the updater!", vbOKOnly Or vbCritical
            EngineRun = False
            IsUnloading = 1
        End If
    Else
        'Didn't find map
        MsgBox "Error! The requested map could not be found! Please run the updater!", vbOKOnly Or vbCritical
        EngineRun = False
        IsUnloading = 1
    End If

End Sub

Sub Data_Map_SendName(ByRef rBuf As DataBuffer)

'*********************************************
'Set the map name and weather
'<Name(S)><Weather(B)><Music(B)>
'*********************************************
Dim Music As Byte

    MapInfo.Name = rBuf.Get_String
    MapInfo.Weather = rBuf.Get_Byte
    
    'Change the music file if we need to
    Music = rBuf.Get_Byte
    If MapInfo.Music <> Music Then
        Engine_Music_Stop 1
        If Music <> 0 Then
            MapInfo.Music = Music
            Engine_Music_Load MusicPath & Music & ".mp3", 1
            Engine_Music_Play 1
            Engine_Music_Volume 96, 1
        End If
    End If
    
End Sub

Sub Data_Send()

'*********************************************
'Send data buffer to the server
'*********************************************
Dim TempBuffer() As Byte

    'Check that we have data to send
    If SocketOpen = 0 Then DoEvents
    If UBound(sndBuf.Get_Buffer) > 0 Then
        If SocketOpen = 0 Then DoEvents
    
        'Assign to the temp buffer
        TempBuffer() = sndBuf.Get_Buffer
        
        'Encrypt the packet
        Select Case PacketEncType
            Case PacketEncTypeXOR
                Encryption_XOR_EncryptByte TempBuffer(), PacketEncKey
            Case PacketEncTypeRC4
                Encryption_RC4_EncryptByte TempBuffer(), PacketEncKey
        End Select
    
        'Send the data
        frmMain.Socket.SendData SoxID, TempBuffer
        
        'Clear the buffer, get it ready for next use
        sndBuf.Clear
        
    End If

End Sub

Sub Data_Server_ChangeChar(ByRef rBuf As DataBuffer)

'*********************************************
'Change a character by the character index
'<CharIndex(I)><Flags(B)>(<Body(I)><Head(I)><Heading(B)><Weapon(I)><Hair(I)><Wings(I)>)
'*********************************************
Dim Flags As Byte
Dim CharIndex As Integer
Dim CharBody As Integer
Dim CharHead As Integer
Dim CharHeading As Byte
Dim CharWeapon As Integer
Dim CharHair As Integer
Dim CharWings As Integer
    
    'Get the character index we are changing
    CharIndex = rBuf.Get_Integer
    
    'Get the flags on what data we need to get
    Flags = rBuf.Get_Byte
    
    'Get the data needed
    If Flags And 1 Then
        CharBody = rBuf.Get_Integer
        CharList(CharIndex).Body = BodyData(CharBody)
    End If
    If Flags And 2 Then
        CharHead = rBuf.Get_Integer
        CharList(CharIndex).Head = HeadData(CharHead)
    End If
    If Flags And 4 Then
        CharHeading = rBuf.Get_Byte
        CharList(CharIndex).Heading = CharHeading
        CharList(CharIndex).HeadHeading = CharHeading
    End If
    If Flags And 8 Then
        CharWeapon = rBuf.Get_Integer
        CharList(CharIndex).Weapon = WeaponData(CharWeapon)
    End If
    If Flags And 16 Then
        CharHair = rBuf.Get_Integer
        CharList(CharIndex).Hair = HairData(CharHair)
    End If
    If Flags And 32 Then
        CharWings = rBuf.Get_Integer
        CharList(CharIndex).Wings = WingData(CharWings)
    End If
    
End Sub

Sub Data_Server_CharHP(ByRef rBuf As DataBuffer)

'*********************************************
'Set the character HP
'<HP(B)><CharIndex(I)>
'*********************************************

Dim CharIndex As Integer
Dim HP As Byte

    HP = rBuf.Get_Byte
    CharIndex = rBuf.Get_Byte

    If CharIndex > LastChar Then Exit Sub

    CharList(CharIndex).HealthPercent = HP

End Sub

Sub Data_Server_CharMP(ByRef rBuf As DataBuffer)

'*********************************************
'Set the character MP
'<MP(B)><CharIndex(I)>
'*********************************************

Dim CharIndex As Integer
Dim MP As Byte

    MP = rBuf.Get_Byte
    CharIndex = rBuf.Get_Byte

    If CharIndex > LastChar Then Exit Sub

    CharList(CharIndex).ManaPercent = MP

End Sub

Sub Data_Server_EraseChar(ByRef rBuf As DataBuffer)

'*********************************************
'Erase a character by the character index
'<CharIndex(I)>
'*********************************************

    Engine_Char_Erase rBuf.Get_Integer

End Sub

Sub Data_Server_EraseObject(ByRef rBuf As DataBuffer)

'*********************************************
'Erase an object on the object layer
'<X(B)><Y(B)>
'*********************************************

Dim j As Integer
Dim X As Byte
Dim Y As Byte

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Loop through until we find the object on (X,Y) then kill it
    For j = 1 To LastObj
        If OBJList(j).Pos.X = X Then
            If OBJList(j).Pos.Y = Y Then
                Engine_OBJ_Erase j
                Exit Sub
            End If
        End If
    Next j

End Sub

Sub Data_Server_IconBlessed(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show blessed icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub
    
    CharList(CharIndex).CharStatus.Blessed = State

End Sub

Sub Data_Server_IconCursed(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show cursed icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub
    
    CharList(CharIndex).CharStatus.Cursed = State

End Sub

Sub Data_Server_IconIronSkin(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show ironskinned icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub
    
    CharList(CharIndex).CharStatus.IronSkinned = State

End Sub

Sub Data_Server_IconProtected(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show protected icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub
    
    CharList(CharIndex).CharStatus.Protected = State

End Sub

Sub Data_Server_IconSpellExhaustion(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show spell exhaustion icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer

    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub

    CharList(CharIndex).CharStatus.Exhausted = State

End Sub

Sub Data_Server_IconStrengthened(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show strengthened icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub
    
    CharList(CharIndex).CharStatus.Strengthened = State

End Sub

Sub Data_Server_IconWarCursed(ByRef rBuf As DataBuffer)

'*********************************************
'Hide/show warcursed icon
'<State(B)><CharIndex(I)>
'*********************************************

Dim State As Byte
Dim CharIndex As Integer

    State = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub

    CharList(CharIndex).CharStatus.WarCursed = State

End Sub

Sub Data_Server_Mailbox(ByRef rBuf As DataBuffer)

'*********************************************
'Recieve the list of messages from a mailbox
'Loop: <New(B)><WriterName(S)><Date(S)><Subject(S)>...<EndFlag(B)>
'*********************************************

Dim NewB As Byte
Dim WName As String
Dim SDate As String
Dim Subj As String

    ShowGameWindow(MailboxWindow) = 1
    
    SelMessage = 0
    LastClickedWindow = MailboxWindow
    MailboxListBuffer = vbNullString
    Do
        NewB = rBuf.Get_Byte
        If NewB = 255 Then Exit Do  'If 1 or 0, it is a message, if 255, it is the EndFlag
        WName = rBuf.Get_String
        SDate = rBuf.Get_String
        Subj = rBuf.Get_String
        MailboxListBuffer = MailboxListBuffer & IIf(NewB, "New - ", "Old - ") & Subj & " - " & WName & " - " & SDate & vbCrLf
    Loop

End Sub

Sub Data_Server_MailItemInfo(ByRef rBuf As DataBuffer)

'*********************************************
'Retrieve information on the selected mail item
'<Name(S)><Amount(I)>
'*********************************************

Dim Name As String
Dim Amount As Integer

    Name = rBuf.Get_String

    Amount = rBuf.Get_Integer
    Engine_SetItemDesc Name, Amount

End Sub

Sub Data_Server_MailItemRemove(ByRef rBuf As DataBuffer)

'*********************************************
'Remove item from mailbox
'<ItemIndex(B)>
'*********************************************

Dim ItemIndex As Byte

    ItemIndex = rBuf.Get_Byte

    ReadMailData.Obj(ItemIndex) = 0

End Sub

Sub Data_Server_MailMessage(ByRef rBuf As DataBuffer)

'*********************************************
'Recieve message that was requested to be read
'<Message(S-EX)><Subject(S)><WriterName(S)> Loop: <ObjGrhIndex(I)>
'*********************************************

Dim X As Byte

    ShowGameWindow(MailboxWindow) = 0

    ShowGameWindow(ViewMessageWindow) = 1
    LastClickedWindow = ViewMessageWindow
    ReadMailData.Message = rBuf.Get_StringEX
    ReadMailData.Message = Engine_WordWrap(ReadMailData.Message, 60)
    ReadMailData.Subject = rBuf.Get_String
    ReadMailData.WriterName = rBuf.Get_String
    For X = 1 To MaxMailObjs
        ReadMailData.Obj(X) = rBuf.Get_Integer
    Next X

End Sub

Sub Data_Server_MakeChar(ByRef rBuf As DataBuffer)

'*********************************************
'Create a character and set their information
'<Body(I)><Head(I)><Heading(B)><CharIndex(I)><X(B)><Y(B)><Speed(B)><Name(S)><Weapon(I)><Hair(I)><Wings(I)><HP%(B)><MP%(B)>
'*********************************************

Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim CharIndex As Integer
Dim X As Byte
Dim Y As Byte
Dim Speed As Byte
Dim Name As String
Dim Weapon As Integer
Dim Hair As Integer
Dim Wings As Integer
Dim HP As Byte
Dim MP As Byte

'Retrieve all the information

    Body = rBuf.Get_Integer
    Head = rBuf.Get_Integer
    Heading = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Speed = rBuf.Get_Byte
    Name = rBuf.Get_String
    Weapon = rBuf.Get_Integer
    Hair = rBuf.Get_Integer
    Wings = rBuf.Get_Integer
    HP = rBuf.Get_Byte
    MP = rBuf.Get_Byte

    'Create the character
    Engine_Char_Make CharIndex, Body, Head, Heading, X, Y, Speed, Name, Weapon, Hair, Wings, HP, MP

End Sub

Sub Data_Server_MakeObject(ByRef rBuf As DataBuffer)

'*********************************************
'Create an object on the object layer
'<GrhIndex(L)><X(B)><Y(B)>
'*********************************************

Dim GrhIndex As Long
Dim X As Byte
Dim Y As Byte

    'Get the values
    GrhIndex = rBuf.Get_Long
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Create the object
    If GrhIndex > 0 Then Engine_OBJ_Create GrhIndex, X, Y

End Sub

Sub Data_Server_MoveChar(ByRef rBuf As DataBuffer)

'*********************************************
'Move a character
'<CharIndex(I)><X(B)><Y(B)><Heading(B)>
'*********************************************

Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer
Dim Heading As Byte

    CharIndex = rBuf.Get_Integer

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Heading = rBuf.Get_Byte
    
    'Make sure the char is the right starting position
    Select Case Heading
        Case NORTH: nX = 0: nY = -1
        Case EAST: nX = 1: nY = 0
        Case SOUTH: nX = 0: nY = 1
        Case WEST: nX = -1: nY = 0
        Case NORTHEAST: nX = 1: nY = -1
        Case SOUTHEAST: nX = 1: nY = 1
        Case SOUTHWEST: nX = -1: nY = 1
        Case NORTHWEST: nX = -1: nY = -1
    End Select
    CharList(CharIndex).Pos.X = X - nX
    CharList(CharIndex).Pos.Y = Y - nY
    
    Engine_Char_Move_ByPos CharIndex, X, Y

End Sub

Sub Data_Server_Ping()

'*********************************************
'We retrieved the ping response, so calculate how long it took
'<>
'*********************************************

    Ping = timeGetTime - PingSTime

    'Reset the unreturned ping count
    NonRetPings = 0

End Sub

Sub Data_Server_PlaySound(ByRef rBuf As DataBuffer)

'*********************************************
'Play a wave file
'<WaveNum(B)>
'*********************************************
Dim WaveNum As Byte

    WaveNum = rBuf.Get_Byte

    Engine_Sound_Play DSBuffer(WaveNum), DSBPLAY_DEFAULT

End Sub

Sub Data_Server_PlaySound3D(ByRef rBuf As DataBuffer)

'*********************************************
'Play a wave file with 3D effect
'<WaveNum(B)><X(B)><Y(B)>
'*********************************************
Dim WaveNum As Byte
Dim X As Integer
Dim Y As Integer

    WaveNum = rBuf.Get_Byte
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    
    Engine_Sound_Play3D WaveNum, X, Y

End Sub

Sub Data_Server_SetCharDamage(ByRef rBuf As DataBuffer)

'*********************************************
'Damage a character and display it
'<CharIndex(I)><Damage(I)>
'*********************************************

Dim CharIndex As Integer
Dim Damage As Integer

    CharIndex = rBuf.Get_Integer
    Damage = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub

    'Create the blood
    Engine_Blood_Create CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y

    'Create the damage
    Engine_Damage_Create CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y, Damage

End Sub

Sub Data_Server_SetUserPosition(ByRef rBuf As DataBuffer)

'*********************************************
'Set the user's position
'<X(B)><Y(B)>
'*********************************************

Dim X As Byte
Dim Y As Byte

'Get the position

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Check for a valid range
    If X < XMinMapSize Then Exit Sub
    If X > XMaxMapSize Then Exit Sub
    If Y < YMinMapSize Then Exit Sub
    If Y > YMaxMapSize Then Exit Sub

    'Update the user's position
    UserPos.X = X
    UserPos.Y = Y
    CharList(UserCharIndex).Pos = UserPos

End Sub

Sub Data_Server_UserCharIndex(ByRef rBuf As DataBuffer)

'*********************************************
'Set the user character index
'<CharIndex(I)>
'*********************************************

    'Retrieve the index of the user's character
    UserCharIndex = rBuf.Get_Integer
    UserPos = CharList(UserCharIndex).Pos
    
    'Update the map-bound sound effects
    Engine_Sound_UpdateMap

End Sub

Sub Data_User_AggressiveFace(ByRef rBuf As DataBuffer)

'*********************************************
'Turn on/off an aggressive face for a character
'<CharIndex(I)><IsOn(B)>
'*********************************************

Dim CharIndex As Integer
Dim IsOn As Byte

    CharIndex = rBuf.Get_Integer

    IsOn = rBuf.Get_Byte
    CharList(CharIndex).Aggressive = IsOn

End Sub

Sub Data_User_Attack(ByRef rBuf As DataBuffer)

'*********************************************
'Change character animation to attack animation
'<CharIndex(I)>
'*********************************************

Dim CharIndex As Integer

    CharIndex = rBuf.Get_Integer
    
    'Check for invalid conditions
    If CharIndex > UBound(CharList()) Then Exit Sub

    CharList(CharIndex).ActionIndex = 2

End Sub

Sub Data_User_BaseStat(ByRef rBuf As DataBuffer)

'*********************************************
'Update base stat
'<StatID(B)><Value(L)>
'*********************************************

Dim StatID As Byte

    StatID = rBuf.Get_Byte
    BaseStats(StatID) = rBuf.Get_Long

End Sub

Sub Data_User_Blink(ByRef rBuf As DataBuffer)

'*********************************************
'Make a character blink
'<CharIndex(I)>
'*********************************************

Dim CharIndex As Integer

    CharIndex = rBuf.Get_Integer

    If CharIndex > LastChar Then Exit Sub
    If CharIndex <= 0 Then Exit Sub

    CharList(CharIndex).StartBlinkTimer = 0
    CharList(CharIndex).BlinkTimer = 0

End Sub

Sub Data_User_CastSkill(ByRef rBuf As DataBuffer)

'*********************************************
'User casted a skill
'<SkillID(B)><CasterIndex(I)><TargetIndex(I)>
'*********************************************

Dim CasterIndex As Integer
Dim TargetIndex As Integer
Dim TempIndex As Integer
Dim SkillID As Byte
Dim X As Long
Dim Y As Long

    SkillID = rBuf.Get_Byte
    CasterIndex = rBuf.Get_Integer
    TargetIndex = rBuf.Get_Integer

    Select Case SkillID

    Case SkID.Heal

        'Set the position
        X = CharList(CasterIndex).RealPos.X + 16
        Y = CharList(CasterIndex).RealPos.Y

        'If not casted on self, bind to character
        If TargetIndex <> CasterIndex Then
            TempIndex = Effect_Heal_Begin(X, Y, 3, 120, 1)
            Effect(TempIndex).BindToChar = TargetIndex
            Effect(TempIndex).BindSpeed = 7
        Else
            TempIndex = Effect_Heal_Begin(X, Y, 3, 120, 0)
        End If

    Case SkID.Protection

        'Create the effect at (not bound to) the target character
        X = CharList(TargetIndex).RealPos.X + 16
        Y = CharList(TargetIndex).RealPos.Y
        Effect_Protection_Begin X, Y, 11, 120, 40, 15

    Case SkID.Strengthen

        'Create the effect at (not bound to) the target character
        X = CharList(TargetIndex).RealPos.X + 16
        Y = CharList(TargetIndex).RealPos.Y
        Effect_Strengthen_Begin X, Y, 12, 120, 40, 15

    Case SkID.Bless

        'Create the effect at (not bound to) the target character
        X = CharList(TargetIndex).RealPos.X + 16
        Y = CharList(TargetIndex).RealPos.Y
        Effect_Bless_Begin X, Y, 3, 120, 40, 15

    Case SkID.SpikeField

        'Create the spike field depending on the direction the user is facing
        X = CharList(CasterIndex).Pos.X
        Y = CharList(CasterIndex).Pos.Y
        If CharList(CasterIndex).HeadHeading = NORTH Then
            Engine_Effect_Create X - 1, Y + 1, 59
            Engine_Effect_Create X, Y + 1, 59
            Engine_Effect_Create X + 1, Y + 1, 59

            Engine_Effect_Create X - 2, Y, 59
            Engine_Effect_Create X - 1, Y, 59
            Engine_Effect_Create X, Y, 59
            Engine_Effect_Create X + 1, Y, 59
            Engine_Effect_Create X + 2, Y, 59

            Engine_Effect_Create X - 2, Y - 1, 59
            Engine_Effect_Create X - 1, Y - 1, 59
            Engine_Effect_Create X, Y - 1, 59
            Engine_Effect_Create X + 1, Y - 1, 59
            Engine_Effect_Create X + 2, Y - 1, 59

            Engine_Effect_Create X - 2, Y - 2, 59
            Engine_Effect_Create X - 1, Y - 2, 59
            Engine_Effect_Create X, Y - 2, 59
            Engine_Effect_Create X + 1, Y - 2, 59
            Engine_Effect_Create X + 2, Y - 2, 59

            Engine_Effect_Create X - 1, Y - 3, 59
            Engine_Effect_Create X, Y - 3, 59
            Engine_Effect_Create X + 1, Y - 3, 59

            Engine_Effect_Create X, Y - 4, 59
        ElseIf CharList(CasterIndex).HeadHeading = EAST Then
            Engine_Effect_Create X - 1, Y - 1, 59
            Engine_Effect_Create X - 1, Y, 59
            Engine_Effect_Create X - 1, Y + 1, 59

            Engine_Effect_Create X, Y - 2, 59
            Engine_Effect_Create X, Y - 1, 59
            Engine_Effect_Create X, Y, 59
            Engine_Effect_Create X, Y + 1, 59
            Engine_Effect_Create X, Y + 2, 59

            Engine_Effect_Create X + 1, Y - 2, 59
            Engine_Effect_Create X + 1, Y - 1, 59
            Engine_Effect_Create X + 1, Y, 59
            Engine_Effect_Create X + 1, Y + 1, 59
            Engine_Effect_Create X + 1, Y + 2, 59

            Engine_Effect_Create X + 2, Y - 2, 59
            Engine_Effect_Create X + 2, Y - 1, 59
            Engine_Effect_Create X + 2, Y, 59
            Engine_Effect_Create X + 2, Y + 1, 59
            Engine_Effect_Create X + 2, Y + 2, 59

            Engine_Effect_Create X + 3, Y - 1, 59
            Engine_Effect_Create X + 3, Y, 59
            Engine_Effect_Create X + 3, Y + 1, 59

            Engine_Effect_Create X + 4, Y, 59
        ElseIf CharList(CasterIndex).HeadHeading = SOUTH Then
            Engine_Effect_Create X - 1, Y - 1, 59
            Engine_Effect_Create X, Y - 1, 59
            Engine_Effect_Create X + 1, Y - 1, 59

            Engine_Effect_Create X - 2, Y, 59
            Engine_Effect_Create X - 1, Y, 59
            Engine_Effect_Create X, Y, 59
            Engine_Effect_Create X + 1, Y, 59
            Engine_Effect_Create X + 2, Y, 59

            Engine_Effect_Create X - 2, Y + 1, 59
            Engine_Effect_Create X - 1, Y + 1, 59
            Engine_Effect_Create X, Y + 1, 59
            Engine_Effect_Create X + 1, Y + 1, 59
            Engine_Effect_Create X + 2, Y + 1, 59

            Engine_Effect_Create X - 2, Y + 2, 59
            Engine_Effect_Create X - 1, Y + 2, 59
            Engine_Effect_Create X, Y + 2, 59
            Engine_Effect_Create X + 1, Y + 2, 59
            Engine_Effect_Create X + 2, Y + 2, 59

            Engine_Effect_Create X - 1, Y + 3, 59
            Engine_Effect_Create X, Y + 3, 59
            Engine_Effect_Create X + 1, Y + 3, 59

            Engine_Effect_Create X, Y + 4, 59
        ElseIf CharList(CasterIndex).HeadHeading = WEST Then
            Engine_Effect_Create X + 1, Y - 1, 59
            Engine_Effect_Create X + 1, Y, 59
            Engine_Effect_Create X + 1, Y + 1, 59

            Engine_Effect_Create X, Y - 2, 59
            Engine_Effect_Create X, Y - 1, 59
            Engine_Effect_Create X, Y, 59
            Engine_Effect_Create X, Y + 1, 59
            Engine_Effect_Create X, Y + 2, 59

            Engine_Effect_Create X - 1, Y - 2, 59
            Engine_Effect_Create X - 1, Y - 1, 59
            Engine_Effect_Create X - 1, Y, 59
            Engine_Effect_Create X - 1, Y + 1, 59
            Engine_Effect_Create X - 1, Y + 2, 59

            Engine_Effect_Create X - 2, Y - 2, 59
            Engine_Effect_Create X - 2, Y - 1, 59
            Engine_Effect_Create X - 2, Y, 59
            Engine_Effect_Create X - 2, Y + 1, 59
            Engine_Effect_Create X - 2, Y + 2, 59

            Engine_Effect_Create X - 3, Y - 1, 59
            Engine_Effect_Create X - 3, Y, 59
            Engine_Effect_Create X - 3, Y + 1, 59

            Engine_Effect_Create X - 4, Y, 59
        End If

    End Select

End Sub

Sub Data_User_Emote(ByRef rBuf As DataBuffer)

'*********************************************
'A character uses an emoticon
'<EmoticonIndex(B)><CharIndex(I)>
'*********************************************

Dim EmoticonIndex As Byte
Dim CharIndex As Integer

    EmoticonIndex = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer

    'Check for valid range (dont worry about emoticon range - it just wont be set if it cant find a valid emoticon ID
    If CharIndex <= 0 Then Exit Sub
    If CharIndex > LastChar Then Exit Sub

    'Reset the fade value
    CharList(CharIndex).EmoFade = 0
    CharList(CharIndex).EmoDir = 1

    'Set the user's emoticon Grh by the emoticon index
    'Grh values are pulled directly from Grh1.raw - refer to that file
    Select Case EmoticonIndex
    Case EmoID.Dots: Engine_Init_Grh CharList(CharIndex).Emoticon, 78
    Case EmoID.Exclimation: Engine_Init_Grh CharList(CharIndex).Emoticon, 81
    Case EmoID.Question: Engine_Init_Grh CharList(CharIndex).Emoticon, 84
    Case EmoID.Surprised: Engine_Init_Grh CharList(CharIndex).Emoticon, 87
    Case EmoID.Heart: Engine_Init_Grh CharList(CharIndex).Emoticon, 90
    Case EmoID.Hearts: Engine_Init_Grh CharList(CharIndex).Emoticon, 93
    Case EmoID.HeartBroken: Engine_Init_Grh CharList(CharIndex).Emoticon, 96
    Case EmoID.Utensils: Engine_Init_Grh CharList(CharIndex).Emoticon, 99
    Case EmoID.Meat: Engine_Init_Grh CharList(CharIndex).Emoticon, 102
    Case EmoID.ExcliQuestion: Engine_Init_Grh CharList(CharIndex).Emoticon, 105
    End Select

End Sub

Sub Data_User_KnownSkills(ByRef rBuf As DataBuffer)

'*********************************************
'Retrieve known skills list
'<KnowSkillList()(B)>
'*********************************************

Dim KnowSkillList() As Long 'Note that each byte holds 8 skills
Dim Index As Long   'Which KnowSkillList array index to use
Dim X As Byte
Dim Y As Byte
Dim i As Byte

    'Retrieve the skill list
    ReDim KnowSkillList(1 To NumBytesForSkills)
    For i = 1 To NumBytesForSkills
        KnowSkillList(i) = rBuf.Get_Byte
    Next i
    
    'Clear the skill list size
    SkillListSize = 0

    'Set the values
    For i = 1 To NumSkills
        
        'Find the index to use
        Index = Int((i - 1) / 8) + 1
    
        'Check if the skill is known
        If KnowSkillList(Index) And (2 ^ (i - ((Index - 1) * 8) - 1)) Then

            'Update the SkillList position and size
            SkillListSize = SkillListSize + 1
            ReDim Preserve SkillList(1 To SkillListSize)

            'Set that the user knows the skill
            UserKnowSkill(i) = 1

            'Update position for skill list
            X = X + 1
            If X > SkillListWidth Then
                X = 1
                Y = Y + 1
            End If

            'Set the skill list ID and Position
            SkillList(SkillListSize).SkillID = i
            SkillList(SkillListSize).X = SkillListX - (X * 32)
            SkillList(SkillListSize).Y = SkillListY - (Y * 32)

        Else
        
            'User does not know the skill
            UserKnowSkill(i) = 0
            
        End If
    Next i

End Sub

Sub Data_User_LookLeft(ByRef rBuf As DataBuffer)

'*********************************************
'Make a character look to the specified direction (Used for LookLeft and LookRight)
'<CharIndex(I)><Heading(B)>
'*********************************************

Dim CharIndex As Integer
Dim Heading As Byte

    CharIndex = rBuf.Get_Integer

    Heading = rBuf.Get_Byte
    CharList(CharIndex).HeadHeading = Heading

End Sub

Sub Data_User_ModStat(ByRef rBuf As DataBuffer)

'*********************************************
'Update mod stat
'<StatID(B)><Value(L)>
'*********************************************

Dim StatID As Byte

    StatID = rBuf.Get_Byte
    ModStats(StatID) = rBuf.Get_Long
    
    'If we get a new speed value, adjust the scroll speed accordingly
    If StatID = SID.Speed Then
        ScrollPixelsPerFrameX = 4
        ScrollPixelsPerFrameY = 4
    End If

End Sub

Sub Data_User_Rotate(ByRef rBuf As DataBuffer)

'*********************************************
'Rotate a character by their CharIndex - works like it does in
' ChangeChar, but used to save ourselves a little bandwidth :)
'<CharIndex(I)><Heading(B)>
'*********************************************

Dim CharIndex As Integer

    CharIndex = rBuf.Get_Integer
    CharList(CharIndex).Heading = rBuf.Get_Byte
    CharList(CharIndex).HeadHeading = CharList(CharIndex).Heading

End Sub

Sub Data_User_SetInventorySlot(ByRef rBuf As DataBuffer)

'*********************************************
'Set an inventory slot's information
'The information in the () is only sent if the ObjIndex <> 0
'<Slot(B)><OBJIndex(L)>(<OBJName(S)><OBJAmount(L)><Equipted(B)><GrhIndex(L)>)
'*********************************************

Dim Slot As Byte

    'Get the slot
    Slot = rBuf.Get_Byte

    'Start gathering the data
    UserInventory(Slot).ObjIndex = rBuf.Get_Long
    
    'If the object index = 0, then we are deleting a slot, so the rest is null
    If UserInventory(Slot).ObjIndex = 0 Then
        UserInventory(Slot).Name = "(None)"
        UserInventory(Slot).Amount = 0
        UserInventory(Slot).Equipped = 0
        UserInventory(Slot).GrhIndex = 0
    Else
        'Index <> 0, so we have to get the information
        UserInventory(Slot).Name = rBuf.Get_String
        UserInventory(Slot).Amount = rBuf.Get_Long
        UserInventory(Slot).Equipped = rBuf.Get_Byte
        UserInventory(Slot).GrhIndex = rBuf.Get_Long
    End If

End Sub

Sub Data_User_Target(ByRef rBuf As DataBuffer)

'*********************************************
'User targets a character
'<CharIndex(I)>
'*********************************************

    TargetCharIndex = rBuf.Get_Integer

End Sub

Sub Data_User_Trade_StartNPCTrade(ByRef rBuf As DataBuffer)

'*********************************************
'Start trading with a NPC
'<NPCName(S)><NumVendItems(I)> Loop: <GrhIndex(L)><Name(S)><Price(L)>
'*********************************************

Dim NPCName As String
Dim NumItems As Integer
Dim Item As Integer

    NPCName = rBuf.Get_String
    NumItems = rBuf.Get_Integer

    ReDim NPCTradeItems(1 To NumItems)
    NPCTradeItemArraySize = NumItems
    For Item = 1 To NumItems
        NPCTradeItems(Item).GrhIndex = rBuf.Get_Long
        NPCTradeItems(Item).Name = rBuf.Get_String
        NPCTradeItems(Item).Price = rBuf.Get_Long
    Next Item
    ShowGameWindow(ShopWindow) = 1
    LastClickedWindow = ShopWindow

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:35)  Decl: 2  Code: 1320  Total: 1322 Lines
':) CommentOnly: 270 (20.4%)  Commented: 5 (0.4%)  Empty: 325 (24.6%)  Max Logic Depth: 4
