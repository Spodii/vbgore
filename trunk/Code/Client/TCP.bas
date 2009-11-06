Attribute VB_Name = "TCP"
Option Explicit
Public PacketOutPos As Byte
Public PacketInPos As Byte

Sub InitSocket()

'*****************************************************************
'Init the sox socket
'*****************************************************************

    'Save the game ini
    Call Engine_Var_Write(DataPath & "Game.ini", "INIT", "Name", UserName)
    If frmConnect.SavePassChk.Value = 0 Then   'If the password wont be saved, clear it out
        Call Engine_Var_Write(DataPath & "Game.ini", "INIT", "Password", "")
    Else
        Call Engine_Var_Write(DataPath & "Game.ini", "INIT", "Password", UserPassword)
    End If
    
    'Clean out the socket so we can make a fresh new connection
    If GOREsock_Loaded Then GOREsock_Terminate

    'Set up the socket
    DoEvents
    GOREsock_Initialize frmMain.hWnd
    DoEvents
    SoxID = GOREsock_Connect("127.0.0.1", 10200)
    
    'If the SoxID = -1, then the connection failed, elsewise, we're good to go! W00t! ^_^
    If SoxID = -1 Then
        MsgBox "Unable to connect to the game server!" & vbCrLf & "Either the server is down or you are not connected to the internet.", vbOKOnly Or vbCritical
    Else
        GOREsock_SetOption SoxID, soxSO_TCP_NODELAY, True
    End If

End Sub

Sub Data_User_Bank_UpdateSlot(ByRef rBuf As DataBuffer)

'*********************************************
'Updates a specific bank item
'<Slot(B)><GrhIndex(L)> If GrhIndex > 0, <Amount(I)>
'*********************************************
Dim GrhIndex As Long
Dim Amount As Integer
Dim Slot As Byte

    'Get the values
    Slot = rBuf.Get_Byte
    GrhIndex = rBuf.Get_Long
    
    'Check if to get the amount
    If GrhIndex > 0 Then Amount = rBuf.Get_Integer

    'Update the item
    UserBank(Slot).Amount = Amount
    UserBank(Slot).GrhIndex = GrhIndex

End Sub

Sub Data_User_Bank_Open(ByRef rBuf As DataBuffer)
'*********************************************
'Sends the list of bank items
'Loop: <Slot(B)><GrhIndex(L)><Amount(I)> until Slot = 255
'*********************************************
Dim GrhIndex As Long
Dim Amount As Integer
Dim Slot As Byte

    'Loop through the items until we get the terminator slot (255)
    Do
        
        'Get the slot
        Slot = rBuf.Get_Byte
        
        'Check if we have acquired the terminator slot
        If Slot = 255 Then Exit Do
        
        'Get the amount and obj index
        GrhIndex = rBuf.Get_Long
        Amount = rBuf.Get_Integer
        
        'Store the values
        UserBank(Slot).Amount = Amount
        UserBank(Slot).GrhIndex = GrhIndex
        
    Loop
    
    'Show the bank window
    ShowGameWindow(BankWindow) = 1
    LastClickedWindow = BankWindow

End Sub

Sub Data_Server_MakeProjectile(ByRef rBuf As DataBuffer)
'*********************************************
'Create a projectile from a ranged weapon
'<AttackerIndex(I)><TargetIndex(I)><GrhIndex(L)><Rotate(B)>
'*********************************************
Dim AttackerIndex As Integer
Dim TargetIndex As Integer
Dim GrhIndex As Long
Dim Rotate As Byte

    AttackerIndex = rBuf.Get_Integer
    TargetIndex = rBuf.Get_Integer
    GrhIndex = rBuf.Get_Long
    Rotate = rBuf.Get_Byte
    
    'If the char doesn't exist, request to create it
    If AttackerIndex > LastChar Or CharList(AttackerIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer AttackerIndex
        Exit Sub
    End If
    If TargetIndex > LastChar Or CharList(TargetIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer TargetIndex
        Exit Sub
    End If
    
    'Create the projectile
    Engine_Projectile_Create AttackerIndex, TargetIndex, GrhIndex, Rotate
    
End Sub

Sub Data_User_SetWeaponRange(ByRef rBuf As DataBuffer)
'*********************************************
'Set the range of the current weapon used so we can do client-side
' distance checks before sending the attack to the server
'<Range(B)>
'*********************************************

    UserAttackRange = rBuf.Get_Byte

End Sub

Sub Data_Server_SetCharSpeed(ByRef rBuf As DataBuffer)
'*********************************************
'Update a char's speed so we can move them the right speed
'<CharIndex(I)><Speed(B)>
'*********************************************
Dim CharIndex As Integer
Dim Speed As Byte

    CharIndex = rBuf.Get_Integer
    Speed = rBuf.Get_Byte
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
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
Dim TempInt As Integer
Dim Str1 As String
Dim Str2 As String
Dim Lng1 As Long
Dim Int1 As Integer
Dim Int2 As Integer
Dim Int3 As Integer
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
            Int3 = rBuf.Get_Integer
            TempStr = Replace$(Message(9), "<amount>", Int1)
            TempStr = Replace$(TempStr, "<npcname>", Str1)
            Engine_AddToChatTextBuffer TempStr, FontColor_Quest
            Engine_MakeChatBubble Int3, Engine_WordWrap(TempStr, BubbleMaxWidth)
        Case 10
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            Int3 = rBuf.Get_Integer
            TempStr = Replace$(Message(10), "<amount>", Int1)
            TempStr = Replace$(TempStr, "<objname>", Str1)
            Engine_AddToChatTextBuffer TempStr, FontColor_Quest
            Engine_MakeChatBubble Int3, Engine_WordWrap(TempStr, BubbleMaxWidth)
        Case 11
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            Int2 = rBuf.Get_Integer
            Str2 = rBuf.Get_String
            Int3 = rBuf.Get_Integer
            TempStr = Replace$(Message(11), "<npcamount>", Int1)
            TempStr = Replace$(TempStr, "<npcname>", Str1)
            TempStr = Replace$(TempStr, "<objamount>", Int2)
            TempStr = Replace$(TempStr, "<objname>", Str2)
            Engine_AddToChatTextBuffer TempStr, FontColor_Quest
            Engine_MakeChatBubble Int3, Engine_WordWrap(TempStr, BubbleMaxWidth)
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
            TempStr = Replace$(Message(41), "<name>", Str1)
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
            LastWhisperName = Str1  'Set the name of the last person to whisper us
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
            Engine_AddToChatTextBuffer Replace$(TempStr, "<cost>", Lng1), FontColor_Info
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
            TempInt = rBuf.Get_Integer
            TempStr = Replace$(Message(76), "<name>", Str1)
            TempStr = Replace$(TempStr, "<message>", Str2)
            Engine_AddToChatTextBuffer TempStr, FontColor_Info
            If TempInt > 0 Then Engine_MakeChatBubble TempInt, Engine_WordWrap(TempStr, BubbleMaxWidth)
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
        'Case 87 to 93 - these are only used by the client
        Case 94
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(94), "<name>", Str1), FontColor_Info
        Case 95
            Int1 = rBuf.Get_Integer
            Engine_AddToChatTextBuffer Replace$(Message(95), "<index>", Int1), FontColor_Info
        Case 96
            Int1 = rBuf.Get_Integer
            Str1 = rBuf.Get_String
            Lng1 = rBuf.Get_Long
            TempStr = Replace$(Message(96), "<amount>", Int1)
            TempStr = Replace$(TempStr, "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<cost>", Lng1), FontColor_Info
        Case 97
            Engine_AddToChatTextBuffer Message(97), FontColor_Info
        Case 98
            Engine_AddToChatTextBuffer Message(98), FontColor_Info
        Case 99
            Engine_AddToChatTextBuffer Message(99), FontColor_Info
        Case 100
            Str1 = rBuf.Get_String
            TempStr = Replace$(Message(100), "<linebreak>", vbCrLf)
            MsgBox Replace$(TempStr, "<reason>", Str1), vbOKOnly Or vbCritical
            IsUnloading = 1
            Engine_UnloadAllForms
        Case 101
            Engine_AddToChatTextBuffer Message(101), FontColor_Info
        Case 102
            Engine_AddToChatTextBuffer Message(102), FontColor_Info
        Case 106
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(106), "<name>", Str1), FontColor_Group
        Case 107
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(107), "<name>", Str1), FontColor_Group
        Case 108
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(108), "<name>", Str1), FontColor_Group
        Case 109
            Engine_AddToChatTextBuffer Message(109), FontColor_Group
        Case 110
            Str1 = rBuf.Get_String
            Engine_AddToChatTextBuffer Replace$(Message(110), "<name>", Str1), FontColor_Group
        Case 111
            Engine_AddToChatTextBuffer Message(111), FontColor_Group
        Case 112
            Engine_AddToChatTextBuffer Message(112), FontColor_Group
        Case 113
            Engine_AddToChatTextBuffer Message(113), FontColor_Group
        Case 114
            Engine_AddToChatTextBuffer Message(114), FontColor_Group
        Case 115
            Str1 = rBuf.Get_String
            Int1 = rBuf.Get_Integer
            TempStr = Replace$(Message(115), "<name>", Str1)
            Engine_AddToChatTextBuffer Replace$(TempStr, "<time>", Int1), FontColor_Group
        Case 116
            Engine_AddToChatTextBuffer Message(116), FontColor_Group
        Case 117
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(117), "<amount>", Lng1), FontColor_Info
        Case 118
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(118), "<amount>", Lng1), FontColor_Info
        Case 119
            Engine_AddToChatTextBuffer Message(119), FontColor_Info
        Case 120
            Lng1 = rBuf.Get_Long
            Engine_AddToChatTextBuffer Replace$(Message(120), "<amount>", Lng1), FontColor_Info
        Case 121
            Engine_AddToChatTextBuffer Message(121), FontColor_Info
        Case 123
            Engine_AddToChatTextBuffer Message(123), FontColor_Group
        Case 125
            Engine_AddToChatTextBuffer Message(125), FontColor_Info
    End Select

End Sub

Sub Data_Server_Connect()
'*********************************************
'Server is telling the client they have successfully logged in
'<>
'*********************************************

    'Set the socket state
    SocketOpen = 1

    If EngineRun = False Then
    
        'Load user config
        Game_Config_Load
    
        'Unload the connect form
        Unload frmConnect
    
        'Load main form
        frmMain.PTDTmr.Enabled = True
        Load frmMain
        frmMain.Visible = True
        frmMain.Show
        frmMain.SetFocus
        DoEvents
            
        'Load the engine
        Engine_Init_TileEngine
    
        'Get the device
        frmMain.Show
        frmMain.SetFocus
        DoEvents
        DIDevice.Acquire
        
        Unload frmNew
        Unload frmConnect
    
    End If
    
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
'<Text(S)><FontColorID(B)>(<CharIndex(B)>)
'*********************************************
Dim CharIndex As Integer
Dim TempStr As String
Dim TempLng As Long
Dim TempByte As Byte

    'Get the text
    TempStr = rBuf.Get_String
    TempByte = rBuf.Get_Byte
    
    'Filter the temp string
    TempStr = Game_FilterString(TempStr)
    
    'See if we have to make a bubble
    If TempByte And DataCode.Comm_UseBubble Then
        
        'We need a char index
        CharIndex = rBuf.Get_Integer
        
        'Split up the string for our chat bubble and assign it to the character
        Engine_MakeChatBubble CharIndex, Engine_WordWrap(TempStr, BubbleMaxWidth)

    End If
    
    'Get the color
    Select Case TempByte
        Case DataCode.Comm_FontType_Fight
            TempLng = FontColor_Fight
        Case DataCode.Comm_FontType_Info
            TempLng = FontColor_Info
        Case DataCode.Comm_FontType_Quest
            TempLng = FontColor_Quest
        Case DataCode.Comm_FontType_Talk
            TempLng = FontColor_Talk
        Case DataCode.Comm_FontType_Group
            TempLng = FontColor_Group
        Case Else
            TempLng = FontColor_Talk
    End Select
    
    'Add the text in the text box
    Engine_AddToChatTextBuffer TempStr, TempLng

End Sub

Sub Data_Map_LoadMap(ByRef rBuf As DataBuffer)

'*********************************************
'Load the map the server told us to load
'<MapNum(I)><ServerSideVersion(I)>
'*********************************************
Dim FileNum As Byte
Dim MapNumInt As Integer
Dim SSV As Integer
Dim Weather As Byte
Dim TempInt As Integer

    'Clear the target character
    TargetCharIndex = 0

    MapNumInt = rBuf.Get_Integer
    SSV = rBuf.Get_Integer

    If Engine_FileExist(MapPath & MapNumInt & ".map", vbNormal) Then  'Get Version Num
        FileNum = FreeFile
        Open MapPath & MapNumInt & ".map" For Binary As #FileNum
            Seek #FileNum, 1
            Get #FileNum, , TempInt
        Close #FileNum
        If TempInt = SSV Then   'Correct Version
            Game_Map_Switch MapNumInt
            sndBuf.Put_Byte DataCode.Map_DoneLoadingMap 'Tell the server we are done loading map
        Else
            'Not correct version
            MsgBox Message(105), vbOKOnly Or vbCritical
            EngineRun = False
            IsUnloading = 1
        End If
    Else
        'Didn't find map
        MsgBox Message(105), vbOKOnly Or vbCritical
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

    MapInfo.name = rBuf.Get_String
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
    If sndBuf.HasBuffer Then
        If SocketOpen = 0 Then DoEvents
    
        'Assign to the temp buffer
        TempBuffer() = sndBuf.Get_Buffer
    
        'Send the data
        GOREsock_SendData SoxID, TempBuffer
        
        'Clear the buffer, get it ready for next use
        sndBuf.Clear
  
    End If

End Sub

Sub Data_Server_ChangeCharType(ByRef rBuf As DataBuffer)

'*********************************************
'Change a character by the character index
'<CharIndex(I)><CharType(B)>
'*********************************************
Dim CharIndex As Integer
Dim CharType As Byte

    CharIndex = rBuf.Get_Integer
    CharType = rBuf.Get_Byte
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
    'Change the character's type
    CharList(CharIndex).CharType = CharType

End Sub

Sub Data_Server_ChangeChar(ByRef rBuf As DataBuffer)

'*********************************************
'Change a character by the character index
'<CharIndex(I)><Flags(B)>(<Body(I)><Head(I)><Heading(B)><Weapon(I)><Hair(I)><Wings(I)>)
'*********************************************
Dim flags As Byte
Dim CharIndex As Integer
Dim CharBody As Integer
Dim CharHead As Integer
Dim CharHeading As Byte
Dim CharWeapon As Integer
Dim CharHair As Integer
Dim CharWings As Integer
Dim DontSetData As Byte
    
    'Get the character index we are changing
    CharIndex = rBuf.Get_Integer
    
    'Get the flags on what data we need to get
    flags = rBuf.Get_Byte
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        DontSetData = 1
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        DontSetData = 1
    End If
    
    'Get the data needed
    If flags And 1 Then
        CharBody = rBuf.Get_Integer
        If DontSetData = 0 Then CharList(CharIndex).Body = BodyData(CharBody)
    End If
    If flags And 2 Then
        CharHead = rBuf.Get_Integer
        If DontSetData = 0 Then CharList(CharIndex).Head = HeadData(CharHead)
    End If
    If flags And 4 Then
        CharHeading = rBuf.Get_Byte
        If DontSetData = 0 Then CharList(CharIndex).Heading = CharHeading
        If DontSetData = 0 Then CharList(CharIndex).HeadHeading = CharHeading
    End If
    If flags And 8 Then
        CharWeapon = rBuf.Get_Integer
        If DontSetData = 0 Then CharList(CharIndex).Weapon = WeaponData(CharWeapon)
    End If
    If flags And 16 Then
        CharHair = rBuf.Get_Integer
        If DontSetData = 0 Then CharList(CharIndex).Hair = HairData(CharHair)
    End If
    If flags And 32 Then
        CharWings = rBuf.Get_Integer
        If DontSetData = 0 Then CharList(CharIndex).Wings = WingData(CharWings)
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

    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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

    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
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

    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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

Sub Data_Server_MailItemRemove(ByRef rBuf As DataBuffer)

'*********************************************
'Remove item from mailbox
'<ItemIndex(B)>
'*********************************************

Dim ItemIndex As Byte

    ItemIndex = rBuf.Get_Byte

    ReadMailData.Obj(ItemIndex) = 0

End Sub

Sub Data_Server_MailObjUpdate(ByRef rBuf As DataBuffer)

'*********************************************
'Updates the objects in a mail message
'<NumObjs(B)> Loop: <ObjGrhIndex(L)>
'*********************************************
Dim NumObjs As Byte
Dim X As Byte

    'Clear the current objects
    For X = 1 To MaxMailObjs
        ReadMailData.Obj(X) = 0
        ReadMailData.ObjName(X) = 0
        ReadMailData.ObjAmount(X) = 0
    Next X
    
    'Get the number of objects
    NumObjs = rBuf.Get_Byte
    
    'Get the mail objects
    For X = 1 To NumObjs
        ReadMailData.Obj(X) = rBuf.Get_Long
        ReadMailData.ObjName(X) = rBuf.Get_String
        ReadMailData.ObjAmount(X) = rBuf.Get_Integer
    Next X

End Sub

Sub Data_Server_MailMessage(ByRef rBuf As DataBuffer)

'*********************************************
'Recieve message that was requested to be read
'<Message(S-EX)><Subject(S)><WriterName(S)><NumObjs(B)> Loop: <ObjGrhIndex(L)>
'*********************************************
Dim NumObjs As Byte
Dim i As Long

    'Clear the current objects
    For i = 1 To MaxMailObjs
        ReadMailData.Obj(i) = 0
        ReadMailData.ObjName(i) = 0
        ReadMailData.ObjAmount(i) = 0
    Next i
    
    'Show the correct windows
    ShowGameWindow(MailboxWindow) = 0
    ShowGameWindow(ViewMessageWindow) = 1
    LastClickedWindow = ViewMessageWindow
    
    'Get the data
    ReadMailData.Message = rBuf.Get_StringEX
    ReadMailData.Message = Engine_WordWrap(ReadMailData.Message, GameWindow.ViewMessage.Message.Width)
    ReadMailData.Subject = rBuf.Get_String
    ReadMailData.WriterName = rBuf.Get_String
    NumObjs = rBuf.Get_Byte
    For i = 1 To NumObjs
        ReadMailData.Obj(i) = rBuf.Get_Long
        ReadMailData.ObjName(i) = rBuf.Get_String
        ReadMailData.ObjAmount(i) = rBuf.Get_Integer
    Next i

End Sub

Sub Data_Server_MakeChar(ByRef rBuf As DataBuffer)

'*********************************************
'Create a character and set their information
'<Body(I)><Head(I)><Heading(B)><CharIndex(I)><X(B)><Y(B)><Speed(B)><Name(S)><Weapon(I)><Hair(I)><Wings(I)><HP%(B)><MP%(B)><ChatID(B)><CharType(B)>
'*********************************************

Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim CharIndex As Integer
Dim X As Byte
Dim Y As Byte
Dim Speed As Byte
Dim name As String
Dim Weapon As Integer
Dim Hair As Integer
Dim Wings As Integer
Dim HP As Byte
Dim MP As Byte
Dim ChatID As Byte
Dim CharType As Byte

    'Retrieve all the information
    Body = rBuf.Get_Integer
    Head = rBuf.Get_Integer
    Heading = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Speed = rBuf.Get_Byte
    name = rBuf.Get_String
    Weapon = rBuf.Get_Integer
    Hair = rBuf.Get_Integer
    Wings = rBuf.Get_Integer
    HP = rBuf.Get_Byte
    MP = rBuf.Get_Byte
    ChatID = rBuf.Get_Byte
    CharType = rBuf.Get_Byte
    
    'Create the character
    Engine_Char_Make CharIndex, Body, Head, Heading, X, Y, Speed, name, Weapon, Hair, Wings, ChatID, CharType, HP, MP

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
Dim Running As Byte
    
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Heading = rBuf.Get_Byte
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
    'Check if running
    If Heading > 128 Then
        Heading = Heading Xor 128
        Running = 1
    End If
    
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
    
    'Move the character
    Engine_Char_Move_ByPos CharIndex, X, Y, Running

End Sub

Sub Data_Server_PTD()

'*********************************************
'We retrieved the PTD response, calculate how long it took
'<>
'*********************************************

    PTD = timeGetTime - PTDSTime

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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

    'Create the blood (if damage)
    If Damage > 0 Then Engine_Blood_Create CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y

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
    
    'Check for a valid UserCharIndex
    If UserCharIndex <= 0 Or UserCharIndex > LastChar Then
    
        'We have an invalid user char index, so we must have the wrong one - request an update on the right one
        sndBuf.Put_Byte DataCode.User_RequestUserCharIndex
        Exit Sub
        
    End If

    'Check if the position is even different
    If X <> UserPos.X Or Y <> UserPos.Y Then
    
        'Update the user's position
        UserPos.X = X
        UserPos.Y = Y
        CharList(UserCharIndex).Pos = UserPos

        'If there is a targeted char, check if the path is valid
        If TargetCharIndex > 0 Then
            ClearPathToTarget = Engine_ClearPath(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(TargetCharIndex).Pos.X, CharList(TargetCharIndex).Pos.Y)
        End If
        
    End If

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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
    'Start the attack animation
    CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).Started = 1
    CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).FrameCounter = 1
    CharList(CharIndex).Body.Attack(CharList(CharIndex).Heading).LastCount = timeGetTime
    CharList(CharIndex).Weapon.Attack(CharList(CharIndex).Heading).FrameCounter = 1
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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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
    
    'If the char doesn't exist, request to create it
    If TargetIndex = 0 Or CasterIndex = 0 Then Exit Sub
    If CasterIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CasterIndex
        Exit Sub
    End If
    If CharList(CasterIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CasterIndex
        Exit Sub
    End If
    If TargetIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer TargetIndex
        Exit Sub
    End If
    If CharList(TargetIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer TargetIndex
        Exit Sub
    End If
    
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
        If CharList(CasterIndex).HeadHeading = NORTH Or CharList(CasterIndex).HeadHeading = NORTHEAST Then
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
        ElseIf CharList(CasterIndex).HeadHeading = EAST Or CharList(CasterIndex).HeadHeading = SOUTHEAST Then
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
        ElseIf CharList(CasterIndex).HeadHeading = SOUTH Or CharList(CasterIndex).HeadHeading = SOUTHWEST Then
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
        ElseIf CharList(CasterIndex).HeadHeading = WEST Or CharList(CasterIndex).HeadHeading = NORTHWEST Then
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

Sub Data_Server_MakeEffect(ByRef rBuf As DataBuffer)

'*********************************************
'Create an effect on the effects layer
'<X(B)><Y(B)><GrhIndex(L)>
'*********************************************
Dim X As Byte
Dim Y As Byte
Dim GrhIndex As Long

    'Get the values
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    GrhIndex = rBuf.Get_Long

    'Create the effect
    Engine_Effect_Create X, Y, GrhIndex, 0, 0, 1
    
End Sub

Sub Data_Server_MakeSlash(ByRef rBuf As DataBuffer)

'*********************************************
'Create a slash effect on the effects layer
'<CharIndex(I)><GrhIndex(L)>
'*********************************************

Dim CharIndex As Integer
Dim GrhIndex As Long
Dim Angle As Single
    
    'Get the values
    CharIndex = rBuf.Get_Integer
    GrhIndex = rBuf.Get_Long
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
    'Get the new heading
    Select Case CharList(CharIndex).Heading
        Case NORTH
            Angle = 0
        Case NORTHEAST
            Angle = 45
        Case EAST
            Angle = 90
        Case SOUTHEAST
            Angle = 135
        Case SOUTH
            Angle = 180
        Case SOUTHWEST
            Angle = 225
        Case WEST
            Angle = 270
        Case NORTHWEST
            Angle = 315
    End Select

    'Create the effect
    Engine_Effect_Create CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y, GrhIndex, Angle, 150, 0
    
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

    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If

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
Dim Heading As Byte
Dim CharIndex As Integer
    
    CharIndex = rBuf.Get_Integer
    Heading = rBuf.Get_Byte
    
    'If the char doesn't exist, request to create it
    If CharIndex > LastChar Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    If CharList(CharIndex).Active = 0 Then
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RequestMakeChar
        sndBuf.Put_Integer CharIndex
        Exit Sub
    End If
    
    CharList(CharIndex).Heading = Heading
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
        UserInventory(Slot).name = "(None)"
        UserInventory(Slot).Amount = 0
        UserInventory(Slot).Equipped = 0
        UserInventory(Slot).GrhIndex = 0
    Else
        'Index <> 0, so we have to get the information
        UserInventory(Slot).name = rBuf.Get_String
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
    
    'Check for a valid UserCharIndex
    If UserCharIndex <= 0 Or UserCharIndex > LastChar Then
    
        'We have an invalid user char index, so we must have the wrong one - request an update on the right one
        sndBuf.Put_Byte DataCode.User_RequestUserCharIndex
        Exit Sub
        
    End If
    
    'Check if the path to the targeted character is valid (if any)
    If TargetCharIndex > 0 Then ClearPathToTarget = Engine_ClearPath(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(TargetCharIndex).Pos.X, CharList(TargetCharIndex).Pos.Y)

End Sub

Sub Data_User_ChangeServer(ByRef rBuf As DataBuffer)

'*********************************************
'Changes a user to a different server
'<Port(I)><IP(S)>
'*********************************************
Dim Port As Integer
Dim IP As String

    'Get the values
    Port = rBuf.Get_Integer
    IP = rBuf.Get_String

    'Clean out the socket so we can make a fresh new connection
    SocketOpen = 0
    GOREsock_Shut SoxID
    GOREsock_Terminate
    CurMap = 0
        
    'Set up the socket
    DoEvents
    GOREsock_Initialize frmMain.hWnd
    DoEvents
    SoxID = GOREsock_Connect(IP, Port)
    
    'If the SoxID = -1, then the connection failed, elsewise, we're good to go! W00t! ^_^
    If SoxID = -1 Then
        MsgBox "Unable to connect to the game server!" & vbCrLf & "Either the server is down or you are not connected to the internet.", vbOKOnly Or vbCritical
        IsUnloading = 1
    Else
        GOREsock_SetOption SoxID, soxSO_TCP_NODELAY, True
    End If

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
        NPCTradeItems(Item).name = rBuf.Get_String
        NPCTradeItems(Item).Price = rBuf.Get_Long
    Next Item
    ShowGameWindow(ShopWindow) = 1
    LastClickedWindow = ShopWindow

End Sub

Sub GOREsock_Close(inSox As Long)

    If Not frmMain.Visible Then MsgBox Message(122), vbOKOnly
        
    If SocketOpen = 1 Then IsUnloading = 1

End Sub

Sub Data_Server_SendQuestInfo(ByRef rBuf As DataBuffer)

'*********************************************
'Server sent the information on a quest
'<QuestID(B)><Name(S)>(<Description(S-EX)>)
'*********************************************
Dim QuestID As Byte
Dim name As String
Dim Desc As String
Dim i As Long
Dim Changed As Byte

    'Get the variables
    QuestID = rBuf.Get_Byte
    name = rBuf.Get_String
    If LenB(name) Then Desc = rBuf.Get_StringEX    'Only get the desc if the name exists

    'Resize the questinfo array if needed
    If QuestID > QuestInfoUBound Then
        QuestInfoUBound = QuestID
        ReDim Preserve QuestInfo(1 To QuestInfoUBound)
    End If
    
    'Store the information
    QuestInfo(QuestID).name = name
    QuestInfo(QuestID).Desc = Desc

    'Loop through the quests, remove any unused slots on the end
    If QuestInfoUBound > 1 Then
        For i = QuestInfoUBound To 1 Step -1
            If QuestInfo(i).name = vbNullString Then
                QuestInfoUBound = QuestInfoUBound - 1
                Changed = 1
            Else
                'Exit on the first section of information
                Exit For
            End If
        Next i
        If Changed Then
            If QuestInfoUBound > 0 Then
                ReDim Preserve QuestInfo(1 To QuestInfoUBound)
            Else
                Erase QuestInfo
            End If
        End If
    Else
        If QuestInfo(1).name = vbNullString Then
            Erase QuestInfo
            QuestInfoUBound = 0
        End If
    End If
    
End Sub

Sub GOREsock_DataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************

Dim rBuf As DataBuffer
Dim CommandID As Byte
Dim BufUBound As Long
Static X As Long

    'Display packet
    If DEBUG_PrintPacket_In Then
        Engine_AddToChatTextBuffer "DataIn: " & StrConv(inData, vbUnicode), -1
    End If

    'Set up the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    BufUBound = UBound(inData)
    
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

    Do
        'Get the Command ID
        CommandID = rBuf.Get_Byte

        'Make the appropriate call based on the Command ID
        With DataCode
            Select Case CommandID

            Case 0
                If DEBUG_PrintPacketReadErrors Then
                    X = X + 1
                    Debug.Print "Empty Command ID #" & X
                End If

            Case .Comm_Talk: Data_Comm_Talk rBuf

            Case .Map_LoadMap: Data_Map_LoadMap rBuf
            Case .Map_SendName:  Data_Map_SendName rBuf

            Case .Server_ChangeChar: Data_Server_ChangeChar rBuf
            Case .Server_ChangeCharType: Data_Server_ChangeCharType rBuf
            Case .Server_CharHP: Data_Server_CharHP rBuf
            Case .Server_CharMP: Data_Server_CharMP rBuf
            Case .Server_Connect: Data_Server_Connect
            Case .Server_Disconnect: Data_Server_Disconnect
            Case .Server_EraseChar: Data_Server_EraseChar rBuf
            Case .Server_EraseObject: Data_Server_EraseObject rBuf
            Case .Server_IconBlessed: Data_Server_IconBlessed rBuf
            Case .Server_IconCursed: Data_Server_IconCursed rBuf
            Case .Server_IconIronSkin: Data_Server_IconIronSkin rBuf
            Case .Server_IconProtected: Data_Server_IconProtected rBuf
            Case .Server_IconStrengthened: Data_Server_IconStrengthened rBuf
            Case .Server_IconWarCursed:  Data_Server_IconWarCursed rBuf
            Case .Server_IconSpellExhaustion: Data_Server_IconSpellExhaustion rBuf
            Case .Server_MailBox: Data_Server_Mailbox rBuf
            Case .Server_MailItemRemove: Data_Server_MailItemRemove rBuf
            Case .Server_MailMessage: Data_Server_MailMessage rBuf
            Case .Server_MailObjUpdate: Data_Server_MailObjUpdate rBuf
            Case .Server_MakeChar: Data_Server_MakeChar rBuf
            Case .Server_MakeEffect: Data_Server_MakeEffect rBuf
            Case .Server_MakeSlash: Data_Server_MakeSlash rBuf
            Case .Server_MakeObject: Data_Server_MakeObject rBuf
            Case .Server_MakeProjectile: Data_Server_MakeProjectile rBuf
            Case .Server_Message: Data_Server_Message rBuf
            Case .Server_MoveChar: Data_Server_MoveChar rBuf
            Case .Server_PTD: Data_Server_PTD
            Case .Server_PlaySound: Data_Server_PlaySound rBuf
            Case .Server_PlaySound3D: Data_Server_PlaySound3D rBuf
            Case .Server_SendQuestInfo: Data_Server_SendQuestInfo rBuf
            Case .Server_SetCharDamage: Data_Server_SetCharDamage rBuf
            Case .Server_SetCharSpeed: Data_Server_SetCharSpeed rBuf
            Case .Server_SetUserPosition: Data_Server_SetUserPosition rBuf
            Case .Server_UserCharIndex: Data_Server_UserCharIndex rBuf

            Case .User_AggressiveFace: Data_User_AggressiveFace rBuf
            Case .User_Attack: Data_User_Attack rBuf
            Case .User_Bank_Open: Data_User_Bank_Open rBuf
            Case .User_Bank_UpdateSlot: Data_User_Bank_UpdateSlot rBuf
            Case .User_BaseStat: Data_User_BaseStat rBuf
            Case .User_Blink: Data_User_Blink rBuf
            Case .User_CastSkill: Data_User_CastSkill rBuf
            Case .User_ChangeServer: Data_User_ChangeServer rBuf
            Case .User_Emote: Data_User_Emote rBuf
            Case .User_KnownSkills: Data_User_KnownSkills rBuf
            Case .User_LookLeft: Data_User_LookLeft rBuf
            Case .User_LookRight: Data_User_LookLeft rBuf
            Case .User_ModStat: Data_User_ModStat rBuf
            Case .User_Rotate: Data_User_Rotate rBuf
            Case .User_SetInventorySlot: Data_User_SetInventorySlot rBuf
            Case .User_SetWeaponRange: Data_User_SetWeaponRange rBuf
            Case .User_Target: Data_User_Target rBuf
            Case .User_Trade_StartNPCTrade: Data_User_Trade_StartNPCTrade rBuf

            Case Else
                If DEBUG_PrintPacketReadErrors Then Debug.Print "Command ID " & CommandID & " caused a premature packet handling abortion!"
                Exit Do 'Something went wrong or we hit the end, either way, RUN!!!!

            End Select
        End With

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do

    Loop

End Sub

Sub GOREsock_Connecting(inSox As Long)

    If SocketOpen = 0 Then

        PacketInPos = 0
        PacketOutPos = 0
        
        Sleep 50
        DoEvents
        
        'Pre-saved character
        If SendNewChar = False Then
            sndBuf.Put_Byte DataCode.User_Login
            sndBuf.Put_String UserName
            sndBuf.Put_String UserPassword
        Else
            'New character
            sndBuf.Put_Byte DataCode.User_NewLogin
            sndBuf.Put_String UserName
            sndBuf.Put_String UserPassword
            sndBuf.Put_Integer UserHead
            sndBuf.Put_Integer UserBody
            sndBuf.Put_Byte UserClass
        End If
    
        'Save Game.ini
        If frmConnect.SavePassChk.Value = 0 Then UserPassword = vbNullString
        Engine_Var_Write DataPath & "Game.ini", "INIT", "Name", UserName
        Engine_Var_Write DataPath & "Game.ini", "INIT", "Password", UserPassword
        
        'Send the data
        Data_Send
        DoEvents
    
    End If
    
End Sub

Sub GOREsock_Connection(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub
