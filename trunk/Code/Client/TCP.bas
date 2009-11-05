Attribute VB_Name = "TCP"
Option Explicit
Public DevMode As Byte

Sub Data_Server_Connect()
'*********************************************
'Server is telling the client they have successfully logged in
'<>
'*********************************************

    If SocketOpen = 0 Then
    
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

    End If

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

Sub Data_Comm_UMsgBox(ByRef rBuf As DataBuffer)

'*********************************************
'Create an urgent messagebox
'<Text(S)>
'*********************************************

    MsgBox rBuf.Get_String, vbApplicationModal
    frmMain.Sox.Shut SoxID

End Sub

Sub Data_Dev_SetMapInfo(ByRef rBuf As DataBuffer)

'*********************************************
'Retrieve map info
'<Name(S)><Version(I)><Weather(B)>
'*********************************************

    DevValue.Name = rBuf.Get_String
    DevValue.Version = rBuf.Get_Integer
    DevValue.Weather = rBuf.Get_Byte

End Sub

Sub Data_Dev_SetMode(ByRef rBuf As DataBuffer)

'*********************************************
'Set the user's dev mode
'<Mode(B)>
'*********************************************

Dim Mode As Byte

    Mode = rBuf.Get_Byte

    DevMode = Mode
    'Exit dev mode
    If Mode = 1 Then
        ShowGameWindow(DevWindow) = 1
        LastClickedWindow = DevWindow
        'Enter dev mode
    Else
        ShowGameWindow(DevWindow) = 0
        If LastClickedWindow = DevWindow Then LastClickedWindow = 0
    End If

End Sub

Sub Data_Map_DoneSwitching()

'*********************************************
'Done switching maps, load engine back up
'<>
'*********************************************

    DownloadingMap = False
    EngineRun = True

End Sub

Sub Data_Map_EndTransfer(ByRef rBuf As DataBuffer)

'*********************************************
'States the end of a map transfer
'<MapNum(I)>
'*********************************************

Dim MapNum As Integer

    MapNum = rBuf.Get_Integer

    If MapNum > NumMaps Then NumMaps = MapNum
    Game_SaveMapData MapNum
    Game_Map_Switch MapNum
    sndBuf.Put_Byte DataCode.Map_DoneLoadingMap

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
            sndBuf.Put_Byte DataCode.Map_RequestUpdate
            sndBuf.Put_Integer MapNumInt
        End If
    Else
        'Didn't find map
        sndBuf.Put_Byte DataCode.Map_RequestUpdate
        sndBuf.Put_Integer MapNumInt
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

Sub Data_Map_StartTransfer(ByRef rBuf As DataBuffer)

'*********************************************
'Start the transfer of a map file
'<Version(I)>
'*********************************************

    MapInfo.MapVersion = rBuf.Get_Integer
    Engine_ClearMapArray

End Sub

Sub Data_Map_UpdateTile(ByRef rBuf As DataBuffer)

'*********************************************
'Update map tile
'<X(B)><Y(B)><ChunkData(I)><Depends on the flags from ChunkData...>
'ChunkData values:
'      1 = Blocked North
'      2 = Blocked East
'      4 = Blocked South
'      8 = Blocked West
'     16 = Is Mailbox
'     32 = Grh(1) <> 0
'     64 = Grh(2) <> 0
'    128 = Grh(3) <> 0
'    256 = Grh(4) <> 0
'    512 = Grh(5) <> 0
'   1024 = Grh(6) <> 0
'   2048 = Light(1 to 4) <> -1
'   4096 = Light(5 to 8) <> -1
'   8192 = Light(9 to 12) <> -1
'  16384 = Light(13 to 16) <> -1
'  32768 = Light(17 to 20) <> -1
'  65536 = Light(21 to 24) <> -1
' 131072 = Shadow(1)
' 262144 = Shadow(2)
' 524288 = Shadow(3)
'1048576 = Shadow(4)
'2097152 = Shadow(5)
'4194304 = Shadow(6)
'8388608 = Sfx
'*********************************************
Dim Grh(1 To 6) As Integer
Dim Light(1 To 24) As Long
Dim ChunkData As Long
Dim ShowRect As RECT
Dim Sfx As Integer
Dim j As Integer
Dim X As Byte
Dim Y As Byte
Dim i As Byte

'Get the values

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    ChunkData = rBuf.Get_Long

    'Check if theres Grh layers
    If ChunkData And 32 Then Grh(1) = rBuf.Get_Integer
    If ChunkData And 64 Then Grh(2) = rBuf.Get_Integer
    If ChunkData And 128 Then Grh(3) = rBuf.Get_Integer
    If ChunkData And 256 Then Grh(4) = rBuf.Get_Integer
    If ChunkData And 512 Then Grh(5) = rBuf.Get_Integer
    If ChunkData And 1024 Then Grh(6) = rBuf.Get_Integer

    'Check if theres lights to get
    For i = 1 To 24
        Light(i) = -1   'Set the default value to -1 and if it is different, then we will set it below
    Next i
    If ChunkData And 2048 Then
        For i = 1 To 4
            Light(i) = rBuf.Get_Long
        Next i
    End If
    If ChunkData And 4096 Then
        For i = 5 To 8
            Light(i) = rBuf.Get_Long
        Next i
    End If
    If ChunkData And 8192 Then
        For i = 9 To 12
            Light(i) = rBuf.Get_Long
        Next i
    End If
    If ChunkData And 16384 Then
        For i = 13 To 16
            Light(i) = rBuf.Get_Long
        Next i
    End If
    If ChunkData And 32768 Then
        For i = 17 To 20
            Light(i) = rBuf.Get_Long
        Next i
    End If
    If ChunkData And 65536 Then
        For i = 21 To 24
            Light(i) = rBuf.Get_Long
        Next i
    End If
    
    'Sfx value
    If ChunkData And 8388608 Then
        Sfx = rBuf.Get_Integer
    End If
    
    'Set the map information
    With MapData(X, Y)

        'Set blocked values
        .Blocked = 0
        If ChunkData And 1 Then .Blocked = .Blocked Or BlockedNorth
        If ChunkData And 2 Then .Blocked = .Blocked Or BlockedEast
        If ChunkData And 4 Then .Blocked = .Blocked Or BlockedSouth
        If ChunkData And 8 Then .Blocked = .Blocked Or BlockedWest
        
        'Set mailbox value
        If ChunkData And 16 Then .Mailbox = 1 Else .Mailbox = 0

        'Set the graphic layers
        For i = 1 To 6
            If Grh(i) > 0 Then Engine_Init_Grh .Graphic(i), Grh(i) Else .Graphic(i).GrhIndex = 0
        Next i

        'Set the lights
        For i = 1 To 24
            .Light(i) = Light(i)
            SaveLightBuffer(X, Y).Light(i) = Light(i)
        Next i
        
        'Set shadows
        If ChunkData And 131072 Then .Shadow(1) = 1 Else .Shadow(1) = 0
        If ChunkData And 262144 Then .Shadow(2) = 1 Else .Shadow(2) = 0
        If ChunkData And 524288 Then .Shadow(3) = 1 Else .Shadow(3) = 0
        If ChunkData And 1048576 Then .Shadow(4) = 1 Else .Shadow(4) = 0
        If ChunkData And 2097152 Then .Shadow(5) = 1 Else .Shadow(5) = 0
        If ChunkData And 4194304 Then .Shadow(6) = 1 Else .Shadow(6) = 0
        
        'Set the sfx
        If ChunkData And 1048576 Then
            'Clear any old sfx
            If Not MapData(X, Y).Sfx Is Nothing Then
                MapData(X, Y).Sfx.Stop
                Set MapData(X, Y).Sfx = Nothing
            End If
            'Set the sfx
            Engine_Sound_SetToMap Sfx, X, Y
        End If
        
    End With

    'Update screen only when percent changes (the more you render, the slower you download)
    If EngineRun = False Then
        If X = 1 Then
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            D3DDevice.BeginScene
            Engine_Render_Text "Downloading Map: " & Y & "%", 600, 584, -1
            D3DDevice.EndScene
            ShowRect.Left = 500
            ShowRect.Top = 584
            ShowRect.Right = 800
            ShowRect.bottom = 600
            D3DDevice.Present ShowRect, ShowRect, 0, ByVal 0    'Only present the text section
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
    
        'Set our temp buffer
        ReDim TempBuffer(UBound(sndBuf.Get_Buffer))
        TempBuffer() = sndBuf.Get_Buffer
        
        'Crop off the last byte, which will always be 0 - bad way to do it, but oh well
        ReDim Preserve TempBuffer(UBound(TempBuffer) - 1)
        
        'Uncomment this to see packets going out from the client
        'Dim i As Long
        'Dim S As String
        'For i = LBound(TempBuffer) To UBound(TempBuffer)
        '    S = S & TempBuffer(i) & " "
        'Next i
        'Debug.Print S
    
        'Encrypt our packet
        Select Case EncryptionType
        Case EncryptionTypeBlowfish
            Encryption_Blowfish_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeCryptAPI
            Encryption_CryptAPI_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeDES
            Encryption_DES_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeGost
            Encryption_Gost_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeRC4
            Encryption_RC4_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeXOR
            Encryption_XOR_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeSkipjack
            Encryption_Skipjack_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeTEA
            Encryption_TEA_EncryptByte TempBuffer(), EncryptionKey
        Case EncryptionTypeTwofish
            Encryption_Twofish_EncryptByte TempBuffer(), EncryptionKey
        End Select
        
        'Display packet
        If DEBUG_PrintPacket_Out Then
            Engine_AddToChatTextBuffer "DataOut: " & StrConv(TempBuffer(), vbUnicode), -1
        End If
    
        'Send the data
        frmMain.Sox.SendData SoxID, TempBuffer()
        
        'Clear the buffer, get it ready for next use
        sndBuf.Clear
        
    End If

End Sub

Sub Data_Server_ChangeChar(ByRef rBuf As DataBuffer)

'*********************************************
'Change a character by the character index
'<CharIndex(I)><Body(I)><Head(I)><Heading(B)><Weapon(I)><Hair(I)>
'*********************************************

Dim CharIndex As Integer

    CharIndex = rBuf.Get_Integer
    
    'Check for invalid instances
    If DownloadingMap Then Exit Sub

    CharList(CharIndex).Body = BodyData(rBuf.Get_Integer)
    CharList(CharIndex).Head = HeadData(rBuf.Get_Integer)
    CharList(CharIndex).Heading = rBuf.Get_Byte
    CharList(CharIndex).HeadHeading = CharList(CharIndex).Heading
    CharList(CharIndex).Weapon = WeaponData(rBuf.Get_Integer)
    CharList(CharIndex).Hair = HairData(rBuf.Get_Integer)

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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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
'<Body(I)><Head(I)><Heading(B)><CharIndex(I)><X(B)><Y(B)><Name(S)><Weapon(I)><Hair(I)><HP%(B)><MP%(B)>
'*********************************************

Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim CharIndex As Integer
Dim X As Byte
Dim Y As Byte
Dim Name As String
Dim Weapon As Integer
Dim Hair As Integer
Dim HP As Byte
Dim MP As Byte

'Retrieve all the information

    Body = rBuf.Get_Integer
    Head = rBuf.Get_Integer
    Heading = rBuf.Get_Byte
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Name = rBuf.Get_String
    Weapon = rBuf.Get_Integer
    Hair = rBuf.Get_Integer
    HP = rBuf.Get_Byte
    MP = rBuf.Get_Byte

    'Create the character
    Engine_Char_Make CharIndex, Body, Head, Heading, X, Y, Name, Weapon, Hair, HP, MP

End Sub

Sub Data_Server_MakeObject(ByRef rBuf As DataBuffer)

'*********************************************
'Create an object on the object layer
'<GrhIndex(I)><X(B)><Y(B)>
'*********************************************

Dim GrhIndex As Integer
Dim X As Byte
Dim Y As Byte

'Get the values

    GrhIndex = rBuf.Get_Integer
    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Create the object
    If GrhIndex > 0 Then Engine_OBJ_Create GrhIndex, X, Y

End Sub

Sub Data_Server_MoveChar(ByRef rBuf As DataBuffer)

'*********************************************
'Move a character
'<CharIndex(I)><X(B)><Y(B)>
'*********************************************

Dim CharIndex As Integer
Dim X As Byte
Dim Y As Byte

    CharIndex = rBuf.Get_Integer

    X = rBuf.Get_Byte
    Y = rBuf.Get_Byte
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
'Play a wave
'<WaveNum(B)>
'*********************************************

Dim WaveNum As Byte

    WaveNum = rBuf.Get_Byte

    'LoadWavetoDSBuffer Str$(WaveNum) & ".wav"

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
    If DownloadingMap Then Exit Sub
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
    If DownloadingMap Then Exit Sub
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

    CharList(CharIndex).BlinkTimer = 300

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
'<KnowSkillList(L)>
'*********************************************

Dim KnowSkillList As Long
Dim X As Byte
Dim Y As Byte
Dim i As Byte

'Retrieve the skill list

    KnowSkillList = rBuf.Get_Long

    'Clear the skill list size
    SkillListSize = 0

    'Set the values
    For i = 1 To NumSkills
        If KnowSkillList And (1 * (2 ^ (i - 1))) Then

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
'<Slot(B)><OBJIndex(L)><OBJName(S)><OBJAmount(L)><Equipted(B)><GrhIndex(I)>
'*********************************************

Dim Slot As Byte

    Slot = rBuf.Get_Byte

    UserInventory(Slot).ObjIndex = rBuf.Get_Long
    UserInventory(Slot).Name = rBuf.Get_String
    UserInventory(Slot).Amount = rBuf.Get_Long
    UserInventory(Slot).Equipped = rBuf.Get_Byte
    UserInventory(Slot).GrhIndex = rBuf.Get_Integer

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
'<NPCName(S)><NumVendItems(I)> Loop: <GrhIndex(I)><Name(S)><Price(L)>
'*********************************************

Dim NPCName As String
Dim NumItems As Integer
Dim Item As Integer

    NPCName = rBuf.Get_String
    NumItems = rBuf.Get_Integer

    ReDim NPCTradeItems(1 To NumItems)
    NPCTradeItemArraySize = NumItems
    For Item = 1 To NumItems
        NPCTradeItems(Item).GrhIndex = rBuf.Get_Integer
        NPCTradeItems(Item).Name = rBuf.Get_String
        NPCTradeItems(Item).Price = rBuf.Get_Long
    Next Item
    ShowGameWindow(ShopWindow) = 1
    LastClickedWindow = ShopWindow

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:35)  Decl: 2  Code: 1320  Total: 1322 Lines
':) CommentOnly: 270 (20.4%)  Commented: 5 (0.4%)  Empty: 325 (24.6%)  Max Logic Depth: 4
