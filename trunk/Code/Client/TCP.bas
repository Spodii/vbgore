Attribute VB_Name = "TCP"
Option Explicit

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
            prgRun = False
            IsUnloading = 1
        End If
    Else
        'Didn't find map
        MsgBox "Error! The requested map could not be found! Please run the updater!", vbOKOnly Or vbCritical
        EngineRun = False
        prgRun = False
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
'<CharIndex(I)><Body(I)><Head(I)><Heading(B)><Weapon(I)><Hair(I)><Wings(I)>
'*********************************************

Dim CharIndex As Integer
Dim CharBody As Integer
Dim CharHead As Integer
Dim CharHeading As Byte
Dim CharWeapon As Integer
Dim CharHair As Integer
Dim CharWings As Integer
    
    'Gather the information - must be done before checking for invalid parameters
    CharIndex = rBuf.Get_Integer
    CharBody = rBuf.Get_Integer
    CharHead = rBuf.Get_Integer
    CharHeading = rBuf.Get_Byte
    CharWeapon = rBuf.Get_Integer
    CharHair = rBuf.Get_Integer
    CharWings = rBuf.Get_Integer

    CharList(CharIndex).Body = BodyData(CharBody)
    CharList(CharIndex).Head = HeadData(CharHead)
    CharList(CharIndex).Heading = CharHeading
    CharList(CharIndex).HeadHeading = CharHeading
    CharList(CharIndex).Weapon = WeaponData(CharWeapon)
    CharList(CharIndex).Hair = HairData(CharHair)
    CharList(CharIndex).Wings = WingData(CharWings)
    
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
Dim x As Byte
Dim Y As Byte

    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Loop through until we find the object on (X,Y) then kill it
    For j = 1 To LastObj
        If OBJList(j).Pos.x = x Then
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

Dim x As Byte

    ShowGameWindow(MailboxWindow) = 0

    ShowGameWindow(ViewMessageWindow) = 1
    LastClickedWindow = ViewMessageWindow
    ReadMailData.Message = rBuf.Get_StringEX
    ReadMailData.Message = Engine_WordWrap(ReadMailData.Message, 60)
    ReadMailData.Subject = rBuf.Get_String
    ReadMailData.WriterName = rBuf.Get_String
    For x = 1 To MaxMailObjs
        ReadMailData.Obj(x) = rBuf.Get_Integer
    Next x

End Sub

Sub Data_Server_MakeChar(ByRef rBuf As DataBuffer)

'*********************************************
'Create a character and set their information
'<Body(I)><Head(I)><Heading(B)><CharIndex(I)><X(B)><Y(B)><Name(S)><Weapon(I)><Hair(I)><Wings(I)><HP%(B)><MP%(B)>
'*********************************************

Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim CharIndex As Integer
Dim x As Byte
Dim Y As Byte
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
    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Name = rBuf.Get_String
    Weapon = rBuf.Get_Integer
    Hair = rBuf.Get_Integer
    Wings = rBuf.Get_Integer
    HP = rBuf.Get_Byte
    MP = rBuf.Get_Byte

    'Create the character
    Engine_Char_Make CharIndex, Body, Head, Heading, x, Y, Name, Weapon, Hair, Wings, HP, MP

End Sub

Sub Data_Server_MakeObject(ByRef rBuf As DataBuffer)

'*********************************************
'Create an object on the object layer
'<GrhIndex(I)><X(B)><Y(B)>
'*********************************************

Dim GrhIndex As Integer
Dim x As Byte
Dim Y As Byte

'Get the values

    GrhIndex = rBuf.Get_Integer
    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Create the object
    If GrhIndex > 0 Then Engine_OBJ_Create GrhIndex, x, Y

End Sub

Sub Data_Server_MoveChar(ByRef rBuf As DataBuffer)

'*********************************************
'Move a character
'<CharIndex(I)><X(B)><Y(B)>
'*********************************************

Dim CharIndex As Integer
Dim x As Byte
Dim Y As Byte

    CharIndex = rBuf.Get_Integer

    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte
    Engine_Char_Move_ByPos CharIndex, x, Y

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
    If CharIndex > UBound(CharList()) Then Exit Sub

    'Create the blood
    Engine_Blood_Create CharList(CharIndex).Pos.x, CharList(CharIndex).Pos.Y

    'Create the damage
    Engine_Damage_Create CharList(CharIndex).Pos.x, CharList(CharIndex).Pos.Y, Damage

End Sub

Sub Data_Server_SetUserPosition(ByRef rBuf As DataBuffer)

'*********************************************
'Set the user's position
'<X(B)><Y(B)>
'*********************************************

Dim x As Byte
Dim Y As Byte

'Get the position

    x = rBuf.Get_Byte
    Y = rBuf.Get_Byte

    'Check for a valid range
    If x < XMinMapSize Then Exit Sub
    If x > XMaxMapSize Then Exit Sub
    If Y < YMinMapSize Then Exit Sub
    If Y > YMaxMapSize Then Exit Sub

    'Update the user's position
    UserPos.x = x
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
Dim x As Long
Dim Y As Long

    SkillID = rBuf.Get_Byte
    CasterIndex = rBuf.Get_Integer
    TargetIndex = rBuf.Get_Integer

    Select Case SkillID

    Case SkID.Heal

        'Set the position
        x = CharList(CasterIndex).RealPos.x + 16
        Y = CharList(CasterIndex).RealPos.Y

        'If not casted on self, bind to character
        If TargetIndex <> CasterIndex Then
            TempIndex = Effect_Heal_Begin(x, Y, 3, 120, 1)
            Effect(TempIndex).BindToChar = TargetIndex
            Effect(TempIndex).BindSpeed = 7
        Else
            TempIndex = Effect_Heal_Begin(x, Y, 3, 120, 0)
        End If

    Case SkID.Protection

        'Create the effect at (not bound to) the target character
        x = CharList(TargetIndex).RealPos.x + 16
        Y = CharList(TargetIndex).RealPos.Y
        Effect_Protection_Begin x, Y, 11, 120, 40, 15

    Case SkID.Strengthen

        'Create the effect at (not bound to) the target character
        x = CharList(TargetIndex).RealPos.x + 16
        Y = CharList(TargetIndex).RealPos.Y
        Effect_Strengthen_Begin x, Y, 12, 120, 40, 15

    Case SkID.Bless

        'Create the effect at (not bound to) the target character
        x = CharList(TargetIndex).RealPos.x + 16
        Y = CharList(TargetIndex).RealPos.Y
        Effect_Bless_Begin x, Y, 3, 120, 40, 15

    Case SkID.SpikeField

        'Create the spike field depending on the direction the user is facing
        x = CharList(CasterIndex).Pos.x
        Y = CharList(CasterIndex).Pos.Y
        If CharList(CasterIndex).HeadHeading = NORTH Then
            Engine_Effect_Create x - 1, Y + 1, 59
            Engine_Effect_Create x, Y + 1, 59
            Engine_Effect_Create x + 1, Y + 1, 59

            Engine_Effect_Create x - 2, Y, 59
            Engine_Effect_Create x - 1, Y, 59
            Engine_Effect_Create x, Y, 59
            Engine_Effect_Create x + 1, Y, 59
            Engine_Effect_Create x + 2, Y, 59

            Engine_Effect_Create x - 2, Y - 1, 59
            Engine_Effect_Create x - 1, Y - 1, 59
            Engine_Effect_Create x, Y - 1, 59
            Engine_Effect_Create x + 1, Y - 1, 59
            Engine_Effect_Create x + 2, Y - 1, 59

            Engine_Effect_Create x - 2, Y - 2, 59
            Engine_Effect_Create x - 1, Y - 2, 59
            Engine_Effect_Create x, Y - 2, 59
            Engine_Effect_Create x + 1, Y - 2, 59
            Engine_Effect_Create x + 2, Y - 2, 59

            Engine_Effect_Create x - 1, Y - 3, 59
            Engine_Effect_Create x, Y - 3, 59
            Engine_Effect_Create x + 1, Y - 3, 59

            Engine_Effect_Create x, Y - 4, 59
        ElseIf CharList(CasterIndex).HeadHeading = EAST Then
            Engine_Effect_Create x - 1, Y - 1, 59
            Engine_Effect_Create x - 1, Y, 59
            Engine_Effect_Create x - 1, Y + 1, 59

            Engine_Effect_Create x, Y - 2, 59
            Engine_Effect_Create x, Y - 1, 59
            Engine_Effect_Create x, Y, 59
            Engine_Effect_Create x, Y + 1, 59
            Engine_Effect_Create x, Y + 2, 59

            Engine_Effect_Create x + 1, Y - 2, 59
            Engine_Effect_Create x + 1, Y - 1, 59
            Engine_Effect_Create x + 1, Y, 59
            Engine_Effect_Create x + 1, Y + 1, 59
            Engine_Effect_Create x + 1, Y + 2, 59

            Engine_Effect_Create x + 2, Y - 2, 59
            Engine_Effect_Create x + 2, Y - 1, 59
            Engine_Effect_Create x + 2, Y, 59
            Engine_Effect_Create x + 2, Y + 1, 59
            Engine_Effect_Create x + 2, Y + 2, 59

            Engine_Effect_Create x + 3, Y - 1, 59
            Engine_Effect_Create x + 3, Y, 59
            Engine_Effect_Create x + 3, Y + 1, 59

            Engine_Effect_Create x + 4, Y, 59
        ElseIf CharList(CasterIndex).HeadHeading = SOUTH Then
            Engine_Effect_Create x - 1, Y - 1, 59
            Engine_Effect_Create x, Y - 1, 59
            Engine_Effect_Create x + 1, Y - 1, 59

            Engine_Effect_Create x - 2, Y, 59
            Engine_Effect_Create x - 1, Y, 59
            Engine_Effect_Create x, Y, 59
            Engine_Effect_Create x + 1, Y, 59
            Engine_Effect_Create x + 2, Y, 59

            Engine_Effect_Create x - 2, Y + 1, 59
            Engine_Effect_Create x - 1, Y + 1, 59
            Engine_Effect_Create x, Y + 1, 59
            Engine_Effect_Create x + 1, Y + 1, 59
            Engine_Effect_Create x + 2, Y + 1, 59

            Engine_Effect_Create x - 2, Y + 2, 59
            Engine_Effect_Create x - 1, Y + 2, 59
            Engine_Effect_Create x, Y + 2, 59
            Engine_Effect_Create x + 1, Y + 2, 59
            Engine_Effect_Create x + 2, Y + 2, 59

            Engine_Effect_Create x - 1, Y + 3, 59
            Engine_Effect_Create x, Y + 3, 59
            Engine_Effect_Create x + 1, Y + 3, 59

            Engine_Effect_Create x, Y + 4, 59
        ElseIf CharList(CasterIndex).HeadHeading = WEST Then
            Engine_Effect_Create x + 1, Y - 1, 59
            Engine_Effect_Create x + 1, Y, 59
            Engine_Effect_Create x + 1, Y + 1, 59

            Engine_Effect_Create x, Y - 2, 59
            Engine_Effect_Create x, Y - 1, 59
            Engine_Effect_Create x, Y, 59
            Engine_Effect_Create x, Y + 1, 59
            Engine_Effect_Create x, Y + 2, 59

            Engine_Effect_Create x - 1, Y - 2, 59
            Engine_Effect_Create x - 1, Y - 1, 59
            Engine_Effect_Create x - 1, Y, 59
            Engine_Effect_Create x - 1, Y + 1, 59
            Engine_Effect_Create x - 1, Y + 2, 59

            Engine_Effect_Create x - 2, Y - 2, 59
            Engine_Effect_Create x - 2, Y - 1, 59
            Engine_Effect_Create x - 2, Y, 59
            Engine_Effect_Create x - 2, Y + 1, 59
            Engine_Effect_Create x - 2, Y + 2, 59

            Engine_Effect_Create x - 3, Y - 1, 59
            Engine_Effect_Create x - 3, Y, 59
            Engine_Effect_Create x - 3, Y + 1, 59

            Engine_Effect_Create x - 4, Y, 59
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
Dim x As Byte
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
            x = x + 1
            If x > SkillListWidth Then
                x = 1
                Y = Y + 1
            End If

            'Set the skill list ID and Position
            SkillList(SkillListSize).SkillID = i
            SkillList(SkillListSize).x = SkillListX - (x * 32)
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
