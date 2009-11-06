VERSION 5.00
Object = "{41E14DE4-8B3E-4E31-89BB-654FD3A78163}#1.0#0"; "goresockmicro.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbGORE Server"
   ClientHeight    =   2640
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   StartUpPosition =   2  'CenterScreen
   Begin GORESOCKmicro.Socket Socket 
      Height          =   660
      Left            =   2520
      Top             =   1800
      Visible         =   0   'False
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
   End
   Begin VB.Timer DataCalcTmr 
      Interval        =   1000
      Left            =   3000
      Top             =   1320
   End
   Begin VB.TextBox BytesOutTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Overall average KBytes/sec uploaded to the clients (41 byte TCP/IP packet headers included)"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox BytesInTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Overall average KBytes/sec downloaded from the clients (41 byte TCP/IP packet headers included)"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2520
      Top             =   1320
   End
   Begin VB.ListBox Userslst 
      Appearance      =   0  'Flat
      Height          =   2130
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2355
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "KBytes Out / Sec:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "KBytes In / Sec:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Users:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public RecoverTimer As Long

Private Sub DataCalcTmr_Timer()

    'Turn bytes into kilobytes
    If DataIn > 1024 Then
        Do While DataIn > 1024
            DataIn = DataIn - 1024
            DataKBIn = DataKBIn + 1
        Loop
    End If

    If DataOut > 1024 Then
        Do While DataOut > 1024
            DataOut = DataOut - 1024
            DataKBOut = DataKBOut + 1
        Loop
    End If

    'Display statistics (KB)
    BytesInTxt.Text = Round((DataKBIn + (DataIn / 1024)) / ((timeGetTime - ServerStartTime) * 0.001), 6)
    BytesOutTxt.Text = Round((DataKBOut + (DataOut / 1024)) / ((timeGetTime - ServerStartTime) * 0.001), 6)

    'Display statistics (Bytes)
    'BytesInTxt.Text = Round(((DataKBIn * 1024) + DataIn) / ((timeGetTime - ServerStartTime) / 1000), 6)
    'BytesOutTxt.Text = Round(((DataKBOut * 1024) + DataOut) / ((timeGetTime - ServerStartTime) / 1000), 6)

End Sub

Private Sub Form_Load()

    'Create conversion buffer
    Set ConBuf = New DataBuffer
    
    'Load MySQL variables
    MySQL_Init

    'Set the file paths
    InitFilePaths

    'Set timeGetTime to a high resolution
    timeBeginPeriod 1

    'Start the server
    StartServer

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case X
        Case MouseMove
        Case LeftUp
        Case LeftDown
            'Return from system tray
            If Me.WindowState = 1 Then
                TrayDelete
                Me.WindowState = 0
                Me.Show
                DataCalcTmr.Enabled = True
            End If
        Case LeftDbClick
        Case RightUp
        Case RightDown
        Case RightDbClick
    End Select

End Sub

Private Sub Form_Resize()

'If the form becomes minimized, move to system tray

    If WindowState = 1 Then
        TrayAdd Me, "Game Server: " & CurrConnections & " connections", MouseMove
        Me.Hide
        DataCalcTmr.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim FileNum As Byte
Dim LoopC As Long
Dim S As String

    If Socket.ShutDown = soxERROR Then 'Terminate will be True if we have ShutDown properly
        If MsgBox("ShutDown procedure has not completed!" & vbCrLf & "(Hint - Select No and Try again!)" & vbCrLf & "Execute Forced ShutDown?", vbApplicationModal + vbCritical + vbYesNo, "UNABLE TO COMPLY!") = vbNo Then
            Let Cancel = True
            Exit Sub
        Else
            Socket.UnHook  'Unfortunately for now, I can't get around doing this automatically for you :( VB crashes if you don't do this!
        End If
    Else
        Socket.UnHook  'The reason is VB closes my Mod which stores the WindowProc function used for SubClassing and VB doesn't know that! So it closes the Mod before the Control!
    End If
    
    'Kill the database connection
    DB_Conn.Close

    'Save the debug packets out
    If DEBUG_RecordPacketsOut Then
        S = vbCrLf
        For LoopC = 0 To 255
            S = S & LoopC & ": " & DebugPacketsOut(LoopC) & vbCrLf
        Next LoopC
        FileNum = FreeFile
        Open App.Path & "\packetsout.txt" For Output As #FileNum
            Write #FileNum, S
        Close #FileNum
    End If

    'Deallocate all arrays to avoid memory leaks
    Erase UserList
    Erase NPCList
    Erase MapData
    Erase MapInfo
    Erase CharList
    Erase ObjData

    'Same with connection Groups
    For LoopC = 1 To NumMaps
        Erase ConnectionGroups(LoopC).UserIndex
    Next LoopC
    Erase ConnectionGroups

    End

End Sub

Private Sub GameTimer_Timer()

'*****************************************************************
'Update world
'*****************************************************************

Static UpdateStats As Long
Dim UserIndex As Integer
Dim NPCIndex As Integer
Dim Recover As Boolean
Dim Update As Boolean

'Update current time

    Elapsed = timeGetTime - LastTime
    LastTime = timeGetTime

    'Check if it is time to recover stats
    If RecoverTimer <= timeGetTime - STAT_RECOVERRATE Then
        Recover = True
        RecoverTimer = timeGetTime
    End If

    'Update Users
    For UserIndex = 1 To LastUser

        'Make sure user is logged on
        If UserList(UserIndex).Flags.UserLogged Then

            'Check if it has been idle for too long
            If UserList(UserIndex).Counters.IdleCount <= timeGetTime - IdleLimit Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 85
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Server_CloseSocket UserIndex
                Exit Sub
            End If
            
            'Check if the user was possible disconnected (or extremely laggy)
            If UserList(UserIndex).Counters.LastPacket <= timeGetTime - LastPacket Then
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.Server_Message
                ConBuf.Put_Byte 85
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Server_CloseSocket UserIndex
                Exit Sub
            End If
            
            'Check if stats need to be recovered
            If Recover = True Then

                'Check if Health needs to be updated
                If UserList(UserIndex).Stats.BaseStat(SID.MinHP) < UserList(UserIndex).Stats.ModStat(SID.MaxHP) Then
                    UserList(UserIndex).Stats.BaseStat(SID.MinHP) = UserList(UserIndex).Stats.BaseStat(SID.MinHP) + 1 + UserList(UserIndex).Stats.ModStat(SID.Str) * 0.5
                End If

                'Check if Stamina needs to be updated
                If UserList(UserIndex).Stats.BaseStat(SID.MinSTA) < UserList(UserIndex).Stats.ModStat(SID.MaxSTA) Then
                    UserList(UserIndex).Stats.BaseStat(SID.MinSTA) = UserList(UserIndex).Stats.BaseStat(SID.MinSTA) + 1 + UserList(UserIndex).Stats.ModStat(SID.Agi) * 0.5
                End If

                'Check if Mana needs to be updated
                If UserList(UserIndex).Stats.BaseStat(SID.MinMAN) < UserList(UserIndex).Stats.ModStat(SID.MaxMAN) Then
                    UserList(UserIndex).Stats.BaseStat(SID.MinMAN) = UserList(UserIndex).Stats.BaseStat(SID.MinMAN) + 1 + UserList(UserIndex).Stats.ModStat(SID.Mag) * 0.5
                End If

            End If

            'Update the spell lengths
            If UserList(UserIndex).Counters.BlessCounter > 0 Then
                UserList(UserIndex).Counters.BlessCounter = UserList(UserIndex).Counters.BlessCounter - Elapsed
                If UserList(UserIndex).Counters.BlessCounter <= 0 Then
                    UserList(UserIndex).Skills.Bless = 0
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_IconBlessed
                    ConBuf.Put_Byte 0
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                End If
            End If
            If UserList(UserIndex).Counters.ProtectCounter > 0 Then
                UserList(UserIndex).Counters.ProtectCounter = UserList(UserIndex).Counters.ProtectCounter - Elapsed
                If UserList(UserIndex).Counters.ProtectCounter <= 0 Then
                    UserList(UserIndex).Skills.Protect = 0
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_IconProtected
                    ConBuf.Put_Byte 0
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                End If
            End If
            If UserList(UserIndex).Counters.StrengthenCounter > 0 Then
                UserList(UserIndex).Counters.StrengthenCounter = UserList(UserIndex).Counters.StrengthenCounter - Elapsed
                If UserList(UserIndex).Counters.StrengthenCounter <= 0 Then
                    UserList(UserIndex).Skills.Strengthen = 0
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_IconStrengthened
                    ConBuf.Put_Byte 0
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                End If
            End If

            'Update spell exhaustion
            If UserList(UserIndex).Counters.SpellExhaustion > 0 Then
                UserList(UserIndex).Counters.SpellExhaustion = UserList(UserIndex).Counters.SpellExhaustion - Elapsed
                If UserList(UserIndex).Counters.SpellExhaustion <= 0 Then
                    UserList(UserIndex).Counters.SpellExhaustion = 0
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
                    ConBuf.Put_Byte 0
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                End If
            End If

            'Update aggressive-face timer
            If UserList(UserIndex).Counters.AggressiveCounter > 0 Then
                UserList(UserIndex).Counters.AggressiveCounter = UserList(UserIndex).Counters.AggressiveCounter - Elapsed
                If UserList(UserIndex).Counters.AggressiveCounter <= 0 Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.User_AggressiveFace
                    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                    ConBuf.Put_Byte 0
                    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
                    UserList(UserIndex).Counters.AggressiveCounter = 0
                End If
            End If
            
            'Check if the packet wait time has passed
            If UserList(UserIndex).HasBuffer Then
                If UserList(UserIndex).PacketWait > 0 Then UserList(UserIndex).PacketWait = UserList(UserIndex).PacketWait - Elapsed
                If UserList(UserIndex).PacketWait <= 0 Then
                    
                    'Send the packet buffer to the user
                    If UserList(UserIndex).PPValue = PP_High Then
                        
                        'High priority - send asap
                        Data_Send_Buffer UserIndex
                        
                    ElseIf UserList(UserIndex).PPValue = PP_Low Then
                        
                        'Low priority - check counter for sending
                        If UserList(UserIndex).PPCount > 0 Then UserList(UserIndex).PPCount = UserList(UserIndex).PPCount - Elapsed
                        If UserList(UserIndex).PPCount <= 0 Then Data_Send_Buffer UserIndex
                    
                    End If
                    
                End If
            End If
            
            UserList(UserIndex).Stats.SendUpdatedStats

        End If

    Next UserIndex

    'Update NPCs
    For NPCIndex = 1 To LastNPC

        'Make sure NPC is active
        If NPCList(NPCIndex).Flags.NPCActive Then

            'See if npc is alive
            If NPCList(NPCIndex).Flags.NPCAlive Then

                'Update warcurse time
                If NPCList(NPCIndex).Skills.WarCurse = 1 Then
                    NPCList(NPCIndex).Counters.WarCurseCounter = NPCList(NPCIndex).Counters.WarCurseCounter - Elapsed
                    If NPCList(NPCIndex).Counters.WarCurseCounter <= 0 Then
                        NPCList(NPCIndex).Counters.WarCurseCounter = 0
                        NPCList(NPCIndex).Skills.WarCurse = 0
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_Message
                        ConBuf.Put_Byte 1
                        ConBuf.Put_String NPCList(NPCIndex).Name
                        Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.Server_IconWarCursed
                        ConBuf.Put_Byte 0
                        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                        Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                    End If
                End If

                'Only update npcs in user populated maps
                If MapInfo(NPCList(NPCIndex).Pos.Map).NumUsers Then

                    'Check to update mod stats
                    If Update = True Then NPC_UpdateModStats NPCIndex

                    'Call the NPC AI
                    NPC_AI NPCIndex

                    'Update aggressive-face timer
                    If NPCList(NPCIndex).Counters.AggressiveCounter > 0 Then
                        NPCList(NPCIndex).Counters.AggressiveCounter = NPCList(NPCIndex).Counters.AggressiveCounter - Elapsed
                        If NPCList(NPCIndex).Counters.AggressiveCounter <= 0 Then
                            ConBuf.Clear
                            ConBuf.Put_Byte DataCode.User_AggressiveFace
                            ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                            ConBuf.Put_Byte 0
                            Data_Send ToMap, NPCIndex, ConBuf.Get_Buffer, NPCList(NPCIndex).Pos.Map
                            NPCList(NPCIndex).Counters.AggressiveCounter = 0
                        End If
                    End If

                End If

            Else

                'Check if it's time to respawn
                If NPCList(NPCIndex).Counters.RespawnCounter <= timeGetTime - NPCList(NPCIndex).RespawnWait Then NPC_Spawn NPCIndex

            End If
        End If
    Next NPCIndex

    'Check if it's time to do a World Save
    If timeGetTime - LastWorldSave >= WORLDSAVE_RATE Then

        'Save all user's data
        For UserIndex = 1 To LastUser
            If UserList(UserIndex).Flags.UserLogged Then Save_User UserList(UserIndex)
        Next UserIndex

        'Reset the save counter
        LastWorldSave = timeGetTime

    End If

End Sub

Private Sub Socket_OnClose(inSox As Long)

'*********************************************
'Socket was closed - make sure the user is logged off and reset the ConnID
'*********************************************

Dim UserIndex As Integer

    UserIndex = User_IndexFromSox(inSox)
    If UserIndex < 0 Then Exit Sub
    If UserList(UserIndex).Flags.UserLogged = 1 Then User_Close UserIndex
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).Flags.Disconnecting = 0

End Sub

Private Sub Socket_OnConnection(inSox As Long)

'*********************************************
'Accepts new user and assigns an open Index
'*********************************************

Dim Index As Integer

    Index = User_NextOpen

    'Check for max users
    If Index > MaxUsers Then Exit Sub
    UserList(Index).ConnID = inSox
    Socket.SetOption inSox, soxSO_TCP_NODELAY, True
    Socket.SetOption inSox, soxSO_RCVBUF, TCPBufferSize
    Socket.SetOption inSox, soxSO_SNDBUF, TCPBufferSize

End Sub

Private Sub Socket_OnDataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************

Dim Index As Integer
Dim rBuf As DataBuffer
Dim BufUBound As Long
Dim CommandID As Byte
Static X As Long

    'Get the UserIndex
    Index = User_IndexFromSox(inSox)
    If Index = -1 Then Exit Sub

    'If it is a character disconnecting, do not check their packets since they're doodie heads
    If UserList(Index).Flags.Disconnecting Then Exit Sub
    
    'Reset the user's packet counter
    UserList(Index).Counters.LastPacket = timeGetTime
    
    'Create the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    BufUBound = UBound(inData)

    'Calculate data transfer rate
    'TCP header = 20 bytes, IPv4 header = 20 bytes, GORESOCK header = 2 bytes
    If CalcTraffic = True Then DataIn = DataIn + BufUBound + 43 '+ 1 because we have to count inData(0)

    'Loop through the data buffer until it's empty
    'If done right, we should use up exactly every byte in the buffer
    Do

        'Get the CommandID
        CommandID = rBuf.Get_Byte

        'Make the appropriate call based on the CommandID
        With DataCode
            
            'Reset idle counter
            If CommandID <> .Server_Ping Then UserList(Index).Counters.IdleCount = timeGetTime
        
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
            
            Case .Map_DoneLoadingMap: Data_Map_DoneLoadingMap Index

            Case .Server_Help: Data_Server_Help Index
            Case .Server_MailCompose: Data_Server_MailCompose rBuf, Index
            Case .Server_MailDelete: Data_Server_MailDelete rBuf, Index
            Case .Server_MailItemInfo: Data_Server_MailItemInfo rBuf, Index
            Case .Server_MailItemTake: Data_Server_MailItemTake rBuf, Index
            Case .Server_MailMessage: Data_Server_MailMessage rBuf, Index
            Case .Server_Ping: Data_Server_Ping Index
            Case .Server_Who: Data_Server_Who Index

            Case .User_Attack: Data_User_Attack Index
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
            Case .User_RightClick: Data_User_RightClick rBuf, Index
            Case .User_Rotate: Data_User_Rotate rBuf, Index
            Case .User_StartQuest: Data_User_StartQuest Index
            Case .User_Trade_BuyFromNPC: Data_User_Trade_BuyFromNPC rBuf, Index
            Case .User_Use: Data_User_Use rBuf, Index
            
            Case Else
                If DEBUG_PrintPacketReadErrors Then Debug.Print "Command ID " & CommandID & " caused a premature packet handling abortion!"
                Exit Do 'Something went wrong or we hit the end, either way, RUN!!!!
                
            End Select

        End With

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do

    Loop

End Sub

Private Sub StartServer()

'*****************************************************************
'Load up server
'*****************************************************************
Dim LoopC As Long

    'Show the form
    Me.Show
    DoEvents

    'Check if server is already started
    If GameTimer.Enabled = True Then Exit Sub
    
    'Set up debug packets out
    If DEBUG_RecordPacketsOut Then ReDim DebugPacketsOut(0 To 255)

    '*** Database ***
    
    'Remove online user states (in case server crashed or something else went wrong)
    Me.Caption = "Removing `online` states..."
    Me.Refresh
    MySQL_RemoveOnline
    
    'Auto-optimize the database
    Me.Caption = "Optimizing database..."
    Me.Refresh
    MySQL_Optimize

    '*** Init vars ***
    
    'Setup Map borders
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
    'Load Data Commands
    InitDataCommands

    'Calculate the max distance between a char and another in it's PC area
    Max_Server_Distance = Fix(Sqr((MinYBorder - 1) ^ 2 + (MinXBorder - 1) ^ 2))

    'Reset User connections
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
    Next LoopC

    'Set up the help lines
    HelpLine(1) = "To move, use W A S D or arrow keys."
    HelpLine(2) = "To attack, use Ctrl, and get objects with Alt."
    HelpLine(3) = "As many help lines as you want can be added..."

    '*** Load data ***
    Load_ServerIni
    Load_OBJs
    Load_Quests
    Load_Maps
    
    '*** Listen ***
    frmMain.Caption = "Loading sockets..."
    frmMain.Refresh
    
    'Change the 127.0.0.1 to 0.0.0.0 or your internal IP to make the server public
    LocalSoxID = Socket.Listen("127.0.0.1", Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "GamePort")))
    Socket.SetOption LocalSoxID, soxSO_TCP_NODELAY, True
    Socket.SetEncryption eRC4, "1!jkl@!$:!(Zxq"
    Socket.SetBufferSize TCPBufferSize
    Socket.SetOption LocalSoxID, soxSO_RCVBUF, TCPBufferSize
    Socket.SetOption LocalSoxID, soxSO_SNDBUF, TCPBufferSize

    '*** Misc ***
    'Calculate data transfer rate
    If CalcTraffic = True Then DataCalcTmr.Enabled = True

    'Show status
    Server_RefreshUserListBox

    'Show local IP/Port
    If frmMain.Socket.Address(LocalSoxID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbCrLf & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbCrLf & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly
    
    'Initialize LastWorldSave
    LastWorldSave = timeGetTime
    RecoverTimer = timeGetTime

    'Start Game timer
    GameTimer.Enabled = True

    'Set the starting time
    ServerStartTime = timeGetTime

    'Done
    Me.Caption = "vbGORE V." & App.Major & "." & App.Minor & "." & App.Revision

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:48)  Decl: 3  Code: 650  Total: 653 Lines
':) CommentOnly: 95 (14.5%)  Commented: 5 (0.8%)  Empty: 144 (22.1%)  Max Logic Depth: 8
Private Sub Userslst_Click()

End Sub
