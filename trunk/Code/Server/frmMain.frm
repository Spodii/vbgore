VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{57B1A102-EC84-4617-AC39-415819BFEC5F}#1.0#0"; "GOREsockServer.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbGORE Server"
   ClientHeight    =   750
   ClientLeft      =   1950
   ClientTop       =   1830
   ClientWidth     =   3480
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   232
   StartUpPosition =   2  'CenterScreen
   Begin GOREsock.GOREsockServer GOREsock 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSWinsockLib.Winsock ServerSocket 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Begin VB.Menu mnudebug 
         Caption         =   "Debugging"
         Begin VB.Menu mnulogs 
            Caption         =   "Logs"
            Begin VB.Menu mnugeneral 
               Caption         =   "&General"
            End
            Begin VB.Menu mnucodetracker 
               Caption         =   "Code &Tracker"
            End
            Begin VB.Menu mnuin 
               Caption         =   "Packets &In"
            End
            Begin VB.Menu mnuout 
               Caption         =   "Packets &Out"
            End
            Begin VB.Menu mnucritical 
               Caption         =   "&Critical"
            End
            Begin VB.Menu mnupacket 
               Caption         =   "Invalid &Packets"
            End
            Begin VB.Menu mnusep2 
               Caption         =   "-"
            End
            Begin VB.Menu mnubrowselog 
               Caption         =   "&Browse..."
            End
         End
         Begin VB.Menu mnupacketout 
            Caption         =   "Packets out count"
         End
         Begin VB.Menu mnupacketin 
            Caption         =   "Packets in count"
         End
         Begin VB.Menu mnufps 
            Caption         =   "Server FPS graph"
         End
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnushutdown 
         Caption         =   "&Shut down"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private Sub Form_Load()
Dim TempSplit() As String

    'Show the form
    Me.Caption = "Creating server..."
    Me.Height = 0
    Me.Width = 5000
    Me.Show
    DoEvents
    
    'This MUST be called before any timeGetTime calls
    InitTimeGetTime

    'Set the server priority
    If RunHighPriority Then
        SetThreadPriority GetCurrentThread, 2       'Reccomended you dont touch these values
        SetPriorityClass GetCurrentProcess, &H80    ' unless you know what you're doing
    End If
    
    'Set the file paths
    InitFilePaths

    'Get the ID of this server
    'Check first if an ID is specified in the ID (ie 1.exe overwrites to ID = 1)
    TempSplit = Split(App.EXEName, ".")
    If IsNumeric(TempSplit(0)) Then
        If Val(TempSplit(0)) > 0 Then ServerID = Val(TempSplit(0))
    End If
    'No server ID defined in the EXE name, get it from the file
    If ServerID = 0 Then ServerID = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "ServerID", 1))
    
    'Kill old log files (only if one server, multiple servers will cause access errors)                                         '//\\LOGLINE//\\
    If DEBUG_UseLogging Then                                                                                                    '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\CodeTracker.log", vbNormal)) Then Kill LogPath & ServerID & "\CodeTracker.log"      '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\CriticalError.log", vbNormal)) Then Kill LogPath & ServerID & "\CriticalError.log"  '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\PacketIn.log", vbNormal)) Then Kill LogPath & ServerID & "\PacketIn.log"            '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\PacketOut.log", vbNormal)) Then Kill LogPath & ServerID & "\PacketOut.log"          '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\packetsin.txt", vbNormal)) Then Kill LogPath & ServerID & "\packetsin.txt"          '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\packetsout.txt", vbNormal)) Then Kill LogPath & ServerID & "\packetsout.txt"        '//\\LOGLINE//\\
        If LenB(Dir$(LogPath & ServerID & "\serverfps.txt", vbNormal)) Then Kill LogPath & ServerID & "\serverfps.txt"          '//\\LOGLINE//\\
    End If                                                                                                                      '//\\LOGLINE//\\

    'Modify the menu
    If Not DEBUG_UseLogging Then            '//\\LOGLINE//\\
        mnugeneral.Enabled = False
        mnuin.Enabled = False
        mnuout.Enabled = False
        mnupacket.Enabled = False
        mnucodetracker.Enabled = False
        mnucritical.Enabled = False
    End If                                  '//\\LOGLINE//\\
    If Not DEBUG_RecordPacketsOut Then mnupacketout.Enabled = False
    If Not DEBUG_RecordPacketsIn Then mnupacketin.Enabled = False
    If Not DEBUG_MapFPS Then mnufps.Enabled = False
    
    'Create conversion buffer
    Set ConBuf = New DataBuffer

    'Load MySQL variables
    Me.Caption = "Connecting to MySQL..."
    Me.Refresh
    MySQL_Init
    
    'Generate the packet keys
    GenerateEncryptionKeys TempSplit
    frmMain.GOREsock.ClearPicture
    frmMain.GOREsock.SetEncryption PacketEncTypeServerIn, PacketEncTypeServerOut, TempSplit
    
    'Start the server
    StartServer

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Show the pop-up menu
    If X = RightUp Then Me.PopupMenu mnu, 0, , , mnushutdown

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnloadServer = 1

End Sub
Private Sub mnufps_Click()
    
    'Save the FPS values
    Save_FPS
        
    'Load the graph
    Shell App.Path & "\ToolServerFPSViewer.exe", vbMaximizedFocus
    
End Sub

Private Sub mnubrowselog_Click()                                                                                    '//\\LOGLINE//\\
    Shell "explorer " & LogPath, vbMaximizedFocus                                                                   '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\
Private Sub mnucritical_Click()                                                                                     '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & ServerID & "\CriticalError.log", vbMaximizedFocus         '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\
Private Sub mnucodetracker_Click()                                                                                  '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & ServerID & "\CodeTracker.log", vbMaximizedFocus           '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\
Private Sub mnugeneral_Click()                                                                                      '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & ServerID & "\General.log", vbMaximizedFocus               '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\
Private Sub mnuin_Click()                                                                                           '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & ServerID & "\PacketIn.log", vbMaximizedFocus              '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\
Private Sub mnuout_Click()                                                                                          '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & ServerID & "\PacketOut.log", vbMaximizedFocus             '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\
Private Sub mnupacket_Click()                                                                                       '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & ServerID & "\InvalidPacketData.log", vbMaximizedFocus     '//\\LOGLINE//\\
End Sub                                                                                                             '//\\LOGLINE//\\

Private Sub mnupacketin_Click()

    'Save the file
    Save_PacketsIn
    
    'Display the file
    Shell "notepad " & LogPath & ServerID & "\packetsin.txt", vbMaximizedFocus

End Sub

Private Sub mnupacketout_Click()

    'Save the file
    Save_PacketsOut
    
    'Display the file
    Shell "notepad " & LogPath & ServerID & "\packetsout.txt", vbMaximizedFocus

End Sub

Private Sub mnushutdown_Click()

'*********************************************
'Shut down the server
'*********************************************
    
    UnloadServer = 1

End Sub

Private Sub StartServer()

'*****************************************************************
'Load up server
'*****************************************************************
Dim LoopC As Long
Dim s() As String
Dim i As Long

    Log "Call StartServer", CodeTracker '//\\LOGLINE//\\
    
    'This holds an array of indicies for us to use - doing it this way is slow, but user-friendly and its done at runtime anyways
    Const cMessages As String = "2,7,8,12,17,20,24,25,26,29,33,34,36,37,38,48,49," & _
        "51,57,60,61,64,69,70,79,81,82,83,84,85,97,98,99,101,102,109,111,112,113,114," & _
        "116,119,121,123,125,127,130,131,132,133"
    
    'Make the server temp path
    MakeSureDirectoryPathExists ServerTempPath
    
    'Set up debug packets out
    If DEBUG_RecordPacketsOut Then ReDim DebugPacketsOut(0 To 255)
    If DEBUG_RecordPacketsIn Then ReDim DebugPacketsIn(0 To 255)

    '*** Database ***
    
    'Remove online user states (in case server crashed or something else went wrong)
    Me.Caption = "Removing `online` states..."
    Me.Refresh
    MySQL_RemoveOnline
    
    'Auto-optimize the database
    If OptimizeDatabase Then
        Me.Caption = "Optimizing database..."
        Me.Refresh
        MySQL_Optimize
    End If
    
    '*** Init vars ***
    
    'How many bytes we need to fit all of our skills
    NumBytesForSkills = Int((NumSkills - 1) / 8) + 1
    
    'Load Data Commands
    InitDataCommands
    
    '*** Build help messages ***
    
    'Get the number of lines
    i = Val(Var_Get(ServerDataPath & "Help.ini", "INIT", "NumHelp"))
    
    'Put all the strings together so it can be sent in one string to the client
    ConBuf.Clear
    For LoopC = 1 To i
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String Trim$(Var_Get(ServerDataPath & "Help.ini", "INIT", Str$(LoopC)))
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Next LoopC
    HelpBuffer = ConBuf.Get_Buffer
    
    '*** Build MOTD messages ***
    
    'Get the number of lines
    i = Val(Var_Get(ServerDataPath & "MOTD.ini", "INIT", "NumLines"))
    
    'Put all the strings together so it can be sent in one string to the client
    ConBuf.Clear
    For LoopC = 1 To i
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String Trim$(Var_Get(ServerDataPath & "MOTD.ini", "INIT", Str$(LoopC)))
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Next LoopC
    MOTDBuffer = ConBuf.Get_Buffer
    
    '*** Build cached messages ***
    frmMain.Caption = "Caching constant packets..."
    frmMain.Refresh
    
    'Split up the messages
    s = Split(cMessages, ",")
    
    'Find the highest index needed and resize the array accordingly
    For LoopC = 0 To UBound(s)
        If Val(s(LoopC)) > i Then i = Val(s(LoopC))
    Next LoopC
    ReDim cMessage(1 To i)
    
    'Loop through the messages, and set the data
    For LoopC = 0 To UBound(s)
        ConBuf.PreAllocate 2
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte CByte(s(LoopC))
        cMessage(Val(s(LoopC))).Data = ConBuf.Get_Buffer
    Next LoopC

    '*** Load data ***
    frmMain.Caption = "Loading maps..."
    frmMain.Refresh
    Load_Maps
    
    frmMain.Caption = "Loading objects..."
    frmMain.Refresh
    Load_OBJs
    
    frmMain.Caption = "Loading quests..."
    frmMain.Refresh
    Load_Quests
    
    frmMain.Caption = "Creating npc files..."
    frmMain.Refresh
    Save_NPCs_Temp
    Load_NPC_Names
    
    '*** Listen (Client-To-Server) ***
    frmMain.Caption = "Loading sockets..."
    frmMain.Refresh
    
    'Change the listen settings in the Server.ini file
    LocalSocketID = frmMain.GOREsock.Listen(ServerInfo(ServerID).IIP, ServerInfo(ServerID).Port)
    frmMain.GOREsock.SetOption LocalSocketID, soxSO_TCP_NODELAY, True

    '*** Listen (Server-To-Server) ***
    If NumServers > 1 Then
    
        'Create the listen socket so we can accept connections from other servers
        ServerSocket(0).RemoteHost = ServerInfo(ServerID).IIP
        ServerSocket(0).LocalPort = ServerInfo(ServerID).ServerPort
        ServerSocket(0).Listen
        
        'Loop through all the servers (skip the ID of this one)
        For i = 1 To NumServers
            If i <> ServerID Then
            
                'Load the socket object
                Load ServerSocket(i)
                
                'Create the connect to the server (if this fails, ie server is not loaded yet, it will connect later)
                ServerSocket(i).RemotePort = ServerInfo(i).ServerPort
                ServerSocket(i).RemoteHost = ServerInfo(i).EIP
                Server_ConnectToServer i
                
            End If
        Next i

    End If

    '*** Misc ***

    'Check for a valid connection
    If frmMain.GOREsock.Address(LocalSocketID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbNewLine & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbNewLine & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

    'Set the starting time
    ServerStartTime = timeGetTime

    'Set the caption
    Me.Caption = "vbGORE v." & App.Major & "." & App.Minor & "." & App.Revision
    
    'Hide the server in the system tray
    TrayAdd Me, Server_BuildToolTipString, MouseMove
    Me.Hide
    Me.Refresh
    DoEvents
    
    'Start the main server loop
    Server_Update

End Sub

Private Sub ServerSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Byte

    'Check for a valid server
    For i = 1 To NumServers
        If i <> ServerID Then
        
            'If the IP and port match, we got a valid connection
            If ServerSocket(i).RemoteHost = ServerInfo(i).EIP Then
                If ServerSocket(i).RemotePort = ServerInfo(i).ServerPort Then
                    
                    'Match according to the corresponding server so the socket index = the server ID
                    ServerSocket(Index).Close
                    ServerSocket(i).Close
                    ServerSocket(i).Accept requestID
                    Exit For
                
                End If
            End If
            
        End If
    Next i

End Sub

Private Sub ServerSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data() As Byte
Dim rBuf As New DataBuffer
Dim CommandID As Byte
Dim BufUBound As Long

    'Get the data
    ServerSocket(Index).GetData Data, vbByte, bytesTotal
    
    'Put the data into the buffer
    rBuf.Set_Buffer Data()
    
    'Hold the buffer ubound
    BufUBound = UBound(Data)
    
    'Loop through the packet and get the commands, just like with the GOREsock DataArrival
    Do
    
        CommandID = rBuf.Get_Byte
        
        With DataCode
            Select Case CommandID
                Case .Comm_Shout: Data_Comm_Shout_ServerToServer rBuf
                Case .Comm_Whisper: Data_Comm_Whisper_ServerToServer rBuf
                Case .Server_Message: Data_Server_Message_ServerToServer rBuf
            End Select
        End With
        
        'Check if the buffer ran out
        If rBuf.Get_ReadPos > BufUBound Then Exit Sub
        
    Loop

End Sub

Private Sub GOREsock_OnClose(inSox As Long)

'*********************************************
'Socket was closed - make sure the user is logged off and reset the ConnID
'*********************************************

    Log "Call frmMain.GOREsock.Close(" & inSox & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the user is in a valid range
    If inSox > LastUser Then Exit Sub
    If inSox < 1 Then Exit Sub
    
    'If the user is logged in still, close them down so they can be removed properly
    If UserList(inSox).flags.UserLogged = 1 Then User_Close inSox

End Sub

Private Sub GOREsock_OnConnecting(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Private Sub GOREsock_OnConnection(inSox As Long)

'*********************************************
'Accepts new user and assigns an open Index
'*********************************************

    Log "Call frmMain.GOREsock.Connection(" & inSox & ")", CodeTracker '//\\LOGLINE//\\

    'Check for max users
    If inSox > MaxUsers Then Exit Sub
    
    'Make sure Nagling is off
    frmMain.GOREsock.SetOption inSox, soxSO_TCP_NODELAY, True

    'Check that the userlist is high enough
    If inSox > LastUser Then
        LastUser = inSox
        ReDim Preserve UserList(1 To inSox)
    End If

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************
Dim Index As Integer
Dim rBuf As DataBuffer
Dim BufUBound As Long
Dim CommandID As Byte
Dim CommandID2 As Byte
Dim i As Byte

    Log "Call frmMain.GOREsock.DataArrival(" & inSox & "," & ByteArrayToStr(inData) & ")", CodeTracker '//\\LOGLINE//\\

    'Get the UserIndex
    Index = inSox
    If Index > LastUser Then Exit Sub
    
    'If it is a character disconnecting, do not check their packets since they're doodie heads
    If UserList(Index).flags.Disconnecting Then Exit Sub
    
    'Reset the user's packet counter
    UserList(Index).Counters.LastPacket = timeGetTime
    
    'Calculate data transfer rate
    'TCP header = 20 bytes, IPv4 header = 20 bytes
    BufUBound = UBound(inData)
    If CalcTraffic Then DataIn = DataIn + BufUBound + 41  '+ 1 because we have to count inData(0)
    
    'Check if to reset the packet flood timer
    If UserList(Index).Counters.PacketsInTime + 1000 < timeGetTime Then
        UserList(Index).Counters.PacketsInTime = timeGetTime
        UserList(Index).Counters.PacketsInCount = 0
    End If

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
    'Debug.Print StrConv(inData, vbUnicode)
    'Debug.Print s
    
    Log "Receive: " & ByteArrayToStr(rBuf.Get_Buffer), PacketIn '//\\LOGLINE//\\

    'Loop through the data buffer until it's empty
    'If done right, we should use up exactly every byte in the buffer
    Do

        'Raise the packets in count and check if the user has been flooding packets
        UserList(Index).Counters.PacketsInCount = UserList(Index).Counters.PacketsInCount + 1
        If UserList(Index).Counters.PacketsInCount > MaxPacketsInPerSec Then Exit Do

        'Get the CommandID
        CommandID = rBuf.Get_Byte
        If DEBUG_RecordPacketsIn Then DebugPacketsIn(CommandID) = DebugPacketsIn(CommandID) + 1
    
        If CommandID >= 100 Then                                                                                '//\\LOGLINE//\\
            Log " * ID: " & CommandID & "  Data Left: " & ByteArrayToStr(rBuf.Get_Buffer_Remainder), PacketIn   '//\\LOGLINE//\\
        ElseIf CommandID >= 10 Then                                                                             '//\\LOGLINE//\\
            Log " * ID: 0" & CommandID & "  Data Left: " & ByteArrayToStr(rBuf.Get_Buffer_Remainder), PacketIn  '//\\LOGLINE//\\
        Else                                                                                                    '//\\LOGLINE//\\
            Log " * ID: 00" & CommandID & "  Data Left: " & ByteArrayToStr(rBuf.Get_Buffer_Remainder), PacketIn '//\\LOGLINE//\\
        End If                                                                                                  '//\\LOGLINE//\\
        
        'Make the appropriate call based on the CommandID
        With DataCode
        
            'Reset idle counter
            UserList(Index).Counters.IdleCount = timeGetTime

            Select Case CommandID
            
            Case 0
                Exit Do
       
            Case .Comm_Emote: Data_Comm_Emote rBuf, Index
            Case .Comm_GroupTalk: Data_Comm_GroupTalk rBuf, Index
            Case .Comm_Shout: Data_Comm_Shout rBuf, Index
            Case .Comm_Talk: Data_Comm_Talk rBuf, Index
            Case .Comm_Whisper: Data_Comm_Whisper rBuf, Index

            Case .GM_Approach: Data_GM_Approach rBuf, Index
            Case .GM_BanIP: Data_GM_BanIP rBuf, Index
            Case .GM_DeThrall: Data_GM_DeThrall rBuf, Index
            Case .GM_Kick: Data_GM_Kick rBuf, Index
            Case .GM_Raise: Data_GM_Raise rBuf, Index
            Case .GM_SetGMLevel: Data_GM_SetGMLevel rBuf, Index
            Case .GM_Summon: Data_GM_Summon rBuf, Index
            Case .GM_Thrall: Data_GM_Thrall rBuf, Index
            Case .GM_UnBanIP: Data_GM_UnBanIP rBuf, Index
            Case .GM_Warp: Data_GM_Warp rBuf, Index
            
            Case .Map_DoneLoadingMap: Data_Map_DoneLoadingMap Index

            Case .Server_Help: Data_Server_Help Index
            Case .Server_MailCompose: Data_Server_MailCompose rBuf, Index
            Case .Server_MailDelete: Data_Server_MailDelete rBuf, Index
            Case .Server_MailItemTake: Data_Server_MailItemTake rBuf, Index
            Case .Server_MailMessage: Data_Server_MailMessage rBuf, Index
            Case .Server_SetUserPosition: Data_Server_SetUserPosition rBuf, Index
            Case .Server_Who: Data_Server_Who Index

            Case .User_Attack: Data_User_Attack rBuf, Index
            Case .User_Bank_Balance: Data_User_Bank_Balance Index
            Case .User_Bank_Deposit: Data_User_Bank_Deposit rBuf, Index
            Case .User_Bank_PutItem: Data_User_Bank_PutItem rBuf, Index
            Case .User_Bank_TakeItem: Data_User_Bank_TakeItem rBuf, Index
            Case .User_Bank_Withdraw: Data_User_Bank_Withdraw rBuf, Index
            Case .User_BaseStat: Data_User_BaseStat rBuf, Index
            Case .User_Blink: Data_User_Blink Index
            Case .User_CancelQuest: Data_User_CancelQuest rBuf, Index
            Case .User_CastSkill: Data_User_CastSkill rBuf, Index
            Case .User_ChangeInvSlot: Data_User_ChangeInvSlot rBuf, Index
            Case .User_Desc: Data_User_Desc rBuf, Index
            Case .User_Drop: Data_User_Drop rBuf, Index
            Case .User_Emote: Data_User_Emote rBuf, Index
            Case .User_Get: Data_User_Get Index
            Case .User_Group_Info: Data_User_Group_Info Index
            Case .User_Group_Invite: Data_User_Group_Invite rBuf, Index
            Case .User_Group_Join: Data_User_Group_Join Index
            Case .User_Group_Leave: Data_User_Group_Leave Index
            Case .User_Group_Make: Data_User_Group_Make Index
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
            Case .User_Target: Data_User_Target rBuf, Index
            Case .User_Trade_Accept: Data_User_Trade_Accept Index
            Case .User_Trade_BuyFromNPC: Data_User_Trade_BuyFromNPC rBuf, Index
            Case .User_Trade_Finish: Data_User_Trade_Finish Index
            Case .User_Trade_Cancel: Data_User_Trade_Cancel Index
            Case .User_Trade_RemoveItem: Data_User_Trade_RemoveItem rBuf, Index
            Case .User_Trade_SellToNPC: Data_User_Trade_SellToNPC rBuf, Index
            Case .User_Trade_Trade: Data_User_Trade_Trade rBuf, Index
            Case .User_Trade_UpdateTrade: Data_User_Trade_UpdateTrade rBuf, Index
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
