VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbGORE Server"
   ClientHeight    =   1365
   ClientLeft      =   1950
   ClientTop       =   1830
   ClientWidth     =   4890
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
   ScaleHeight     =   91
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   StartUpPosition =   2  'CenterScreen
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

    'Show the form
    Me.Caption = "Creating server..."
    Me.Height = 460
    Me.Width = 5000
    Me.Show
    DoEvents
    
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
    
    'Set the timer accuracy
    timeBeginPeriod 1
    
    'Set the file paths
    InitFilePaths
    
    'Set the server priority
    If RunHighPriority Then
        SetThreadPriority GetCurrentThread, 2       'Reccomended you dont touch these values
        SetPriorityClass GetCurrentProcess, &H80    ' unless you know what you're doing
    End If
    
    'Load MySQL variables
    Me.Caption = "Connecting to MySQL..."
    Me.Refresh
    MySQL_Init
    
    'Generate the packet keys
    GenerateEncryptionKeys
    
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

Private Sub mnubrowselog_Click()                                                                        '//\\LOGLINE//\\
    Shell "explorer " & LogPath, vbMaximizedFocus                                                       '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\
Private Sub mnucritical_Click()                                                                         '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "CriticalError.log", vbMaximizedFocus         '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\
Private Sub mnucodetracker_Click()                                                                      '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "CodeTracker.log", vbMaximizedFocus           '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\
Private Sub mnugeneral_Click()                                                                          '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "General.log", vbMaximizedFocus               '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\
Private Sub mnuin_Click()                                                                               '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "PacketIn.log", vbMaximizedFocus              '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\
Private Sub mnuout_Click()                                                                              '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "PacketOut.log", vbMaximizedFocus             '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\
Private Sub mnupacket_Click()                                                                           '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "InvalidPacketData.log", vbMaximizedFocus     '//\\LOGLINE//\\
End Sub                                                                                                 '//\\LOGLINE//\\

Private Sub mnupacketin_Click()

    'Save the file
    Save_PacketsIn
    
    'Display the file
    Shell "notepad " & LogPath & "packetsin.txt", vbMaximizedFocus

End Sub

Private Sub mnupacketout_Click()

    'Save the file
    Save_PacketsOut
    
    'Display the file
    Shell "notepad " & LogPath & "packetsout.txt", vbMaximizedFocus

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
        "51,57,60,61,64,69,70,79,81,82,83,84,85,97,98,99,101,102"
    
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
    
    'Setup Map borders
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
    'Load Data Commands
    InitDataCommands
    
    '*** Build help messages ***
    
    'Get the number of lines
    i = Val(Var_Get(ServerDataPath & "Help.ini", "INIT", "NumHelp"))
    
    'Put all the strings together so it can be sent in one string to the client
    ConBuf.Clear
    For LoopC = 1 To i
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String Trim$(Var_Get(ServerDataPath & "Help.ini", "INIT", str$(LoopC)))
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
        ConBuf.Put_String Trim$(Var_Get(ServerDataPath & "MOTD.ini", "INIT", str$(LoopC)))
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
    
    GOREsock_Initialize Me.hwnd
    
    'Change the listen settings in the Server.ini file
    LocalSocketID = GOREsock_Listen(ServerInfo(ServerID).IIP, ServerInfo(ServerID).Port)
    GOREsock_SetOption LocalSocketID, soxSO_TCP_NODELAY, True

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
    If GOREsock_Address(LocalSocketID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbNewLine & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbNewLine & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

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
            End Select
        End With
        
        'Check if the buffer ran out
        If rBuf.Get_ReadPos > BufUBound Then Exit Sub
        
    Loop

End Sub
