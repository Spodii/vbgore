VERSION 5.00
Object = "{5D0A4D01-701F-4AEF-8518-952FB5EC23FF}#1.0#0"; "vbgoresocketbinary.ocx"
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
   Begin SoxOCXBinary.SoxBinary Socket 
      Height          =   420
      Left            =   2520
      Top             =   1440
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
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
      ToolTipText     =   "Overall average KBytes/sec uploaded to the clients (40 byte TCP/IPv4 packet headers included)"
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
      Top             =   1080
      Width           =   1695
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
      Top             =   840
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

Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private Sub Form_Load()

    'Create conversion buffer
    Set ConBuf = New DataBuffer
    
    'Load MySQL variables
    MySQL_Init

    'Set the file paths
    InitFilePaths
    
    'Set the server priority
    If RunHighPriority Then
        SetThreadPriority GetCurrentThread, 2       'Reccomended you dont touch these values
        SetPriorityClass GetCurrentProcess, &H80    ' unless you know what you're doing
    End If
    
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
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim FileNum As Byte
Dim LoopC As Long
Dim S As String

    On Error Resume Next

    Log "Call Form_Unload(" & Cancel & ")", CodeTracker '//\\LOGLINE//\\

    If Socket.ShutDown = soxERROR Then 'Terminate will be True if we have ShutDown properly
        If MsgBox("ShutDown procedure has not completed!" & vbNewLine & "(Hint - Select No and Try again!)" & vbNewLine & "Execute Forced ShutDown?", vbApplicationModal + vbCritical + vbYesNo, "UNABLE TO COMPLY!") = vbNo Then
            Let Cancel = True
            Exit Sub
        Else
            Socket.UnHook  'Unfortunately for now, I can't get around doing this automatically for you :( VB crashes if you don't do this!
        End If
    Else
        Socket.UnHook  'The reason is VB closes my Mod which stores the WindowProc function used for SubClassing and VB doesn't know that! So it closes the Mod before the Control!
    End If
    
    'Stop the server loop
    ServerRunning = 0
    
    'Kill the database connection
    DB_Conn.Close

    'Save the debug packets out
    If DEBUG_RecordPacketsOut Then
        S = vbNewLine
        For LoopC = 0 To 255
            S = S & LoopC & ": " & DebugPacketsOut(LoopC) & vbNewLine
        Next LoopC
        FileNum = FreeFile
        Open App.Path & "\packetsout.txt" For Output As #FileNum
            Write #FileNum, S
        Close #FileNum
    End If
    
    'Kill the temp files
    Kill ServerTempPath & "*"
    
    'Close the log files                                                                                            '//\\LOGLINE//\\
    If LogFileNumGeneral Then Close #LogFileNumGeneral                                                              '//\\LOGLINE//\\
    If LogFileNumCodeTracker Then Close #LogFileNumCodeTracker                                                      '//\\LOGLINE//\\
    If LogFileNumPacketIn Then Close #LogFileNumPacketIn                                                            '//\\LOGLINE//\\
    If LogFileNumPacketOut Then Close #LogFileNumPacketOut                                                          '//\\LOGLINE//\\
    If LogFileNumCriticalError Then Close #LogFileNumCriticalError                                                  '//\\LOGLINE//\\
    If LogFileNumInvalidPacketData Then Close #LogFileNumInvalidPacketData                                          '//\\LOGLINE//\\

    'Deallocate all arrays to avoid memory leaks
    Erase UserList
    Erase NPCList
    Erase MapInfo
    Erase CharList
    Erase ObjData

    'Same with connection Groups
    For LoopC = 1 To NumMaps
        Erase MapUsers(LoopC).Index
    Next LoopC
    Erase MapUsers

    End

End Sub

Private Sub Socket_OnClose(inSox As Long)

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
    If UserList(UserIndex).Flags.UserLogged = 1 Then User_Close UserIndex

End Sub

Private Sub Socket_OnConnection(inSox As Long)

'*********************************************
'Accepts new user and assigns an open Index
'*********************************************

Dim Index As Integer

    Log "Call Socket_OnConnection(" & inSox & ")", CodeTracker '//\\LOGLINE//\\

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

    Log "Call Socket_OnDataArrival(" & inSox & "," & ByteArrayToStr(inData) & ")", CodeTracker '//\\LOGLINE//\\

    'Get the UserIndex
    Index = User_IndexFromSox(inSox)
    If Index = -1 Then Exit Sub

    'If it is a character disconnecting, do not check their packets since they're doodie heads
    If UserList(Index).Flags.Disconnecting Then Exit Sub
    
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

Private Sub StartServer()

'*****************************************************************
'Load up server
'*****************************************************************
Dim LoopC As Long

    Log "Call StartServer", CodeTracker '//\\LOGLINE//\\

    'Show the form
    Me.Show
    DoEvents
    
    'Make the server temp path
    MakeSureDirectoryPathExists ServerTempPath
    
    'Set up debug packets out
    If DEBUG_RecordPacketsOut Then ReDim DebugPacketsOut(0 To 255)

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
    Save_NPCs_Temp
    Load_NPC_Names
    
    '*** Listen ***
    frmMain.Caption = "Loading sockets..."
    frmMain.Refresh
    
    'Change the 127.0.0.1 to 0.0.0.0 or your internal IP to make the server public
    LocalSoxID = Socket.Listen("127.0.0.1", Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "GamePort")))
    Socket.SetOption LocalSoxID, soxSO_TCP_NODELAY, True

    '*** Misc ***

    'Show status
    Server_RefreshUserListBox

    'Show local IP/Port
    If frmMain.Socket.Address(LocalSoxID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbNewLine & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbNewLine & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

    'Set the starting time
    ServerStartTime = CurrentTime

    'Set the caption
    Me.Caption = "vbGORE v." & App.Major & "." & App.Minor & "." & App.Revision
    
    'Start the main server loop
    Server_Update

End Sub
