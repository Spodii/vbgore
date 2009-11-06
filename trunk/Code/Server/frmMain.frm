VERSION 5.00
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
   Begin VB.Menu mnu 
      Caption         =   "menu"
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

Private Sub mnubrowselog_Click()                                                    '//\\LOGLINE//\\
    Shell "explorer " & LogPath                                                     '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\
Private Sub mnucritical_Click()                                                     '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "CriticalError.log"       '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\
Private Sub mnucodetracker_Click()                                                  '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "CodeTracker.log"         '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\
Private Sub mnugeneral_Click()                                                      '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "General.log"             '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\
Private Sub mnuin_Click()                                                           '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "PacketIn.log"            '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\
Private Sub mnuout_Click()                                                          '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "PacketOut.log"           '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\
Private Sub mnupacket_Click()                                                       '//\\LOGLINE//\\
    If DEBUG_UseLogging Then Shell "notepad " & LogPath & "InvalidPacketData.log"   '//\\LOGLINE//\\
End Sub                                                                             '//\\LOGLINE//\\

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

    Log "Call StartServer", CodeTracker '//\\LOGLINE//\\
    
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
    frmMain.Caption = "Loading configuration..."
    frmMain.Refresh
    Load_ServerIni
    
    frmMain.Caption = "Loading objects..."
    frmMain.Refresh
    Load_OBJs
    
    frmMain.Caption = "Loading quests..."
    frmMain.Refresh
    Load_Quests
    
    frmMain.Caption = "Loading maps..."
    frmMain.Refresh
    Load_Maps
    
    frmMain.Caption = "Creating npc files..."
    frmMain.Refresh
    Save_NPCs_Temp
    Load_NPC_Names
    
    '*** Listen ***
    frmMain.Caption = "Loading sockets..."
    frmMain.Refresh
    
    GOREsock_Initialize Me.hWnd
    
    'Change the 127.0.0.1 to 0.0.0.0 or your internal IP to make the server public
    LocalSoxID = GOREsock_Listen("127.0.0.1", Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "GamePort")))
    GOREsock_SetOption LocalSoxID, soxSO_TCP_NODELAY, True

    '*** Misc ***

    'Show local IP/Port
    If GOREsock_Address(LocalSoxID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbNewLine & "Make sure you use your INTERNAL IP, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig" & vbNewLine & "Finally, make sure you are NOT running another instance of the server, since two applications can not bind to the same port. If problems persist, you can try changing the port.", vbOKOnly

    'Set the starting time
    ServerStartTime = CurrentTime

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
