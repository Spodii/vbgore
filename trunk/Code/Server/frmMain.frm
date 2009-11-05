VERSION 5.00
Object = "{9842967E-F54F-4981-93DF-0772B2672E38}#1.0#0"; "vbgoresocketbinary.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbGORE Server"
   ClientHeight    =   6150
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   8415
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   Begin SoxOCX.Sox Sox 
      Height          =   420
      Left            =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer DataCalcTmr 
      Interval        =   1000
      Left            =   1560
      Top             =   1440
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   2175
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.Timer AutoMapTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   600
      Top             =   1440
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   1440
   End
   Begin VB.ListBox Userslst 
      Appearance      =   0  'Flat
      Height          =   3810
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   3435
   End
   Begin VB.TextBox LocalAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox PortTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txStatus 
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
      Height          =   2010
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4080
      Width           =   8175
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   720
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Server Port:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   675
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   60
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

Private Sub AutoMaptimer_Timer()

'*****************************************************************
'Send out map updates if needed
'*****************************************************************

Dim UserIndex As Long
Dim LoopX As Long
Dim LoopC As Long
Dim TempInt As Integer
Dim ChunkData As Integer
Dim MapPiece As MapBlock

    If UBound(ConnectionGroups(0).UserIndex()) > 0 Then
        For UserIndex = 1 To UBound(ConnectionGroups(0).UserIndex())

            'Clear buffer
            ConBuf.Clear

            'Send 20 tiles at a time
            For LoopX = 1 To 20

                If UserList(UserIndex).Flags.DownloadingMap Then

                    'Done sending map
                    If UserList(UserIndex).Counters.SendMapCounter.Y > YMaxMapSize Then
                        ConBuf.Put_Byte DataCode.Map_EndTransfer
                        ConBuf.Put_Integer UserList(UserIndex).Counters.SendMapCounter.Map
                        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                        UserList(UserIndex).Flags.DownloadingMap = 0
                        UserList(UserIndex).Counters.SendMapCounter.X = 0
                        UserList(UserIndex).Counters.SendMapCounter.Y = 0
                        UserList(UserIndex).Counters.SendMapCounter.Map = 0
                        TempInt = UBound(ConnectionGroups(0).UserIndex()) - 1
                        If TempInt = 0 Then
                            ReDim ConnectionGroups(0).UserIndex(0)
                            frmMain.AutoMapTimer.Enabled = False
                        Else
                            For LoopC = 1 To TempInt - 1
                                If ConnectionGroups(0).UserIndex(LoopC) = UserIndex Then Exit For
                            Next LoopC
                            For LoopC = LoopC To TempInt - 1
                                ConnectionGroups(0).UserIndex(LoopC) = ConnectionGroups(0).UserIndex(LoopC + 1)
                            Next LoopC
                            ReDim ConnectionGroups(0).UserIndex(1 To TempInt - 1)
                        End If
                        Exit For
                    Else

                        'Build the map tile into the buffer
                        Server_UpdateMapTile UserIndex, UserList(UserIndex).Counters.SendMapCounter.Map, UserList(UserIndex).Counters.SendMapCounter.X, UserList(UserIndex).Counters.SendMapCounter.Y

                        'Update which tile we're on
                        UserList(UserIndex).Counters.SendMapCounter.X = UserList(UserIndex).Counters.SendMapCounter.X + 1
                        If UserList(UserIndex).Counters.SendMapCounter.X > XMaxMapSize Then
                            UserList(UserIndex).Counters.SendMapCounter.X = XMinMapSize
                            UserList(UserIndex).Counters.SendMapCounter.Y = UserList(UserIndex).Counters.SendMapCounter.Y + 1
                        End If

                    End If
                End If

            Next LoopX

            'Send the chunk to the user
            Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

        Next UserIndex

    End If

End Sub

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
    BytesInTxt.Text = Round((DataKBIn + (DataIn / 1024)) / ((timeGetTime - ServerStartTime) * 0.001), 4)
    BytesOutTxt.Text = Round((DataKBOut + (DataOut / 1024)) / ((timeGetTime - ServerStartTime) * 0.001), 4)

    'Display statistics (Bytes)
    'BytesInTxt.Text = Round(((DataKBIn * 1024) + DataIn) / ((timeGetTime - ServerStartTime) / 1000), 4)
    'BytesOutTxt.Text = Round(((DataKBOut * 1024) + DataOut) / ((timeGetTime - ServerStartTime) / 1000), 4)

End Sub

Private Sub Form_Load()

'Create conversion buffer

    Set ConBuf = New DataBuffer

    'Initialize our encryption
    Encryption_Misc_Init

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

Dim LoopC As Long

    If Sox.ShutDown = soxERROR Then 'Terminate will be True if we have ShutDown properly
        If MsgBox("ShutDown procedure has not completed!" & vbCrLf & "(Hint - Select No and Try again!)" & vbCrLf & "Execute Forced ShutDown?", vbApplicationModal + vbCritical + vbYesNo, "UNABLE TO COMPLY!") = vbNo Then
            Let Cancel = True
        Else
            Sox.UnHook  'Unfortunately for now, I can't get around doing this automatically for you :( VB crashes if you don't do this!
        End If
    Else
        Sox.UnHook  'The reason is VB closes my Mod which stores the WindowProc function used for SubClassing and VB doesn't know that! So it closes the Mod before the Control!
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
                ConBuf.Put_Byte DataCode.Comm_UMsgbox
                ConBuf.Put_String "Sorry you have been idle to long. Disconnected."
                Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
                Server_CloseSocket UserIndex
                Exit Sub
            End If

            'Check if stats need to be recovered
            If Recover = True Then

                'Check if Health needs to be updated
                If UserList(UserIndex).Stats.ModStat(SID.MinHP) < UserList(UserIndex).Stats.ModStat(SID.MaxHP) Then
                    UserList(UserIndex).Stats.ModStat(SID.MinHP) = UserList(UserIndex).Stats.ModStat(SID.MinHP) + UserList(UserIndex).Stats.ModStat(SID.Regen) * 0.5
                    If UserList(UserIndex).Stats.ModStat(SID.MinHP) > UserList(UserIndex).Stats.ModStat(SID.MaxHP) Then UserList(UserIndex).Stats.ModStat(SID.MinHP) = UserList(UserIndex).Stats.ModStat(SID.MaxHP)
                End If

                'Check if Stamina needs to be updated
                If UserList(UserIndex).Stats.ModStat(SID.MinSTA) < UserList(UserIndex).Stats.ModStat(SID.MaxSTA) Then
                    UserList(UserIndex).Stats.ModStat(SID.MinSTA) = UserList(UserIndex).Stats.ModStat(SID.MinSTA) + UserList(UserIndex).Stats.ModStat(SID.Rest) * 0.5
                    If UserList(UserIndex).Stats.ModStat(SID.MinSTA) > UserList(UserIndex).Stats.ModStat(SID.MaxSTA) Then UserList(UserIndex).Stats.ModStat(SID.MinSTA) = UserList(UserIndex).Stats.ModStat(SID.MaxSTA)
                End If

                'Check if Mana needs to be updated
                If UserList(UserIndex).Stats.ModStat(SID.MinMAN) < UserList(UserIndex).Stats.ModStat(SID.MaxMAN) Then
                    UserList(UserIndex).Stats.ModStat(SID.MinMAN) = UserList(UserIndex).Stats.ModStat(SID.MinMAN) + UserList(UserIndex).Stats.ModStat(SID.Meditate) * 0.5
                    If UserList(UserIndex).Stats.ModStat(SID.MinMAN) > UserList(UserIndex).Stats.ModStat(SID.MaxMAN) Then UserList(UserIndex).Stats.ModStat(SID.MinMAN) = UserList(UserIndex).Stats.ModStat(SID.MaxMAN)
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

            'Update blink timer
            UserList(UserIndex).Counters.BlinkCounter = UserList(UserIndex).Counters.BlinkCounter - Elapsed
            If UserList(UserIndex).Counters.BlinkCounter <= 0 Then
                UserList(UserIndex).Counters.BlinkCounter = 3000 + Int(Rnd * 7000)
                ConBuf.Clear
                ConBuf.Put_Byte DataCode.User_Blink
                ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
                Data_Send ToPCArea, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
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

            'Send data buffer if there is anything left
            Data_SendBuffer UserIndex
            
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
                        ConBuf.Put_Byte DataCode.Comm_Talk
                        ConBuf.Put_String NPCList(NPCIndex).Name & " appears stronger."
                        ConBuf.Put_Byte DataCode.Comm_FontType_Fight
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

                    'Update blink timer
                    NPCList(NPCIndex).Counters.BlinkCounter = NPCList(NPCIndex).Counters.BlinkCounter - Elapsed
                    If NPCList(NPCIndex).Counters.BlinkCounter <= 0 Then
                        NPCList(NPCIndex).Counters.BlinkCounter = 3000 + Int(Rnd * 7000)
                        ConBuf.Clear
                        ConBuf.Put_Byte DataCode.User_Blink
                        ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                        Data_Send ToNPCArea, NPCIndex, ConBuf.Get_Buffer
                    End If

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

        'Save all maps
        Save_MapData

        'Save all user's data
        For UserIndex = 1 To LastUser
            If UserList(UserIndex).Flags.UserLogged Then Save_User UserList(UserIndex), CharPath & UCase$(UserList(UserIndex).Name) & ".chr"
        Next UserIndex

        'Reset the save counter
        LastWorldSave = timeGetTime

    End If

End Sub

Private Sub Sox_OnClose(inSox As Long)

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

Private Sub Sox_OnConnection(inSox As Long)

'*********************************************
'Accepts new user and assigns an open Index
'*********************************************

Dim Index As Integer

    Index = User_NextOpen

    'Check for max users
    If Index > MaxUsers Then Exit Sub
    UserList(Index).ConnID = inSox

End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte)

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
    
    'Display the packet
    If DEBUG_PrintPacket_In Then
        frmMain.txStatus.Text = frmMain.txStatus.Text & "DataIn: " & StrConv(inData, vbUnicode) & " " & vbCrLf
        frmMain.txStatus.SelStart = Len(frmMain.txStatus.Text)
    End If
    
    'Decrypt our packet
    Select Case EncryptionType
        Case EncryptionTypeBlowfish
            Encryption_Blowfish_DecryptByte inData, EncryptionKey
        Case EncryptionTypeCryptAPI
            Encryption_CryptAPI_DecryptByte inData, EncryptionKey
        Case EncryptionTypeDES
            Encryption_DES_DecryptByte inData, EncryptionKey
        Case EncryptionTypeGost
            Encryption_Gost_DecryptByte inData, EncryptionKey
        Case EncryptionTypeRC4
            Encryption_RC4_DecryptByte inData, EncryptionKey
        Case EncryptionTypeXOR
            Encryption_XOR_DecryptByte inData, EncryptionKey
        Case EncryptionTypeSkipjack
            Encryption_Skipjack_DecryptByte inData, EncryptionKey
        Case EncryptionTypeTEA
            Encryption_TEA_DecryptByte inData, EncryptionKey
        Case EncryptionTypeTwofish
            Encryption_Twofish_DecryptByte inData, EncryptionKey
    End Select

    'Create the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    BufUBound = UBound(inData)

    'Calculate data transfer rate
    If CalcTraffic = True Then DataIn = DataIn + BufUBound + 1

    'Loop through the data buffer until it's empty
    'If done right, we should use up exactly every byte in the buffer
    Do

        'Get the CommandID
        CommandID = rBuf.Get_Byte

        'Make the appropriate call based on the CommandID
        With DataCode

            'Reset idle counter (sloppy method, clean up later)
            If CommandID <> .Server_Ping Then UserList(Index).Counters.IdleCount = timeGetTime

            Select Case CommandID
            Case 0
                X = X + 1
                Debug.Print "---Blank Byte #" & X
            Case .Comm_Emote: Data_Comm_Emote rBuf, Index
            Case .Comm_Shout: Data_Comm_Shout rBuf, Index
            Case .Comm_Talk: Data_Comm_Talk rBuf, Index
            Case .Comm_Whisper: Data_Comm_Whisper rBuf, Index

            Case .Dev_Save_Map: Data_Dev_Save_Map Index
            Case .Dev_SetBlocked: Data_Dev_SetBlocked rBuf, Index
            Case .Dev_SetExit: Data_Dev_SetExit rBuf, Index
            Case .Dev_SetLight: Data_Dev_SetLight rBuf, Index
            Case .Dev_SetMailbox: Data_Dev_SetMailbox rBuf, Index
            Case .Dev_SetMapInfo: Data_Dev_SetMapInfo rBuf, Index
            Case .Dev_SetMode: Data_Dev_SetMode Index
            Case .Dev_SetNPC: Data_Dev_SetNPC rBuf, Index
            Case .Dev_SetObject: Data_Dev_SetObject rBuf, Index
            Case .Dev_SetSurface: Data_Dev_SetSurface rBuf, Index
            Case .Dev_SetTile: Data_Dev_SetTile rBuf, Index
            Case .Dev_UpdateTile: Data_Dev_UpdateTile Index

            Case .GM_Approach: Data_GM_Approach rBuf, Index
            Case .GM_Kick: Data_GM_Kick rBuf, Index
            Case .GM_Raise: Data_GM_Raise rBuf, Index
            Case .GM_Summon: Data_GM_Summon rBuf, Index

            Case .Map_DoneLoadingMap: Data_Map_DoneLoadingMap Index
            Case .Map_RequestUpdate: Data_Map_RequestUpdate rBuf, Index

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

                'Case Else: Exit Sub 'Something went wrong or we hit the end, either way, RUN!!!!
            End Select

        End With

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos >= BufUBound Then Exit Do

    Loop

End Sub

Private Sub Sox_OnError(inSox As Long, inError As Long, inDescription As String, inSource As String, inSnipet As String)

    With txStatus
        Let .Text = .Text & "Error: SocketID " & inSox & ": Error = " & inError & " (Description) " & inDescription & " (Source) " & inSource & " (Area) " & inSnipet & vbCrLf
        Let .SelStart = Len(.Text) 'Just makes our new message visible
    End With

End Sub

Private Sub StartServer()

'*****************************************************************
'Load up server
'*****************************************************************

Dim LoopC As Long

'Check if server is already started

    If GameTimer.Enabled = True Then Exit Sub

    '*** Init vars ***
    Me.Caption = Me.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\Data\"
    SIniPath = App.Path & "\ServerData\"
    CharPath = App.Path & "\Charfile\"

    'Setup Map borders
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)

    'Resize our AI arrays
    ReDim Nodes((XMaxMapSize - (MinXBorder * 2)) * (YMaxMapSize - (MinYBorder * 2)))
    ReDim Field(MinXBorder To MaxXBorder, MinYBorder To MaxYBorder)
    FieldSize = (Abs(MaxXBorder - MinXBorder) + 1) * (Abs(MaxYBorder - MinYBorder) + 1) 'Calculates the size of our 2d array

    'Load Data Commands
    Server_InitDataCommands

    'Calculate the max distance between a char and another in it's PC area
    Max_Server_Distance = Fix(Sqr((MinYBorder - 1) ^ 2 + (MinXBorder - 1) ^ 2))

    'Reset User connections
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
    Next LoopC

    'Set up the help lines
    HelpLine(1) = "To move, use W A S D or arrow keys."
    HelpLine(2) = "To enter map editing mode, type /devmode."
    HelpLine(3) = "To attack, use Ctrl, and get objects with Alt"

    '*** Load data ***
    Load_ServerIni
    Load_Maps
    Load_OBJs
    Load_Quests

    '*** Listen ***
    LocalSoxID = Sox.Listen(Var_Get(SIniPath & "Server.ini", "INIT", "IP"), 10200)
    Sox.SetOption LocalSoxID, soxSO_TCP_NODELAY, True

    '*** Misc ***
    'Calculate data transfer rate
    If CalcTraffic = True Then DataCalcTmr.Enabled = True

    'Show status
    Server_RefreshUserListBox

    'Show local IP/Port
    LocalAdd.Text = frmMain.Sox.Address(LocalSoxID)
    PortTxt.Text = Sox.Port(LocalSoxID)
    If frmMain.Sox.Address(LocalSoxID) = "-1" Then MsgBox "Error while creating server connection. Please make sure you are connected to the internet and supplied a valid IP" & vbCrLf & "Make sure you use your internal IP if you are on a router, which can be found by Start -> Run -> 'Cmd' (Enter) -> IPConfig", vbOKOnly

    'Initialize LastWorldSave
    LastWorldSave = timeGetTime
    RecoverTimer = timeGetTime

    'Start Game timer
    GameTimer.Enabled = True

    'Set the starting time
    ServerStartTime = timeGetTime

    'Show
    Me.Show

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:48)  Decl: 3  Code: 650  Total: 653 Lines
':) CommentOnly: 95 (14.5%)  Commented: 5 (0.8%)  Empty: 144 (22.1%)  Max Logic Depth: 8
