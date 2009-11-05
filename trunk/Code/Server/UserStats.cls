VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public UserIndex As Integer     'Index of the user the class belongs to
Public LastHPPercent As Byte    'Last percentage of HP when last sent
Public LastMPPercent As Byte    'Last percentage of MP when last sent

Private Update As Byte  'States if we need to update the modstats

Private BaseStats(1 To NumStats) As Long
Private ModStats(1 To NumStats) As Long
Private UpdateBaseStats(1 To NumStats) As Byte
Private UpdateModStats(1 To NumStats) As Byte

Public Sub ForceFullUpdate()
Dim i As Long

    'Force update of every stat
    LastHPPercent = 255
    LastMPPercent = 255
    For i = 1 To NumStats
        UpdateBaseStats(i) = 1
        UpdateModStats(i) = 1
    Next i

End Sub

Public Property Get BaseStat(ByVal StatID As Byte) As Long

'Get the stat value

    BaseStat = BaseStats(StatID)

End Property

Public Property Let BaseStat(ByVal StatID As Byte, ByVal Value As Long)

'Let the stat value

    If BaseStats(StatID) = Value Then Exit Property
    BaseStats(StatID) = Value
    UpdateBaseStats(StatID) = 1
    Update = 1

    'Apply limit on HP/Mana/Stamina
    Select Case StatID
    Case SID.MinHP
        If Value > ModStat(SID.MaxHP) Then BaseStat(SID.MinHP) = ModStat(SID.MaxHP)
    Case SID.MinMAN
        If Value > ModStat(SID.MaxMAN) Then BaseStat(SID.MinMAN) = ModStat(SID.MaxMAN)
    Case SID.MinSTA
        If Value > ModStat(SID.MaxSTA) Then BaseStat(SID.MinSTA) = ModStat(SID.MaxSTA)
    End Select

End Property

Public Property Get ModStat(ByVal StatID As Byte) As Long

'Get the stat value

    ModStat = ModStats(StatID)

End Property

Public Property Let ModStat(ByVal StatID As Byte, ByVal Value As Long)

    'Check for a reference to a stat that can not be modded
    If StatID = SID.MinHP Or StatID = SID.MinMAN Or StatID = SID.MinSTA Or StatID = SID.Gold Or _
        StatID = SID.Points Or StatID = SID.EXP Or StatID = SID.ELU Or StatID = SID.ELV Then
        If DEBUG_DebugMode Then MsgBox "Invalid stat ID was referenced for ModStat. Please refer to UserStats.ModStat[PropertyLet] for this error!" & vbCrLf & "SID of the stat referenced was: " & StatID, vbOKOnly
        Exit Property
    End If
    
    'Let the stat value
    If ModStat(StatID) = Value Then Exit Property
    ModStats(StatID) = Value
    UpdateModStats(StatID) = 1

End Property

Public Sub SendUpdatedStats()

Dim i As Long
Dim j As Byte   'States if we send the packet (if stats have changed)

'Update modstats

    If Update = 1 Then
        User_UpdateModStats UserIndex
        Update = 0
    End If

    'Get the buffer ready
    ConBuf.Clear

    'Send updated user stats
    For i = 1 To NumStats

        'Update base stats
        If UpdateBaseStats(i) Then
            ConBuf.Put_Byte DataCode.User_BaseStat
            ConBuf.Put_Byte i
            ConBuf.Put_Long BaseStat(i)
            UpdateBaseStats(i) = 0
            j = 1
        End If

        'Update mod stats
        If UpdateModStats(i) Then
            ConBuf.Put_Byte DataCode.User_ModStat
            ConBuf.Put_Byte i
            ConBuf.Put_Long ModStat(i)
            UpdateModStats(i) = 0
            j = 1
        End If

    Next i

    'Send the stats
    If j Then Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer

    'Check the HP Percent
    If LastHPPercent <> CByte((BaseStats(SID.MinHP) / ModStats(SID.MaxHP)) * 100) Then
        LastHPPercent = CByte((BaseStats(SID.MinHP) / ModStats(SID.MaxHP)) * 100)
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_CharHP
        ConBuf.Put_Byte LastHPPercent
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If

    'Check the MP Percent
    If LastMPPercent <> CByte((BaseStats(SID.MinMAN) / ModStats(SID.MaxMAN)) * 100) Then
        LastMPPercent = CByte((BaseStats(SID.MinMAN) / ModStats(SID.MaxMAN)) * 100)
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_CharMP
        ConBuf.Put_Byte LastMPPercent
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:46)  Decl: 11  Code: 146  Total: 157 Lines
':) CommentOnly: 17 (10.8%)  Commented: 5 (3.2%)  Empty: 43 (27.4%)  Max Logic Depth: 3
