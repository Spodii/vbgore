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
Public LastSpeed As Long        'Value of the speed since last check (we do not stick this directly in the Let ModStat routine
                                ' because the modstat often stacks before the final result is acquired

Public Update As Byte           'States if we need to update the modstats

Private BaseStats(1 To NumStats) As Long
Private ModStats(FirstModStat To NumStats) As Long
Private UpdateBaseStats(1 To NumStats) As Byte
Private UpdateModStats(FirstModStat To NumStats) As Byte

Public Sub ForceFullUpdate()
'*****************************************************************
'Forces all of the user's stats (base and mod) to be updated on the client
'More info: http://www.vbgore.com/GameServer.UserStats.ForceFullUpdate
'*****************************************************************
Dim i As Long

    Log "Call ForceFullUpdate", CodeTracker '//\\LOGLINE//\\

    'Force update of every stat
    LastHPPercent = 255
    LastMPPercent = 255
    For i = 1 To NumStats
        UpdateBaseStats(i) = 1
        If i >= FirstModStat Then UpdateModStats(i) = 1
    Next i

End Sub

Public Property Get BaseStat(ByVal StatID As Byte) As Long

    Log "Get BaseStat(" & StatID & ")", CodeTracker '//\\LOGLINE//\\
    
    'Get the stat value
    BaseStat = BaseStats(StatID)

End Property

Public Property Let BaseStat(ByVal StatID As Byte, ByVal Value As Long)

    Log "Let BaseStat(" & StatID & "," & Value & ")", CodeTracker '//\\LOGLINE//\\

    'Let the stat value
    If BaseStats(StatID) = Value Then Exit Property
    BaseStats(StatID) = Value
    UpdateBaseStats(StatID) = 1
    
    'Force update of mod stats if the base stat is a stat that has a mod stat
    If StatID >= FirstModStat Then Update = 1
    
    'Apply limit on HP/Mana/Stamina
    If BaseStat(SID.MinHP) > ModStat(SID.MaxHP) Then BaseStat(SID.MinHP) = ModStat(SID.MaxHP)
    If BaseStat(SID.MinMAN) > ModStat(SID.MaxMAN) Then BaseStat(SID.MinMAN) = ModStat(SID.MaxMAN)
    If BaseStat(SID.MinSTA) > ModStat(SID.MaxSTA) Then BaseStat(SID.MinSTA) = ModStat(SID.MaxSTA)
    
    'Keep the HP, Mana and Stamina above or equal to 0
    If BaseStat(SID.MinHP) < 0 Then BaseStat(SID.MinHP) = 0
    If BaseStat(SID.MinMAN) < 0 Then BaseStat(SID.MinMAN) = 0
    If BaseStat(SID.MinSTA) < 0 Then BaseStat(SID.MinSTA) = 0

End Property

Public Property Get ModStat(ByVal StatID As Byte) As Long

    Log "Get ModStat(" & StatID & ")", CodeTracker '//\\LOGLINE//\\

    'Get the stat value
    ModStat = ModStats(StatID)

End Property

Public Property Let ModStat(ByVal StatID As Byte, ByVal Value As Long)

    Log "Let ModStat(" & StatID & "," & Value & ")", CodeTracker '//\\LOGLINE//\\

    'Check for a reference to a stat that can not be modded
    If StatID < FirstModStat Then
        Log "ModStat(Let): Invalid stat ID (" & StatID & ") was referenced for ModStat. This stat does not use a mod value.", CriticalError '//\\LOGLINE//\\
        Exit Property
    End If
    
    'Let the stat value
    If ModStat(StatID) = Value Then Exit Property
    ModStats(StatID) = Value
    UpdateModStats(StatID) = 1
    
    'Apply limit on HP/Mana/Stamina
    If BaseStat(SID.MinHP) > ModStat(SID.MaxHP) Then BaseStat(SID.MinHP) = ModStat(SID.MaxHP)
    If BaseStat(SID.MinMAN) > ModStat(SID.MaxMAN) Then BaseStat(SID.MinMAN) = ModStat(SID.MaxMAN)
    If BaseStat(SID.MinSTA) > ModStat(SID.MaxSTA) Then BaseStat(SID.MinSTA) = ModStat(SID.MaxSTA)
    
End Property

Public Sub SendUpdatedStats()
'*****************************************************************
'Sends the updated user stats to the client
'More info: http://www.vbgore.com/GameServer.UserStats.ForceFullUpdate
'*****************************************************************
Dim SendSize As Long
Dim i As Long

    Log "Call SendUpdatedStats", CodeTracker '//\\LOGLINE//\\

    'Update modstats
    If Update Then
        User_UpdateModStats UserIndex
        Update = 0
    End If
    
    'Update the speed
    If UpdateModStats(SID.Speed) Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_SetCharSpeed
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte ModStat(SID.Speed)
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, , PP_High
    End If
    
    'Figure out the send size
    For i = 1 To NumStats
        If UpdateBaseStats(i) Then SendSize = SendSize + 6
        If i >= FirstModStat Then
            If UpdateModStats(i) Then SendSize = SendSize + 6
        End If
    Next i

    'Don't even check if we need to send anything if SendSize = 0
    If SendSize Then

        'Allocate the memory needed
        ConBuf.PreAllocate SendSize
    
        'Send updated user stats
        For i = 1 To NumStats
            
            'Update base stats
            If UpdateBaseStats(i) Then
                ConBuf.Put_Byte DataCode.User_BaseStat
                ConBuf.Put_Byte i
                ConBuf.Put_Long BaseStat(i)
                UpdateBaseStats(i) = 0
            End If
        
            'Update mod stats
            If i >= FirstModStat Then
                If UpdateModStats(i) Then
                    ConBuf.Put_Byte DataCode.User_ModStat
                    ConBuf.Put_Byte i
                    ConBuf.Put_Long ModStat(i)
                    UpdateModStats(i) = 0
                End If
            End If
    
        Next i
    
        'Send the stats
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer, , PP_Stat
        
    End If

    'Check the HP Percent
    If ModStat(SID.MaxHP) > 0 Then  'Check against division by 0
        If LastHPPercent <> CByte((BaseStats(SID.MinHP) / ModStats(SID.MaxHP)) * 100) Then
            LastHPPercent = CByte((BaseStats(SID.MinHP) / ModStats(SID.MaxHP)) * 100)
            ConBuf.PreAllocate 4
            ConBuf.Put_Byte DataCode.Server_CharHP
            ConBuf.Put_Byte LastHPPercent
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_StatPercent
        End If
    End If
    
    'Check the MP Percent
    If ModStat(SID.MaxMAN) > 0 Then  'Check against division by 0
        If LastMPPercent <> CByte((BaseStats(SID.MinMAN) / ModStats(SID.MaxMAN)) * 100) Then
            LastMPPercent = CByte((BaseStats(SID.MinMAN) / ModStats(SID.MaxMAN)) * 100)
            ConBuf.PreAllocate 4
            ConBuf.Put_Byte DataCode.Server_CharMP
            ConBuf.Put_Byte LastMPPercent
            ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
            Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_StatPercent
        End If
    End If
    
    'Check the speed
    If LastSpeed <> ModStats(SID.Speed) Then
        LastSpeed = ModStat(SID.Speed)
        ConBuf.PreAllocate 7
        ConBuf.Put_Byte DataCode.Server_SetCharSpeed
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Long ModStats(SID.Speed)
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If
    
End Sub
