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
Public UserIndex As Integer             'Index of the user the class belongs to
Public LastHPPercent As Byte            'Last percentage of HP when last sent

Private BaseStats(1 To NumStats) As Long
Private ModStats(1 To NumStats) As Long
Private UpdateBaseStats(1 To NumStats) As Byte
Private UpdateModStats(1 To NumStats) As Byte

Public Property Get BaseStat(ByVal StatID As Byte) As Long

'Get the stat value

    BaseStat = BaseStats(StatID)

End Property

Public Property Let BaseStat(ByVal StatID As Byte, ByVal Value As Long)

'Let the stat value

    BaseStats(StatID) = Value
    UpdateBaseStats(StatID) = 1

End Property

Public Sub LoadClass(ByVal FileIndex As Byte)

Dim i As Long

'Load the class values

    Get #FileIndex, , BaseStats()

    'Set the mod values to the base values
    For i = 1 To NumStats
        ModStats(i) = BaseStats(i)
    Next i

End Sub

Public Property Let ModStat(ByVal StatID As Byte, ByVal Value As Long)

'Let the stat value

    ModStats(StatID) = Value
    UpdateModStats(StatID) = 1

End Property

Public Property Get ModStat(ByVal StatID As Byte) As Long

'Get the stat value

    ModStat = ModStats(StatID)

End Property

Public Sub SaveClass(ByVal FileIndex As Byte)

'Save the class values

    Put #FileIndex, , BaseStats()

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:44)  Decl: 8  Code: 60  Total: 68 Lines
':) CommentOnly: 7 (10.3%)  Commented: 2 (2.9%)  Empty: 21 (30.9%)  Max Logic Depth: 2
