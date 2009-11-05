Attribute VB_Name = "FileIO"
Option Explicit

Sub Load_Mail(ByVal MailIndex As Integer, ByRef MailHandler As MailData)

Dim FileNum As Byte
Dim LengthI As Integer
Dim LengthB As Byte

'Open the file and retrieve the data

    FileNum = FreeFile
    Open MailPath & MailIndex & ".mail" For Binary As FileNum

    Get FileNum, , LengthI
    MailHandler.Message = Space$(LengthI)
    Get FileNum, , MailHandler.Message

    Get FileNum, , LengthB
    MailHandler.Subject = Space$(LengthB)
    Get FileNum, , MailHandler.Subject

    Get FileNum, , LengthB
    MailHandler.WriterName = Space$(LengthB)
    Get FileNum, , MailHandler.WriterName

    Get FileNum, , MailHandler.New
    Get FileNum, , MailHandler.Obj
    Get FileNum, , MailHandler.RecieveDate
    Close FileNum

End Sub

Sub Load_Maps()

'*****************************************************************
'Loads the MapX.X files
'*****************************************************************
Dim TempSplit() As String
Dim FileNumMap As Byte
Dim FileNumInf As Byte
Dim CharIndex As Integer
Dim NPCIndex As Integer
Dim TempInt As Integer
Dim ByFlags As Long
Dim BxFlags As Byte
Dim LoopC As Long
Dim Map As Long
Dim x As Long
Dim Y As Long
Dim i As Long

    NumMaps = Val(Var_Get(DataPath & "Map.dat", "INIT", "NumMaps"))
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo

    'Create ConnectionGroups
    ReDim ConnectionGroups(1 To NumMaps)
    For LoopC = 1 To NumMaps
        ReDim ConnectionGroups(LoopC).UserIndex(0)
    Next LoopC

    For Map = 1 To NumMaps

        'Map
        FileNumMap = FreeFile
        Open MapPath & Map & ".map" For Binary As #FileNumMap
        Seek #FileNumMap, 1

        'Inf
        FileNumInf = FreeFile
        Open MapEXPath & Map & ".inf" For Binary As #FileNumInf
        Seek #FileNumInf, 1

        'Map header
        Get #FileNumMap, , MapInfo(Map).MapVersion

        'Load arrays
        For Y = YMinMapSize To YMaxMapSize
            For x = XMinMapSize To XMaxMapSize

                'Get tile's flags
                Get #FileNumMap, , ByFlags

                'Blocked
                If ByFlags And 1 Then Get #FileNumMap, , MapData(Map, x, Y).Blocked Else MapData(Map, x, Y).Blocked = 0

                'Graphic layers
                If ByFlags And 2 Then Get #FileNumMap, , MapData(Map, x, Y).Graphic(1)
                If ByFlags And 4 Then Get #FileNumMap, , MapData(Map, x, Y).Graphic(2)
                If ByFlags And 8 Then Get #FileNumMap, , MapData(Map, x, Y).Graphic(3)
                If ByFlags And 16 Then Get #FileNumMap, , MapData(Map, x, Y).Graphic(4)
                If ByFlags And 32 Then Get #FileNumMap, , MapData(Map, x, Y).Graphic(5)
                If ByFlags And 64 Then Get #FileNumMap, , MapData(Map, x, Y).Graphic(6)

                'Set light to default (-1) - it will be set again if it is not -1 from the code below
                For i = 1 To 24
                    MapData(Map, x, Y).Light(i) = -1
                Next i

                'Get lighting values
                If ByFlags And 128 Then
                    For i = 1 To 4
                        Get #FileNumMap, , MapData(Map, x, Y).Light(i)
                    Next i
                End If
                If ByFlags And 256 Then
                    For i = 5 To 8
                        Get #FileNumMap, , MapData(Map, x, Y).Light(i)
                    Next i
                End If
                If ByFlags And 512 Then
                    For i = 9 To 12
                        Get #FileNumMap, , MapData(Map, x, Y).Light(i)
                    Next i
                End If
                If ByFlags And 1024 Then
                    For i = 13 To 16
                        Get #FileNumMap, , MapData(Map, x, Y).Light(i)
                    Next i
                End If
                If ByFlags And 2048 Then
                    For i = 17 To 20
                        Get #FileNumMap, , MapData(Map, x, Y).Light(i)
                    Next i
                End If
                If ByFlags And 4096 Then
                    For i = 21 To 24
                        Get #FileNumMap, , MapData(Map, x, Y).Light(i)
                    Next i
                End If

                'Mailbox
                If ByFlags And 8192 Then MapData(Map, x, Y).Mailbox = 1 Else MapData(Map, x, Y).Mailbox = 0

                'Shadows
                If ByFlags And 16384 Then MapData(Map, x, Y).Shadow(1) = 1 Else MapData(Map, x, Y).Shadow(1) = 0
                If ByFlags And 32768 Then MapData(Map, x, Y).Shadow(2) = 1 Else MapData(Map, x, Y).Shadow(2) = 0
                If ByFlags And 65536 Then MapData(Map, x, Y).Shadow(3) = 1 Else MapData(Map, x, Y).Shadow(3) = 0
                If ByFlags And 131072 Then MapData(Map, x, Y).Shadow(4) = 1 Else MapData(Map, x, Y).Shadow(4) = 0
                If ByFlags And 262144 Then MapData(Map, x, Y).Shadow(5) = 1 Else MapData(Map, x, Y).Shadow(5) = 0
                If ByFlags And 524288 Then MapData(Map, x, Y).Shadow(6) = 1 Else MapData(Map, x, Y).Shadow(6) = 0
                
                'Set the sfx
                If ByFlags And 1048576 Then
                    Get #FileNumMap, , MapData(Map, x, Y).Sfx
                End If
                
                '.inf file

                'Get flag's byte
                Get #FileNumInf, , BxFlags

                'Load Tile Exit
                If BxFlags And 1 Then
                    Get #FileNumInf, , MapData(Map, x, Y).TileExit.Map
                    Get #FileNumInf, , MapData(Map, x, Y).TileExit.x
                    Get #FileNumInf, , MapData(Map, x, Y).TileExit.Y
                End If

                'Load NPC
                If BxFlags And 2 Then
                    Get #FileNumInf, , TempInt

                    'Set up pos and startup pos
                    NPCIndex = Load_NPC(TempInt)
                    NPCList(NPCIndex).Pos.Map = Map
                    NPCList(NPCIndex).Pos.x = x
                    NPCList(NPCIndex).Pos.Y = Y
                    NPCList(NPCIndex).StartPos = NPCList(NPCIndex).Pos

                    'Place it on the map
                    MapData(Map, x, Y).NPCIndex = NPCIndex

                    'Give it a char index
                    CharIndex = Server_NextOpenCharIndex
                    NPCList(NPCIndex).Char.CharIndex = CharIndex
                    CharList(CharIndex).Index = NPCIndex
                    CharList(CharIndex).CharType = CharType_NPC

                    'Set alive flag
                    NPCList(NPCIndex).Flags.NPCAlive = 1

                End If

                'Item
                If BxFlags And 4 Then
                    Get #FileNumInf, , MapData(Map, x, Y).ObjInfo.ObjIndex
                    Get #FileNumInf, , MapData(Map, x, Y).ObjInfo.Amount
                End If

            Next x
        Next Y

        'Close files
        Close #FileNumMap
        Close #FileNumInf

        'Other Room Data
        MapInfo(Map).Name = Var_Get(MapEXPath & Map & ".dat", "1", "Name")
        MapInfo(Map).Weather = Val(Var_Get(MapEXPath & Map & ".dat", "1", "Weather"))
        MapInfo(Map).Music = Val(Var_Get(MapEXPath & Map & ".dat", "1", "Music"))

    Next Map

End Sub

Function Load_NPC(ByVal NPCNumber As Integer) As Integer

'*****************************************************************
'Loads a NPC and returns its index
'*****************************************************************

Dim NPCIndex As Integer
Dim FileNum As Byte

'Check for valid NPCNumber

    If NPCNumber <= 0 Then Exit Function
    If Server_FileExist(NPCsPath & NPCNumber & ".npc", vbNormal) = False Then Exit Function

    'Find next open NPCindex
    NPCIndex = NPC_NextOpen

    'Update NPC counters
    If NPCIndex > LastNPC Then
        LastNPC = NPCIndex
        If LastNPC <> 0 Then ReDim Preserve NPCList(1 To LastNPC)
    End If
    NumNPCs = NumNPCs + 1

    'Load stats from file
    FileNum = FreeFile

    Open NPCsPath & NPCNumber & ".npc" For Binary As FileNum
        Get FileNum, , NPCList(NPCIndex)
    Close FileNum

    'Set the temp mod stats
    NPCList(NPCIndex).ModStat(SID.MinHP) = NPCList(NPCIndex).BaseStat(SID.MinHP)
    NPC_UpdateModStats NPCIndex

    'Setup NPC
    NPCList(NPCIndex).Flags.NPCActive = 1

    'Save NPCNumber
    NPCList(NPCIndex).NPCNumber = NPCNumber

    'Return new NPCIndex
    Load_NPC = NPCIndex

End Function

Sub Load_OBJs()

Dim Object As Long
Dim FileNum As Byte

'Get the number of objects

    FileNum = FreeFile
    Open OBJsPath & "Count.obj" For Binary As FileNum
    Get FileNum, , NumObjDatas
    Close FileNum
    ReDim ObjData(0 To NumObjDatas) As ObjData  'Leave slot 0 open for a blank slot

    'Fill Object List
    For Object = 1 To NumObjDatas
        Open OBJsPath & Object & ".obj" For Binary As FileNum
        Get FileNum, , ObjData(Object)
        Close FileNum
    Next Object

End Sub

Public Sub Load_Quests()

Dim Quest As Long
Dim FileNum As Byte

'Get Number of Quests

    FileNum = FreeFile
    Open QuestsPath & "Count.quest" For Binary As FileNum
    Get FileNum, , NumQuests
    Close FileNum
    ReDim QuestData(1 To NumQuests) As Quest

    'Fill Object List
    For Quest = 1 To NumQuests
        Open QuestsPath & Quest & ".quest" For Binary As FileNum
        Get FileNum, , QuestData(Quest)
        Close FileNum
    Next Quest

End Sub

Sub Load_ServerIni()

'*****************************************************************
'Loads the Server.ini
'*****************************************************************
Dim TempSplit() As String

'Misc

    IdleLimit = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "IdleLimit"))

    'Res pos
    TempSplit() = Split(Var_Get(ServerDataPath & "Server.ini", "INIT", "ResPos"), "-")
    ResPos.Map = Val(TempSplit(0))
    ResPos.x = Val(TempSplit(1))
    ResPos.Y = Val(TempSplit(2))

    'Max users
    MaxUsers = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "MaxUsers"))
    ReDim UserList(1 To MaxUsers) As User

End Sub

Sub Load_User(UserChar As User, UserFile As String)

Dim FileNum As Byte
Dim i As Integer

'Load the user character

    UserChar.Password = Var_Get(UserFile & ".pass", "A", "A")
    FileNum = FreeFile
    Open UserFile For Binary As FileNum
    Get FileNum, , UserChar.ArmorEqpSlot
    Get FileNum, , UserChar.WingsEqpSlot
    Get FileNum, , UserChar.Char
    Get FileNum, , i
    UserChar.CompletedQuests = Space$(i)
    Get FileNum, , UserChar.CompletedQuests
    Get FileNum, , i
    UserChar.Desc = Space$(i)
    Get FileNum, , UserChar.Desc
    Get FileNum, , UserChar.Object
    Get FileNum, , UserChar.Pos
    Get FileNum, , UserChar.Quest()
    Get FileNum, , UserChar.Skills
    Get FileNum, , UserChar.WeaponEqpSlot
    Get FileNum, , UserChar.WeaponType
    Get FileNum, , UserChar.MailID
    Get FileNum, , UserChar.KnownSkills
    UserChar.Stats.LoadClass FileNum

    'Equipt items
    If UserChar.WeaponEqpSlot > 0 Then UserChar.WeaponEqpObjIndex = UserChar.Object(UserChar.WeaponEqpSlot).ObjIndex
    If UserChar.ArmorEqpSlot > 0 Then UserChar.ArmorEqpObjIndex = UserChar.Object(UserChar.ArmorEqpSlot).ObjIndex
    If UserChar.WingsEqpSlot > 0 Then UserChar.WingsEqpObjIndex = UserChar.Object(UserChar.WingsEqpSlot).ObjIndex

End Sub

Sub Save_Mail(ByVal MailIndex As Integer, ByRef MailData As MailData)

Dim FileNum As Byte
Dim LengthI As Integer  'Length of a string as an integer
Dim LengthB As Byte     'Length of a string as a byte

'Open the file and save the data

    FileNum = FreeFile
    Open MailPath & MailIndex & ".mail" For Binary As FileNum
    LengthI = Len(MailData.Message)
    Put FileNum, , LengthI
    Put FileNum, , MailData.Message

    LengthB = Len(MailData.Subject)
    Put FileNum, , LengthB
    Put FileNum, , MailData.Subject

    LengthB = Len(MailData.WriterName)
    Put FileNum, , LengthB
    Put FileNum, , MailData.WriterName

    Put FileNum, , MailData.New
    Put FileNum, , MailData.Obj
    Put FileNum, , MailData.RecieveDate

    Close FileNum

End Sub

Sub Save_MapData()

'*****************************************************************
'Saves the MapX.inf files (all others don't need back up)
'*****************************************************************

Dim Map As Long
Dim x As Long
Dim Y As Long
Dim ByFlags As Byte
Dim FileNum As Byte

    NumMaps = Val(Var_Get(DataPath & "Map.dat", "INIT", "NumMaps"))

    'Get the next free file slot
    FileNum = FreeFile

    For Map = 1 To NumMaps

        'Open files and save updated version

        'inf
        Open MapEXPath & Map & ".inf" For Binary As #FileNum
        Seek #FileNum, 1

        'Save arrays
        For Y = YMinMapSize To YMaxMapSize
            For x = XMinMapSize To XMaxMapSize
                '.inf file

                '#############################
                'Set up flag's byte
                '#############################
                'Reset it
                ByFlags = 0

                'Tile exits
                If MapData(Map, x, Y).TileExit.Map Then ByFlags = ByFlags Xor 1

                'NPC
                If MapData(Map, x, Y).NPCIndex Then ByFlags = ByFlags Xor 2

                'OBJs
                If MapData(Map, x, Y).ObjInfo.ObjIndex Then ByFlags = ByFlags Xor 4

                'Store flag's byte
                Put #FileNum, , ByFlags

                'Tile exit
                If MapData(Map, x, Y).TileExit.Map Then
                    Put #FileNum, , MapData(Map, x, Y).TileExit.Map
                    Put #FileNum, , MapData(Map, x, Y).TileExit.x
                    Put #FileNum, , MapData(Map, x, Y).TileExit.Y
                End If

                'Store NPC
                If MapData(Map, x, Y).NPCIndex Then
                    Put #FileNum, , NPCList(MapData(Map, x, Y).NPCIndex).NPCNumber
                End If

                'Get and make Object
                If MapData(Map, x, Y).ObjInfo.ObjIndex Then
                    Put #FileNum, , MapData(Map, x, Y).ObjInfo.ObjIndex
                    Put #FileNum, , MapData(Map, x, Y).ObjInfo.Amount
                End If
            Next x
        Next Y

        'Close files
        Close #FileNum
    Next Map

End Sub

Sub Save_User(UserChar As User, UserFile As String)

'*****************************************************************
'Saves a user's data to a .chr file
'*****************************************************************

Dim FileNum As Byte
Dim i As Integer

'Save the user character

    FileNum = FreeFile
    Var_Write UserFile & ".pass", "A", "A", UserChar.Password
    Open UserFile & ".ip" For Append Shared As FileNum
    Print #FileNum, UserChar.IP
    Close FileNum
    Open UserFile For Binary As FileNum
    Put FileNum, , UserChar.ArmorEqpSlot
    Put FileNum, , UserChar.WingsEqpSlot
    Put FileNum, , UserChar.Char
    i = CInt(Len(UserChar.CompletedQuests))
    Put FileNum, , i
    Put FileNum, , UserChar.CompletedQuests
    i = CInt(Len(UserChar.Desc))
    Put FileNum, , i
    Put FileNum, , UserChar.Desc
    Put FileNum, , UserChar.Object
    Put FileNum, , UserChar.Pos
    Put FileNum, , UserChar.Quest
    Put FileNum, , UserChar.Skills
    Put FileNum, , UserChar.WeaponEqpSlot
    Put FileNum, , UserChar.WeaponType
    Put FileNum, , UserChar.MailID
    Put FileNum, , UserChar.KnownSkills
    UserChar.Stats.SaveClass FileNum

End Sub

Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = vbNullString

    sSpaces = Space$(1000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    Var_Get = RTrim$(sSpaces)
    Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:48)  Decl: 1  Code: 656  Total: 657 Lines
':) CommentOnly: 130 (19.8%)  Commented: 6 (0.9%)  Empty: 151 (23%)  Max Logic Depth: 6
