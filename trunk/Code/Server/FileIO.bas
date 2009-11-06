Attribute VB_Name = "FileIO"
Option Explicit

Function Load_Mail(ByVal MailIndex As Long) As MailData
Dim DataSplit() As String
Dim ObjSplit() As String
Dim ObjStr As String
Dim S As String
Dim i As Long

    'Open the database
    DB_RS.Open "SELECT * FROM mail WHERE id=" & MailIndex, DB_Conn, adOpenStatic, adLockOptimistic
    
    'Make sure we have a valid mail index
    If DB_RS.EOF = False Then
        
        'Apply the values
        Load_Mail.Subject = Trim$(DB_RS!sub)
        Load_Mail.WriterName = Trim$(DB_RS!By)
        Load_Mail.RecieveDate = DB_RS!Date
        Load_Mail.Message = Trim$(DB_RS!msg)
        Load_Mail.New = Val(DB_RS!New)
        ObjStr = Trim$(DB_RS!objs)
    
        'Check for a valid object string
        If ObjStr <> "" Then
        
            'Split the objects up from the object string
            ObjSplit = Split(ObjStr, vbCrLf)
            
            'Loop through the objects
            For i = 0 To UBound(ObjSplit)
            
                'Split up the index and amount
                DataSplit = Split(ObjSplit(i), " ")
                
                'Set the data
                Load_Mail.Obj(i + 1).ObjIndex = Val(DataSplit(0))
                Load_Mail.Obj(i + 1).Amount = Val(DataSplit(1))
                
            Next i
        
        End If
        
    End If
   
    'Close the database
    DB_RS.Close

End Function

Sub Load_Maps()

'*****************************************************************
'Loads the MapX.X files
'*****************************************************************
Dim TempSplit() As String
Dim FileNumMap As Byte
Dim FileNumInf As Byte
Dim CharIndex As Integer
Dim NPCIndex As Integer
Dim TempLng As Long
Dim TempInt As Integer
Dim ByFlags As Long
Dim BxFlags As Byte
Dim LoopC As Long
Dim Map As Long
Dim X As Long
Dim Y As Long
Dim i As Long

    frmMain.Caption = "Loading maps..."
    frmMain.Refresh

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
            For X = XMinMapSize To XMaxMapSize

                'Get tile's flags
                Get #FileNumMap, , ByFlags

                'Blocked
                If ByFlags And 1 Then Get #FileNumMap, , MapData(Map, X, Y).Blocked Else MapData(Map, X, Y).Blocked = 0

                'Graphic layers (values dont need to be stored)
                If ByFlags And 2 Then Get #FileNumMap, , TempLng
                If ByFlags And 4 Then Get #FileNumMap, , TempLng
                If ByFlags And 8 Then Get #FileNumMap, , TempLng
                If ByFlags And 16 Then Get #FileNumMap, , TempLng
                If ByFlags And 32 Then Get #FileNumMap, , TempLng
                If ByFlags And 64 Then Get #FileNumMap, , TempLng

                'Get lighting values (values dont need to be stored)
                If ByFlags And 128 Then
                    For i = 1 To 4
                        Get #FileNumMap, , TempLng
                    Next i
                End If
                If ByFlags And 256 Then
                    For i = 5 To 8
                        Get #FileNumMap, , TempLng
                    Next i
                End If
                If ByFlags And 512 Then
                    For i = 9 To 12
                        Get #FileNumMap, , TempLng
                    Next i
                End If
                If ByFlags And 1024 Then
                    For i = 13 To 16
                        Get #FileNumMap, , TempLng
                    Next i
                End If
                If ByFlags And 2048 Then
                    For i = 17 To 20
                        Get #FileNumMap, , TempLng
                    Next i
                End If
                If ByFlags And 4096 Then
                    For i = 21 To 24
                        Get #FileNumMap, , TempLng
                    Next i
                End If

                'Mailbox
                If ByFlags And 8192 Then MapData(Map, X, Y).Mailbox = 1 Else MapData(Map, X, Y).Mailbox = 0
                
                'Sfx (value doesn't need to be stored)
                If ByFlags And 1048576 Then
                    Get #FileNumMap, , TempInt
                End If
                
                '.inf file

                'Get flag's byte
                Get #FileNumInf, , BxFlags

                'Load Tile Exit
                If BxFlags And 1 Then
                    Get #FileNumInf, , MapData(Map, X, Y).TileExit.Map
                    Get #FileNumInf, , MapData(Map, X, Y).TileExit.X
                    Get #FileNumInf, , MapData(Map, X, Y).TileExit.Y
                End If

                'Load NPC
                If BxFlags And 2 Then
                    Get #FileNumInf, , TempInt

                    'Set up pos and startup pos
                    NPCIndex = Load_NPC(TempInt)
                    NPCList(NPCIndex).Pos.Map = Map
                    NPCList(NPCIndex).Pos.X = X
                    NPCList(NPCIndex).Pos.Y = Y
                    NPCList(NPCIndex).StartPos = NPCList(NPCIndex).Pos

                    'Place it on the map
                    MapData(Map, X, Y).NPCIndex = NPCIndex

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
                    Get #FileNumInf, , MapData(Map, X, Y).ObjInfo.ObjIndex
                    Get #FileNumInf, , MapData(Map, X, Y).ObjInfo.Amount
                End If

            Next X
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
Dim ShopStr As String
Dim ItemSplit() As String
Dim TempSplit() As String
Dim S As String
Dim i As Long

    'Check for valid NPCNumber
    If NPCNumber <= 0 Then Exit Function

    'Find next open NPCindex
    NPCIndex = NPC_NextOpen

    'Update NPC counters
    If NPCIndex > LastNPC Then
        LastNPC = NPCIndex
        If LastNPC <> 0 Then ReDim Preserve NPCList(1 To LastNPC)
    End If
    NumNPCs = NumNPCs + 1

    'Load the NPC information from the database
    DB_RS.Open "SELECT * FROM npcs WHERE id=" & NPCNumber, DB_Conn, adOpenStatic, adLockOptimistic
    
    'Make sure the NPC exists
    If DB_RS.EOF Then
        If DEBUG_DebugMode Then MsgBox "Error loading NPC " & NPCIndex & " with NPCNumber " & NPCNumber & " - no NPC by the number found!", vbOKOnly
        Exit Function
    End If

    'Loop through every field - match up the names then set the data accordingly
    With NPCList(NPCIndex)
        .Name = Trim$(DB_RS!Name)
        .Desc = Trim$(DB_RS!Desc)
        .Movement = Val(DB_RS!Movement)
        .RespawnWait = Val(DB_RS!RespawnWait)
        .Attackable = Val(DB_RS!Attackable)
        .Hostile = Val(DB_RS!Hostile)
        .Quest = Val(DB_RS!Quest)
        .GiveEXP = Val(DB_RS!give_exp)
        .GiveGLD = Val(DB_RS!give_gold)
        .Char.Hair = Val(DB_RS!char_hair)
        .Char.Head = Val(DB_RS!char_head)
        .Char.Body = Val(DB_RS!char_body)
        .Char.Weapon = Val(DB_RS!char_weapon)
        .Char.Wings = Val(DB_RS!char_wings)
        .Char.Heading = Val(DB_RS!char_heading)
        .Char.HeadHeading = Val(DB_RS!char_headheading)
        .BaseStat(SID.Mag) = Val(DB_RS!stat_mag)
        .BaseStat(SID.DEF) = Val(DB_RS!stat_def)
        .BaseStat(SID.MinHIT) = Val(DB_RS!stat_hit_min)
        .BaseStat(SID.MaxHIT) = Val(DB_RS!stat_hit_max)
        .BaseStat(SID.MaxHP) = Val(DB_RS!stat_hp)
        .BaseStat(SID.MaxMAN) = Val(DB_RS!stat_mp)
        .BaseStat(SID.MaxSTA) = Val(DB_RS!stat_sp)
        ShopStr = Trim$(DB_RS!objs_shop)
        
        'Create the shop list
        If ShopStr <> "" Then
            TempSplit = Split(ShopStr, vbCrLf)
            ReDim .VendItems(1 To UBound(TempSplit) + 1)
            .NumVendItems = UBound(TempSplit) + 1
            For i = 0 To UBound(TempSplit)
                ItemSplit = Split(TempSplit(i), " ")
                If UBound(ItemSplit) = 1 Then   'If ubound <> 1, we have an invalid item entry
                    .VendItems(i + 1).ObjIndex = Val(ItemSplit(0))
                    .VendItems(i + 1).Amount = Val(ItemSplit(1))
                Else
                    If DEBUG_DebugMode Then MsgBox "Invalid shop/vending item entry found in the database!" & vbCrLf & "NPC: " & NPCNumber & vbCrLf & "Slot: " & i, vbOKOnly
                End If
            Next i
        End If
        
        'Set up the NPC
        .NPCNumber = NPCNumber
        .Flags.NPCActive = 1
        .BaseStat(SID.MinHP) = .BaseStat(SID.MaxHP)
        .BaseStat(SID.MinMAN) = .BaseStat(SID.MaxMAN)
        .BaseStat(SID.MinSTA) = .BaseStat(SID.MaxSTA)
                
    End With
    
    'Close the recordset
    DB_RS.Close

    'Set the temp mod stats
    NPC_UpdateModStats NPCIndex

    'Return new NPCIndex
    Load_NPC = NPCIndex

End Function

Sub Load_OBJs()

    frmMain.Caption = "Loading objects..."
    frmMain.Refresh
    
    'Get the number of objects (Sort by id, descending, only get 1 entry, only return id)
    DB_RS.Open "SELECT id FROM objects ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    NumObjDatas = DB_RS(0)
    DB_RS.Close
    
    'Resize the objects array
    ReDim ObjData(1 To NumObjDatas)
    
    'Retrieve the objects from the database
    DB_RS.Open "SELECT * FROM objects", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Fill the object list
    Do While DB_RS.EOF = False  'Loop until we reach the end of the recordset
        With ObjData(DB_RS!id)
            .Name = Trim$(DB_RS!Name)
            .Price = Val(DB_RS!Price)
            .ObjType = Val(DB_RS!ObjType)
            .WeaponType = Val(DB_RS!WeaponType)
            .GrhIndex = Val(DB_RS!GrhIndex)
            .SpriteBody = Val(DB_RS!sprite_body)
            .SpriteWeapon = Val(DB_RS!sprite_weapon)
            .SpriteHair = Val(DB_RS!sprite_hair)
            .SpriteHead = Val(DB_RS!sprite_head)
            .SpriteWings = Val(DB_RS!sprite_wings)
            .RepHP = Val(DB_RS!replenish_hp)
            .RepMP = Val(DB_RS!replenish_mp)
            .RepSP = Val(DB_RS!replenish_sp)
            .RepHPP = Val(DB_RS!replenish_hp_percent)
            .RepMPP = Val(DB_RS!replenish_mp_percent)
            .RepSPP = Val(DB_RS!replenish_sp_percent)
            .AddStat(SID.Str) = Val(DB_RS!stat_str)
            .AddStat(SID.Agi) = Val(DB_RS!stat_agi)
            .AddStat(SID.Mag) = Val(DB_RS!stat_mag)
            .AddStat(SID.DEF) = Val(DB_RS!stat_def)
            .AddStat(SID.MinHIT) = Val(DB_RS!stat_hit_min)
            .AddStat(SID.MaxHIT) = Val(DB_RS!stat_hit_max)
            .AddStat(SID.MaxHP) = Val(DB_RS!stat_hp)
            .AddStat(SID.MaxMAN) = Val(DB_RS!stat_mp)
            .AddStat(SID.MaxSTA) = Val(DB_RS!stat_sp)
            .AddStat(SID.EXP) = Val(DB_RS!stat_exp)
            .AddStat(SID.Points) = Val(DB_RS!stat_points)
            .AddStat(SID.Gold) = Val(DB_RS!stat_gold)
        End With
        DB_RS.MoveNext
    Loop
    
    'Close the recordset
    DB_RS.Close
    
End Sub

Public Sub Load_Quests()
Dim LoopQuest As Long
Dim S As String
Dim i As Long

    frmMain.Caption = "Loading quests..."
    frmMain.Refresh
    
    'Get the number of quests (Sort by id, descending, only get 1 entry, only return id)
    DB_RS.Open "SELECT id FROM quests ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    NumQuests = DB_RS(0)
    DB_RS.Close
    
    'Resize the quests array
    ReDim QuestData(1 To NumQuests)

    'Retrieve the data from the database
    DB_RS.Open "SELECT * FROM quests", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Fill in the information
    Do While DB_RS.EOF = False  'Loop until we reach the end of the recordset
        With QuestData(DB_RS!id)
            .Name = Trim$(DB_RS!Name)
            .Redoable = Val(DB_RS!Redoable)
            .StartTxt = Trim$(DB_RS!text_start)
            .AcceptTxt = Trim$(DB_RS!text_accept)
            .IncompleteTxt = Trim$(DB_RS!text_incomplete)
            .FinishTxt = Trim$(DB_RS!text_finish)
            .AcceptReqLvl = Val(DB_RS!accept_req_level)
            .AcceptReqObj = Val(DB_RS!accept_req_obj)
            .AcceptReqObjAmount = Val(DB_RS!accept_req_objamount)
            .AcceptRewExp = Val(DB_RS!accept_reward_exp)
            .AcceptRewGold = Val(DB_RS!accept_reward_gold)
            .AcceptRewObj = Val(DB_RS!accept_reward_obj)
            .AcceptRewObjAmount = Val(DB_RS!accept_reward_objamount)
            .AcceptLearnSkill = Val(DB_RS!accept_reward_learnskill)
            .FinishReqObj = Val(DB_RS!finish_req_obj)
            .FinishReqObjAmount = Val(DB_RS!finish_req_objamount)
            .FinishReqNPC = Val(DB_RS!finish_req_killnpc)
            .FinishReqNPCAmount = Val(DB_RS!finish_req_killnpcamount)
            .FinishRewExp = Val(DB_RS!finish_reward_exp)
            .FinishRewGold = Val(DB_RS!finish_reward_gold)
            .FinishRewObj = Val(DB_RS!finish_reward_obj)
            .FinishRewObjAmount = Val(DB_RS!finish_reward_objamount)
            .FinishLearnSkill = Val(DB_RS!finish_reward_learnskill)
        End With
        DB_RS.MoveNext
    Loop
    
    'Close the recordset
    DB_RS.Close
    
End Sub

Sub Load_ServerIni()

'*****************************************************************
'Loads the Server.ini
'*****************************************************************
Dim TempSplit() As String

    frmMain.Caption = "Loading configuration..."
    frmMain.Refresh

    'Misc
    IdleLimit = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "IdleLimit"))
    LastPacket = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "LastPacket"))

    'Start pos
    TempSplit() = Split(Var_Get(ServerDataPath & "Server.ini", "INIT", "StartPos"), "-")
    StartPos.Map = Val(TempSplit(0))
    StartPos.X = Val(TempSplit(1))
    StartPos.Y = Val(TempSplit(2))

    'Res pos
    TempSplit() = Split(Var_Get(ServerDataPath & "Server.ini", "INIT", "ResPos"), "-")
    ResPos.Map = Val(TempSplit(0))
    ResPos.X = Val(TempSplit(1))
    ResPos.Y = Val(TempSplit(2))

    'Max users
    MaxUsers = Val(Var_Get(ServerDataPath & "Server.ini", "INIT", "MaxUsers"))
    ReDim UserList(1 To MaxUsers) As User

End Sub

Sub Load_User(UserChar As User, UserName As String)
Dim TempStr() As String
Dim TempStr2() As String
Dim InvStr As String
Dim MailStr As String
Dim KSStr As String
Dim CurQStr As String
Dim S As String
Dim i As Long

    'Retrieve the user from the database
    DB_RS.Open "SELECT * FROM users WHERE `name`='" & UserName & "'", DB_Conn, adOpenStatic, adLockOptimistic

    'Make sure the character exists
    If DB_RS.EOF = True Then
        DB_RS.Close
        Exit Sub
    End If
    
    'Loop through every field - match up the names then set the data accordingly
    If DB_PasswordKey <> "" Then
        UserChar.Password = Encryption_RC4_DecryptString(Trim$(DB_RS!Password), DB_PasswordKey)
    Else
        UserChar.Password = Trim$(DB_RS!Password)
    End If
    UserChar.Desc = Trim$(DB_RS!Desc)
    UserChar.Flags.GMLevel = DB_RS!gm
    InvStr = DB_RS!inventory
    MailStr = DB_RS!mail
    KSStr = DB_RS!KnownSkills
    UserChar.CompletedQuests = Trim$(DB_RS!CompletedQuests)
    CurQStr = Trim$(DB_RS!currentquest)
    UserChar.Pos.X = Val(DB_RS!pos_x)
    UserChar.Pos.Y = Val(DB_RS!pos_y)
    UserChar.Pos.Map = Val(DB_RS!pos_map)
    UserChar.Char.Hair = Val(DB_RS!char_hair)
    UserChar.Char.Head = Val(DB_RS!char_head)
    UserChar.Char.Body = Val(DB_RS!char_body)
    UserChar.Char.Weapon = Val(DB_RS!char_weapon)
    UserChar.Char.Wings = Val(DB_RS!char_wings)
    UserChar.Char.Heading = Val(DB_RS!char_heading)
    UserChar.Char.HeadHeading = Val(DB_RS!char_headheading)
    UserChar.WeaponEqpSlot = Val(DB_RS!eq_weapon)
    UserChar.ArmorEqpSlot = Val(DB_RS!eq_armor)
    UserChar.WingsEqpSlot = Val(DB_RS!eq_wings)
    UserChar.Stats.BaseStat(SID.Str) = Val(DB_RS!stat_str)
    UserChar.Stats.BaseStat(SID.Agi) = Val(DB_RS!stat_agi)
    UserChar.Stats.BaseStat(SID.Mag) = Val(DB_RS!stat_mag)
    UserChar.Stats.BaseStat(SID.DEF) = Val(DB_RS!stat_def)
    UserChar.Stats.BaseStat(SID.Gold) = Val(DB_RS!stat_gold)
    UserChar.Stats.BaseStat(SID.EXP) = Val(DB_RS!stat_exp)
    UserChar.Stats.BaseStat(SID.ELV) = Val(DB_RS!stat_elv)
    UserChar.Stats.BaseStat(SID.ELU) = Val(DB_RS!stat_elu)
    UserChar.Stats.BaseStat(SID.Points) = Val(DB_RS!stat_points)
    UserChar.Stats.BaseStat(SID.MinHIT) = Val(DB_RS!stat_hit_min)
    UserChar.Stats.BaseStat(SID.MaxHIT) = Val(DB_RS!stat_hit_max)
    UserChar.Stats.BaseStat(SID.MaxHP) = Val(DB_RS!stat_hp_max) 'Max HP/SP/MP MUST be loaded before the mins!
    UserChar.Stats.BaseStat(SID.MaxMAN) = Val(DB_RS!stat_mp_max)
    UserChar.Stats.BaseStat(SID.MaxSTA) = Val(DB_RS!stat_sp_max)
    UserChar.Stats.ModStat(SID.MaxHP) = UserChar.Stats.BaseStat(SID.MaxHP)
    UserChar.Stats.ModStat(SID.MaxMAN) = UserChar.Stats.BaseStat(SID.MaxMAN)
    UserChar.Stats.ModStat(SID.MaxSTA) = UserChar.Stats.BaseStat(SID.MaxSTA)
    UserChar.Stats.BaseStat(SID.MinHP) = Val(DB_RS!stat_hp_min)
    UserChar.Stats.BaseStat(SID.MinMAN) = Val(DB_RS!stat_mp_min)
    UserChar.Stats.BaseStat(SID.MinSTA) = Val(DB_RS!stat_sp_min)
    
    'Update the user as being online
    If MySQLUpdate_Online Then
        DB_RS!online = 1
        DB_RS.Update
    End If
    
    'Close the recordset
    DB_RS.Close

    'Inventory string
    If InvStr <> "" Then
        TempStr = Split(InvStr, vbCrLf) 'Split up the inventory slots
        For i = 0 To UBound(TempStr)    'Loop through the slots
            TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex, amount and equipted (in that order)
            If Val(TempStr2(0)) <= MAX_INVENTORY_SLOTS Then
                UserChar.Object(Val(TempStr2(0))).ObjIndex = Val(TempStr2(1))
                UserChar.Object(Val(TempStr2(0))).Amount = Val(TempStr2(2))
                UserChar.Object(Val(TempStr2(0))).Equipped = Val(TempStr2(3))
            End If
        Next i
    End If
    
    'Mail string
    If MailStr <> "" Then
        TempStr = Split(MailStr, vbCrLf)    'Split up the mail indexes
        For i = 0 To UBound(TempStr)
            If i <= MaxMailPerUser Then UserChar.MailID(i + 1) = Val(TempStr(i))
        Next i
    End If
    
    'Known skills string (if the index is stored, then that skill is known - if not stored, then unknown)
    If KSStr <> "" Then
        TempStr = Split(KSStr, vbCrLf)      'Split up the known skill indexes
        For i = 0 To UBound(TempStr)
            If Val(TempStr(i)) <= NumSkills Then UserChar.KnownSkills(Val(TempStr(i))) = 1
        Next i
    End If
    
    'Current quest string
    If CurQStr <> "" Then
        TempStr = Split(CurQStr, vbCrLf)    'Split up the quests
        For i = 0 To UBound(TempStr)
            If i + 1 < MaxQuests Then 'Make sure we are within limit
                TempStr2 = Split(TempStr(i), " ")   'Split up the QuestID and NPCKills (in that order)
                UserChar.Quest(i + 1) = Val(TempStr2(0))
                UserChar.QuestStatus(i + 1).NPCKills = Val(TempStr2(1))
            End If
        Next i
    End If
    
    'Equipt items
    If UserChar.WeaponEqpSlot > 0 Then UserChar.WeaponEqpObjIndex = UserChar.Object(UserChar.WeaponEqpSlot).ObjIndex
    If UserChar.ArmorEqpSlot > 0 Then UserChar.ArmorEqpObjIndex = UserChar.Object(UserChar.ArmorEqpSlot).ObjIndex
    If UserChar.WingsEqpSlot > 0 Then UserChar.WingsEqpObjIndex = UserChar.Object(UserChar.WingsEqpSlot).ObjIndex

    'Force stat updates
    UserChar.Stats.ForceFullUpdate

    'Misc values
    UserChar.Name = UserName
    
End Sub

Sub Save_Mail(ByVal MailIndex As Long, ByRef MailData As MailData)
Dim ObjStr As String
Dim S As String
Dim i As Long

    'Build the object string
    For i = 1 To MaxMailObjs
        If MailData.Obj(i).ObjIndex > 0 Then
            If MailData.Obj(i).Amount > 0 Then
                If S <> "" Then S = S & vbCrLf  'Split the line, but make sure we dont add a split on first entry
                S = S & MailData.Obj(i).ObjIndex & " " & MailData.Obj(i).Amount
            End If
        End If
    Next i
    
    'If we are updating the mail, then the record must be deleted, so make sure it isn't there (or else we get a duplicate key entry error)
    DB_Conn.Execute "DELETE FROM mail WHERE id=" & MailIndex

    'Open the database with an empty table
    DB_RS.Open "SELECT * FROM mail WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.AddNew
    
    'Put the data in the recordset
    DB_RS!id = Str$(MailIndex)
    DB_RS!sub = MailData.Subject
    DB_RS!By = MailData.WriterName
    DB_RS!Date = MailData.RecieveDate
    DB_RS!msg = MailData.Message
    DB_RS!New = Str$(MailData.New)
    DB_RS!objs = S
    
    'Update the database with the new piece of mail
    DB_RS.Update
   
    'Close the database
    DB_RS.Close

End Sub

Sub Save_MapData()

'*****************************************************************
'Saves the MapX.inf files (all others don't need back up)
'*****************************************************************

Dim Map As Long
Dim X As Long
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
            For X = XMinMapSize To XMaxMapSize
                '.inf file

                '#############################
                'Set up flag's byte
                '#############################
                'Reset it
                ByFlags = 0

                'Tile exits
                If MapData(Map, X, Y).TileExit.Map Then ByFlags = ByFlags Xor 1

                'NPC
                If MapData(Map, X, Y).NPCIndex Then ByFlags = ByFlags Xor 2

                'OBJs
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then ByFlags = ByFlags Xor 4

                'Store flag's byte
                Put #FileNum, , ByFlags

                'Tile exit
                If MapData(Map, X, Y).TileExit.Map Then
                    Put #FileNum, , MapData(Map, X, Y).TileExit.Map
                    Put #FileNum, , MapData(Map, X, Y).TileExit.X
                    Put #FileNum, , MapData(Map, X, Y).TileExit.Y
                End If

                'Store NPC
                If MapData(Map, X, Y).NPCIndex Then
                    Put #FileNum, , NPCList(MapData(Map, X, Y).NPCIndex).NPCNumber
                End If

                'Get and make Object
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                    Put #FileNum, , MapData(Map, X, Y).ObjInfo.ObjIndex
                    Put #FileNum, , MapData(Map, X, Y).ObjInfo.Amount
                End If
            Next X
        Next Y

        'Close files
        Close #FileNum
    Next Map

End Sub

Sub Save_User(UserChar As User)

'*****************************************************************
'Saves a user's data to a .chr file
'*****************************************************************
Dim InvStr As String
Dim MailStr As String
Dim KSStr As String
Dim CurQStr As String
Dim i As Long

    With UserChar
    
        'Make sure we are trying to save a valid user by testing a few variables first
        If Len(.Name) < 3 Then Exit Sub
        If Len(.Name) > 10 Then Exit Sub
        If Len(.Password) < 3 Then Exit Sub
        If Len(.Password) > 10 Then Exit Sub
    
        'If we are updating the user, then the record must be deleted, so make sure it isn't there (or else we get a duplicate key entry error)
        Server_UserExist .Name, True
            
        'Build the inventory string
        For i = 1 To MAX_INVENTORY_SLOTS
            If .Object(i).ObjIndex > 0 Then
                If InvStr <> "" Then InvStr = InvStr & vbCrLf   'Add the line break, but dont add it to first entry
                InvStr = InvStr & i & " " & .Object(i).ObjIndex & " " & .Object(i).Amount & " " & .Object(i).Equipped
            End If
        Next i
        
        'Build mail string
        For i = 1 To MaxMailPerUser
            If .MailID(i) > 0 Then
                If MailStr <> "" Then MailStr = MailStr & vbCrLf
                MailStr = MailStr & .MailID(i)
            End If
        Next i
        
        'Build known skills string
        For i = 1 To NumSkills
            If .KnownSkills(i) > 0 Then
                If KSStr <> "" Then KSStr = KSStr & vbCrLf
                KSStr = KSStr & i
            End If
        Next i
        
        'Build current quest string
        For i = 1 To MaxQuests
            If .Quest(i) > 0 Then
                If CurQStr <> "" Then CurQStr = CurQStr & vbCrLf
                CurQStr = CurQStr & .Quest(i) & " " & .QuestStatus(i).NPCKills
            End If
        Next i
    
        'Open the database with an empty table
        DB_RS.Open "SELECT * FROM users WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
        DB_RS.AddNew
        
        'Put the data in the recordset
        If DB_PasswordKey <> "" Then
            DB_RS!Password = Encryption_RC4_EncryptString(.Password, DB_PasswordKey)
        Else
            DB_RS!Password = .Password
        End If
        DB_RS!Name = .Name
        DB_RS!gm = .Flags.GMLevel
        DB_RS!Desc = .Desc
        DB_RS!inventory = InvStr
        DB_RS!mail = MailStr
        DB_RS!KnownSkills = KSStr
        DB_RS!CompletedQuests = .CompletedQuests
        DB_RS!currentquest = CurQStr
        DB_RS!pos_x = .Pos.X
        DB_RS!pos_y = .Pos.Y
        DB_RS!pos_map = .Pos.Map
        DB_RS!char_hair = .Char.Hair
        DB_RS!char_head = .Char.Head
        DB_RS!char_body = .Char.Body
        DB_RS!char_weapon = .Char.Weapon
        DB_RS!char_wings = .Char.Wings
        DB_RS!char_heading = .Char.Heading
        DB_RS!char_headheading = .Char.HeadHeading
        DB_RS!eq_weapon = .WeaponEqpSlot
        DB_RS!eq_armor = .ArmorEqpSlot
        DB_RS!eq_wings = .WingsEqpSlot
        DB_RS!stat_str = .Stats.BaseStat(SID.Str)
        DB_RS!stat_agi = .Stats.BaseStat(SID.Agi)
        DB_RS!stat_mag = .Stats.BaseStat(SID.Mag)
        DB_RS!stat_def = .Stats.BaseStat(SID.DEF)
        DB_RS!stat_gold = .Stats.BaseStat(SID.Gold)
        DB_RS!stat_exp = .Stats.BaseStat(SID.EXP)
        DB_RS!stat_elv = .Stats.BaseStat(SID.ELV)
        DB_RS!stat_elu = .Stats.BaseStat(SID.ELU)
        DB_RS!stat_points = .Stats.BaseStat(SID.Points)
        DB_RS!stat_hit_min = .Stats.BaseStat(SID.MinHIT)
        DB_RS!stat_hit_max = .Stats.BaseStat(SID.MaxHIT)
        DB_RS!stat_hp_min = .Stats.BaseStat(SID.MinHP)
        DB_RS!stat_hp_max = .Stats.BaseStat(SID.MaxHP)
        DB_RS!stat_mp_min = .Stats.BaseStat(SID.MinMAN)
        DB_RS!stat_mp_max = .Stats.BaseStat(SID.MaxMAN)
        DB_RS!stat_sp_min = .Stats.BaseStat(SID.MinSTA)
        DB_RS!stat_sp_max = .Stats.BaseStat(SID.MaxSTA)
        DB_RS!online = 0
            
    End With
    
    'Update the database
    DB_RS.Update
    
    'Close the recordset
    DB_RS.Close

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
