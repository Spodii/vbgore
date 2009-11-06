Attribute VB_Name = "General"
Option Explicit

Sub Main()
Dim I As Long

    InitFilePaths               'Load the server data path
    
    'Use the vbgore root directories if in the 3rd party tools folder
    I = InStr(1, UCase$(DataPath), UCase$("\3rd Party Tools\Database Editors"))
    If I > 0 Then DataPath = Left$(DataPath, I) & Right$(DataPath, Len(DataPath) - I - Len("\3rd Party Tools\Database Editors"))
    I = InStr(1, UCase$(ServerDataPath), UCase$("\3rd Party Tools\Database Editors"))
    If I > 0 Then ServerDataPath = Left$(ServerDataPath, I) & Right$(ServerDataPath, Len(ServerDataPath) - I - Len("\3rd Party Tools\Database Editors"))
    I = InStr(1, UCase$(GrhPath), UCase$("\3rd Party Tools\Database Editors"))
    If I > 0 Then GrhPath = Left$(GrhPath, I) & Right$(GrhPath, Len(GrhPath) - I - Len("\3rd Party Tools\Database Editors"))
    
    MySQL_Init                  'initialize sql
    Editor_LoadOBJs             'Load the objects array
    Editor_LoadNPCsCombo        'Load the NPC array
    CharList(1).Heading = 1     'Set the heading to 1 to start with
    CharList(1).HeadHeading = 1 'Same for the head heading
    ReDim ShopObjs(0)           'Set the shop array to something
    ReDim DropObjs(0)           'Set the drop array to something
    
    frmMain.Show
    Engine_Init_TileEngine frmMain.PreviewPic.hwnd, frmMain.PreviewPic.ScaleWidth, frmMain.PreviewPic.ScaleHeight, 32, 32, 1, 0.011

End Sub

Sub Editor_LoadOBJs()
'*****************************************************************
'Loads all the objects! Useing it for the drops/shops
'*****************************************************************
    Dim NumObjDatas
    'Get the number of objects (Sort by id, descending, only get 1 entry, only return id)
    DB_RS.Open "SELECT id FROM objects ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    NumObjDatas = DB_RS(0)
    DB_RS.Close
    
    'Resize the objects array
    ReDim ObjData(1 To NumObjDatas)
    'Retrieve the objects from the database
    DB_RS.Open "SELECT * FROM objects", DB_Conn, adOpenStatic, adLockOptimistic
    
    'load the objects
    Do While DB_RS.EOF = False  'Loop until we reach the end of the recordset
        With ObjData(DB_RS!id)
            .Name = Trim$(DB_RS!Name)
        End With
        DB_RS.MoveNext
    Loop
    'Close the recordset
    DB_RS.Close
End Sub

Sub Editor_LoadNPCsCombo()
'*****************************************************************
'Loads all the NPCs names and places them in the newlbl on frmMain
'*****************************************************************
    Dim NumNPCsDatas
    'clear anything in the list already
    frmMain.SelectNpcCombo.Clear
    'Retrieve the list of from the database
    DB_RS.Open "SELECT * FROM npcs", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Fill the npc list
    Do While DB_RS.EOF = False  'Loop until we reach the end of the recordset
        frmMain.SelectNpcCombo.AddItem DB_RS!id & "- " & DB_RS!Name
        DB_RS.MoveNext
    Loop
    'Close the recordset
    DB_RS.Close
End Sub

Sub Editor_SaveNPC(ByVal NpcNum As Integer)
    Dim I As Integer
    Dim here As Integer
    Dim TempStr As String
 
    With frmMain
            'If we are updating the user, then the record must be deleted, so make sure it isn't there (or else we get a duplicate key entry error)
            NpcExist NpcNum, True
        
            'Open the database with an empty table
            DB_RS.Open "SELECT * FROM npcs WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
            DB_RS.AddNew
            
            'Put the data in the recordset
            DB_RS!id = NpcNum
            DB_RS!Name = .NameTxt
            DB_RS!descr = .DescTxt
            DB_RS!ai = .AITxt
            DB_RS!chat = .ChatTxt
            DB_RS!RespawnWait = .RespTxt
            DB_RS!Attackable = .AtkChk
            DB_RS!attackgrh = .AtkGrhTxt
            DB_RS!attackrange = .AtkRngTxt
            DB_RS!attacksfx = .AtkSfxTxt
            DB_RS!projectilerotatespeed = .ProRotSpeedTxt
            DB_RS!Hostile = .HostileChk
            DB_RS!Quest = .QuestTxt
            DB_RS!give_exp = .GExpTxt
            DB_RS!give_gold = .GGoldTxt

            'Save the sprite info
            For I = 0 To .CharLbl.UBound - 1
                If .CharLbl.Item(I).Visible = True Then
                    DB_RS("char_" & .CharLbl(I).Caption) = .CharTxt(I).Text
                End If
            Next I
        
            'Save the stat info
            For I = 0 To .StatLbl.UBound
                If .StatLbl.Item(I).Visible = True Then
                    DB_RS("stat_" & .StatLbl(I).Caption) = .StatTxt(I).Text
                End If
            Next I
        
            'Save the shop info
            TempStr = ""
            For I = 0 To UBound(ShopObjs) - 1
                TempStr = TempStr & vbCrLf & ShopObjs(I).OBJIndex & " " & ShopObjs(I).Amount
            Next I
            If TempStr = "" Then TempStr = " "
            DB_RS!objs_shop = Right$(TempStr, Len(TempStr) - 1)
            
            'save the drop info
            TempStr = ""
            For I = 0 To UBound(DropObjs) - 1
            TempStr = TempStr & vbCrLf & DropObjs(I).OBJIndex & " " & DropObjs(I).Amount & " " & DropObjs(I).DropC
            Next I
            If TempStr = "" Then TempStr = " "
            DB_RS!drops = Right$(TempStr, Len(TempStr) - 1)
        
    End With
    'Update the database
    DB_RS.Update
    'Close the recordset
    DB_RS.Close
    
    MsgBox "NPC " & NpcNum & " successfully saved!", vbOKOnly
    
    frmMain.SelectNpcCombo.Clear
    Editor_LoadNPCsCombo
End Sub

Sub Editor_OpenNPC(ByVal NpcNum As Integer)
'*****************************************************************
'Loads a NPC and returns its index
'*****************************************************************
    Dim ItemSplit() As String
    Dim TempSplit() As String
    Dim I As Long
    Dim here As Integer
    
    'Check for valid NPCNumber
    If NpcNum <= 0 Then Exit Sub
    
    Npcnumber = NpcNum
    
    'Load the NPC information from the database
    DB_RS.Open "SELECT * FROM npcs WHERE id=" & NpcNum, DB_Conn, adOpenStatic, adLockOptimistic
    
    'Make sure the NPC exists
    If DB_RS.EOF Then
        Exit Sub
    End If
    
    'Loop through every field - match up the names then set the data accordingly
    With frmMain
    
        Erase ShopObjs
        Erase DropObjs
        
        For I = .StatLbl.LBound To .StatLbl.UBound
            .StatLbl(I).Visible = False
            .StatTxt(I).Visible = False
            .StatTxt(I).Text = vbNullString
            .StatLbl(I).Caption = vbNullString
        Next I
        For I = .CharLbl.LBound To .CharLbl.UBound
            .CharLbl(I).Visible = False
            .CharTxt(I).Visible = False
            .CharLbl(I).Caption = vbNullString
            .CharTxt(I).Text = vbNullString
        Next I
        .OBJDropList.Clear
        .OBJList.Clear
    
        'Load stats
            here = 0
            For I = 0 To DB_RS.Fields.Count - 1
                If InStr(1, DB_RS.Fields.Item(I).Name, "stat_", vbTextCompare) Then
                    .StatLbl.Item(here + I).Caption = Replace(DB_RS.Fields.Item(I).Name, "stat_", "") '!stat_min_atack)
                    .StatTxt.Item(here + I).Text = Val(DB_RS(I))
                    .StatLbl.Item(here + I).Visible = True
                    .StatTxt.Item(here + I).Visible = True
                Else
                    here = here - 1
                End If
            Next I
            
        'Load char
            here = 0
            For I = 0 To DB_RS.Fields.Count - 1
                If InStr(1, DB_RS.Fields.Item(I).Name, "char_", vbTextCompare) Then
                    .CharLbl.Item(here + I).Caption = Replace(DB_RS.Fields.Item(I).Name, "char_", "") '!stat_min_atack)
                    .CharTxt.Item(here + I).Text = Val(DB_RS(I))
                    .CharLbl.Item(here + I).Visible = True
                    .CharTxt.Item(here + I).Visible = True
                    Editor_SetNPCGrhs here + I
                Else
                    here = here - 1
                End If
            Next I
            
        'Load other info
        
        .NameTxt = DB_RS!Name
        .DescTxt = DB_RS!descr
        .AITxt = DB_RS!ai
        .ChatTxt = DB_RS!chat
        .RespTxt = DB_RS!RespawnWait
        .AtkChk = DB_RS!Attackable
        .AtkGrhTxt = DB_RS!attackgrh
        .AtkRngTxt = DB_RS!attackrange
        .AtkSfxTxt = DB_RS!attacksfx
        .ProRotSpeedTxt = DB_RS!projectilerotatespeed
        .HostileChk = DB_RS!Hostile
        .QuestTxt = DB_RS!Quest
        .GExpTxt = DB_RS!give_exp
        .GGoldTxt = DB_RS!give_gold
        
        'load the drops
        If Trim$(DB_RS!drops) <> "" Then
            ItemSplit = Split(DB_RS!drops, vbCrLf)
            ReDim DropObjs(UBound(ItemSplit))
                For I = 0 To UBound(ItemSplit)
                    TempSplit = Split(ItemSplit(I), " ")
                    DropObjs(I).OBJIndex = TempSplit(0)
                    DropObjs(I).Amount = TempSplit(1)
                    DropObjs(I).DropC = TempSplit(2)
                    .OBJDropList.AddItem ObjData(DropObjs(I).OBJIndex).Name & " / " & DropObjs(I).Amount & " / " & DropObjs(I).DropC, I
            Next I
        End If
        
        
        'Load the shop
        If Trim$(DB_RS!objs_shop) <> "" Then
            ItemSplit = Split(DB_RS!objs_shop, vbCrLf)
            ReDim ShopObjs(UBound(ItemSplit))
                For I = 0 To UBound(ItemSplit)
                    TempSplit = Split(ItemSplit(I), " ")
                    ShopObjs(I).OBJIndex = TempSplit(0)
                    ShopObjs(I).Amount = TempSplit(1)
                    .OBJList.AddItem ObjData(ShopObjs(I).OBJIndex).Name & " / " & ShopObjs(I).Amount, I
            Next I
        End If
    End With
    DB_RS.Close
End Sub

Function NpcExist(ByVal Npcnumber As Integer, Optional ByVal DeleteIfExists As Boolean = False) As Boolean
'*****************************************************************
'Checks the database for if a user exists by the specified name
'*****************************************************************
    'Make the query
    DB_RS.Open "SELECT * FROM npcs WHERE id=" & Npcnumber, DB_Conn, adOpenStatic, adLockOptimistic

    'If End Of File = true, then the user doesn't exist
    If DB_RS.EOF = True Then NpcExist = False Else NpcExist = True
    
    'Close the recordset
    DB_RS.Close
    
    'Delete the npc so we can update it if it exists.
    If DeleteIfExists Then
        If NpcExist Then DB_Conn.Execute "DELETE FROM npcs WHERE id=" & Npcnumber
    End If

End Function

Sub Engine_UnloadAllForms()
'*****************************************************************
'Unloads all forms
'*****************************************************************
Dim frm As Form
    Erase ShopObjs
    Erase DropObjs
    For Each frm In VB.Forms
        Unload frm
    Next
End Sub

Public Sub Server_Unload()
'*****************************************************************
'Unload the server and all the variables
'*****************************************************************
Engine_UnloadAllForms
End Sub
