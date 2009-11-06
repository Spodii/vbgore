Attribute VB_Name = "General"
Option Explicit

Sub ShowFrmSetTile()
    frmSetTile.Visible = True
    frmSetTile.Show , frmMain
    SetTilesChkValue = 1
    frmMain.SetTilesPic.Picture = LoadPicture(GrhMapPath & "settile.bmp")
End Sub

Sub HideFrmSetTile()
    frmSetTile.Visible = False
    frmSetTile.Hide
    SetTilesChkValue = 0
    frmMain.SetTilesPic.Picture = LoadPicture(GrhMapPath & "settileg.bmp")
End Sub

Sub ShowFrmTile()
    frmTile.Visible = True
    frmTile.Show , frmMain
    ViewTilesChkValue = 1
    frmMain.ViewTilesPic.Picture = LoadPicture(GrhMapPath & "viewtiles.bmp")
End Sub

Sub HideFrmTile()
    frmTile.Visible = False
    frmTile.Hide
    ViewTilesChkValue = 0
    frmMain.ViewTilesPic.Picture = LoadPicture(GrhMapPath & "viewtilesg.bmp")
End Sub

Sub ShowFrmNPCs()
    frmNPCs.Visible = True
    frmNPCs.Show , frmMain
    ShowNPCsChkValue = 1
    frmMain.ShowNPCsPic.Picture = LoadPicture(GrhMapPath & "npc.bmp")
End Sub

Sub HideFrmNPCs()
    frmNPCs.Visible = False
    frmNPCs.Hide
    ShowNPCsChkValue = 0
    frmMain.ShowNPCsPic.Picture = LoadPicture(GrhMapPath & "npcg.bmp")
End Sub

Sub ShowFrmMapInfo()
    frmMapInfo.Visible = True
    frmMapInfo.Show , frmMain
    ShowMapInfoChkValue = 1
    frmMain.ShowMapInfoPic.Picture = LoadPicture(GrhMapPath & "mapinfo.bmp")
End Sub

Sub HideFrmMapInfo()
    frmMapInfo.Visible = False
    frmMapInfo.Hide
    ShowMapInfoChkValue = 0
    frmMain.ShowMapInfoPic.Picture = LoadPicture(GrhMapPath & "mapinfog.bmp")
End Sub

Sub ShowFrmParticles()
    frmParticles.Visible = True
    frmParticles.Show , frmMain
    PartChkValue = 1
    frmMain.PartPic.Picture = LoadPicture(GrhMapPath & "particles.bmp")
End Sub

Sub HideFrmParticles()
    frmParticles.Visible = False
    frmParticles.Hide
    PartChkValue = 0
    frmMain.PartPic.Picture = LoadPicture(GrhMapPath & "particlesg.bmp")
End Sub

Sub ShowFrmFloods()
    frmFloods.Visible = True
    frmFloods.Show , frmMain
    FloodsChkValue = 1
    frmMain.FloodsPic.Picture = LoadPicture(GrhMapPath & "floods.bmp")
End Sub

Sub HideFrmFloods()
    frmFloods.Visible = False
    frmFloods.Hide
    FloodsChkValue = 0
    frmMain.FloodsPic.Picture = LoadPicture(GrhMapPath & "floodsg.bmp")
End Sub

Sub ShowFrmObj()
    frmObj.Visible = True
    frmObj.Show , frmMain
    ObjEditChkValue = 1
    frmMain.ObjEditPic.Picture = LoadPicture(GrhMapPath & "objects.bmp")
End Sub

Sub HideFrmObj()
    frmObj.Visible = False
    frmObj.Hide
    ObjEditChkValue = 0
    frmMain.ObjEditPic.Picture = LoadPicture(GrhMapPath & "objectsg.bmp")
End Sub

Sub ShowFrmOptimizeStart()
    frmOptimizeStart.Visible = True
    frmOptimizeStart.Show , frmMain
End Sub

Sub HideFrmOptimizeStart()
    frmOptimizeStart.Visible = False
    frmOptimizeStart.Hide
End Sub

Sub ShowFrmReport()
    frmReport.Visible = True
    frmReport.Show , frmMain
End Sub

Sub HideFrmReport()
    frmReport.Visible = False
    frmReport.Hide
End Sub

Sub ShowFrmSfx()
    frmSfx.Visible = True
    frmSfx.Show , frmMain
    SfxChkValue = 1
    frmMain.SetSfxPic.Picture = LoadPicture(GrhMapPath & "sounds.bmp")
End Sub

Sub HideFrmSfx()
    frmSfx.Visible = False
    frmSfx.Hide
    SfxChkValue = 0
    frmMain.SetSfxPic.Picture = LoadPicture(GrhMapPath & "soundsg.bmp")
End Sub

Sub ShowFrmExit()
    frmExit.Visible = True
    frmExit.Show , frmMain
    ExitsChkValue = 1
    frmMain.ExitsPic.Picture = LoadPicture(GrhMapPath & "exits.bmp")
End Sub

Sub HideFrmExit()
    frmExit.Visible = False
    frmExit.Hide
    ExitsChkValue = 0
    frmMain.ExitsPic.Picture = LoadPicture(GrhMapPath & "exitsg.bmp")
End Sub

Sub ShowFrmBlock()
    frmBlock.Visible = True
    frmBlock.Show , frmMain
    BlocksChkValue = 1
    frmMain.BlocksPic.Picture = LoadPicture(GrhMapPath & "blocks.bmp")
End Sub

Sub HideFrmBlock()
    frmBlock.Visible = False
    frmBlock.Hide
    BlocksChkValue = 0
    frmMain.BlocksPic.Picture = LoadPicture(GrhMapPath & "blocksg.bmp")
End Sub

Sub ShowFrmARGB(ByRef tTxtBox As TextBox)
    frmARGB.Visible = True
    frmARGB.Show , frmMain
    frmARGB.LongTxt.Text = tTxtBox.Text
End Sub

Sub HideFrmARGB(ByRef tTxtBox As TextBox)
    frmARGB.Visible = False
    frmARGB.Hide
End Sub

Sub ShowFrmTileSelect(ByVal stBoxIDx As Byte)
    frmTileSelect.Visible = True
    frmTileSelect.Show , frmMain
    tsDrawAll = 1
    stBoxID = stBoxIDx
    If tsStart = 0 Then tsStart = 1
    If tsTileWidth = 0 Then tsTileWidth = 32
    If tsTileHeight = 0 Then tsTileHeight = 32
    tsWidth = CLng(frmTileSelect.ScaleWidth / tsTileWidth)  'Use clng to make sure we round down
    tsHeight = CLng(frmTileSelect.ScaleHeight / tsTileHeight)
    ReDim PreviewGrhList(tsWidth * tsHeight)    'Resize our array accordingly to fit all our Grhs
    Engine_SetTileSelectionArray
End Sub

Sub HideFrmTileSelect()
    frmTileSelect.Visible = False
    frmTileSelect.Hide
End Sub

Sub ShowFrmTSOpt()
    frmTSOpt.Visible = True
    frmTSOpt.Show , frmMain
    frmTileSelect.Enabled = False
    frmTSOpt.WidthTxt.Text = Engine_Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "W")
    frmTSOpt.HeightTxt.Text = Engine_Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "H")
    frmTSOpt.StartTxt.Text = Engine_Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "S")
End Sub

Sub HideFrmTSOpt()
    frmTSOpt.Visible = False
    frmTSOpt.Hide
    frmTileSelect.Enabled = True
End Sub

Sub SetTile(ByVal tX As Byte, ByVal tY As Byte, ByVal Button As Integer, ByVal Shift As Integer)

'*****************************************************************
'Updates the marked tile with the new graphics/lights/etc
'*****************************************************************
Dim TempLng As Long
Dim TempNPC As NPC
Dim l1(1 To 4) As Byte
Dim l2(1 To 4) As Byte
Dim l3(1 To 4) As Byte
Dim i As Integer
Dim b As Byte
Dim X As Byte
Dim Y As Byte

    If tX < XMinMapSize Then Exit Sub
    If tX > XMaxMapSize Then Exit Sub
    If tY < YMinMapSize Then Exit Sub
    If tY > YMaxMapSize Then Exit Sub

    'Check to get tile information
    If frmTile.Visible = True Then
        If Button = vbRightButton Then frmTile.SetInfo tX, tY
    End If
    
    'Check to place/erase a tile
    If frmSetTile.Visible = True Then
        If Button = vbLeftButton Then
            With MapData(tX, tY)
                For i = 1 To 6
                    If frmSetTile.LayerChk(i).Value = 1 Then    'Graphic layer
                        If Val(frmSetTile.GrhTxt(i).Text) > 0 Then
                            Engine_Init_Grh .Graphic(i), Val(frmSetTile.GrhTxt(i).Text)
                        Else
                            .Graphic(i).GrhIndex = 0
                        End If
                    End If
                    If frmSetTile.LightChk(i).Value = 1 Then    'Light layer
                        .Light((i - 1) * 4 + 1) = Val(frmSetTile.LightTxt((i - 1) * 4 + 1).Text)
                        .Light((i - 1) * 4 + 2) = Val(frmSetTile.LightTxt((i - 1) * 4 + 2).Text)
                        .Light((i - 1) * 4 + 3) = Val(frmSetTile.LightTxt((i - 1) * 4 + 3).Text)
                        .Light((i - 1) * 4 + 4) = Val(frmSetTile.LightTxt((i - 1) * 4 + 4).Text)
                        SaveLightBuffer(tX, tY).Light((i - 1) * 4 + 1) = .Light((i - 1) * 4 + 1)
                        SaveLightBuffer(tX, tY).Light((i - 1) * 4 + 2) = .Light((i - 1) * 4 + 2)
                        SaveLightBuffer(tX, tY).Light((i - 1) * 4 + 3) = .Light((i - 1) * 4 + 3)
                        SaveLightBuffer(tX, tY).Light((i - 1) * 4 + 4) = .Light((i - 1) * 4 + 4)
                    End If
                    If frmSetTile.ShadowChk(i).Value = 1 Then   'Shadow layer
                        .Shadow(i) = Val(frmSetTile.ShadowTxt(i).Text)
                    End If
                Next i
            End With
        End If
    End If
    
    'Check to erase a tile
    If (Shift <> 0) Or (GetAsyncKeyState(vbKeyControl) <> 0) Then
        If frmSetTile.Visible = True Then
            If Button = vbRightButton Then
                For i = 1 To 6
                    If frmSetTile.LayerChk(i).Value = 1 Then
                        MapData(tX, tY).Graphic(i).GrhIndex = 0
                    End If
                Next i
            End If
        End If
    End If
                        
    'Check to place/erase a sound effect
    If frmSfx.Visible = True Then
        If Button = vbLeftButton Then
            MapData(tX, tY).Sfx = Val(frmSfx.SfxTxt.Text)
        End If
    End If
    
    'Check to place/erase a NPC
    If frmNPCs.Visible = True Then
        If Button = vbLeftButton Then
            If tY > 1 Then  'Dont place NPCs on tiles y = 1, since their head goes onto tile 0, then uhoh! :o
                If Not Shift Then
                    If frmNPCs.SetOpt.Value = True Then
                        If MapData(tX, tY).NPCIndex = 0 Then
                            DB_RS.Open "SELECT id,char_body,char_hair,char_head,char_heading,name,char_weapon,char_hair FROM npcs WHERE id=" & frmNPCs.NPCList.ListIndex + 1, DB_Conn, adOpenStatic, adLockOptimistic
                            Engine_Char_Make NextOpenCharIndex, DB_RS!char_body, DB_RS!char_head, DB_RS!char_heading, tX, tY, Trim$(DB_RS!Name), DB_RS!char_weapon, DB_RS!char_hair, DB_RS!id
                            DB_RS.Close
                        End If
                    End If
                    If frmNPCs.EraseOpt.Value = True Then
                        If MapData(tX, tY).NPCIndex <> 0 Then Engine_Char_Erase MapData(tX, tY).NPCIndex
                    End If
                End If
            End If
        End If
    End If
    
    'Check to place/erase an exit
    If frmExit.Visible = True Then
        If Button = vbLeftButton Then
            If Not Shift Then
                If frmExit.SetOpt.Value = True Then
                    MapData(tX, tY).TileExit.Map = Val(frmExit.MapTxt.Text)
                    MapData(tX, tY).TileExit.X = Val(frmExit.XTxt.Text)
                    MapData(tX, tY).TileExit.Y = Val(frmExit.YTxt.Text)
                End If
                If frmExit.EraseOpt.Value = True Then
                    MapData(tX, tY).TileExit.Map = 0
                    MapData(tX, tY).TileExit.X = 0
                    MapData(tX, tY).TileExit.Y = 0
                End If
            End If
        End If
    End If
    
    'Check to place a block
    If frmBlock.Visible = True Then
        If Button = vbLeftButton Then
            If Not Shift Then
                If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then  'Worthless to block in the border, waste of map space
                    b = 0   'Build the blocked value
                    If frmBlock.BlockChk(1).Value = 1 Then b = b Or 1
                    If frmBlock.BlockChk(2).Value = 1 Then b = b Or 2
                    If frmBlock.BlockChk(3).Value = 1 Then b = b Or 4
                    If frmBlock.BlockChk(4).Value = 1 Then b = b Or 8
                    MapData(tX, tY).Blocked = b
                End If
            End If
        End If
    End If
                
    'Check to place/erase an Object
    If frmObj.Visible = True Then
        If Button = vbLeftButton Then
            If Not Shift Then
                If frmObj.SetOpt.Value = True Then
                    If MapData(tX, tY).ObjInfo.ObjIndex = 0 Then
                        DB_RS.Open "SELECT grhindex FROM objects WHERE id=" & frmObj.OBJList.ListIndex + 1, DB_Conn, adOpenStatic, adLockOptimistic
                        TempLng = DB_RS!GrhIndex
                        DB_RS.Close
                        Engine_OBJ_Create TempLng, tX, tY
                        MapData(tX, tY).ObjInfo.ObjIndex = frmObj.OBJList.ListIndex + 1
                        MapData(tX, tY).ObjInfo.Amount = Val(frmObj.AmountTxt.Text)
                    End If
                End If
                If frmObj.EraseOpt.Value = True Then
                    For i = 1 To LastObj
                        If OBJList(i).Pos.X = tX Then
                            If OBJList(i).Pos.Y = tY Then
                                Engine_OBJ_Erase i
                                MapData(tX, tY).ObjInfo.ObjIndex = 0
                                MapData(tX, tY).ObjInfo.Amount = 0
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
        End If
    End If
    
    'Move a particle effect
    If frmParticles.Visible = True Then
        If Button = vbRightButton Then
            If Shift Then
                If frmParticles.ParticlesList.ListIndex + 1 >= LBound(Effect) Then
                    If frmParticles.ParticlesList.ListIndex + 1 <= UBound(Effect) Then
                        If Effect(frmParticles.ParticlesList.ListIndex + 1).Used = True Then
                            For i = 0 To Effect(frmParticles.ParticlesList.ListIndex + 1).ParticleCount
                                Effect(frmParticles.ParticlesList.ListIndex + 1).Particles(i).sngA = 0
                            Next i
                            Effect(frmParticles.ParticlesList.ListIndex + 1).X = HoverX - (ParticleOffsetX - 288)
                            Effect(frmParticles.ParticlesList.ListIndex + 1).Y = HoverY - (ParticleOffsetY - 288)
                        End If
                    End If
                End If
            End If
        End If
    End If
            
End Sub

Function NextOpenCharIndex() As Integer

'*****************************************************************
'Finds the next open CharIndex in Charlist
'*****************************************************************

Dim LoopC As Long

'Check for the first char creation

    If LastChar = 0 Then
        ReDim CharList(1 To 1)
        LastChar = 1
        NextOpenCharIndex = 1
        Exit Function
    End If

    'Loop through the character slots
    For LoopC = 1 To LastChar + 1

        'We need to create a new slot
        If LoopC > LastChar Then
            LastChar = LoopC
            NextOpenCharIndex = LoopC
            ReDim Preserve CharList(1 To LastChar)
            Exit Function
        End If

        'Re-use an old slot that is not being used
        If CharList(LoopC).Active = 0 Then
            NextOpenCharIndex = LoopC
            Exit Function
        End If

    Next LoopC

End Function

Sub Game_Map_Switch(Map As Integer)

'*****************************************************************
'Loads and switches to a new map
'*****************************************************************
Dim GetEffectNum As Byte
Dim GetX As Integer
Dim GetY As Integer
Dim GetParticleCount As Integer
Dim GetGfx As Byte
Dim GetDirection As Integer
Dim ByFlags As Long
Dim BxFlags As Byte
Dim MapNum As Byte
Dim InfNum As Byte
Dim TempInt As Integer
Dim TempNPC As NPC
Dim i As Integer
Dim Y As Byte
Dim X As Byte

    'Make sure the map exists
    If Engine_FileExist(MapPath & Map & ".map", vbNormal) = False Then
        MsgBox "Error! Map path (" & MapPath & Map & ".map) could not be found!", vbOKOnly
        Exit Sub
    End If
    
    'Clear the offset values for the particle engine
    ParticleOffsetX = 0
    ParticleOffsetY = 0
    LastOffsetX = 0
    LastOffsetY = 0
    
    'Change caption
    frmMain.MapLbl.Caption = "Map: " & Map
    
    '*** Misc ***

    'Erase characters
    For i = 1 To LastChar
        If CharList(i).Active Then Engine_Char_Erase i
    Next i

    'Erase objects
    For i = 1 To LastObj
        OBJList(i).Grh.GrhIndex = 0
    Next i
    
    'Erase map-bound particle effects
    For i = 1 To NumEffects
        If Effect(i).Used Then
            If Effect(i).BoundToMap Then Effect_Kill i
        End If
    Next i
    Effect_Kill 0, True

    'Clear out old mapinfo variables
    MapInfo.Name = ""
    
    'Open the files
    MapNum = FreeFile
    Open MapPath & Map & ".map" For Binary As #MapNum
    InfNum = FreeFile
    Open MapEXPath & Map & ".inf" For Binary As #InfNum

    '*** Map File ***
    Seek #MapNum, 1

    'Map Header
    Get #MapNum, , MapInfo.MapVersion

    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            'Clear the graphic layers
            For i = 1 To 6
                MapData(X, Y).Graphic(i).GrhIndex = 0
            Next i

            'Get tile's flags
            Get #MapNum, , ByFlags

            'Blocked
            If ByFlags And 1 Then Get #MapNum, , MapData(X, Y).Blocked Else MapData(X, Y).Blocked = 0
            
            'Graphic layers
            If ByFlags And 2 Then
                Get #MapNum, , MapData(X, Y).Graphic(1).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            End If
            If ByFlags And 4 Then
                Get #MapNum, , MapData(X, Y).Graphic(2).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            End If
            If ByFlags And 8 Then
                Get #MapNum, , MapData(X, Y).Graphic(3).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            End If
            If ByFlags And 16 Then
               Get #MapNum, , MapData(X, Y).Graphic(4).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            End If
            If ByFlags And 32 Then
                Get #MapNum, , MapData(X, Y).Graphic(5).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(5), MapData(X, Y).Graphic(5).GrhIndex
            End If
            If ByFlags And 64 Then
                Get #MapNum, , MapData(X, Y).Graphic(6).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(6), MapData(X, Y).Graphic(6).GrhIndex
            End If
            
            'Set light to default (-1) - it will be set again if it is not -1 from the code below
            For i = 1 To 24
                MapData(X, Y).Light(i) = -1
            Next i
            
            'Get lighting values
            If ByFlags And 128 Then
                For i = 1 To 4
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    Get #MapNum, , MapData(X, Y).Light(i)
                Next i
            End If
            
            'Store the lighting in the SaveLightBuffer
            For i = 1 To 24
                SaveLightBuffer(X, Y).Light(i) = MapData(X, Y).Light(i)
            Next i

            'Mailbox
            If ByFlags And 8192 Then MapData(X, Y).Mailbox = 1 Else MapData(X, Y).Mailbox = 0
            
            'Shadows
            If ByFlags And 16384 Then MapData(X, Y).Shadow(1) = 1 Else MapData(X, Y).Shadow(1) = 0
            If ByFlags And 32768 Then MapData(X, Y).Shadow(2) = 1 Else MapData(X, Y).Shadow(2) = 0
            If ByFlags And 65536 Then MapData(X, Y).Shadow(3) = 1 Else MapData(X, Y).Shadow(3) = 0
            If ByFlags And 131072 Then MapData(X, Y).Shadow(4) = 1 Else MapData(X, Y).Shadow(4) = 0
            If ByFlags And 262144 Then MapData(X, Y).Shadow(5) = 1 Else MapData(X, Y).Shadow(5) = 0
            If ByFlags And 524288 Then MapData(X, Y).Shadow(6) = 1 Else MapData(X, Y).Shadow(6) = 0
            
            'Sfx
            MapData(X, Y).Sfx = 0
            If ByFlags And 1048576 Then Get #MapNum, , MapData(X, Y).Sfx
            
            '*** Inf File ***

            'Flags
            Get #InfNum, , BxFlags
            
            'Load Tile Exit
            If BxFlags And 1 Then
                Get #InfNum, , MapData(X, Y).TileExit.Map
                Get #InfNum, , MapData(X, Y).TileExit.X
                Get #InfNum, , MapData(X, Y).TileExit.Y
            Else
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
            End If
            
            'Load NPC
            If BxFlags And 2 Then
                Get #InfNum, , TempInt

                'Set up pos and startup pos
                DB_RS.Open "SELECT id,char_body,char_hair,char_head,char_heading,name,char_weapon,char_hair FROM npcs WHERE id=" & TempInt, DB_Conn, adOpenStatic, adLockOptimistic
                Engine_Char_Make NextOpenCharIndex, DB_RS!char_body, DB_RS!char_head, DB_RS!char_heading, X, Y, Trim$(DB_RS!Name), DB_RS!char_weapon, DB_RS!char_hair, DB_RS!id
                DB_RS.Close
        
            End If
        
            'Item
            If BxFlags And 4 Then
                Get #InfNum, , MapData(X, Y).ObjInfo.ObjIndex
                Get #InfNum, , MapData(X, Y).ObjInfo.Amount
            Else
                MapData(X, Y).ObjInfo.ObjIndex = 0
                MapData(X, Y).ObjInfo.Amount = 0
            End If

        Next X
    Next Y
    
    'Get the number of effects
    Get #MapNum, , Y
    
    'Store the individual particle effect types
    If Y > 0 Then
        For X = 1 To Y
            Get #MapNum, , GetEffectNum
            Get #MapNum, , GetX
            Get #MapNum, , GetY
            Get #MapNum, , GetParticleCount
            Get #MapNum, , GetGfx
            Get #MapNum, , GetDirection
            Effect_Begin GetEffectNum, GetX, GetY, GetGfx, GetParticleCount, GetDirection
        Next X
    End If
    
    'Close the map and inf files
    Close #MapNum
    Close #InfNum

    'Get info
    MapInfo.Name = Engine_Var_Get(MapEXPath & Map & ".dat", "1", "Name")
    MapInfo.Weather = Val(Engine_Var_Get(MapEXPath & Map & ".dat", "1", "Weather"))
    MapInfo.Music = Val(Engine_Var_Get(MapEXPath & Map & ".dat", "1", "Music"))
    
    'Display info
    With frmMapInfo
        .MapNameTxt.Text = MapInfo.Name
        .VersionTxt.Text = MapInfo.MapVersion
        .WeatherTxt.Text = MapInfo.Weather
        .MusicTxt.Text = MapInfo.Music
    End With

    'Set current map
    CurMap = Map
    
    'Update effects
    UpdateEffectList

End Sub

Sub Game_SaveMapData(MapNum As Integer)

'*****************************************************************
'Saves all info of a specific map (used for live-editing)
'*****************************************************************

Dim FileNumMap As Byte
Dim FileNumInf As Byte
Dim ByFlags As Long
Dim BxFlags As Byte
Dim LoopC As Long
Dim Y As Byte
Dim X As Byte
Dim i As Integer

    'Check for bright mode
    If BrightChkValue = 1 Then
        MsgBox "Error! Can not save a map while in Bright Mode!", vbOKOnly
        Exit Sub
    End If
    
    'Change caption
    frmMain.MapLbl.Caption = "Current Map: " & MapNum

    'Erase old files if the exist
    If Engine_FileExist(MapPath & MapNum & ".map", vbNormal) Then Kill MapPath & MapNum & ".map"
    If Engine_FileExist(MapEXPath & MapNum & ".inf", vbNormal) Then Kill MapEXPath & MapNum & ".inf"

    'Make sure effects list is updated
    UpdateEffectList

    'Open .map file
    FileNumMap = FreeFile
    Open MapPath & MapNum & ".map" For Binary As #FileNumMap
    Seek #FileNumMap, 1

    'Open .inf file
    FileNumInf = FreeFile
    Open MapEXPath & MapNum & ".inf" For Binary As #FileNumInf
    Seek #FileNumInf, 1

    'map Header
    Put #FileNumMap, , MapInfo.MapVersion

    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            '#######################
            '.map file
            '#######################

            '***********************
            'Prepare flag's bytes
            '***********************
            
            'Clear our flags
            ByFlags = 0
            
            'Blocked
            If MapData(X, Y).Blocked > 0 Then ByFlags = ByFlags Or 1
            
            'Graphic layers
            If MapData(X, Y).Graphic(1).GrhIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(2).GrhIndex Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(3).GrhIndex Then ByFlags = ByFlags Or 8
            If MapData(X, Y).Graphic(4).GrhIndex Then ByFlags = ByFlags Or 16
            If MapData(X, Y).Graphic(5).GrhIndex Then ByFlags = ByFlags Or 32
            If MapData(X, Y).Graphic(6).GrhIndex Then ByFlags = ByFlags Or 64
            
            'Light 1-4 used
            For i = 1 To 4
                If MapData(X, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 128
            Next i
            'Light 5-8 used
            For i = 5 To 8
                If MapData(X, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 256
            Next i
            'Light 9-12 used
            For i = 9 To 12
                If MapData(X, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 512
            Next i
            'Light 13-16 used
            For i = 13 To 16
                If MapData(X, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 1024
            Next i
            'Light 17-20 used
            For i = 17 To 20
                If MapData(X, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 2048
            Next i
            'Light 21-24 used
            For i = 21 To 24
                If MapData(X, Y).Light(i) <> -1 Then ByFlags = ByFlags Or 4096
            Next i
            
            'Mailbox
            If MapData(X, Y).Mailbox = 1 Then ByFlags = ByFlags Or 8192

            'Shadows
            If MapData(X, Y).Shadow(1) = 1 Then ByFlags = ByFlags Or 16384
            If MapData(X, Y).Shadow(2) = 1 Then ByFlags = ByFlags Or 32768
            If MapData(X, Y).Shadow(3) = 1 Then ByFlags = ByFlags Or 65536
            If MapData(X, Y).Shadow(4) = 1 Then ByFlags = ByFlags Or 131072
            If MapData(X, Y).Shadow(5) = 1 Then ByFlags = ByFlags Or 262144
            If MapData(X, Y).Shadow(6) = 1 Then ByFlags = ByFlags Or 524288
            
            'Sfx
            If MapData(X, Y).Sfx > 0 Then ByFlags = ByFlags Or 1048576

            '**********************
            'Store data
            '**********************
            'Save the flags
            Put #FileNumMap, , ByFlags
            
            'Save blocked value
            If MapData(X, Y).Blocked > 0 Then Put #FileNumMap, , MapData(X, Y).Blocked

            'Save needed grh indexes
            For LoopC = 1 To 6
                If MapData(X, Y).Graphic(LoopC).GrhIndex > 0 Then
                    Put #FileNumMap, , MapData(X, Y).Graphic(LoopC).GrhIndex
                End If
            Next LoopC
            
            'Save needed lights
            If ByFlags And 128 Then
                For i = 1 To 4
                    Put #FileNumMap, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    Put #FileNumMap, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    Put #FileNumMap, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    Put #FileNumMap, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    Put #FileNumMap, , MapData(X, Y).Light(i)
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    Put #FileNumMap, , MapData(X, Y).Light(i)
                Next i
            End If
            
            'Save Sfx
            If ByFlags And 1048576 Then Put #FileNumMap, , MapData(X, Y).Sfx

            '#######################
            '.inf file
            '#######################
            '***********************
            'Prepare flag's bytes
            '***********************
            'Reset flags
            BxFlags = 0
            
            'Tile Exit
            If MapData(X, Y).TileExit.Map Then BxFlags = BxFlags Or 1
            
            'NPC
            If MapData(X, Y).NPCIndex Then BxFlags = BxFlags Or 2
            
            'Item
            If MapData(X, Y).ObjInfo.ObjIndex Then BxFlags = BxFlags Or 4

            '**********************
            'Store data
            '**********************
            'Save the flags
            Put #FileNumInf, , BxFlags

            'Save Tile exits
            If MapData(X, Y).TileExit.Map Then
                Put #FileNumInf, , MapData(X, Y).TileExit.Map
                Put #FileNumInf, , MapData(X, Y).TileExit.X
                Put #FileNumInf, , MapData(X, Y).TileExit.Y
            End If

            'Save NPCs
            If MapData(X, Y).NPCIndex Then
                Put #FileNumInf, , CharList(MapData(X, Y).NPCIndex).NPCNumber
            End If

            'Item
            If MapData(X, Y).ObjInfo.ObjIndex Then
                Put #FileNumInf, , MapData(X, Y).ObjInfo.ObjIndex
                Put #FileNumInf, , MapData(X, Y).ObjInfo.Amount
            End If
            
        Next X
    Next Y
    
    'Get the number of map-bound effects and store that number
    Y = 0
    For X = 1 To NumEffects
        If Effect(X).Used Then
            If Effect(X).BoundToMap = 1 Then Y = Y + 1
        End If
    Next X
    Put #FileNumMap, , Y

    'Store the individual particle effect types
    For X = 1 To NumEffects
        If Effect(X).Used Then
            If Effect(X).BoundToMap Then
                If Effect(X).EffectNum = EffectNum_Waterfall Or Effect(X).EffectNum = EffectNum_Fire Then
                    Put #FileNumMap, , Effect(X).EffectNum
                    i = Effect(X).X + ParticleOffsetX   'Store as integer instead of single to save room
                    Put #FileNumMap, , i
                    i = Effect(X).Y + ParticleOffsetY   'Store as integer instead of single to save room
                    Put #FileNumMap, , i
                    Put #FileNumMap, , Effect(X).ParticleCount
                    Put #FileNumMap, , Effect(X).Gfx
                    i = Effect(X).Direction  'Store as integer instead of single to save room
                    Put #FileNumMap, , i
                End If
            End If
        End If
    Next X

    'Close the map and inf files
    Close #FileNumMap
    Close #FileNumInf
    
    'Update the NumMaps file
    i = Engine_Var_Get(DataPath & "Map.dat", "INIT", "NumMaps")
    If MapNum > i Then
        Engine_Var_Write DataPath & "Map.dat", "INIT", "NumMaps", CStr(MapNum)
    End If

    'Write .dat file
    Engine_Var_Write MapEXPath & MapNum & ".dat", "1", "Name", MapInfo.Name
    Engine_Var_Write MapEXPath & MapNum & ".dat", "1", "Weather", Str$(MapInfo.Weather)
    Engine_Var_Write MapEXPath & MapNum & ".dat", "1", "Music", Str$(MapInfo.Music)

    'Map saved
    MsgBox "Map #" & MapNum & " (" & MapInfo.Name & ") successfully saved!", vbOKOnly

End Sub

Sub Main()

'*****************************************************************
'Main
'*****************************************************************
Dim FilePath As String
Dim i As Integer

    'Init vars
    DataPath = App.Path & "\Data\"
    Data2Path = App.Path & "\Data2\"
    MapPath = App.Path & "\Maps\"
    MapEXPath = App.Path & "\MapsEX\"
    
    'Load MySQL
    MySQL_Init
    
    'Show the screens
    frmMain.Show
    DoEvents
    
    'Load DirectX
    Engine_Init_TileEngine frmMain.ScreenPic.hwnd, 32, 32, 18, 25, 10, 0.011
    
    'Check for the first map
    If Command$ = "" Then
        If Engine_FileExist(MapPath & "1.map", vbNormal) Then Game_Map_Switch 1
    Else
        FilePath = Mid$(Command$, 2, Len(Command$) - 2) 'Retrieve the filepath from Command$ and crop off the "'s
        Game_Map_Switch Val(Right$(FilePath, Len(FilePath) - Len(MapPath)))
    End If
    
    'Set default preview tile to 1,1
    SetTile 1, 1, vbRightButton, 0

    'Main Loop
    prgRun = True
    Do While prgRun

        'Don't draw frame is window is minimized or there is no map loaded
        If frmMain.WindowState <> 1 Then
            
            'Draw tile selection screen
            If frmTileSelect.Visible = True Then
                Engine_Render_TileSelection
            Else
            
                If CurMap > 0 Then
    
                    'Show the next frame
                    Engine_ShowNextFrame
                    
                    Engine_Input_CheckKeys
    
                    'Check to unload surfaces
                    For i = 1 To NumGrhFiles
    
                        'Only update surfaces in use
                        If SurfaceTimer(i) > 0 Then
    
                            'Lower the counter
                            SurfaceTimer(i) = SurfaceTimer(i) - ElapsedTime
    
                            'Unload the surface
                            If SurfaceTimer(i) <= 0 Then
                                Set SurfaceDB(i) = Nothing
                                SurfaceTimer(i) = 0
                            End If
    
                        End If
    
                    Next i
    
                End If
                
            End If
        End If

        'Do other events
        DoEvents

        'Check if unloading
        If IsUnloading Then
            prgRun = False
            Exit Do
        End If

    Loop

    'Close Down
    Engine_Init_UnloadTileEngine
    Engine_UnloadAllForms
    End

End Sub

Function NextOpenNPC() As Integer

'*****************************************************************
'Finds the next open NPC Index in NPCList
'*****************************************************************

Dim LoopC As Long

    Do
        LoopC = LoopC + 1
        If LoopC > LastChar Then
            LoopC = LastChar + 1
            Exit Do
        End If
    Loop While CharList(LoopC).NPCNumber > 0

    NextOpenNPC = LoopC

End Function

Sub UpdateEffectList()

'*****************************************************************
'Update the map's effect list
'*****************************************************************
Dim i As Byte

    On Error GoTo ErrOut

    frmParticles.ParticlesList.Clear
    For i = 1 To NumEffects
        If Effect(i).Used = False Then
            frmParticles.ParticlesList.AddItem "Empty"
        Else
            frmParticles.ParticlesList.AddItem "ID: " & Effect(i).EffectNum & " X: " & Effect(i).X & " Y: " & Effect(i).Y & " P#: " & Effect(i).ParticleCount
        End If
    Next i
    
    Exit Sub
    
ErrOut:

    MsgBox "Error updating the active effects list!", vbOKOnly

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 11  Code: 602  Total: 613 Lines
':) CommentOnly: 103 (16.8%)  Commented: 4 (0.7%)  Empty: 108 (17.6%)  Max Logic Depth: 7
