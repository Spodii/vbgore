Attribute VB_Name = "General"
Option Explicit

Public Type STGG
    GrhIndex As Long
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type
Public Type STGA
    GrhIndex As Long
    Grh As Grh
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type
Public Type STG
    NumGrhs As Integer
    Grh() As STGG
End Type
Public Type STA
    NumGrhs As Integer
    Grh() As STGA
End Type
Public ShownTextureGrhs As STG
Public ShownTextureAnims As STA

Public STAWidth As Long
Public STAHeight As Long

Public PreviewGrh As Grh

Public Sub Input_Keys_ClearQueue()

'*****************************************************************
'Clears the GetAsyncKeyState queue to prevent key presses from a long time
' ago falling into "have been pressed"
'*****************************************************************
Dim i As Long

    For i = 0 To 255
        GetAsyncKeyState i
    Next i

End Sub

Sub SetInfo(ByVal s As String, Optional ByVal Critical As Byte = 0)
    
    If Critical Then
        frmMain.CritTimer.Enabled = False
        frmMain.CritTimer.Enabled = True
        frmMain.InfoLbl.Caption = s
    Else
        If Not frmMain.CritTimer.Enabled Then frmMain.InfoLbl.Caption = s
    End If

End Sub

Sub SetLayer(ByVal Layer As Byte)
Dim i As Long

    For i = 1 To 6
        If i = Layer Then
            frmSetTile.LayerPic(i).Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\layer" & i & "s.*"))
        Else
            frmSetTile.LayerPic(i).Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\layer" & i & ".*"))
        End If
    Next i
    DrawLayer = Layer
    
End Sub

Sub Server_Unload()
    'Dummy sub
End Sub

Sub ShowFrmARGB()
    frmARGB.Visible = True
    frmARGB.Show
    frmARGB.SetFocus
End Sub

Sub HideFrmARGB()
    frmARGB.Visible = False
    frmARGB.Hide
End Sub

Sub DrawPreview()
Dim i As Byte
Dim TempRect As RECT

    'Set the map set preview
    If Val(frmSetTile.GrhTxt.Text) < 1 Then
        PreviewMapGrh.GrhIndex = 0
    Else
        If PreviewMapGrh.GrhIndex <> Val(frmSetTile.GrhTxt.Text) Then
            Engine_Init_Grh PreviewMapGrh, Val(frmSetTile.GrhTxt.Text)
        End If
    End If
    
    'Set the view area
    TempRect.bottom = frmPreview.ScaleHeight
    TempRect.Right = frmPreview.ScaleWidth
    
    If Not Engine_ValidateDevice Then Exit Sub

    'Draw the preview
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
    
        'Draw the grhs
        Engine_Render_Grh PreviewGrh, 0, 0, 0, 1, True, Val(frmSetTile.LightTxt(1).Text), Val(frmSetTile.LightTxt(2).Text), Val(frmSetTile.LightTxt(3).Text), Val(frmSetTile.LightTxt(4).Text)
    
    D3DDevice.EndScene
    D3DDevice.Present TempRect, TempRect, frmPreview.hWnd, ByVal 0
    
    frmPreview.Caption = "Grh: " & frmSetTile.GrhTxt.Text

End Sub

Sub DrawTileInfoPreview()
Dim i As Byte
Dim TempRect As RECT
Dim TempRect2 As RECT

    If Val(frmTile.XLbl.Caption) < 1 Then Exit Sub
    If Val(frmTile.XLbl.Caption) > MapInfo.Width Then Exit Sub
    If Val(frmTile.YLbl.Caption) < 1 Then Exit Sub
    If Val(frmTile.YLbl.Caption) > MapInfo.Height Then Exit Sub
    
    On Error Resume Next
    
    'Set the view area
    TempRect.Top = 280
    TempRect.Left = 40
    TempRect.bottom = 256 + 280 'frmTile.GrhPic.Width
    TempRect.Right = 256 + 40 'frmTile.GrhPic.Height
    
    TempRect2.bottom = 256
    TempRect2.Right = 256
    
    If Not Engine_ValidateDevice Then Exit Sub

    'Draw the preview
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 255, 255, 255), 1#, 0
    D3DDevice.BeginScene
    
        'Draw the grhs
        Engine_Render_Grh MapData(Val(frmTile.XLbl.Caption), Val(frmTile.YLbl.Caption)).Graphic(SelectedLayer), 0, 0, 0, 1, True, Val(frmTile.LightTxt(1).Text), Val(frmTile.LightTxt(2).Text), Val(frmTile.LightTxt(3).Text), Val(frmTile.LightTxt(4).Text)
    
    D3DDevice.EndScene
    D3DDevice.Present TempRect2, TempRect, frmTile.hWnd, ByVal 0
    
    On Error GoTo 0

End Sub

Sub ShowFrmTileSelect(ByVal stBoxIDx As Byte)
    frmTileSelect.Visible = True
    frmTileSelect.Show , frmMain
    frmTileSelect.SetFocus
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
    frmTSOpt.Show , frmTileSelect
    frmTSOpt.SetFocus
    frmTileSelect.Enabled = False
    frmTSOpt.WidthTxt.Text = Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "W")
    frmTSOpt.HeightTxt.Text = Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "H")
    frmTSOpt.StartTxt.Text = Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "S")
    frmTSOpt.Show
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
Dim AB As Byte
Dim AC As Byte

    If tX < 1 Then Exit Sub
    If tX > MapInfo.Width Then Exit Sub
    If tY < 1 Then Exit Sub
    If tY > MapInfo.Height Then Exit Sub

    'Check to get tile information
    If frmTile.Visible = True Then
        If Button = vbRightButton Then frmTile.SetTileInfo tX, tY
    End If
    
    'Check to place/erase a tile
    If frmSetTile.Visible = True Then
        If Button = vbLeftButton Then
            With MapData(tX, tY)
            
                'Graphics
                If frmSetTile.LayerChk.Value = 1 Then
                    If Val(frmSetTile.GrhTxt.Text) > 0 Then
                        If .Graphic(DrawLayer).GrhIndex <> Val(frmSetTile.GrhTxt.Text) Then
                            Engine_Init_Grh .Graphic(DrawLayer), Val(frmSetTile.GrhTxt.Text)
                            AC = 1
                        End If
                    Else
                        If GetAsyncKeyState(vbKeyShift) Then
                            For i = 1 To 6
                                If .Graphic(i).GrhIndex <> 0 Then
                                    .Graphic(i).GrhIndex = 0
                                    AC = 1
                                End If
                            Next i
                        Else
                            If .Graphic(DrawLayer).GrhIndex <> 0 Then
                                .Graphic(DrawLayer).GrhIndex = 0
                                AC = 1
                            End If
                        End If
                    End If
                End If
                
                'Lights
                If frmSetTile.LightChk.Value = 1 Then
                    If .Light((DrawLayer - 1) * 4 + 1) <> Val(frmSetTile.LightTxt(1).Text) Then
                        If .Light((DrawLayer - 1) * 4 + 2) <> Val(frmSetTile.LightTxt(2).Text) Then
                            If .Light((DrawLayer - 1) * 4 + 3) <> Val(frmSetTile.LightTxt(3).Text) Then
                                If .Light((DrawLayer - 1) * 4 + 4) <> Val(frmSetTile.LightTxt(4).Text) Then
                                    .Light((DrawLayer - 1) * 4 + 1) = Val(frmSetTile.LightTxt(1).Text)
                                    .Light((DrawLayer - 1) * 4 + 2) = Val(frmSetTile.LightTxt(2).Text)
                                    .Light((DrawLayer - 1) * 4 + 3) = Val(frmSetTile.LightTxt(3).Text)
                                    .Light((DrawLayer - 1) * 4 + 4) = Val(frmSetTile.LightTxt(4).Text)
                                    SaveLightBuffer(tX, tY).Light((DrawLayer - 1) * 4 + 1) = .Light((DrawLayer - 1) * 4 + 1)
                                    SaveLightBuffer(tX, tY).Light((DrawLayer - 1) * 4 + 2) = .Light((DrawLayer - 1) * 4 + 2)
                                    SaveLightBuffer(tX, tY).Light((DrawLayer - 1) * 4 + 3) = .Light((DrawLayer - 1) * 4 + 3)
                                    SaveLightBuffer(tX, tY).Light((DrawLayer - 1) * 4 + 4) = .Light((DrawLayer - 1) * 4 + 4)
                                    'Check if in bright mode
                                    If BrightChkValue Then
                                        .Light((DrawLayer - 1) * 4 + 1) = -1
                                        .Light((DrawLayer - 1) * 4 + 2) = -1
                                        .Light((DrawLayer - 1) * 4 + 3) = -1
                                        .Light((DrawLayer - 1) * 4 + 4) = -1
                                    End If
                                    AC = 1
                                End If
                            End If
                        End If
                    End If
                End If
                
                'Shadows
                If frmSetTile.ShadowChk.Value = 1 Then
                    .Shadow(DrawLayer) = Val(frmSetTile.ShadowTxt.Text)
                End If
                
            End With
        End If
    End If
    
    'Check to erase a tile
    If (Shift <> 0) Or (GetAsyncKeyState(vbKeyControl) <> 0) Then
        If frmSetTile.Visible = True Then
            If Button = vbRightButton Then
                If frmSetTile.LayerChk.Value = 1 Then
                    If MapData(tX, tY).Graphic(DrawLayer).GrhIndex <> 0 Then
                        MapData(tX, tY).Graphic(DrawLayer).GrhIndex = 0
                        AB = 1
                    End If
                End If
            End If
        End If
    End If
                        
    'Check to place/erase a sound effect
    If frmSfx.Visible = True Then
        If Button = vbLeftButton Then
            MapData(tX, tY).Sfx = Val(frmSfx.SfxTxt.Text)
            AB = 1
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
                            AB = 1
                        End If
                    End If
                    If frmNPCs.EraseOpt.Value = True Then
                        If MapData(tX, tY).NPCIndex <> 0 Then
                            Engine_Char_Erase MapData(tX, tY).NPCIndex
                            AB = 1
                        End If
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
                    AB = 1
                End If
                If frmExit.EraseOpt.Value = True Then
                    MapData(tX, tY).TileExit.Map = 0
                    MapData(tX, tY).TileExit.X = 0
                    MapData(tX, tY).TileExit.Y = 0
                    AB = 1
                End If
            End If
        End If
    End If
    
    'Check to place a block
    If frmBlock.Visible = True Then
        If Button = vbLeftButton Then
            If Not Shift Then
                If frmBlock.SetWalkChk.Value = 1 Then
                    b = 0   'Build the blocked value
                    If frmBlock.BlockChk(1).Value = 1 Then b = b Or 1
                    If frmBlock.BlockChk(2).Value = 1 Then b = b Or 2
                    If frmBlock.BlockChk(3).Value = 1 Then b = b Or 4
                    If frmBlock.BlockChk(4).Value = 1 Then b = b Or 8
                    If MapData(tX, tY).Blocked <> b Then
                        MapData(tX, tY).Blocked = b
                        AB = 1
                    End If
                End If
                If frmBlock.SetAttackChk.Value = 1 Then
                    If MapData(tX, tY).BlockedAttack <> frmBlock.BlockAttackChk.Value Then
                        MapData(tX, tY).BlockedAttack = frmBlock.BlockAttackChk.Value
                        AB = 1
                    End If
                End If
            End If
        End If
    End If
    
    If Button = vbLeftButton Then
        If AB = 1 Then
            If ShowMiniMap Then Engine_BuildMiniMap
        End If
        If AC = 1 Or AB = 1 Then Engine_CreateTileLayers
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

Public Function LoadTextureToForm(ByVal frm As Form, ByVal TextureNum As Long, Optional ByVal Resize As Byte = 1)

'*****************************************************************
'Loads a file into a hWnd
'*****************************************************************
Dim SrcBitmapWidth As Long
Dim SrcBitmapHeight As Long

    Engine_Init_Texture TextureNum
    
    'Set the bitmap dimensions if needed
    SrcBitmapWidth = SurfaceSize(TextureNum).X + 8  'Compensates for border sizes
    SrcBitmapHeight = SurfaceSize(TextureNum).Y + 25
    
    frm.Caption = "Texture: " & TextureNum
    frm.Show
    
    CreateShownGrhs TextureNum
    
        'Size
    If Resize Then
        frm.Width = SrcBitmapWidth * Screen.TwipsPerPixelX
        frm.Height = SrcBitmapHeight * Screen.TwipsPerPixelY
    End If
    
    LoadTextureToForm = 1

End Function

Public Sub CreateShownGrhs(ByVal TextureNum As Long)
Dim i As Long
Dim j As Long
Dim k As Long
Dim XOffset As Long

    ShownTextureGrhs.NumGrhs = 0
    Erase ShownTextureGrhs.Grh

    For i = 1 To UBound(GrhData)
        If GrhData(i).FileNum = TextureNum Then
            ShownTextureGrhs.NumGrhs = ShownTextureGrhs.NumGrhs + 1
            ReDim Preserve ShownTextureGrhs.Grh(1 To ShownTextureGrhs.NumGrhs)
            With ShownTextureGrhs.Grh(ShownTextureGrhs.NumGrhs)
                .X = GrhData(i).sX
                .Y = GrhData(i).sY
                .Width = GrhData(i).pixelWidth
                .Height = GrhData(i).pixelHeight
                .GrhIndex = i
            End With
        End If
    Next i
    
    ShownTextureAnims.NumGrhs = 0
    Erase ShownTextureAnims.Grh
    XOffset = 0
    STAWidth = 0
    STAHeight = 0
    
    For i = 1 To UBound(GrhData)
        If GrhData(i).NumFrames > 1 Then
            For j = 1 To GrhData(i).NumFrames
                For k = 1 To ShownTextureGrhs.NumGrhs
                    If GrhData(i).Frames(j) = ShownTextureGrhs.Grh(k).GrhIndex Then
                        ShownTextureAnims.NumGrhs = ShownTextureAnims.NumGrhs + 1
                        ReDim Preserve ShownTextureAnims.Grh(1 To ShownTextureAnims.NumGrhs)
                        With ShownTextureAnims.Grh(ShownTextureAnims.NumGrhs)
                            .X = XOffset
                            .Width = GrhData(GrhData(i).Frames(j)).pixelWidth
                            .Height = GrhData(GrhData(i).Frames(j)).pixelHeight
                            .Y = 0
                            .GrhIndex = i
                            Engine_Init_Grh .Grh, i
                            XOffset = XOffset + .Width
                            If .Height > STAHeight Then STAHeight = .Height
                        End With
                        GoTo NextI
                    End If
                Next k
            Next j
        End If
    
NextI:
        
    Next i
    
    STAWidth = XOffset
    
    If STAWidth = 0 Then
        frmSearchAnim.Visible = False
    Else
        frmSearchAnim.Width = (STAWidth + 8) * Screen.TwipsPerPixelX
        frmSearchAnim.Height = (STAHeight + 25) * Screen.TwipsPerPixelY
        frmSearchAnim.Visible = True
        frmSearchAnim.SetFocus
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
Dim LargestTileSize As Long
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
            If Effect(i).BoundToMap = 1 Then Effect_Kill i
        End If
    Next i
    Effect_Kill 0, True

    'Clear out old mapinfo variables
    MapInfo.Name = vbNullString
    
    'Open the files
    MapNum = FreeFile
    Open MapPath & Map & ".map" For Binary As #MapNum
    InfNum = FreeFile
    Open MapEXPath & Map & ".inf" For Binary As #InfNum

    '*** Map File ***
    Seek #MapNum, 1

    'Map Header
    Get #MapNum, , MapInfo.MapVersion
    Get #MapNum, , MapInfo.Width
    Get #MapNum, , MapInfo.Height
    
    frmMapInfo.WidthTxt.Text = MapInfo.Width
    frmMapInfo.HeightTxt.Text = MapInfo.Height
    
    ReDim SaveLightBuffer(1 To MapInfo.Width, 1 To MapInfo.Height)
    
    'Setup borders
    MinXBorder = 1 + (WindowTileWidth \ 2)
    MaxXBorder = MapInfo.Width - (WindowTileWidth \ 2)
    MinYBorder = 1 + (WindowTileHeight \ 2)
    MaxYBorder = MapInfo.Height - (WindowTileHeight \ 2)

    'Resize mapdata array
    ReDim MapData(1 To MapInfo.Width, 1 To MapInfo.Height) As MapBlock
    
    'Load arrays
    For Y = 1 To MapInfo.Height
        For X = 1 To MapInfo.Width

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
                
                'Find the size of the largest tile used
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelHeight
                End If
                
            End If
            If ByFlags And 4 Then
                Get #MapNum, , MapData(X, Y).Graphic(2).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 8 Then
                Get #MapNum, , MapData(X, Y).Graphic(3).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 16 Then
               Get #MapNum, , MapData(X, Y).Graphic(4).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 32 Then
                Get #MapNum, , MapData(X, Y).Graphic(5).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(5), MapData(X, Y).Graphic(5).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 64 Then
                Get #MapNum, , MapData(X, Y).Graphic(6).GrhIndex
                Engine_Init_Grh MapData(X, Y).Graphic(6), MapData(X, Y).Graphic(6).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelHeight
                End If
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
            
            'Blocked attack tiles
            If ByFlags And 2097152 Then MapData(X, Y).BlockedAttack = 1 Else MapData(X, Y).BlockedAttack = 0
            
            'Sign
            If ByFlags And 4194304 Then Get #MapNum, , MapData(X, Y).Sign Else MapData(X, Y).Sign = 0
            
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
            Effect_Begin GetEffectNum, GetX, GetY, GetGfx, GetParticleCount, GetDirection, True
        Next X
    End If
    
    'Close the map and inf files
    Close #MapNum
    Close #InfNum

    'Get info
    MapInfo.Name = Var_Get(MapEXPath & Map & ".dat", "1", "Name")
    MapInfo.Weather = Val(Var_Get(MapEXPath & Map & ".dat", "1", "Weather"))
    MapInfo.Music = Val(Var_Get(MapEXPath & Map & ".dat", "1", "Music"))
    
    'Display info
    With frmMapInfo
        .MapNameTxt.Text = MapInfo.Name
        .VersionTxt.Text = MapInfo.MapVersion
        .WeatherTxt.Text = MapInfo.Weather
        .MusicTxt.Text = MapInfo.Music
    End With
    
    'Change caption
    frmMain.MapNameLbl.Caption = MapInfo.Name & " (" & Map & ")"

    'Set current map
    CurMap = Map
    
    'Update effects
    UpdateEffectList
    
    'Build the mini-map
    Engine_BuildMiniMap
    
    'Auto-calculate the maximum size to set the tile buffer
    LargestTileSize = LargestTileSize + (32 - (LargestTileSize Mod 32)) 'Round to the next highest factor of 32
    TileBufferSize = (LargestTileSize \ 32) 'Divide into tiles
    
    'Force to 2 to draw characters since they are 2 tiles tall
    'If you have characters or paperdoll parts > 64 pixels in width or high, you need to increase this
    If TileBufferSize < 2 Then TileBufferSize = 2
    
    Engine_CreateTileLayers

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
    
    'Check for any NPCs on blocked tiles
    For X = 1 To MapInfo.Width
        For Y = 1 To MapInfo.Height
            If MapData(X, Y).NPCIndex > 0 Then
                If MapData(X, Y).Blocked > 0 Then
                    MsgBox "Warning! You have a NPC placed on a blocked tile, which can lead to problem!" & vbNewLine & "Please correct this error if possible!", vbOKOnly
                    GoTo SkipCheck  'Only show the error once
                End If
            End If
        Next Y
    Next X

SkipCheck:

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
    Put #FileNumMap, , MapInfo.Width
    Put #FileNumMap, , MapInfo.Height

    'Write .map file
    For Y = 1 To MapInfo.Height
        For X = 1 To MapInfo.Width
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
            
            'Blocked attack tile
            If MapData(X, Y).BlockedAttack = 1 Then ByFlags = ByFlags Or 2097152
            
            'Signs
            If MapData(X, Y).Sign > 0 Then ByFlags = ByFlags Or 4194304
            
            'If there is a warp
            If MapData(X, Y).TileExit.Map > 0 Then ByFlags = ByFlags Or 8388608

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
            
            'Save sign
            If ByFlags And 4194304 Then Put #FileNumMap, , MapData(X, Y).Sign

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
            If Effect(X).BoundToMap = 1 Then
                If Effect(X).EffectNum = EffectNum_Waterfall Or Effect(X).EffectNum = EffectNum_Fire Then
                    Put #FileNumMap, , Effect(X).EffectNum
                    i = Effect(X).X + ParticleOffsetX   'Store as integer instead of single to save room
                    Put #FileNumMap, , i
                    i = Effect(X).Y + ParticleOffsetY   'Store as integer instead of single to save room
                    Put #FileNumMap, , i
                    Put #FileNumMap, , Effect(X).ParticleCount
                    Put #FileNumMap, , Effect(X).Gfx
                    i = Effect(X).Direction             'Store as integer instead of single to save room
                    Put #FileNumMap, , i
                End If
            End If
        End If
    Next X

    'Close the map and inf files
    Close #FileNumMap
    Close #FileNumInf
    
    'Update the NumMaps file
    i = Var_Get(DataPath & "Map.dat", "INIT", "NumMaps")
    If MapNum > i Then
        Var_Write DataPath & "Map.dat", "INIT", "NumMaps", CStr(MapNum)
    End If

    'Write .dat file
    Var_Write MapEXPath & MapNum & ".dat", "1", "Name", MapInfo.Name
    Var_Write MapEXPath & MapNum & ".dat", "1", "Weather", Str$(MapInfo.Weather)
    Var_Write MapEXPath & MapNum & ".dat", "1", "Music", Str$(MapInfo.Music)
    
    'Change caption
    frmMain.MapNameLbl.Caption = MapInfo.Name & " (" & MapNum & ")"

    'Map saved
    MsgBox "Map #" & MapNum & " (" & MapInfo.Name & ") successfully saved!", vbOKOnly
    
    Engine_CreateTileLayers

End Sub

Sub Main()

'*****************************************************************
'Main
'*****************************************************************
Dim LastUnloadTime As Long
Dim StartTime As Long
Dim FilePath As String
Dim i As Integer

    InitManifest

    'Init vars
    InitFilePaths

    'Load MySQL
    MySQL_Init
    
    'Show the screens
    frmMain.Show
    DoEvents
    
    DrawLayer = 1
    
    'Load DirectX
    Engine_Init_TileEngine frmScreen.hWnd, 32, 32, 18, 25, 10, 0.011
    
    'Check for the first map
    If LenB(Command$) = 0 Then
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
    
        StartTime = timeGetTime
        
        Input_Keys_ClearQueue

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
                    
                    If LastUnloadTime < timeGetTime Then
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
                        LastUnloadTime = timeGetTime + 5000
                    End If
        
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
        
        'Do sleep event - force FPS at ~60 (62.5) average (prevents extensive processing)
        If (timeGetTime - StartTime) < 16 Then  'If Elapsed Time < Time required for 60 FPS
            Sleep 16 - (timeGetTime - StartTime)
        End If

    Loop

    'Close Down
    frmMain.Timer1.Enabled = False
    frmMain.Timer2.Enabled = False
    frmMain.CritTimer.Enabled = False
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
Dim i As Long
Dim j As Long
Dim HighIndex As Long

    If Not frmParticles.Visible Then Exit Sub

    On Error GoTo ErrOut
    
    j = frmParticles.ParticlesList.ListIndex
    frmParticles.ParticlesList.Clear

    'Get the highest used index
    HighIndex = NumEffects
    Do While Not Effect(HighIndex).Used
        HighIndex = HighIndex - 1
        If HighIndex <= 0 Then Exit Sub
    Loop

    For i = 1 To HighIndex
        If Effect(i).Used = False Then
            frmParticles.ParticlesList.AddItem "Empty"
        Else
            If i = WeatherEffectIndex Then
                frmParticles.ParticlesList.AddItem "Reserved for weather"
            Else
                frmParticles.ParticlesList.AddItem "ID: " & Effect(i).EffectNum & " X: " & Effect(i).X + ParticleOffsetX & " Y: " & Effect(i).Y + ParticleOffsetY & " P: " & Effect(i).ParticleCount
            End If
        End If
    Next i
    frmParticles.ParticlesList.Refresh
    frmParticles.ParticlesList.ListIndex = j
    
    Exit Sub
    
ErrOut:

    SetInfo "Error updating the active effects list!", 1

End Sub
