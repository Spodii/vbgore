VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "vbGORE Map Editor"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   1
      Top             =   10815
      Width           =   15390
      Begin VB.Label InfoLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Random information of goodie-ness!"
         Top             =   0
         Width           =   930
      End
      Begin VB.Line LineName 
         X1              =   560
         X2              =   560
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Label MapNameLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Map Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8520
         TabIndex        =   5
         ToolTipText     =   "Name of your currently loaded map"
         Top             =   0
         Width           =   2010
      End
      Begin VB.Line LineTile 
         X1              =   704
         X2              =   704
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Line LineMouse 
         X1              =   768
         X2              =   768
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Line LineFPS 
         X1              =   856
         X2              =   856
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Label TileLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10680
         TabIndex        =   4
         ToolTipText     =   "Tile the cursor is hovering over"
         Top             =   0
         Width           =   675
      End
      Begin VB.Label MouseLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11640
         TabIndex        =   3
         ToolTipText     =   "Pixel the cursor is hovering over"
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label FPSLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "FPS: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12960
         TabIndex        =   2
         ToolTipText     =   "Frames per second"
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1026
      TabIndex        =   0
      Top             =   0
      Width           =   15390
      Begin VB.Image BlocksPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5760
         Top             =   0
         Width           =   480
      End
      Begin VB.Image SaveAsPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1920
         Top             =   0
         Width           =   480
      End
      Begin VB.Image SavePic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1440
         Top             =   0
         Width           =   480
      End
      Begin VB.Image LoadPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   960
         Top             =   0
         Width           =   480
      End
      Begin VB.Image PartPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   8160
         Top             =   0
         Width           =   480
      End
      Begin VB.Image ShowNPCsPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   7680
         Top             =   0
         Width           =   480
      End
      Begin VB.Image ExitsPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   7200
         Top             =   0
         Width           =   480
      End
      Begin VB.Image ViewTilesPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6720
         Top             =   0
         Width           =   480
      End
      Begin VB.Image FloodsPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6240
         Top             =   0
         Width           =   480
      End
      Begin VB.Image SetTilesPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5280
         Top             =   0
         Width           =   480
      End
      Begin VB.Image InfoPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   4560
         Top             =   0
         Width           =   480
      End
      Begin VB.Image GridPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   4080
         Top             =   0
         Width           =   480
      End
      Begin VB.Image BrightPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   3600
         Top             =   0
         Width           =   480
      End
      Begin VB.Image CharsPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   3120
         Top             =   0
         Width           =   480
      End
      Begin VB.Image WeatherPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   2640
         Top             =   0
         Width           =   480
      End
      Begin VB.Image OptimizePic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   480
         Top             =   0
         Width           =   480
      End
      Begin VB.Image NewPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Image SetSfxPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   8640
         Top             =   0
         Width           =   480
      End
      Begin VB.Image ShowMapInfoPic 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   9120
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   12360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BlocksPic_DblClick()
    BlocksPic_Click
End Sub

Private Sub BrightPic_DblClick()
    BrightPic_Click
End Sub

Private Sub CharsPic_DblClick()
    CharsPic_Click
End Sub

Private Sub ExitsPic_DblClick()
    ExitsPic_Click
End Sub

Private Sub FloodsPic_DblClick()
    FloodsPic_Click
End Sub

Private Sub FPSLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo FPSLbl.ToolTipText
End Sub

Private Sub Image5_Click()

End Sub

Private Sub MapNameLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo MapNameLbl.ToolTipText
End Sub

Private Sub MDIForm_Load()
Dim F As Form

    GrhMapPath = App.Path & "\FormSkins\" & Skin_GetCurrent & "\mapeditor\"

    'Load preferences
    frmMain.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "MAIN", "X"))
    frmMain.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "MAIN", "Y"))
    frmTile.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "TILE", "X"))
    frmTile.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "TILE", "Y"))
    frmSetTile.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "SETTILE", "X"))
    frmSetTile.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "SETTILE", "Y"))
    frmNPCs.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "NPCS", "X"))
    frmNPCs.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "NPCS", "Y"))
    frmMapInfo.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "MAPINFO", "X"))
    frmMapInfo.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "MAPINFO", "Y"))
    frmParticles.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "PART", "X"))
    frmParticles.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "PART", "Y"))
    frmFloods.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "FLOODS", "X"))
    frmFloods.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "FLOODS", "Y"))
    frmExit.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "EXIT", "X"))
    frmExit.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "EXIT", "Y"))
    frmBlock.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "BLOCK", "X"))
    frmBlock.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "BLOCK", "Y"))
    frmSfx.Left = Val(Var_Get(Data2Path & "MapEditor.ini", "SFX", "X"))
    frmSfx.Top = Val(Var_Get(Data2Path & "MapEditor.ini", "SFX", "Y"))
    tsTileWidth = Val(Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "W"))
    tsTileHeight = Val(Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "H"))
    tsStart = Val(Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "S"))
    
    'Set the tools
    WeatherChkValue = 0
    WeatherPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\weatherg.*"))
    CharsChkValue = 1
    CharsPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\shownpc.*"))
    BrightChkValue = 0
    BrightPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\brightg.*"))
    GridChkValue = 0
    GridPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\gridg.*"))
    InfoChkValue = 1
    InfoPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\info.*"))
    SfxChkValue = 0
    SetSfxPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\soundsg.*"))
    
    OptimizePic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\optimize.*"))
    LoadPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\load.*"))
    SavePic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\save.*"))
    SaveAsPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\saveas.*"))
    NewPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\newbutton.*"))
    
    SetLayer 1
    
    'Skin settings
    Skin_InitStructure Var_Get(App.Path & "\FormSkins\CurrentSkin.ini", "INIT", "CurrentSkin")
    For Each F In VB.Forms
        Skin_SetForm F
    Next F
    
    'Override settings
    Me.BackColor = &H8000000C
    frmPreview.BackColor = vbBlack
    frmPreview.Width = 128 * Screen.TwipsPerPixelX
    frmPreview.Height = 128 * Screen.TwipsPerPixelY
    
    'Show/hide all the other forms
    HideFrmTile
    HideFrmSetTile
    HideFrmMapInfo
    HideFrmNPCs
    HideFrmParticles
    HideFrmFloods
    HideFrmExit
    HideFrmBlock
    HideFrmSfx
    frmPreview.Show
    
    '//TEMP
    MsgBox "This map editor is far from complete, so either wait for it to be finished or use a copy from an older version.", vbOKOnly
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo vbNullString
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then
    
        'Ask if they want to save
        Select Case MsgBox("Are you sure you wish to quit?" & vbCrLf & "All unsaved changes will be lost!", vbYesNo)
            Case vbNo
            
                'Cancel the quitting
                Cancel = 1
                Exit Sub
                
        End Select
        
        'Unload the engine
        IsUnloading = 1
        
        'Save positions
        Var_Write Data2Path & "MapEditor.ini", "MAIN", "X", frmMain.Left
        Var_Write Data2Path & "MapEditor.ini", "MAIN", "Y", frmMain.Top

    End If

End Sub

Private Sub MDIForm_Resize()

    If picInfo.ScaleWidth < 380 Then Exit Sub
    
    FPSLbl.Left = picInfo.ScaleWidth - 56
    LineFPS.X1 = picInfo.ScaleWidth - 64
    LineFPS.X2 = picInfo.ScaleWidth - 64
    MouseLbl.Left = picInfo.ScaleWidth - 144
    LineMouse.X1 = picInfo.ScaleWidth - 152
    LineMouse.X2 = picInfo.ScaleWidth - 152
    TileLbl.Left = picInfo.ScaleWidth - 208
    LineTile.X1 = picInfo.ScaleWidth - 216
    LineTile.X2 = picInfo.ScaleWidth - 216
    MapNameLbl.Left = picInfo.ScaleWidth - 350
    LineName.X1 = picInfo.ScaleWidth - 358
    LineName.X2 = picInfo.ScaleWidth - 358
    InfoLbl.Width = picInfo.ScaleWidth - 374

End Sub

Private Sub GridPic_DblClick()
    GridPic_Click
End Sub

Private Sub InfoPic_DblClick()
    InfoPic_Click
End Sub

Private Sub MouseLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo MouseLbl.ToolTipText
End Sub

Private Sub NewPic_Click()
Dim tX As Long
Dim tY As Long
Dim i As Long

    'Confirm
    If MsgBox("Are you sure you wish to clear the current map?", vbYesNo) = vbNo Then Exit Sub
    
    'Turn off bright mode
    If BrightChkValue = 1 Then BrightPic_Click
    
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
    
    'Clear the map info
    MapInfo.MapVersion = 1
    MapInfo.Music = 0
    MapInfo.Name = "New map"
    MapInfo.Weather = 0
    
    'Clear the map tiles
    For tX = XMinMapSize To XMaxMapSize
        For tY = YMinMapSize To YMaxMapSize
        
            'Erase graphics / shadows
            For i = 1 To 6
                MapData(tX, tY).Graphic(i).GrhIndex = 0
                MapData(tX, tY).Shadow(i) = 0
            Next i
            
            'Erase lights
            For i = 1 To 24
                MapData(tX, tY).Light(i) = -1
            Next i
            
            'Erase NPCs
            If MapData(tX, tY).NPCIndex > 0 Then
                Engine_Char_Erase MapData(tX, tY).NPCIndex
                MapData(tX, tY).NPCIndex = 0
            End If
            
            'Erase objects
            If MapData(tX, tY).ObjInfo.ObjIndex > 0 Then
                Engine_OBJ_Erase MapData(tX, tY).ObjInfo.ObjIndex
                MapData(tX, tY).ObjInfo.ObjIndex = 0
                MapData(tX, tY).ObjInfo.ObjIndex = 0
            End If
            
            'Erase flags
            MapData(tX, tY).Blocked = 0
            MapData(tX, tY).Mailbox = 0
            MapData(tX, tY).Sfx = 0
            MapData(tX, tY).TileExit.Map = 0
            MapData(tX, tY).TileExit.X = 0
            MapData(tX, tY).TileExit.Y = 0
            MapData(tX, tY).UserIndex = 0
            MapData(tX, tY).BlockedAttack = 0
            MapData(tX, tY).Sign = 0
            
        Next tY
    Next tX

End Sub

Private Sub PartPic_DblClick()
    PartPic_Click
End Sub

Private Sub SaveAsPic_Click()
Dim NewMapVal As Integer

    'Confirm
    If MsgBox("Are you sure you wish to save the current map as a new map?", vbYesNo) = vbNo Then Exit Sub
    
    'Get value
    NewMapVal = InputBox("Please enter the map number for the new map.")
    
    'Check if the file already exists
    If Engine_FileExist(MapPath & NewMapVal & ".map", vbNormal) Then
        If MsgBox("Map " & NewMapVal & " already exists, do you wish to save over the current map? If so, the old map will not be able to be restored!", vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Save the current map
    Game_SaveMapData NewMapVal
    CurMap = NewMapVal

End Sub

Private Sub SavePic_Click()

    'Confirm
    If MsgBox("Are you sure you wish to save the current map?", vbYesNo) = vbNo Then Exit Sub

    'Save the current map
    Game_SaveMapData CurMap
    
End Sub


Private Sub BlocksPic_Click()

    'Show/hide frmBlock
    If BlocksChkValue = 1 Then HideFrmBlock Else ShowFrmBlock

End Sub

Private Sub BrightPic_Click()
Dim X As Byte
Dim Y As Byte
Dim i As Byte

    If BrightChkValue = 1 Then
        BrightChkValue = 0
        BrightPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\brightg.*"))
    Else
        BrightChkValue = 1
        BrightPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\bright.*"))
    End If

    'Turn on
    If BrightChkValue = 1 Then
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                For i = 1 To 24
                    MapData(X, Y).Light(i) = -1
                Next i
            Next Y
        Next X
    
    'Turn off
    ElseIf BrightChkValue = 0 Then
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                For i = 1 To 24
                    MapData(X, Y).Light(i) = SaveLightBuffer(X, Y).Light(i)
                Next i
            Next Y
        Next X
    
    End If

End Sub

Private Sub CharsPic_Click()

    If CharsChkValue = 1 Then
        CharsChkValue = 0
        CharsPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\shownpcg.*"))
    Else
        CharsChkValue = 1
        CharsPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\shownpc.*"))
    End If

End Sub

Private Sub ExitsPic_Click()

    'Show/hide frmExits
    If ExitsChkValue Then HideFrmExit Else ShowFrmExit

End Sub

Private Sub FloodsPic_Click()

    'Show/hide frmFloods
    If FloodsChkValue = 1 Then HideFrmFloods Else ShowFrmFloods

End Sub


Private Sub GridPic_Click()

    If GridChkValue = 1 Then
        GridChkValue = 0
        GridPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\gridg.*"))
    Else
        GridChkValue = 1
        GridPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\grid.*"))
    End If

End Sub

Private Sub InfoPic_Click()

    If InfoChkValue = 1 Then
        InfoChkValue = 0
        InfoPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\infog.*"))
    Else
        InfoChkValue = 1
        InfoPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\info.*"))
    End If

End Sub

Private Sub LoadPic_Click()
Dim FileName As String

    On Error GoTo ErrOut

    'Confirm
    If MsgBox("Are you sure you wish to load another map?" & vbCrLf & "Any changes made to the current map will be lost!", vbYesNo) = vbNo Then Exit Sub
    
    'Load map
    With frmMain.CD
        .Filter = "Maps|*.map"
        .DialogTitle = "Load"
        .FileName = ""
        .InitDir = MapPath
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End With
    FileName = Right$(frmMain.CD.FileName, Len(frmMain.CD.FileName) - Len(MapPath))
    Game_Map_Switch CInt(Left$(FileName, Len(FileName) - 4))
    
ErrOut:

End Sub

Private Sub OptimizePic_Click()

    'Show optimization form
    ShowFrmOptimizeStart

End Sub

Private Sub PartPic_Click()

    'Show/hide frmParticles
    If PartChkValue = 1 Then HideFrmParticles Else ShowFrmParticles

End Sub

Private Sub SetSfxPic_Click()

    'Show/hide frmSfx
    If SfxChkValue = 1 Then HideFrmSfx Else ShowFrmSfx

End Sub

Private Sub SetSfxPic_DblClick()
    SetSfxPic_Click
End Sub

Private Sub SetTilesPic_Click()

    'Show/hide frmSetTile
    If SetTilesChkValue = 1 Then HideFrmSetTile Else ShowFrmSetTile

End Sub

Private Sub SetTilesPic_DblClick()
    SetTilesPic_Click
End Sub

Private Sub ShowMapInfoPic_Click()

    'Show/hide frmMapInfo
    If ShowMapInfoChkValue = 1 Then HideFrmMapInfo Else ShowFrmMapInfo

End Sub

Private Sub ShowMapInfoPic_DblClick()
    ShowMapInfoPic_Click
End Sub

Private Sub ShowNPCsPic_Click()

    'Show/hide frmNPCs
    If ShowNPCsChkValue = 1 Then HideFrmNPCs Else ShowFrmNPCs

End Sub

Private Sub ShowNPCsPic_DblClick()

    ShowNPCsPic_Click
    
End Sub

Private Sub TileLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo TileLbl.ToolTipText
End Sub

Private Sub ViewTilesPic_Click()

    'Show/hide frmViewTiles
    If ViewTilesChkValue = 1 Then HideFrmTile Else ShowFrmTile

End Sub

Private Sub ViewTilesPic_DblClick()
    ViewTilesPic_Click
End Sub

Private Sub WeatherPic_Click()

    If WeatherChkValue = 1 Then
        WeatherChkValue = 0
        WeatherPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\weatherg.*"))
    Else
        WeatherChkValue = 1
        WeatherPic.Picture = LoadPicture(GrhMapPath & Dir$(GrhMapPath & "\weather.*"))
    End If

End Sub

Private Sub WeatherPic_DblClick()
    WeatherPic_Click
End Sub
