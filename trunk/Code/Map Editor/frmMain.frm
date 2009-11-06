VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "vbGORE Map Editor"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1200
      Top             =   0
   End
   Begin VB.Timer CritTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1680
      Top             =   0
   End
   Begin MapEditor.ucToolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1014
      TabIndex        =   6
      Top             =   10380
      Width           =   15240
      Begin VB.CommandButton ARGBLongCmd 
         Caption         =   "ARGB <-> Long Tool"
         Height          =   315
         Left            =   7080
         TabIndex        =   11
         Top             =   30
         Width           =   1815
      End
      Begin VB.CommandButton SheetCmd 
         Caption         =   "View Sheet"
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Top             =   30
         Width           =   1215
      End
      Begin VB.ComboBox SearchCmb 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3120
         TabIndex        =   9
         Top             =   30
         Width           =   1575
      End
      Begin VB.TextBox SearchTxt 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "Enter your search here..."
         Top             =   30
         Width           =   2895
      End
      Begin VB.CommandButton SearchBtn 
         Caption         =   "Search"
         Height          =   315
         Left            =   4800
         TabIndex        =   7
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      FillColor       =   &H80000009&
      ForeColor       =   &H80000009&
      Height          =   225
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1014
      TabIndex        =   0
      Top             =   10815
      Width           =   15240
      Begin VB.Label InfoLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Random information of goodie-ness!"
         Top             =   0
         Width           =   930
      End
      Begin VB.Line LineName 
         BorderColor     =   &H80000009&
         X1              =   560
         X2              =   560
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Label MapNameLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Map Name"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   8520
         TabIndex        =   4
         ToolTipText     =   "Name of your currently loaded map"
         Top             =   0
         Width           =   2010
      End
      Begin VB.Line LineTile 
         BorderColor     =   &H80000009&
         X1              =   704
         X2              =   704
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Line LineMouse 
         BorderColor     =   &H80000009&
         X1              =   768
         X2              =   768
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Line LineFPS 
         BorderColor     =   &H80000009&
         X1              =   856
         X2              =   856
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Label TileLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   10680
         TabIndex        =   3
         ToolTipText     =   "Tile the cursor is hovering over"
         Top             =   0
         Width           =   675
      End
      Begin VB.Label MouseLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   11640
         TabIndex        =   2
         ToolTipText     =   "Pixel the cursor is hovering over"
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label FPSLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "FPS: 0"
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   12960
         TabIndex        =   1
         ToolTipText     =   "Frames per second"
         Top             =   0
         Width           =   780
      End
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

Private Sub BlocksPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the map 'blocked tiles' and 'no-attack tiles' editing form."

End Sub

Private Sub BrightPic_DblClick()
    BrightPic_Click
End Sub

Private Sub BrightPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Turn on / off bright mode (turns all tiles to brightest lighting)."

End Sub

Private Sub CharsPic_DblClick()
    CharsPic_Click
End Sub

Private Sub CharsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Show / hide characters on the map."

End Sub

Private Sub Command1_Click()

End Sub

Private Sub ARGBLongCmd_Click()

    ShowFrmARGB

End Sub

Private Sub CritTimer_Timer()
Static i As Long
    
    If InfoLbl.ForeColor = vbRed Then InfoLbl.ForeColor = &H80000009 Else InfoLbl.ForeColor = vbRed
    i = i + 1
    If i > 7 Then
        i = 0
        CritTimer.Enabled = False
    End If
    
End Sub

Private Sub ExitsPic_DblClick()
    ExitsPic_Click
End Sub

Private Sub ExitsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the exits (also known as warps) form."

End Sub

Private Sub FloodsPic_DblClick()
    FloodsPic_Click
End Sub

Private Sub FloodsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the 'floods' form (simulates screen click events over large areas)."

End Sub

Private Sub FPSLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo FPSLbl.ToolTipText
End Sub

Private Sub GridPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Hide / show the tile (32x32) grid."

End Sub

Private Sub InfoPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Hide / show tile information and attributes."

End Sub

Private Sub LoadPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Load an existing map."

End Sub

Private Sub MapNameLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo MapNameLbl.ToolTipText
End Sub

Private Sub MDIForm_Load()
Dim F As Form

    'Load the forms
    Load frmARGB
    Load frmBlock
    Load frmExit
    Load frmFloods
    Load frmMapInfo
    Load frmNPCs
    Load frmOptimizeStart
    Load frmParticles
    Load frmPreview
    Load frmScreen
    Load frmSearchAnim
    Load frmSearchTexture
    Load frmSetTile
    Load frmSfx
    Load frmTile
    Load frmTSOpt
    Load frmSearchList
    
    With Toolbar
        .Initialize 16, True, False, True
        .AddBitmap LoadResPicture("TOOLBARICONS", vbResBitmap), vbMagenta
 
        .AddButton , 7, "New"
        .AddButton , 10, "Load"
        .AddButton , 15, "Save"
        .AddButton , 16, "Save As"
        .AddButton , 11, "Optimize"
        
        .AddButton , , , eSeparator
        
        .AddButton , 12, "Set Tiles"
        .AddButton , 0, "Blocks"
        .AddButton , 13, "Floods"
        .AddButton , 6, "Tile Info"
        .AddButton , 1, "Exits"
        .AddButton , 8, "NPCs"
        .AddButton , 14, "Particles"
        .AddButton , 18, "Sfx"
        .AddButton , 4, "Map Info"
        
        .AddButton , , , eSeparator
        
        .AddButton , 20, "Toggle Weather"
        .AddButton , 9, "Toggle Characters"
        .AddButton , 5, "Toggle Bright Mode"
        .AddButton , 3, "Toggle Grid"
        .AddButton , 4, "Toggle Tile Info"
        .AddButton , 21, "Toggle Mini-map"

    End With

    'Load preferences
    On Error Resume Next
    For Each F In VB.Forms
        If UCase$(F.Name) <> "FRMTILESELECT" And UCase$(F.Name) <> "FRMMAIN" And UCase$(F.Name) <> "FRMTSOPT" Then
            F.Top = Val(Var_Get(Data2Path & "MapEditor.ini", F.Name, "Y"))
            F.Left = Val(Var_Get(Data2Path & "MapEditor.ini", F.Name, "X"))
            If UCase$(F.Name) = "FRMPREVIEW" Or UCase$(F.Name) = "FRMSEARCHLIST" Then
                F.Width = Val(Var_Get(Data2Path & "MapEditor.ini", F.Name, "W"))
                F.Height = Val(Var_Get(Data2Path & "MapEditor.ini", F.Name, "H"))
            End If
            F.Visible = Var_Get(Data2Path & "MapEditor.ini", F.Name, "V")
            If F.Visible Then F.Show Else F.Hide
        End If
    Next F
    On Error GoTo 0
    frmSearchAnim.Visible = False
    frmSearchTexture.Visible = False
    frmSearchList.Visible = False
    
    WeatherPic_Click
    CharsPic_Click
    
    SearchCmb.Clear
    SearchCmb.AddItem "Grh Index", 0
    SearchCmb.AddItem "File Number", 1
    SearchCmb.AddItem "Description", 2
    SearchCmb.ListIndex = 0

    tsTileWidth = Val(Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "W"))
    tsTileHeight = Val(Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "H"))
    tsStart = Val(Var_Get(Data2Path & "MapEditor.ini", "TSOPT", "S"))
    
    SetLayer 1
    
    'Show/hide all the other forms
    frmPreview.Show

    MDIForm_Resize

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo vbNullString
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim F As Form

    If IsUnloading = 0 Then
    
        'Ask if they want to save
        Select Case MsgBox("Are you sure you wish to quit?" & vbCrLf & "All unsaved changes will be lost!", vbYesNo)
            Case vbNo
            
                'Cancel the quitting
                If IsUnloading = 0 Then Cancel = 1
                Exit Sub
                
        End Select
        
        'Save config
        For Each F In VB.Forms
            If UCase$(F.Name) <> "FRMTILESELECT" And UCase$(F.Name) <> "FRMMAIN" And UCase$(F.Name) <> "FRMTSOPT" Then
                Var_Write Data2Path & "MapEditor.ini", F.Name, "X", F.Left
                Var_Write Data2Path & "MapEditor.ini", F.Name, "Y", F.Top
                Var_Write Data2Path & "MapEditor.ini", F.Name, "W", F.Width
                Var_Write Data2Path & "MapEditor.ini", F.Name, "H", F.Height
                Var_Write Data2Path & "MapEditor.ini", F.Name, "V", F.Visible
            End If
        Next F
        
        'Unload the engine
        IsUnloading = 1

    End If

End Sub

Private Sub MDIForm_Resize()

    If picInfo.ScaleWidth < 380 Then Exit Sub
    
    FPSLbl.Left = picInfo.ScaleWidth - 56
    LineFPS.x1 = picInfo.ScaleWidth - 64
    LineFPS.x2 = picInfo.ScaleWidth - 64
    MouseLbl.Left = picInfo.ScaleWidth - 144
    LineMouse.x1 = picInfo.ScaleWidth - 152
    LineMouse.x2 = picInfo.ScaleWidth - 152
    TileLbl.Left = picInfo.ScaleWidth - 208
    LineTile.x1 = picInfo.ScaleWidth - 216
    LineTile.x2 = picInfo.ScaleWidth - 216
    MapNameLbl.Left = picInfo.ScaleWidth - 350
    LineName.x1 = picInfo.ScaleWidth - 358
    LineName.x2 = picInfo.ScaleWidth - 358
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
    If MsgBox("Are you sure you wish to make a new map?" & vbNewLine & "All unsaved changes will be lost.", vbYesNo) = vbNo Then Exit Sub
    
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
    For tX = 1 To MapInfo.Width
        For tY = 1 To MapInfo.Height
        
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

Private Sub NewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Create a new map."

End Sub

Private Sub OptimizePic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Run the map optimizer to clean up unused information."

End Sub

Private Sub PartPic_DblClick()
    PartPic_Click
End Sub

Private Sub PartPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the particle effect placement form."

End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

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

Private Sub SaveAsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Save the current map as a new number."

End Sub

Private Sub SavePic_Click()

    'Confirm
    If MsgBox("Are you sure you wish to save the current map?", vbYesNo) = vbNo Then Exit Sub

    'Save the current map
    Game_SaveMapData CurMap
    
End Sub


Private Sub BlocksPic_Click()

    'Show/hide frmBlock
    frmBlock.Visible = (Not frmBlock.Visible)

End Sub

Private Sub BrightPic_Click()
Dim X As Byte
Dim Y As Byte
Dim i As Byte

    If BrightChkValue = 1 Then
        BrightChkValue = 0
    Else
        BrightChkValue = 1
    End If

    'Turn on
    If BrightChkValue = 1 Then
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
                For i = 1 To 24
                    MapData(X, Y).Light(i) = -1
                Next i
            Next Y
        Next X
    
    'Turn off
    ElseIf BrightChkValue = 0 Then
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
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
    Else
        CharsChkValue = 1
    End If

End Sub

Private Sub ExitsPic_Click()

    'Show/hide frmExits
    frmExit.Visible = (Not frmExit.Visible)

End Sub

Private Sub FloodsPic_Click()

    'Show/hide frmFloods
    frmFloods.Visible = (Not frmFloods.Visible)

End Sub


Private Sub GridPic_Click()

    If GridChkValue = 1 Then
        GridChkValue = 0
    Else
        GridChkValue = 1
    End If

End Sub

Private Sub InfoPic_Click()

    If InfoChkValue = 1 Then
        InfoChkValue = 0
    Else
        InfoChkValue = 1
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
    frmOptimizeStart.Visible = (Not frmOptimizeStart.Visible)

End Sub

Private Sub PartPic_Click()

    'Show/hide frmParticles
    frmParticles.Visible = (Not frmParticles.Visible)

End Sub

Private Sub SavePic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Save the current map."

End Sub

Private Sub SearchBtn_Click()
Dim s As String
Dim i As Long
Dim j As Long
    'On Error GoTo ErrOut

    Select Case SearchCmb.ListIndex
        
        'Search by Grh
        Case 0
        
            'Load the texture that contains the grh they entered
            i = Val(SearchTxt.Text)
            If i <= 0 Then GoTo ErrOut
            If i > UBound(GrhData()) Then GoTo ErrOut
            j = GrhData(i).FileNum
            If j <= 0 Then GoTo ErrOut
            If j > NumGrhFiles Then GoTo ErrOut
            If LoadTextureToForm(frmSearchTexture, j) = 0 Then GoTo ErrOut
            SearchTextureFileNum = j
            
        'Search by file
        Case 1
            
            i = Val(SearchTxt.Text)
            If i <= 0 Then GoTo ErrOut
            If i > NumGrhFiles Then GoTo ErrOut
            If LoadTextureToForm(frmSearchTexture, i) = 0 Then GoTo ErrOut
            SearchTextureFileNum = i
            
        'Search by description
        Case 2
            
            NumDescResults = 0
            Erase DescResults
            frmSearchList.Caption = "Search: """ & SearchTxt.Text & """"
            s = UCase$(SearchTxt.Text)
            For i = 1 To NumTextureDesc
                If InStr(1, UCase$(TextureDesc(i)), s) Then
                    NumDescResults = NumDescResults + 1
                    ReDim Preserve DescResults(1 To NumDescResults)
                    DescResults(NumDescResults) = i
                    frmSearchList.SearchLst.AddItem i & " - " & TextureDesc(i)
                End If
            Next i
            frmSearchList.Visible = True
            frmSearchList.Show
            
    End Select
    
    Exit Sub
    
ErrOut:

    SetInfo "Invalid or unknown search value entered!", 1

End Sub

Private Sub SearchBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Perform the selected graphic search"

End Sub

Private Sub SearchTxt_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SearchBtn_Click
        KeyAscii = 0
    End If
    
    If SearchCmb.ListIndex <> 2 Then
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
            SetInfo "You may only enter alphanumeric searches when performing a description search."
        End If
    End If
    
End Sub

Private Sub SearchTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Value that you wish to search for. What to enter depends on what search method is used."

End Sub

Private Sub SearchTxt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If SearchTxt.ForeColor = &HC0C0C0 Then
        SearchTxt.ForeColor = &H80000008
        SearchTxt.Text = vbNullString
    End If
    
End Sub

Private Sub SetSfxPic_Click()

    'Show/hide frmSfx
    frmSfx.Visible = (Not frmSfx.Visible)

End Sub

Private Sub SetSfxPic_DblClick()
    SetSfxPic_Click
End Sub

Private Sub SetSfxPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the map-bound sound effect placement form."

End Sub

Private Sub SetTilesPic_Click()

    'Show/hide frmSetTile
    frmSetTile.Visible = (Not frmSetTile.Visible)

End Sub

Private Sub SetTilesPic_DblClick()
    SetTilesPic_Click
End Sub

Private Sub SetTilesPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the map tile editing form."

End Sub

Private Sub SheetCmd_Click()

    ShowFrmTileSelect 1

End Sub

Private Sub ShowMapInfoPic_Click()

    'Show/hide frmMapInfo
    frmMapInfo.Visible = (Not frmMapInfo.Visible)

End Sub

Private Sub ShowMapInfoPic_DblClick()
    ShowMapInfoPic_Click
End Sub

Private Sub ShowMapInfoPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the general map attributes and information form."

End Sub

Private Sub ShowNPCsPic_Click()

    'Show/hide frmNPCs
    frmNPCs.Visible = (Not frmNPCs.Visible)

End Sub

Private Sub ShowNPCsPic_DblClick()

    ShowNPCsPic_Click
    
End Sub

Private Sub ShowNPCsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the NPC placement form."

End Sub

Private Sub SheetCmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "View the tile selection sheet."

End Sub

Private Sub TileLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetInfo TileLbl.ToolTipText
End Sub

Private Sub Timer1_Timer()

    MDIForm_Resize
    
End Sub

Private Sub Timer2_Timer()
Static LastThingy As Long   'Yes, I named a variable "thingy", wanna fight about it!? >:|
Dim i As Long
Dim R As RECT
Dim j As Long

    'This timer brought to you by a lazy ass programmer :)
    'Fuck speed, this is a tool!

    If D3DDevice Is Nothing Then Exit Sub
    If SearchTextureFileNum > 0 Then
    
        If ShownTextureGrhs.NumGrhs > 0 Then

            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            D3DDevice.BeginScene
            
            If LastThingy <> SearchTextureFileNum Then
                frmSearchTexture.Refresh
                LastThingy = SearchTextureFileNum
            End If
            Engine_Render_FullTexture frmSearchTexture.hWnd, SearchTextureFileNum
            
            j = D3DColorARGB(150, 255, 255, 255)
            For i = 1 To ShownTextureGrhs.NumGrhs
                With ShownTextureGrhs.Grh(i)
                    Engine_Render_Rectangle .X, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X + .Width, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y + .Height, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                End With
            Next i
            
            D3DDevice.EndScene
            
            R.Left = 0
            R.Top = 0
            R.Right = SurfaceSize(SearchTextureFileNum).X
            R.bottom = SurfaceSize(SearchTextureFileNum).Y
            
            D3DDevice.Present R, R, frmSearchTexture.hWnd, ByVal 0
        
        End If
        
        If ShownTextureAnims.NumGrhs > 0 Then
            
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            D3DDevice.BeginScene
            
            If LastThingy <> SearchTextureFileNum Then
                frmSearchTexture.Refresh
                LastThingy = SearchTextureFileNum
            End If
            
            For i = 1 To ShownTextureAnims.NumGrhs
                With ShownTextureAnims.Grh(i)
                    Engine_Render_Grh .Grh, .X, .Y, 0, 1, True
                End With
            Next i
            
            j = D3DColorARGB(150, 255, 255, 255)
            For i = 1 To ShownTextureGrhs.NumGrhs
                With ShownTextureGrhs.Grh(i)
                    Engine_Render_Rectangle .X, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X + .Width, .Y, 1, .Height, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                    Engine_Render_Rectangle .X, .Y + .Height, .Width, 1, 0, 0, 1, 1, 1, 1, 0, 0, j, j, j, j
                End With
            Next i
            
            D3DDevice.EndScene
            
            R.Left = 0
            R.Top = 0
            R.Right = STAWidth
            R.bottom = STAHeight
            
            D3DDevice.Present R, R, frmSearchAnim.hWnd, ByVal 0
            
        End If
        
    End If
    
    DrawPreview
    DrawTileInfoPreview

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As Long)

    Select Case Button
        Case 1: NewPic_Click
        Case 2: LoadPic_Click
        Case 3: SavePic_Click
        Case 4: SaveAsPic_Click
        Case 5: frmOptimizeStart.Visible = True: frmOptimizeStart.Show
        
        Case 7: SetTilesPic_Click
        Case 8: BlocksPic_Click
        Case 9: FloodsPic_Click
        Case 10: ViewTilesPic_Click
        Case 11: ExitsPic_Click
        Case 12: ShowNPCsPic_Click
        Case 13: PartPic_Click
        Case 14: SetSfxPic_Click
        Case 15: ShowMapInfoPic_Click
        
        Case 17: WeatherPic_Click
        Case 18: CharsPic_Click
        Case 19: BrightPic_Click
        Case 20: GridPic_Click
        Case 21: InfoPic_Click
        Case 22: If ShowMiniMap = 0 Then ShowMiniMap = 1 Else ShowMiniMap = 0: Engine_BuildMiniMap
        
    End Select

End Sub

Private Sub Toolbar_ButtonEnter(ByVal Button As Long)

    Select Case Button
        Case 1: SetInfo "Create a new map."
        Case 2: SetInfo "Load an existing map."
        Case 3: SetInfo "Save the current map over the existing number."
        Case 4: SetInfo "Save the current map as a new number."
        Case 5: SetInfo "Run the map optimizer to clean up unused information."
        
        'Case 6: Sep
        
        Case 7: SetInfo "Display the map tile editing form."
        Case 8: SetInfo "Display the map 'blocked tiles' and 'no-attack tiles' editing form."
        Case 9: SetInfo "Display the 'floods' form (simulates screen click events over large areas)."
        Case 10: SetInfo "Display the tile information form (right-click the game screen to set the tile to view)."
        Case 11: SetInfo "Display the exits (also known as warps) form."
        Case 12: SetInfo "Display the NPC placement form."
        Case 13: SetInfo "Display the particle effect placement form."
        Case 14: SetInfo "Display the map-bound sound effect placement form."
        Case 15: SetInfo "Display the general map attributes and information form."
        
        'Case 16: Sep
        
        Case 17: SetInfo "Hide / show map weather."
        Case 18: SetInfo "Show / hide characters on the map."
        Case 19: SetInfo "Turn on / off bright mode (turns all tiles to brightest lighting)."
        Case 20: SetInfo "Hide / show the tile (32x32) grid."
        Case 21: SetInfo "Hide / show tile information and attributes."
        Case 22: SetInfo "Hide / show the mini-map."
        
    End Select

End Sub

Private Sub ViewTilesPic_Click()

    'Show/hide frmViewTiles
    frmTile.Visible = (Not frmTile.Visible)

End Sub

Private Sub ViewTilesPic_DblClick()
    ViewTilesPic_Click
End Sub

Private Sub ViewTilesPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Display the tile information form (right-click the game screen to set the tile to view)."

End Sub

Private Sub WeatherPic_Click()

    If WeatherChkValue = 1 Then
        WeatherChkValue = 0
    Else
        WeatherChkValue = 1
    End If

End Sub

Private Sub WeatherPic_DblClick()
    WeatherPic_Click
End Sub

Private Sub WeatherPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Hide / show map weather."

End Sub
