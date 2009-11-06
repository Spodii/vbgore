VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Map Editor"
   ClientHeight    =   10440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   12000
   Begin VB.PictureBox L2Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4380
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   23
      Top             =   1245
      Width           =   120
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   11400
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ScreenPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   1440
      Width           =   12000
   End
   Begin VB.PictureBox ToolbarPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox NewPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   34
         ToolTipText     =   "New: Clear the current map and make it a new map"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox L6Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   5340
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   27
         Top             =   1245
         Width           =   120
      End
      Begin VB.PictureBox L5Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   5340
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   26
         Top             =   1020
         Width           =   120
      End
      Begin VB.PictureBox L4Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   4860
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   25
         Top             =   1245
         Width           =   120
      End
      Begin VB.PictureBox L3Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   4860
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   24
         Top             =   1020
         Width           =   120
      End
      Begin VB.PictureBox L1Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   4380
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   22
         Top             =   1020
         Width           =   120
      End
      Begin VB.PictureBox ExitPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   11520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   21
         ToolTipText     =   "Quit: Exit vbGORE Map Editor"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox OptimizePic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9600
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         ToolTipText     =   "Optimize: Perform automatic map performance/size optimization check algorithm"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox LoadPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   19
         ToolTipText     =   "Load: Load an existing map file"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox SavePic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         ToolTipText     =   "Save: Save currently displayed map as the current map number"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox SaveAsPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   11040
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         ToolTipText     =   "Save As: Save currently displayed map as a different map number (new number or overwrite existing map)"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox WeatherPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   16
         ToolTipText     =   "Hide/Show weather effects"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox CharsPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         ToolTipText     =   "Hide/Show NPCs placed on the map"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox ObjPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         ToolTipText     =   "Hide/Show objects placed on the map"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox BrightPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         ToolTipText     =   $"frmMain.frx":17D2A
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox GridPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         ToolTipText     =   "Hide/Show the 32x32 grid"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox InfoPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2400
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         ToolTipText     =   "Hide/Show information flag squares on tiles"
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox SetTilesPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   3360
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   10
         ToolTipText     =   "Hide/Show tile placement form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox ViewTilesPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   6240
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   9
         ToolTipText     =   "Hide/Show selected tile information form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox ShowMapInfoPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   11040
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   8
         ToolTipText     =   "Hide/Show map information form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox ShowNPCsPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   8160
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   7
         ToolTipText     =   "Hide/Show NPC placement/removal form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox PartPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   9120
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   6
         ToolTipText     =   "Hide/Show particle effect placement/removal form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox FloodsPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   4320
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   5
         ToolTipText     =   "Hide/Show map flooding form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox ExitsPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   7200
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   4
         ToolTipText     =   "Hide/Show exit placement/removal form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox BlocksPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   5280
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   3
         ToolTipText     =   "Hide/Show blocked tile placement/removal form"
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox SetSfxPic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   10080
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   2
         ToolTipText     =   "Hide/Show map-based sound effects placement/removal form"
         Top             =   0
         Width           =   960
      End
      Begin VB.Label FPSLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FPS: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   255
         Left            =   7920
         TabIndex        =   33
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label YLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   195
         Left            =   7080
         TabIndex        =   32
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label XLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   195
         Left            =   6120
         TabIndex        =   31
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label TileYLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TileY: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   195
         Left            =   7080
         TabIndex        =   30
         Top             =   960
         Width           =   675
      End
      Begin VB.Label TileXLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TileX: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   195
         Left            =   6120
         TabIndex        =   29
         Top             =   960
         Width           =   675
      End
      Begin VB.Label MapLbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Map: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   255
         Left            =   7920
         TabIndex        =   28
         Top             =   960
         Width           =   1095
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

Private Sub BrightPic_DblClick()
    BrightPic_Click
End Sub

Private Sub CharsPic_DblClick()
    CharsPic_Click
End Sub

Private Sub ExitPic_Click()
    
    'Ask if they want to save
    Select Case MsgBox("Are you sure you wish to quit?" & vbCrLf & "All unsaved changes will be lost!", vbYesNo)
        Case vbNo
        
            'Cancel the quitting
            Exit Sub
            
    End Select
    
    'Unload the engine
    IsUnloading = 1
    
    'Save positions
    Var_Write Data2Path & "MapEditor.ini", "MAIN", "X", frmMain.Left
    Var_Write Data2Path & "MapEditor.ini", "MAIN", "Y", frmMain.Top

End Sub

Private Sub ExitsPic_DblClick()
    ExitsPic_Click
End Sub

Private Sub FloodsPic_DblClick()
    FloodsPic_Click
End Sub

Private Sub Form_Load()

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
    
    'Set the toolbar
    ToolbarPic.Picture = LoadPicture(GrhMapPath & "toolbar.bmp")
    
    'Set the tools
    WeatherChkValue = 0
    WeatherPic.Picture = LoadPicture(GrhMapPath & "weatherg.bmp")
    CharsChkValue = 1
    CharsPic.Picture = LoadPicture(GrhMapPath & "shownpc.bmp")
    ObjChkValue = 1
    ObjPic.Picture = LoadPicture(GrhMapPath & "showobj.bmp")
    BrightChkValue = 0
    BrightPic.Picture = LoadPicture(GrhMapPath & "brightg.bmp")
    GridChkValue = 0
    GridPic.Picture = LoadPicture(GrhMapPath & "gridg.bmp")
    InfoChkValue = 1
    InfoPic.Picture = LoadPicture(GrhMapPath & "info.bmp")
    SfxChkValue = 0
    SetSfxPic.Picture = LoadPicture(GrhMapPath & "soundsg.bmp")
    
    OptimizePic.Picture = LoadPicture(GrhMapPath & "optimize.bmp")
    LoadPic.Picture = LoadPicture(GrhMapPath & "load.bmp")
    SavePic.Picture = LoadPicture(GrhMapPath & "save.bmp")
    SaveAsPic.Picture = LoadPicture(GrhMapPath & "saveas.bmp")
    ExitPic.Picture = LoadPicture(GrhMapPath & "exitbutton.bmp")
    NewPic.Picture = LoadPicture(GrhMapPath & "newbutton.bmp")
    
    'Set the layer checks
    L1ChkValue = 1
    L2ChkValue = 1
    L3ChkValue = 1
    L4ChkValue = 1
    L5ChkValue = 1
    L6ChkValue = 1
    L1Pic.Picture = LoadPicture(GrhMapPath & "topbuttonselected.bmp")
    L2Pic.Picture = LoadPicture(GrhMapPath & "bottombuttonselected.bmp")
    L3Pic.Picture = LoadPicture(GrhMapPath & "topbuttonselected.bmp")
    L4Pic.Picture = LoadPicture(GrhMapPath & "bottombuttonselected.bmp")
    L5Pic.Picture = LoadPicture(GrhMapPath & "topbuttonselected.bmp")
    L6Pic.Picture = LoadPicture(GrhMapPath & "bottombuttonselected.bmp")
    
    'Show/hide all the other forms
    HideFrmTile
    HideFrmSetTile
    HideFrmMapInfo
    HideFrmNPCs
    HideFrmParticles
    HideFrmFloods
    HideFrmExit
    HideFrmBlock
    
End Sub

Private Sub GridPic_DblClick()
    GridPic_Click
End Sub

Private Sub InfoPic_DblClick()
    InfoPic_Click
End Sub

Private Sub L1Pic_Click()

    If L1ChkValue = 1 Then
        L1ChkValue = 0
        L1Pic.Picture = LoadPicture(GrhMapPath & "\topbuttonselectedg.bmp")
    Else
        L1ChkValue = 1
        L1Pic.Picture = LoadPicture(GrhMapPath & "\topbuttonselected.bmp")
    End If

End Sub

Private Sub L1Pic_DblClick()
    L1Pic_Click
End Sub

Private Sub L2Pic_Click()

    If L2ChkValue = 1 Then
        L2ChkValue = 0
        L2Pic.Picture = LoadPicture(GrhMapPath & "\bottombuttonselectedg.bmp")
    Else
        L2ChkValue = 1
        L2Pic.Picture = LoadPicture(GrhMapPath & "\bottombuttonselected.bmp")
    End If
    
End Sub

Private Sub L2Pic_DblClick()
    L2Pic_Click
End Sub

Private Sub L3Pic_Click()

    If L3ChkValue = 1 Then
        L3ChkValue = 0
        L3Pic.Picture = LoadPicture(GrhMapPath & "\topbuttonselectedg.bmp")
    Else
        L3ChkValue = 1
        L3Pic.Picture = LoadPicture(GrhMapPath & "\topbuttonselected.bmp")
    End If

End Sub

Private Sub L3Pic_DblClick()
    L3Pic_Click
End Sub

Private Sub L4Pic_Click()

    If L4ChkValue = 1 Then
        L4ChkValue = 0
        L4Pic.Picture = LoadPicture(GrhMapPath & "\bottombuttonselectedg.bmp")
    Else
        L4ChkValue = 1
        L4Pic.Picture = LoadPicture(GrhMapPath & "\bottombuttonselected.bmp")
    End If
    
End Sub

Private Sub L4Pic_DblClick()
    L4Pic_Click
End Sub

Private Sub L5Pic_Click()

    If L5ChkValue = 1 Then
        L5ChkValue = 0
        L5Pic.Picture = LoadPicture(GrhMapPath & "\topbuttonselectedg.bmp")
    Else
        L5ChkValue = 1
        L5Pic.Picture = LoadPicture(GrhMapPath & "\topbuttonselected.bmp")
    End If

End Sub

Private Sub L5Pic_DblClick()
    L5Pic_Click
End Sub

Private Sub L6Pic_Click()

    If L6ChkValue = 1 Then
        L6ChkValue = 0
        L6Pic.Picture = LoadPicture(GrhMapPath & "\bottombuttonselectedg.bmp")
    Else
        L6ChkValue = 1
        L6Pic.Picture = LoadPicture(GrhMapPath & "\bottombuttonselected.bmp")
    End If
    
End Sub

Private Sub L6Pic_DblClick()
    L6Pic_Click
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

Private Sub ScreenPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Call the MouseMove event
    ScreenPic_MouseMove Button, Shift, X, Y

End Sub

Private Sub ScreenPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tX As Integer
Dim tY As Integer

    'Convert the click position to tile position
    Engine_ConvertCPtoTP 0, 0, X, Y, tX, tY
    HovertX = tX
    HovertY = tY

    'Update caption
    HoverX = X + ParticleOffsetX - 288
    HoverY = Y + ParticleOffsetY - 288
    XLbl.Caption = "X: " & HoverX
    YLbl.Caption = "Y: " & HoverY
    TileXLbl.Caption = "TileX: " & HovertX
    TileYLbl.Caption = "TileY: " & HovertY

    'Click the tile
    SetTile tX, tY, Button, Shift
             
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
        BrightPic.Picture = LoadPicture(GrhMapPath & "brightg.bmp")
    Else
        BrightChkValue = 1
        BrightPic.Picture = LoadPicture(GrhMapPath & "bright.bmp")
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
        CharsPic.Picture = LoadPicture(GrhMapPath & "shownpcg.bmp")
    Else
        CharsChkValue = 1
        CharsPic.Picture = LoadPicture(GrhMapPath & "shownpc.bmp")
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
        GridPic.Picture = LoadPicture(GrhMapPath & "gridg.bmp")
    Else
        GridChkValue = 1
        GridPic.Picture = LoadPicture(GrhMapPath & "grid.bmp")
    End If

End Sub

Private Sub InfoPic_Click()

    If InfoChkValue = 1 Then
        InfoChkValue = 0
        InfoPic.Picture = LoadPicture(GrhMapPath & "infog.bmp")
    Else
        InfoChkValue = 1
        InfoPic.Picture = LoadPicture(GrhMapPath & "info.bmp")
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

Private Sub ToolbarPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

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
        WeatherPic.Picture = LoadPicture(GrhMapPath & "weatherg.bmp")
    Else
        WeatherChkValue = 1
        WeatherPic.Picture = LoadPicture(GrhMapPath & "weather.bmp")
    End If

End Sub

Private Sub WeatherPic_DblClick()
    WeatherPic_Click
End Sub
