VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Map Capture Tool"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   6480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select Map"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Map"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox ScreenPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   9000
      Left            =   120
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   796
      TabIndex        =   0
      Top             =   600
      Width           =   12000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Command2_Click()
Dim x As Long
Dim y As Long
Dim o As Long

    For y = 1 To MapInfo.Height \ (600 \ 32) + 1
        For x = 1 To MapInfo.Width \ (800 \ 32) + 1
            DrawScreen (x - 1) * (800 \ 32), (y - 1) * (600 \ 32)
            SaveBackBuffer App.Path & "\" & y & "x" & x & ".bmp"
        Next x
    Next y
    
    ScreenPic.Cls
    ScreenPic.Width = (MapInfo.Width * 32 * Screen.TwipsPerPixelX) + 64
    ScreenPic.Height = (MapInfo.Height * 32 * Screen.TwipsPerPixelY) + 64
    For y = 1 To MapInfo.Height \ (600 \ 32) + 1
        For x = 1 To MapInfo.Width \ (800 \ 32) + 1
            Me.Picture = LoadPicture(App.Path & "\" & y & "x" & x & ".bmp")
            BitBlt ScreenPic.hDC, (x - 1) * 800, (y - 1) * (600 - 32), 800, 600 - 32, Me.hDC, 0, 0, vbSrcCopy
            Me.Cls
            Kill App.Path & "\" & y & "x" & x & ".bmp"
        Next x
    Next y
    Me.Picture = Nothing
    Me.Refresh
    SavePicture ScreenPic.Image, App.Path & "\map" & MapNum2 & ".bmp"
    ScreenPic.Width = 800 * Screen.TwipsPerPixelX
    ScreenPic.Height = 600 * Screen.TwipsPerPixelY
    MsgBox "Map successfully saved!", vbOKOnly

End Sub

Private Sub Command3_Click()
Dim FileName As String

    'Load map
    With frmMain.CD
        .Filter = "Maps|*.map"
        .DialogTitle = "Load"
        .FileName = vbNullString
        .InitDir = GetRootPath & "Maps\"
        .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        .ShowOpen
    End With
    FileName = Right$(frmMain.CD.FileName, Len(frmMain.CD.FileName) - Len(GetRootPath & "Maps\"))
    Game_Map_Switch CInt(Left$(FileName, Len(FileName) - 4))

End Sub

Private Sub Form_Load()
    
    Game_Map_Switch 1
    ScreenPic.Width = 800 * 15
    ScreenPic.Height = 600 * 15

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    UnloadProject

End Sub
