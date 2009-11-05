VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Grh3RawMaker"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8070
   Icon            =   "Grh3RawMaker.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Grh3RawMaker.frx":17D2A
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtImgFileNum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.CheckBox chkAnimated 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Animated"
      CausesValidation=   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommDlg 
      Left            =   7320
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFilePath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox txtStartGrh 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Label cmdStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make Grh3.raw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3000
      TabIndex        =   18
      Top             =   1680
      Width           =   1830
   End
   Begin VB.Label cmdBrowse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6240
      TabIndex        =   6
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label lblImgNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Img File Num:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblRows 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rows:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   16
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label lblTilesInRow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles in Row:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   15
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label lblTiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiles:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   510
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Top             =   2040
      Width           =   7815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Image File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Grh. Num:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Height:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tile Width:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FilePath As String
Private TileWidth As Long
Private FileName As String
Private TileHeight As Long
Private TilesInRow As Long
Private Rows As Long
Private Tiles As Long
Private StartGrh As Long
Private ImgFileNum As Long
Private Animated As Byte
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Sub cmdBrowse_Click()
On Error Resume Next

    TileWidth = txtWidth.Text
    TileHeight = txtHeight.Text
    StartGrh = txtStartGrh.Text

    'Let the user browse searching for image files
    With CommDlg
        .Filter = "Bitmap Images (*.bmp)|*.bmp|Png Images (*.png)|*.png"
        .DialogTitle = "Load image file"
        .FileName = ""
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End With
    
    'Get the filepath
    FilePath = CommDlg.FileName
    
    'Check to see if the path is valid
    If FilePath = "" Then Exit Sub
    
    'Set the right filepath in the textbox
    txtFilePath.Text = FilePath
    
    'Display the image information
    ReadImageInfo (txtFilePath.Text)
    lblWidth.Caption = "Width: " & ImageWidth
    lblHeight.Caption = "Height: " & ImageHeight
    lblSize.Caption = "Size: " & FileSize & " bytes."
    
    'Calculate Tiles
    TilesInRow = (ImageWidth / TileWidth)
    Rows = (ImageHeight / TileHeight)
    Tiles = (TilesInRow * Rows)
    
    lblTiles.Caption = "Tiles: " & Tiles
    lblTilesInRow.Caption = "Tiles in Row: " & TilesInRow
    lblRows.Caption = "Rows: " & Rows

End Sub

Private Sub cmdStart_Click()

On Error Resume Next

Dim i As Long
Dim StartX As Long
Dim StartY As Long
Dim CurrentTile As Long

    'Check if they have entered the correct info
    If txtHeight.Text = "" Or txtWidth.Text = "" Or txtStartGrh.Text = "" Or txtImgFileNum.Text = "" Then
        MsgBox ("Please enter the correct information!")
        Exit Sub
    End If
    
    'Check to see if they have entered a image file
    If txtFilePath.Text = "" Then
        MsgBox ("Choose an image file first!")
        Exit Sub
    End If
    
    'Set a few variables
    CurrentTile = 1
    ImgFileNum = txtImgFileNum.Text
    FileName = App.Path & "\Grh3.raw"

    StartGrh = txtStartGrh.Text
    
    f = FreeFile
    
    If Not Dir(FileName) Then
        Open FileName For Output As #f
        Close #f
    End If
    
    Open FileName For Append As #f
    
        For i = StartGrh To (TilesInRow * Rows + StartGrh - 1)
            If CurrentTile > TilesInRow Then
                CurrentTile = 1
                StartY = StartY + TileHeight
                StartX = 0
            End If
        
            Print #f, "Grh" & i & "=" & 1 & "-" & ImgFileNum & "-" & StartX & "-" & StartY & "-" & TileWidth & "-" & TileHeight
            
            StartX = StartX + TileWidth
            CurrentTile = CurrentTile + 1
        Next i
        
        MsgBox ("Grh3.raw successfully created!")
        On Error Resume Next
        Unload frmImgPrev
        Unload Me
    
    Close #f
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

    'Close form
    If Button = vbLeftButton Then
        If X >= Me.ScaleWidth - 23 Then
            If X <= Me.ScaleWidth - 10 Then
                If Y <= 26 Then
                    If Y >= 11 Then
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub
