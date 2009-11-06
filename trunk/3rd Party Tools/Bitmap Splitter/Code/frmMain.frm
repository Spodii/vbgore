VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Bitmap Splitter"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BackPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2520
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton SplitCmd 
      Caption         =   "Split"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Source Bitmap"
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
      Begin VB.PictureBox ScreenPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   120
         ScaleHeight     =   95
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   840
         Width           =   3855
      End
      Begin VB.CommandButton OpenCmd 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox FileTxt 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Preview:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Selection Information"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox HTxt 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Text            =   "24"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WTxt 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Text            =   "24"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox YTxt 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox XTxt 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start X:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Form_Load()

    FileTxt.Text = App.Path

End Sub

Private Sub OpenCmd_Click()

    'Load map
    With frmMain.CD
        .Filter = "Bitmap (BMP)|*.bmp"
        .DialogTitle = "Load"
        .FileName = vbNullString
        .InitDir = App.Path & "\"
        .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        .ShowOpen
        FileTxt.Text = .FileName
    End With

End Sub

Private Sub SplitCmd_Click()
Dim ImgInfo As CImageInfo
Dim FileName As String
Dim s() As String
Dim x As Long
Dim y As Long

    'Confirm a valid file
    If Dir$(FileTxt.Text, vbNormal) = vbNullString Then
        MsgBox "File not found!", vbOKOnly
        Exit Sub
    End If

    'Get the image info
    Set ImgInfo = New CImageInfo
    ImgInfo.ReadImageInfo FileTxt.Text
    If ImgInfo.Width <= 1 Or ImgInfo.Height <= 1 Then
        MsgBox "Invalid image!", vbOKOnly
        Exit Sub
    End If
    
    'Check for valid size text
    If Not IsNumeric(XTxt.Text) Then
        MsgBox "Invalid X value!", vbOKOnly
        Exit Sub
    End If
    If Not IsNumeric(YTxt.Text) Then
        MsgBox "Invalid Y value!", vbOKOnly
        Exit Sub
    End If
    If Not IsNumeric(WTxt.Text) Then
        MsgBox "Invalid Width value!", vbOKOnly
        Exit Sub
    End If
    If Not IsNumeric(HTxt.Text) Then
        MsgBox "Invalid Height value!", vbOKOnly
        Exit Sub
    End If
    
    'Get the file name
    s = Split(FileTxt.Text, "\")
    s = Split(s(UBound(s)), ".")
    FileName = s(0)
    
    'Create the directory
    MakeSureDirectoryPathExists App.Path & "\" & FileName & "\"
    
    'Load the start picture
    BackPic.Width = ImgInfo.Width + 20
    BackPic.Height = ImgInfo.Height + 20
    BackPic.Picture = LoadPicture(FileTxt.Text)
    
    'Resize the screen
    ScreenPic.Width = Val(WTxt.Text) * Screen.TwipsPerPixelX
    ScreenPic.Height = Val(HTxt.Text) * Screen.TwipsPerPixelY
    
    'Loop through the image
    For y = Val(YTxt.Text) To (ImgInfo.Height - YTxt.Text) \ Val(HTxt.Text)
        For x = Val(XTxt.Text) To (ImgInfo.Width - XTxt.Text) \ Val(WTxt.Text)
            ScreenPic.Cls
            BitBlt ScreenPic.hDC, 0, 0, Val(WTxt.Text), Val(HTxt.Text), BackPic.hDC, x * Val(WTxt.Text), y * Val(HTxt.Text), vbSrcCopy
            SavePicture ScreenPic.Image, App.Path & "\" & FileName & "\y" & Format(y, "00") & "x" & Format(x, "00") & ".bmp"
        Next x
    Next y


End Sub
