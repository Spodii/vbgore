VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Compress"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox DataList 
      Height          =   3765
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   6375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client Graphic Files"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton GfxD 
         Caption         =   "Decompress"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox ClientGfxPath 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "C:\Programming\LoM\Client\Grh"
         Top             =   240
         Width           =   6135
      End
      Begin VB.CommandButton GfxC 
         Caption         =   "Compress"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton GfxK 
         Caption         =   "Kill *.cgf's"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton GfxF 
         Caption         =   "Kill *.bmp's"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Map Files"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton MapD 
         Caption         =   "Decompress"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton MapF 
         Caption         =   "Kill *.map's"
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton MapK 
         Caption         =   "Kill *.cmf's"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton MapC 
         Caption         =   "Compress"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox ServerMapPathTxt 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\Programming\LoM\Server\Maps"
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileList() As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Private Sub Form_Load()

    MsgBox "This program has lost it's use. Use it only for a reference - it is not a useful part of vbGORE anymore.", vbOKOnly
    timeBeginPeriod 1

End Sub

Private Sub GetFiles(FilePath As String, Extension As String)

Dim TempList() As String
Dim i As Integer
Dim k As Integer
Dim l As Integer

    ReDim TempList(0)
    ReDim FileList(0)
    TempList = AllFilesInFolders(FilePath, False)
    k = UBound(TempList)
    For i = 0 To k
        If UCase$(Right$(TempList(i), 3)) = UCase$(Extension) Then
            l = l + 1
            ReDim Preserve FileList(l)
            FileList(l) = TempList(i)
        End If
    Next i

End Sub

Private Sub GfxC_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ClientGfxPath.Text, "bmp"
    k = UBound(FileList)
    Me.Caption = "Compressing Graphic File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Compression_Compress FileList(i), Left$(FileList(i), Len(FileList(i)) - 3) & "cgf", RLE
        DataList.AddItem "Compressed: " & Left$(FileList(i), Len(FileList(i)) - 3) & "cgf (" & (timeGetTime - a) & "ms)"
        Me.Caption = "Compressing Graphic File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "Compressing Graphic Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub GfxD_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ClientGfxPath.Text, "cgf"
    k = UBound(FileList)
    Me.Caption = "DeCompressing Graphic File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Compression_DeCompress FileList(i), Left$(FileList(i), Len(FileList(i)) - 3) & "bmp", RLE
        DataList.AddItem "DeCompressed: " & Left$(FileList(i), Len(FileList(i)) - 3) & "bmp (" & (timeGetTime - a) & "ms)"
        Me.Caption = "DeCompressing Graphic File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "DeCompressing Graphic Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub GfxF_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ClientGfxPath.Text, "bmp"
    k = UBound(FileList)
    Me.Caption = "Deleting *.bmp File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Kill FileList(i)
        DataList.AddItem "Deleted: " & Left$(FileList(i), Len(FileList(i)) - 3) & "bmp (" & (timeGetTime - a) & "ms)"
        Me.Caption = "Deleting *.bmp File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "Deleting *.bmp Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub Gfxk_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ClientGfxPath.Text, "cgf"
    k = UBound(FileList)
    Me.Caption = "Deleting *.cgf File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Kill FileList(i)
        DataList.AddItem "Deleted: " & Left$(FileList(i), Len(FileList(i)) - 3) & "cgf (" & (timeGetTime - a) & "ms)"
        Me.Caption = "Deleting *.cgf File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "Deleting *.cgf Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub MapC_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ServerMapPathTxt.Text, "map"
    k = UBound(FileList)
    Me.Caption = "Compressing Map File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Compression_Compress FileList(i), Left$(FileList(i), Len(FileList(i)) - 3) & "cmf", LZW
        DataList.AddItem "Compressed: " & Left$(FileList(i), Len(FileList(i)) - 3) & "cmf (" & (timeGetTime - a) & "ms)"
        Me.Caption = "Compressing Map File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "Compressing Map Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub MapD_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ServerMapPathTxt.Text, "cmf"
    k = UBound(FileList)
    Me.Caption = "DeCompressing Map File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Compression_DeCompress FileList(i), Left$(FileList(i), Len(FileList(i)) - 3) & "map", LZW
        DataList.AddItem "DeCompressed: " & Left$(FileList(i), Len(FileList(i)) - 3) & "map (" & (timeGetTime - a) & "ms)"
        Me.Caption = "DeCompressing Map File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "DeCompressing Map Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub MapF_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ServerMapPathTxt.Text, "map"
    k = UBound(FileList)
    Me.Caption = "Deleting *.map File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Kill FileList(i)
        DataList.AddItem "Deleted: " & Left$(FileList(i), Len(FileList(i)) - 3) & "map (" & (timeGetTime - a) & "ms)"
        Me.Caption = "Deleting *.map File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "Deleting *.map Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

Private Sub MapK_Click()

Dim i As Integer
Dim k As Integer
Dim a As Long
Dim b As Long

    DataList.Clear
    GetFiles ServerMapPathTxt.Text, "cmf"
    k = UBound(FileList)
    Me.Caption = "Deleting *.cmf File 0 of " & k & " (0%)"
    Me.Refresh
    b = timeGetTime
    For i = 1 To k
        a = timeGetTime
        Kill FileList(i)
        DataList.AddItem "Deleted: " & Left$(FileList(i), Len(FileList(i)) - 3) & "cmf (" & (timeGetTime - a) & "ms)"
        Me.Caption = "Deleting *.cmf File " & i & " of " & k & " (" & (Round(i / k, 2) * 100) & "%)"
        DataList.Refresh
        Me.Refresh
        DoEvents
    Next i
    Me.Caption = "Deleting *.cmf Files Complete (" & (timeGetTime - b) & "ms)"

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:50)  Decl: 60  Code: 238  Total: 298 Lines
':) CommentOnly: 55 (18.5%)  Commented: 0 (0%)  Empty: 40 (13.4%)  Max Logic Depth: 3
