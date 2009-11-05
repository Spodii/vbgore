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
'**       ____        _________   ______   ______  ______   _______           **
'**       \   \      /   /     \ /  ____\ /      \|      \ |   ____|          **
'**        \   \    /   /|      |  /     |        |       ||  |____           **
'***        \   \  /   / |     /| |  ___ |        |      / |   ____|         ***
'****        \   \/   /  |     \| |  \  \|        |   _  \ |  |____         ****
'******       \      /   |      |  \__|  |        |  | \  \|       |      ******
'********      \____/    |_____/ \______/ \______/|__|  \__\_______|    ********
'*******************************************************************************
'*******************************************************************************
'************ vbGORE - Visual Basic 6.0 Graphical Online RPG Engine ************
'************            Official Release: Version 0.1.2            ************
'************                 http://www.vbgore.com                 ************
'*******************************************************************************
'*******************************************************************************
'***** Source Distribution Information: ****************************************
'*******************************************************************************
'** If you wish to distribute this source code, you must distribute as-is     **
'** from the vbGORE website unless permission is given to do otherwise. This  **
'** comment block must remain in-tact in the distribution. If you wish to     **
'** distribute modified versions of vbGORE, please contact Spodi (info below) **
'** before distributing the source code. You may never label the source code  **
'** as the "Official Release" or similar unless the code and content remains  **
'** unmodified from the version downloaded from the official website.         **
'** You may also never sale the source code without permission first. If you  **
'** want to sell the code, please contact Spodi (below). This is to prevent   **
'** people from ripping off other people by selling an insignificantly        **
'** modified version of open-source code just to make a few quick bucks.      **
'*******************************************************************************
'***** Creating Engines With vbGORE: *******************************************
'*******************************************************************************
'** If you plan to create an engine with vbGORE that, please contact Spodi    **
'** before doing so. You may not sell the engine unless told elsewise (the    **
'** engine must has substantial modifications), and you may not claim it as   **
'** all your own work - credit must be given to vbGORE, along with a link to  **
'** the vbGORE homepage. Failure to gain approval from Spodi directly to      **
'** make a new engine with vbGORE will result in first a friendly reminder,   **
'** followed by much more drastic measures.                                   **
'*******************************************************************************
'***** Helping Out vbGORE: *****************************************************
'*******************************************************************************
'** If you want to help out with vbGORE's progress, theres a few things you   **
'** can do:                                                                   **
'**  *Donate - Great way to keep a free project going. :) Info and benifits   **
'**        for donating can be found at:                                      **
'**        http://www.vbgore.com/en/index.php?title=Donate                    **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        help expend the wiki pages!                                        **
'**  *Link To Us - Creating a link to vbGORE, whether it is on your own web   **
'**        page or a link to vbGORE in a forum you visit, every link helps    **
'**        spread the word of vbGORE's existance! Buttons and banners for     **
'**        linking to vbGORE can be found on the following page:              **
'**        http://www.vbgore.com/en/index.php?title=Buttons_and_Banners       **
'*******************************************************************************
'***** Conact Information: *****************************************************
'*******************************************************************************
'** Please contact the creator of vbGORE (Spodi) directly with any questions: **
'** AIM: Spodii                          Yahoo: Spodii                        **
'** MSN: Spodii@hotmail.com              Email: spodi@vbgore.com              **
'** 2nd Email: spodii@hotmail.com        Website: http://www.vbgore.com       **
'*******************************************************************************
'***** Credits: ****************************************************************
'*******************************************************************************
'** Below are credits to those who have helped with the project or who have   **
'** distributed source code which has help this project's creation. The below **
'** is listed in no particular order of significance:                         **
'**                                                                           **
'** ORE (Aaron Perkins): Used as base engine and for learning experience      **
'**   http://www.baronsoft.com/                                               **
'** SOX (Trevor Herselman): Used for all the networking                       **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=35239&lngWId=1      **
'** Compression Methods (Marco v/d Berg): Provided compression algorithms     **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1      **
'** All Files In Folder (Jorge Colaccini): Algorithm implimented into engine  **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51435&lngWId=1      **
'** Game Programming Wiki (All community): Help on many different subjects    **
'**   http://wwww.gpwiki.org/                                                 **
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

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
