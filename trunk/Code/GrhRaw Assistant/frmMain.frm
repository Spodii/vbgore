VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "GrhRaw Assistant"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BackBufferPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3360
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      Caption         =   "Options"
      Height          =   1695
      Left            =   9600
      TabIndex        =   33
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton OpenCmd 
         Caption         =   "Open GrhRaw.txt"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Open the GrhRaw.txt file"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton ViewCmd 
         Caption         =   "View Grhs"
         Height          =   315
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "View the new Grh entries in Notepad"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton AppendCmd 
         Caption         =   "Append"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Append the Grh values directly to GrhRaw.txt at the end"
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      Caption         =   "Grid-Created Grhs"
      Height          =   3015
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   3015
      Begin VB.ListBox GrhLst 
         Height          =   2595
         ItemData        =   "frmMain.frx":17D2A
         Left            =   120
         List            =   "frmMain.frx":17D31
         TabIndex        =   29
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Existing Grhs"
      Height          =   2775
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   3015
      Begin VB.ListBox OldGrhLst 
         Height          =   2400
         ItemData        =   "frmMain.frx":17D3D
         Left            =   120
         List            =   "frmMain.frx":17D3F
         TabIndex        =   28
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Texture Information"
      Height          =   1695
      Left            =   7800
      TabIndex        =   14
      Top             =   120
      Width           =   1695
      Begin VB.TextBox MaxColumnsTxt 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         ToolTipText     =   "Maximum columns with the Start Y and Grid Height in consideration"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox MaxRowsTxt 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         ToolTipText     =   "Maximum rows with the Start X and Grid Width in consideration"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TexHeightTxt 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         ToolTipText     =   "Width of each grid segment in pixels"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TexWidthTxt 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         ToolTipText     =   "Width of each grid segment in pixels"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Columns:"
         Height          =   195
         Left            =   45
         TabIndex        =   23
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rows:"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Grid and Numbering"
      Height          =   1695
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox GridChk 
         BackColor       =   &H80000005&
         Caption         =   "Grid"
         Height          =   195
         Left            =   3720
         TabIndex        =   37
         ToolTipText     =   "Display the grid"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.TextBox TexturePathTxt 
         Height          =   285
         Left            =   720
         TabIndex        =   31
         ToolTipText     =   "File path to the texture to work on (must be PNG format)"
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton BrowseCmd 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         ToolTipText     =   "Browse to the texture file"
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox FreeOnlyChk 
         BackColor       =   &H80000005&
         Caption         =   "Free Grhs Only"
         Height          =   195
         Left            =   2160
         TabIndex        =   17
         ToolTipText     =   "Tick to use only Grh values not already used in GrhRaw.txt"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox StartGrhTxt 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Text            =   "1"
         ToolTipText     =   "The Grh number to start the numbering at"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox RowsTxt 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Text            =   "-1"
         ToolTipText     =   "Number of rows to use (-1 for as many as possible)"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox ColumnsTxt 
         Height          =   285
         Left            =   3720
         TabIndex        =   10
         Text            =   "-1"
         ToolTipText     =   "Number of columns to use (-1 for as many as possible)"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox StartYTxt 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Pixel Y co-ordinate to start at (first pixel starts at 0)"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox StartXTxt 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Pixel X co-ordinate to start at (first pixel starts at 0)"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox GridHeightTxt 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Text            =   "32"
         ToolTipText     =   "Height of each grid segment in pixels"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox GridWidthTxt 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "32"
         ToolTipText     =   "Width of each grid segment in pixels"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Texture:"
         Height          =   195
         Left            =   0
         TabIndex        =   32
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Grh:"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rows:"
         Height          =   195
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Columns:"
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Y:"
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start X:"
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   3240
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   535
      TabIndex        =   0
      Top             =   1920
      Width           =   8055
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   -240
      Top             =   -240
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

Private Sub AppendCmd_Click()
Dim FileNum As Byte
Dim s As String

    'Confirm
    If MsgBox("Are you sure you wish to append this Grh list to the end of GrhRaw.txt?", vbYesNo) = vbNo Then Exit Sub
    
    'Open the file
    FileNum = FreeFile
    Open Data2Path & "GrhRaw.txt" For Binary Access Write As #FileNum

        'Add the text
        s = vbNewLine & BuildGrhString  'We must hold it in a string first, since calling from the function makes strange results
        Put #FileNum, LOF(FileNum) + 1, s
        
    Close #FileNum
        
End Sub

Private Sub BrowseCmd_Click()

    'Bring up the common dialog browse
    With CD
        .Filter = "Portable Network Graphics (PNG)|*.png"
        .DialogTitle = "Load"
        .FileName = vbNullString
        .InitDir = TexturePathTxt.Text
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        .ShowOpen
        TexturePathTxt.Text = .FileName
    End With
    
    'Scroll the text box to the right (so you can see the file name)
    TexturePathTxt.SelStart = Len(TexturePathTxt.Text)

End Sub

Private Sub Form_Load()
Dim s() As String
Dim i As Long
Dim p As String

    'Set the default path
    s = Split(App.Path, "\")
    For i = 0 To UBound(s) - 2
        p = p & s(i) & "\"
    Next i
    TexturePathTxt.Text = GrhPath & "27.png"
    TexturePathTxt.SelStart = Len(TexturePathTxt.Text)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Make sure the frmView is closed, too
    Unload frmView

End Sub

Private Sub Form_Resize()

    'Keep the form above a specific size
    If Me.Width < 11550 Then Me.Width = 11550
    If Me.Height < 4530 Then Me.Height = 4530
    
    'Scale the picturebox
    PreviewPic.Width = Me.Width - PreviewPic.Left - (160 * 2) + 64
    PreviewPic.Height = Me.Height - PreviewPic.Top - (160 * 4)
    
    'Scale the old grh list frame and list box
    If Me.Height - 4235 > 600 Then
        Frame3.Visible = True
        OldGrhLst.Visible = True
        Frame3.Height = Me.Height - 4235
        If Frame3.Height > 2775 Then Frame3.Height = 2775
        OldGrhLst.Height = Frame3.Height - 320
    Else
        Frame3.Visible = False
        OldGrhLst.Visible = False
    End If
    
    'Scale the new grh list frame and list box
    Frame4.Top = IIf(Frame3.Visible, Frame3.Top + Frame3.Height, -120) + 240
    Frame4.Height = Me.Height - Frame4.Top - (160 * 4) + 16
    GrhLst.Height = Frame4.Height - 320

End Sub

Private Sub OldGrhLst_KeyDown(KeyCode As Integer, Shift As Integer)

    'Deselect the selected item
    If KeyCode = vbKeyEscape Then
        If OldGrhLst.ListIndex >= 0 Then OldGrhLst.Selected(GrhLst.ListIndex) = False
    End If
    
    'Check to move up or down
    If KeyCode = vbKeyDown Then
        If OldGrhLst.ListIndex >= OldGrhLst.ListCount - 1 Then Exit Sub
        OldGrhLst.ListIndex = OldGrhLst.ListIndex + 1
        KeyCode = 0
        RefreshImage False
    End If
    If KeyCode = vbKeyUp Then
        If OldGrhLst.ListIndex <= 0 Then Exit Sub
        OldGrhLst.ListIndex = OldGrhLst.ListIndex - 1
        KeyCode = 0
        RefreshImage False
    End If
    
    'Check to refresh
    If LastOldGrhLstIndex <> OldGrhLst.ListIndex Then RefreshImage False

End Sub

Private Sub GrhLst_KeyDown(KeyCode As Integer, Shift As Integer)

    'Deselect the selected item
    If KeyCode = vbKeyEscape Then
        If GrhLst.ListIndex >= 0 Then GrhLst.Selected(GrhLst.ListIndex) = False
    End If
    
    'Check to move up or down
    If KeyCode = vbKeyDown Then
        If GrhLst.ListIndex >= GrhLst.ListCount - 1 Then Exit Sub
        GrhLst.ListIndex = GrhLst.ListIndex + 1
        KeyCode = 0
        RefreshImage False
    End If
    If KeyCode = vbKeyUp Then
        If GrhLst.ListIndex <= 0 Then Exit Sub
        GrhLst.ListIndex = GrhLst.ListIndex - 1
        KeyCode = 0
        RefreshImage False
    End If
    
    'Check to refresh
    If LastGrhLstIndex <> GrhLst.ListIndex Then RefreshImage False

End Sub

Private Sub GrhLst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Deselect the selected item
    If Button = vbRightButton Then
        If GrhLst.ListIndex >= 0 Then GrhLst.Selected(GrhLst.ListIndex) = False
    End If
    
    'Check to refresh
    If LastGrhLstIndex <> GrhLst.ListIndex Then RefreshImage False
    
End Sub

Private Sub OldGrhLst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Deselect the selected item
    If Button = vbRightButton Then
        If OldGrhLst.ListIndex >= 0 Then OldGrhLst.Selected(OldGrhLst.ListIndex) = False
    End If
    
    'Check to refresh
    If LastOldGrhLstIndex <> OldGrhLst.ListIndex Then RefreshImage False
    
End Sub

Private Sub GridChk_Click()

    'Refresh the image
    RefreshImage
    
End Sub

Private Sub GridHeightTxt_Change()

    'Update the maximum number of rows and columns
    If Val(GridHeightTxt.Text) > 1 Then UpdateMaxRowsColumns
    
End Sub

Private Sub GridWidthTxt_Change()

    'Update the maximum number of rows and columns
    If Val(GridWidthTxt.Text) > 1 Then UpdateMaxRowsColumns

End Sub

Private Sub OpenCmd_Click()

    ShellExecute Me.hwnd, "open", Data2Path & "GrhRaw.txt", vbNullString, vbNullString, 1

End Sub

Private Sub PreviewPic_KeyDown(KeyCode As Integer, Shift As Integer)

    'Deselect the selected item
    If KeyCode = vbKeyEscape Then
        If GrhLst.ListIndex >= 0 Then
            GrhLst.Selected(GrhLst.ListIndex) = False
            RefreshImage False
        End If
    End If

End Sub

Private Sub PreviewPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim GridX As Long
Dim GridY As Long
Dim Columns As Long
Dim Rows As Long
Dim Index As Long

    'Select a tile
    If Button = vbLeftButton Then
    
        'Find the grid position
        GridX = (x - Val(StartXTxt.Text)) \ (Val(GridWidthTxt.Text))
        GridY = (y - Val(StartYTxt.Text)) \ (Val(GridHeightTxt.Text))
        
        'Find the number of rows and columns
        Rows = Val(RowsTxt.Text)
        Columns = Val(ColumnsTxt.Text)
        If Rows = -1 Then Rows = Val(MaxRowsTxt.Text)
        If Rows <= 0 Then Exit Sub
        If Columns = -1 Then Columns = Val(MaxColumnsTxt.Text)
        If Columns <= 0 Then Exit Sub
        
        'Check if in range of the image
        If GridX > Rows Then Exit Sub
        If GridY >= Columns Then Exit Sub
        
        'Convert the grid position to list index
        Index = GridX + (Rows * GridY)
        
        'Make sure the index is valid
        If Index >= GrhLst.ListCount Then Exit Sub
        
        'Set the list index
        GrhLst.ListIndex = Index
        
        'Redraw with the selection
        If LastGrhLstIndex <> GrhLst.ListIndex Then RefreshImage False
        
    'Deselect the selected item
    ElseIf Button = vbRightButton Then
    
        If GrhLst.ListIndex >= 0 Then
            GrhLst.Selected(GrhLst.ListIndex) = False
            RefreshImage False
        End If
        
    End If
    
End Sub

Private Sub StartXTxt_Change()

    'Update the maximum number of rows and columns
    UpdateMaxRowsColumns
    
End Sub

Private Sub StartYTxt_Change()

    'Update the maximum number of rows and columns
    UpdateMaxRowsColumns
    
End Sub

Private Sub TexturePathTxt_Change()

    'Check for a valid length
    If Len(TexturePathTxt.Text) < 4 Then Exit Sub
    
    'Check if the PNG suffix is there
    If LCase$(Right$(TexturePathTxt.Text, 4)) <> ".png" Then Exit Sub

    'Check if the file exists
    If Not FileExist(TexturePathTxt.Text, vbNormal) Then Exit Sub
    
    'Try to display the file
    LoadTexture TexturePathTxt.Text
    
End Sub

Private Sub ViewCmd_Click()

    'Show the grh list view form
    Load frmView
    frmView.Show
    
    'Display the text
    frmView.GrhTxt.Text = BuildGrhString

End Sub
