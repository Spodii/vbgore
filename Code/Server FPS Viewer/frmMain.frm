VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "vbGORE Server FPS / Population Viewer"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Graph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox EndTxt 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   4680
         TabIndex        =   6
         Text            =   "1"
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox FPSChk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "FPS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox UsersChk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Users"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox NPCsChk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NPCs"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox StartTxt 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   3360
         TabIndex        =   1
         Text            =   "1"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End:"
         Height          =   195
         Left            =   4200
         TabIndex        =   7
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start:"
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumData As Long
Private Type ServerFPS
    FPS As Long         'FPS
    Users As Integer    'Number of users
    NPCs As Integer     'Number of NPCs
End Type
Private Data() As ServerFPS

Private MaxFPS As Long
Private MaxUsers As Long
Private MaxNPCs As Long

Private MouseX As Long
Private MouseY As Long

Private gWidth As Long

Private Sub EndTxt_Change()

    If Val(EndTxt.Text) < Val(StartTxt.Text) + 1 Then EndTxt.Text = Val(StartTxt.Text) + 1
    If Val(EndTxt.Text) > NumData Then EndTxt.Text = NumData
    DrawData

End Sub

Private Sub Form_Load()

    'Load the file paths
    InitFilePaths

    'Load the FPS data
    LoadData

    'Draw the data
    DrawData

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Control
    
    For Each c In Me
        If TypeName(c) = "cButton" Then
            c.Refresh
            c.DrawState = 0
        End If
    Next c
    Set c = Nothing
    
End Sub

Private Sub DrawData()

On Error Resume Next

Const Start As Long = 25
Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long
Dim i As Long
Dim j As Long
Dim b As Byte

    'Calculate the width
    gWidth = Graph.ScaleWidth / (Val(EndTxt.Text) - Val(StartTxt.Text))
    
    'Clear the old graph
    Graph.Cls
    
    'Draw the grid
    For i = 1 To (MaxFPS / 10)
        
        'Draw the line
        Graph.ForeColor = &HC0C0C0
        j = Graph.ScaleHeight - (Graph.ScaleHeight * (i / (MaxFPS / 10)))
        Graph.Line (Graph.ScaleWidth, j)-(0, j)
        
        'Draw the text
        If FPSChk.Value = 1 Then
            Graph.ForeColor = vbBlack
            Graph.Print i * Round(MaxFPS / 10)
        End If
        If UsersChk.Value = 1 Then
            Graph.ForeColor = vbGreen
            Graph.Print i * Round(MaxUsers / 10)
        End If
        If NPCsChk.Value = 1 Then
            Graph.ForeColor = vbRed
            Graph.Print i * Round(MaxNPCs / 10)
        End If
    Next i
    
    '*** FPS ***
    b = 0
    If FPSChk.Value = 1 Then
    
        'Loop through all the data
        For i = 1 To Val(EndTxt.Text)
            Graph.ForeColor = vbBlack
            
            If i > 1 Then
                
                'Draw the line
                
                X1 = (i - 1) * gWidth + Start
                Y1 = Graph.ScaleHeight - (Graph.ScaleHeight * (Data(i - 1).FPS / MaxFPS))
                X2 = i * gWidth + Start
                Y2 = Graph.ScaleHeight - (Graph.ScaleHeight * (Data(i).FPS / MaxFPS))
                Graph.Line (X1, Y1)-(X2, Y2)
                Graph.Circle (X2, Y2), 1
            End If

            'Draw the text
            If b = 0 Then
                If Abs(X2 - MouseX) < 4 Then
                    If Abs(Y2 - MouseY) < 4 Then
                        Graph.Font.Size = 16
                        Graph.Font.Bold = True
                        Graph.Circle (X2, Y2), 2
                        Graph.Print Data(i).FPS
                        Graph.Font.Size = 8
                        Graph.Font.Bold = False
                        b = 1
                    End If
                End If
            End If
            
        Next i
        
    End If
    
    '*** Users ***
    b = 0
    If UsersChk.Value = 1 Then
        For i = 1 To Val(EndTxt.Text)
            Graph.ForeColor = vbGreen
            
            If i > 1 Then
                X1 = (i - 1) * gWidth + Start
                Y1 = Graph.ScaleHeight - (Graph.ScaleHeight * (Data(i - 1).Users / MaxUsers))
                X2 = i * gWidth + Start
                Y2 = Graph.ScaleHeight - (Graph.ScaleHeight * (Data(i).Users / MaxUsers))
                Graph.Line (X1, Y1)-(X2, Y2)
                Graph.Circle (X2, Y2), 1
            End If
            
            If b = 0 Then
                If Abs(X2 - MouseX) < 4 Then
                    If Abs(Y2 - MouseY) < 4 Then
                        Graph.Font.Size = 16
                        Graph.Font.Bold = True
                        Graph.Circle (X2, Y2), 2
                        Graph.Print Data(i).Users
                        Graph.Font.Size = 8
                        Graph.Font.Bold = False
                        b = 1
                    End If
                End If
            End If
            
        Next i
    End If
    
    '*** NPCs ***
    b = 0
    If NPCsChk.Value = 1 Then
        For i = 1 To Val(EndTxt.Text)
            Graph.ForeColor = vbRed
            
            If i > 1 Then
                X1 = (i - 1) * gWidth + Start
                Y1 = Graph.ScaleHeight - (Graph.ScaleHeight * (Data(i - 1).NPCs / MaxNPCs))
                X2 = i * gWidth + Start
                Y2 = Graph.ScaleHeight - (Graph.ScaleHeight * (Data(i).NPCs / MaxNPCs))
                Graph.Line (X1, Y1)-(X2, Y2)
                Graph.Circle (X2, Y2), 1
            End If

            If b = 0 Then
                If Abs(X2 - MouseX) < 4 Then
                    If Abs(Y2 - MouseY) < 4 Then
                        Graph.Font.Size = 16
                        Graph.Font.Bold = True
                        Graph.Circle (X2, Y2), 2
                        Graph.Print Data(i).NPCs
                        Graph.Font.Size = 8
                        Graph.Font.Bold = False
                        b = 1
                    End If
                End If
            End If
            
        Next i
    End If
    
End Sub

Private Sub LoadData()
Dim FileNum As Byte
Dim i As Long

    'On Error GoTo ErrOut

    'Get the data
    FileNum = FreeFile
    If Len(Command$) < 4 Then
        Open LogPath & "1\serverfps.txt" For Binary As #FileNum
    Else
        Open Mid$(Command$, 2, Len(Command$) - 2) For Binary As #FileNum
    End If
        Get #FileNum, , NumData
        If NumData = 0 Then GoTo ErrOut
        ReDim Data(1 To NumData)
        ReDim ClickArea(1 To NumData)
        Get #FileNum, , Data()
    Close #FileNum
    
    'Get the maxes
    MaxFPS = 0
    MaxUsers = 0
    MaxNPCs = 0
    For i = 1 To NumData
        If MaxFPS < Data(i).FPS Then MaxFPS = Data(i).FPS
        If MaxUsers < Data(i).Users Then MaxUsers = Data(i).Users
        If MaxNPCs < Data(i).NPCs Then MaxNPCs = Data(i).NPCs
    Next i
    MaxFPS = MaxFPS + 5
    MaxUsers = MaxUsers + 5
    MaxNPCs = MaxNPCs + 5
    EndTxt.Text = NumData
    
    Exit Sub
    
ErrOut:

    MsgBox "Error loading the server FPS logs!", vbOKOnly
    Unload Me
    End
        
End Sub

Private Sub Form_Resize()

    Graph.Width = Me.ScaleWidth
    Graph.Height = Me.ScaleHeight
    DrawData

End Sub

Private Sub FPSChk_Click()

    DrawData

End Sub

Private Sub Graph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Control
    
    For Each c In Me
        If TypeName(c) = "cButton" Then
            c.Refresh
            c.DrawState = 0
        End If
    Next c
    Set c = Nothing
    
    MouseX = X
    MouseY = Y
    DrawData

End Sub

Private Sub NPCsChk_Click()

    DrawData

End Sub

Private Sub StartTxt_Change()

    If Val(StartTxt.Text) < 0 Then StartTxt.Text = 0
    If Val(StartTxt.Text) - 2 > NumData Then StartTxt.Text = NumData - 2
    DrawData

End Sub

Private Sub UsersChk_Click()

    DrawData

End Sub
