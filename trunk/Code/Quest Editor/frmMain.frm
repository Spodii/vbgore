VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Quest Editor"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox RedoChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Redoable"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   52
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox IncompleteTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   360
      TabIndex        =   48
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox FRSkillTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4560
      TabIndex        =   46
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox ARSkillTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1680
      TabIndex        =   44
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox FRObjAmountTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4200
      TabIndex        =   42
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox FRObjTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4200
      TabIndex        =   40
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox ARObjAmountTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1320
      TabIndex        =   38
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox ARObjTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1320
      TabIndex        =   36
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox ARGoldTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   960
      TabIndex        =   34
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox ARExpTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   960
      TabIndex        =   33
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox FRGoldTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   3840
      TabIndex        =   29
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox FRExpTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   3840
      TabIndex        =   28
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox FANPCAmountTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4200
      TabIndex        =   27
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox FANPCTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4200
      TabIndex        =   26
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox FAObjAmountTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4200
      TabIndex        =   25
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox FAObjtxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4200
      TabIndex        =   24
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox AAAmountTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1680
      TabIndex        =   23
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox AAObjTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox AALvlTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1080
      TabIndex        =   21
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox FinishTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   2880
      Width           =   4815
   End
   Begin VB.TextBox AcceptTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox StartTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   960
      TabIndex        =   17
      Top             =   960
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2520
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LoadLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4560
      TabIndex        =   51
      Top             =   600
      Width           =   540
   End
   Begin VB.Label SaveAsLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save As"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3360
      TabIndex        =   50
      Top             =   600
      Width           =   885
   End
   Begin VB.Label SaveLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2520
      TabIndex        =   49
      Top             =   600
      Width           =   555
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incompleted Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   360
      TabIndex        =   47
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Learn Skill:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   3480
      TabIndex        =   45
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Learn Skill:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   600
      TabIndex        =   43
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   3480
      TabIndex        =   41
      Top             =   5400
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Object:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   3480
      TabIndex        =   39
      Top             =   5160
      Width           =   630
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   600
      TabIndex        =   37
      Top             =   5400
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Object:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   600
      TabIndex        =   35
      Top             =   5160
      Width           =   630
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accept Reward:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   240
      TabIndex        =   32
      Top             =   4440
      Width           =   1650
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   480
      TabIndex        =   31
      Top             =   4680
      Width           =   435
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   480
      TabIndex        =   30
      Top             =   4920
      Width           =   465
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   3360
      TabIndex        =   16
      Top             =   4920
      Width           =   465
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   3360
      TabIndex        =   15
      Top             =   4680
      Width           =   435
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Reward:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   3120
      TabIndex        =   14
      Top             =   4440
      Width           =   1545
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   960
      TabIndex        =   13
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Have Object:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   480
      TabIndex        =   12
      Top             =   3600
      Width           =   1140
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accept Requirements:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   3480
      TabIndex        =   9
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kill NPC Index:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   2880
      TabIndex        =   8
      Top             =   3840
      Width           =   1290
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3480
      TabIndex        =   7
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Object Index:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   3000
      TabIndex        =   6
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Requirements:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   3120
      Width           =   2190
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   900
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accept Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

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

Private Sub LoadLbl_Click()
Dim FileName As String
Dim TempNum As Integer
    On Error GoTo ErrOut

    'Confirm
    If MsgBox("Are you sure you wish to load another Quest?" & vbCrLf & "Any changes made to the current Quest will be lost!", vbYesNo) = vbNo Then Exit Sub
    
    'Load map
    With frmMain.CD
        .Filter = "Quests|*.quest"
        .DialogTitle = "Load"
        .FileName = ""
        .InitDir = QuestPath
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End With
    FileName = Right$(frmMain.CD.FileName, Len(frmMain.CD.FileName) - Len(QuestPath))
    Editor_LoadQuest Val(FileName)
    
ErrOut:
    
End Sub

Private Sub SaveAsLbl_Click()
Dim RetNumber As Integer

    'Confirm
    If MsgBox("Are you sure you wish to save changes to Quest " & QuestNum & " as a new number?", vbYesNo) = vbNo Then Exit Sub
    
    'Get number
    RetNumber = Val(InputBox("Please enter the number to save the Quest as."))
    If RetNumber = 0 Then Exit Sub
    
    'Check for overwrite
    If Engine_FileExist(App.Path & "\Quests\" & RetNumber & ".quest", vbNormal) Then
        If MsgBox("Quest number " & RetNumber & " already exists, are you sure you wish to overwrite it?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Save
    Editor_SaveQuest RetNumber

End Sub

Private Sub SaveLbl_Click()

    'Confirm
    If MsgBox("Are you sure you wish to save changes to quest " & QuestNum & "?", vbYesNo) = vbNo Then Exit Sub
    
    'Save changes
    Editor_SaveQuest QuestNum

End Sub
