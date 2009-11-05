VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mail Editor"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   15
      Left            =   6000
      TabIndex        =   33
      Text            =   "0"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   14
      Left            =   4680
      TabIndex        =   31
      Text            =   "0"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   13
      Left            =   3360
      TabIndex        =   29
      Text            =   "0"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   27
      Text            =   "0"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   11
      Left            =   720
      TabIndex        =   25
      Text            =   "0"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   10
      Left            =   6000
      TabIndex        =   23
      Text            =   "0"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   21
      Text            =   "0"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   8
      Left            =   3360
      TabIndex        =   19
      Text            =   "0"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   17
      Text            =   "0"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   15
      Text            =   "0"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   5
      Left            =   6000
      TabIndex        =   13
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   11
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   5400
      TabIndex        =   32
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   4080
      TabIndex        =   30
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   2760
      TabIndex        =   28
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   1440
      TabIndex        =   26
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   24
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   5400
      TabIndex        =   22
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   4080
      TabIndex        =   20
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   2760
      TabIndex        =   18
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1440
      TabIndex        =   16
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5400
      TabIndex        =   12
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   10
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox ItemTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommDlg 
      Left            =   2520
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox NewTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      MaxLength       =   1
      TabIndex        =   34
      Text            =   "0"
      Top             =   6480
      Width           =   255
   End
   Begin VB.TextBox DateTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   2160
      Width           =   6375
   End
   Begin VB.TextBox WriterTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   1560
      Width           =   6375
   End
   Begin VB.TextBox SubjectTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   6375
   End
   Begin VB.TextBox MessageTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label LoadCmd 
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
      Left            =   5880
      TabIndex        =   43
      Top             =   6480
      Width           =   540
   End
   Begin VB.Label SaveCmd 
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
      Left            =   5040
      TabIndex        =   42
      Top             =   6480
      Width           =   555
   End
   Begin VB.Label Command1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   4080
      TabIndex        =   41
      Top             =   6480
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Items:"
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
      Left            =   120
      TabIndex        =   40
      Top             =   5040
      Width           =   630
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Is New:"
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
      Left            =   120
      TabIndex        =   39
      Top             =   6480
      Width           =   765
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recieve Date:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   38
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Writer Name:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   720
      Width           =   855
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   2520
      Width           =   1035
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

Private Sub AmountTxt_Change(Index As Integer)

    If Index > MaxMailObjs Then Exit Sub
    MailData.Obj(Index).Amount = AmountTxt(Index).Text

End Sub

Private Sub Command1_Click()

Dim EmptyMail As MailData

    On Error GoTo ErrOut

    'Confirm
    If MsgBox("Are you sure you wish to deleted the selected mail?", vbYesNo) = vbNo Then Exit Sub

    'Clear
    MailData = EmptyMail
    FillInInformation
    Kill FilePath

Exit Sub

    'Error
ErrOut:
    MsgBox "Error deleting the mail!", vbOKOnly

End Sub

Private Sub DateTxt_Change()

    MailData.RecieveDate = CDate(DateTxt.Text)

End Sub

Private Sub ItemTxt_Change(Index As Integer)

    If Index > MaxMailObjs Then Exit Sub
    MailData.Obj(Index).ObjIndex = ItemTxt(Index).Text

End Sub

Private Sub LoadCmd_Click()

'Prepare common dialog to show existing .chr files to load

    With CommDlg
        .Filter = "Mail Files|*.mail"
        .DialogTitle = "Load mail file"
        .FileName = ""
        .Flags = cdlOFNFileMustExist
        .InitDir = App.Path & "\Mail\"
        .ShowOpen
    End With

    'Get the file path
    FilePath = CommDlg.FileName

    'Check for valid path
    If FilePath = "" Then Exit Sub
    If Right$(FilePath, 5) <> ".mail" Then Exit Sub

    'Open the character file
    LoadMail FilePath, MailData

    'Fill in all the information
    FillInInformation

End Sub

Private Sub MessageTxt_Change()

    MailData.Message = MessageTxt.Text

End Sub

Private Sub NewTxt_Change()

    MailData.New = NewTxt.Text

End Sub

Private Sub SaveCmd_Click()

'Confirm

    If MsgBox("Are you sure you wish to save the mail?" & vbCrLf & "Changes are irreverseable!", vbYesNo) = vbNo Then Exit Sub

    'Save the mail
    SaveMail FilePath, MailData

    'Done
    MsgBox "Mail saved successfully."

End Sub

Private Sub SubjectTxt_Change()

    MailData.Subject = SubjectTxt.Text

End Sub

Private Sub WriterTxt_Change()

    MailData.WriterName = WriterTxt.Text

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

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:56)  Decl: 57  Code: 130  Total: 187 Lines
':) CommentOnly: 70 (37.4%)  Commented: 0 (0%)  Empty: 47 (25.1%)  Max Logic Depth: 2
