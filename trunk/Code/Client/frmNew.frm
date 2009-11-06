VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Login"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox HeadCmb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox ClassCmb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox BodyCmb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Top             =   660
      Width           =   1275
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1080
      TabIndex        =   1
      Top             =   210
      Width           =   1635
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Head:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label CreateLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2880
      Width           =   705
   End
   Begin VB.Label CancelLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   690
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelLbl_Click()

    'Show the connect screen
    frmConnect.Visible = True
    
    'Hide this screen
    Me.Visible = False

End Sub

Private Sub CreateLbl_Click()

    'Set the variables
    UserName = NameTxt.Text
    UserPassword = PasswordTxt.Text
    UserBody = BodyCmb.ListIndex
    UserHead = HeadCmb.ListIndex
    UserClass = ClassCmb.ListIndex
    
    'Convert the body by listbox index to the body number
    Select Case UserBody
        Case 0: UserBody = 1
        Case Else: UserBody = 1
    End Select
    
    'Convert the head by listbox index to the head number
    Select Case UserHead
        Case 0: UserHead = 1
        Case Else: UserHead = 1
    End Select
    
    'Convert the class by listbox index to the class number
    Select Case UserClass
        Case 0: UserClass = ClassID.Warrior
        Case 1: UserClass = ClassID.Mage
        Case 2: UserClass = ClassID.Rogue
        Case Else: UserClass = ClassID.Warrior
    End Select
    
    'Connect
    If Game_CheckUserData Then
        SendNewChar = True
        InitSocket
    End If

End Sub

Private Sub Form_Load()

    'Load up the head, body and class values you can select
    'For the head and body, to add more, you have to edit it accordingly in the server
    ' under User_ConnectNew on this line:
    '    'Check for a valid body and head
    '    If Head <> 1 Then Exit Sub
    '    If Body <> 1 Then Exit Sub
    'Or something similar. It will appear at the top of the routine, and is pretty much the
    ' only thing that makes reference to the body or head in that sub, so it is easy to find.
    'Failure to do this will make the server reject the character. This is to prevent people from
    ' editing the packets to make their body or head whatever they want it to be.
    
    'Create the heads
    With HeadCmb
        .Clear
        .AddItem "Head 1", 0
        .ListIndex = 0
    End With
    
    'Create the bodies
    With BodyCmb
        .Clear
        .AddItem "Body 1", 0
        .ListIndex = 0
    End With
    
    'Create the classes
    With ClassCmb
        .Clear
        .AddItem "Warrior", 0
        .AddItem "Mage", 1
        .AddItem "Rogue", 2
        .ListIndex = 0
    End With
    
End Sub
