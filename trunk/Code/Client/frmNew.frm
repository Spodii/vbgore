VERSION 5.00
Begin VB.Form frmNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Login"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox HeadCmb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2355
      Width           =   2055
   End
   Begin VB.ComboBox ClassCmb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1530
      Width           =   2055
   End
   Begin VB.ComboBox BodyCmb 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1935
      Width           =   2055
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   1875
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   780
      Width           =   1875
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClickCancel()
'*****************************************************************
'Hides frmNew and displays frmConnect
'More info: http://www.vbgore.com/GameClient.frmNew.ClickCancel
'*****************************************************************

    'Show the connect screen
    frmConnect.Visible = True
    
    'Hide this screen
    Me.Visible = False

End Sub

Private Sub ClickCreate()
'*****************************************************************
'Sends the packet to the server requesting to create a new user
'More info: http://www.vbgore.com/GameClient.frmNew.ClickCancel
'*****************************************************************

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

Private Sub Form_Unload(Cancel As Integer)
'*****************************************************************
'Unloads the picture textboxes
'More info: http://www.vbgore.com/GameClient.frmNew.Form_Unload
'*****************************************************************

    FreePictureTextboxes Me.hwnd

End Sub

Private Sub NameTxt_KeyPress(KeyAscii As Integer)
'*****************************************************************
'Create new character when return is pressed
'More info: http://www.vbgore.com/GameClient.frmNew.NameTxt_KeyPress
'*****************************************************************

    If KeyAscii = Asc(vbNewLine) Then
        KeyAscii = 0
        ClickCreate
    End If

End Sub

Private Sub PasswordTxt_KeyPress(KeyAscii As Integer)
'*****************************************************************
'Create new character when return is pressed
'More info: http://www.vbgore.com/GameClient.frmNew.PasswordTxt_KeyPress
'*****************************************************************

    If KeyAscii = Asc(vbNewLine) Then
        KeyAscii = 0
        ClickCreate
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'*****************************************************************
'Create new character when return is pressed
'More info: http://www.vbgore.com/GameClient.frmNew.Form_KeyPress
'*****************************************************************

    If KeyAscii = Asc(vbNewLine) Then ClickCreate

End Sub

Private Sub Form_Load()
'*****************************************************************
'Loads up the values for frmNew and creates the listbox values and pictures
'More info: http://www.vbgore.com/GameClient.frmNew.Form_Load
'*****************************************************************

    'Set the background picture
    Me.Picture = LoadPicture(GrhPath & "New.bmp")

    'Set the text boxes to transparent
    SetPictureTextboxes Me.hwnd

    'Load up the head, body and class values you can select
    'For the head and body, to add more, you have to edit it accordingly in the server
    ' under User_ConnectNew on this line:
    '
    '    'Check for a valid body and head
    '    If Head <> 1 Then Exit Sub
    '    If Body <> 1 Then Exit Sub
    '
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************************
'Check if the buttons on the form were clicked
'More info: http://www.vbgore.com/GameClient.frmNew.Form_MouseDown
'*****************************************************************

    'Click on "Create"
    If Engine_Collision_Rect(X, Y, 1, 1, 5, 189, 66, 15) Then ClickCreate
    
    'Click on "Cancel"
    If Engine_Collision_Rect(X, Y, 1, 1, 118, 190, 66, 15) Then ClickCancel

End Sub
