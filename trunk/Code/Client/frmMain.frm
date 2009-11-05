VERSION 5.00
Object = "{9842967E-F54F-4981-93DF-0772B2672E38}#1.0#0"; "vbgoresocketbinary.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Client"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SoxOCX.Sox Sox 
      Height          =   420
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer PingTmr 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements DirectXEvent8

Private NC As Byte

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
Dim DevData(1 To BufferSize) As DIDEVICEOBJECTDATA
Dim NumEvents As Long
Dim LoopC As Long
Dim Moved As Byte

    On Error GoTo ErrOut

    'Check if message is for us
    If eventid <> MouseEvent Then Exit Sub

    'Retrieve data
    NumEvents = DIDevice.GetDeviceData(DevData, DIGDD_DEFAULT)

    'Loop through data
    For LoopC = 1 To NumEvents
        Select Case DevData(LoopC).lOfs

            'Move on X axis
        Case DIMOFS_X
            MousePosAdd.x = (DevData(LoopC).lData * MouseSpeed)
            MousePos.x = MousePos.x + MousePosAdd.x
            If MousePos.x < 0 Then MousePos.x = 0
            If MousePos.x > frmMain.ScaleWidth Then MousePos.x = frmMain.ScaleWidth
            Moved = 1

            'Move on Y axis
        Case DIMOFS_Y
            MousePosAdd.Y = (DevData(LoopC).lData * MouseSpeed)
            MousePos.Y = MousePos.Y + MousePosAdd.Y
            If MousePos.Y < 0 Then MousePos.Y = 0
            If MousePos.Y > frmMain.ScaleHeight Then MousePos.Y = frmMain.ScaleHeight
            Moved = 1

            'Left button pressed
        Case DIMOFS_BUTTON0
            If DevData(LoopC).lData = 0 Then
                MouseLeftDown = 0
                SelGameWindow = 0
            Else
                If MouseLeftDown = 0 Then   'Clicked down
                    MouseLeftDown = 1
                    Engine_Input_Mouse_LeftClick
                End If
            End If

            'Right button pressed
        Case DIMOFS_BUTTON1
            If DevData(LoopC).lData = 0 Then
                MouseRightDown = 0
                Engine_Input_Mouse_RightRelease
            Else
                If MouseRightDown = 0 Then  'Clicked down
                    MouseRightDown = 1
                    Engine_Input_Mouse_RightClick
                End If
            End If

        End Select

        'Update movement
        If Moved Then
            Engine_Input_Mouse_Move

            'Reset move variables
            Moved = 0
            MousePosAdd.x = 0
            MousePosAdd.Y = 0
        End If

    Next LoopC

Exit Sub

ErrOut:
    NC = 1

End Sub

Private Sub Form_Click()

'Regain focus to DImouse

    If NC Then
        DIDevice.Acquire
        NC = 0
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim TempS() As String
Dim S As String
Dim i As Byte
'Attack key

    If KeyCode = vbKeyControl Then sndBuf.Put_Byte DataCode.User_Attack
    
    'Chat buffer stuff
    If KeyCode = vbKeyPageUp Then
        If ShowGameWindow(ChatWindow) Then
            ChatBufferChunk = ChatBufferChunk + 1
            Engine_UpdateChatArray
        End If
    End If
    If KeyCode = vbKeyPageDown Then
        If ShowGameWindow(ChatWindow) Then
            If ChatBufferChunk > 1 Then
                ChatBufferChunk = ChatBufferChunk - 1
                Engine_UpdateChatArray
            End If
        End If
    End If
    
    'Enter text
    If KeyCode = vbKeyReturn Then
        If LastClickedWindow = AmountWindow Then
            'Use the amount window
            If AmountWindowDropIndex Then
                If AmountWindowValue <> "" Then
                    If IsNumeric(AmountWindowValue) Then
                        sndBuf.Put_Byte DataCode.User_Drop
                        sndBuf.Put_Byte AmountWindowDropIndex
                        sndBuf.Put_Integer CInt(AmountWindowValue)
                    Else
                        AmountWindowValue = vbNullString
                    End If
                End If
            End If
        ElseIf LastClickedWindow = WriteMessageWindow Then
            'Send message
            If Len(WriteMailData.Subject) > 0 Then
                If Len(WriteMailData.Message) > 0 Then
                    If Len(WriteMailData.RecieverName) > 0 Then
                        sndBuf.Put_Byte DataCode.Server_MailCompose
                        sndBuf.Put_String WriteMailData.RecieverName
                        sndBuf.Put_String WriteMailData.Subject
                        sndBuf.Put_StringEX WriteMailData.Message
                        S = vbNullString
                        For i = 1 To MaxMailObjs
                            S = S & WriteMailData.ObjIndex(i)
                            WriteMailData.ObjIndex(i) = 0
                        Next i
                        sndBuf.Put_String S
                        S = vbNullString
                        For i = 1 To MaxMailObjs
                            S = S & WriteMailData.ObjAmount(i)
                            WriteMailData.ObjAmount(i) = 0
                        Next i
                        sndBuf.Put_String S
                        WriteMailData.Message = vbNullString
                        WriteMailData.RecieverName = vbNullString
                        WriteMailData.Subject = vbNullString
                        ShowGameWindow(WriteMessageWindow) = 0
                        LastClickedWindow = 0
                    End If
                End If
            End If
        Else
            If EnterText = True Then
                EnterText = False
                If EnterTextBuffer <> "" Then
                    'Check for commands
                    If UCase$(Left$(EnterTextBuffer, 4)) = "/BLI" Then
                        sndBuf.Put_Byte DataCode.User_Blink
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/LOOKL" Then
                        sndBuf.Put_Byte DataCode.User_LookLeft
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/LOOKR" Then
                        sndBuf.Put_Byte DataCode.User_LookRight
                    ElseIf UCase$(Left$(EnterTextBuffer, 4)) = "/WHO" Then
                        sndBuf.Put_Byte DataCode.Server_Who
                    ElseIf UCase$(Left$(EnterTextBuffer, 3)) = "/SH" Then
                        sndBuf.Put_Byte DataCode.Comm_Shout
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/TELL" Then
                        sndBuf.Put_Byte DataCode.Comm_Whisper
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(Left$(EnterTextBuffer, 3)) = "/ME" Then
                        sndBuf.Put_Byte DataCode.Comm_Emote
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(Left$(EnterTextBuffer, 3)) = "/EM" Then
                        sndBuf.Put_Byte DataCode.Comm_Emote
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(EnterTextBuffer) = "/QUIT" Then
                        IsUnloading = 1
                    ElseIf UCase$(EnterTextBuffer) = "/ACCEPT" Then
                        sndBuf.Put_Byte DataCode.User_StartQuest
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/DESC" Then
                        sndBuf.Put_Byte DataCode.User_Desc
                        sndBuf.Put_Byte SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(EnterTextBuffer) = "/HELP" Then
                        sndBuf.Put_Byte DataCode.Server_Help
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/APPR" Then
                        sndBuf.Put_Byte DataCode.GM_Approach
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(Left$(EnterTextBuffer, 4)) = "/SUM" Then
                        sndBuf.Put_Byte DataCode.GM_Summon
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/KICK" Then
                        sndBuf.Put_Byte DataCode.GM_Kick
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                    ElseIf UCase$(EnterTextBuffer) = "/DEVMODE" Then
                        sndBuf.Put_Byte DataCode.Dev_SetMode
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/RAISE" Then
                        TempS() = Split(SplitCommandFromString(EnterTextBuffer), " ")
                        If IsNumeric(TempS(1)) Then
                            sndBuf.Put_Byte DataCode.GM_Raise
                            sndBuf.Put_String TempS(0)
                            sndBuf.Put_Long CLng(TempS(1))
                        End If
                    Else
                        sndBuf.Put_Byte DataCode.Comm_Talk
                        sndBuf.Put_String EnterTextBuffer
                    End If
                    EnterTextBuffer = vbNullString
                    EnterTextBufferWidth = 10
                End If
            Else
                EnterText = True
            End If
        End If
    End If
    'Hide/show windows
    If GetAsyncKeyState(vbKeyControl) Then
        If KeyCode = vbKeyW Then
            If ShowGameWindow(InventoryWindow) Then
                ShowGameWindow(InventoryWindow) = 0
            Else
                ShowGameWindow(InventoryWindow) = 1
                LastClickedWindow = InventoryWindow
            End If
        ElseIf KeyCode = vbKeyQ Then
            If ShowGameWindow(QuickBarWindow) Then
                ShowGameWindow(QuickBarWindow) = 0
            Else
                ShowGameWindow(QuickBarWindow) = 1
                LastClickedWindow = QuickBarWindow
            End If
        ElseIf KeyCode = vbKeyC Then
            If ShowGameWindow(ChatWindow) Then
                ShowGameWindow(ChatWindow) = 0
            Else
                ShowGameWindow(ChatWindow) = 1
                LastClickedWindow = ChatWindow
            End If
        End If
    End If
    
    'Close screen
    If KeyCode = vbKeyEscape Then
        If LastClickedWindow = 0 Then
            If EnterText Then
                EnterTextBuffer = vbNullString
                EnterText = False
            Else
                ShowGameWindow(MenuWindow) = 1
                LastClickedWindow = MenuWindow
            End If
        Else
            If ShowGameWindow(LastClickedWindow) Then
                ShowGameWindow(LastClickedWindow) = 0
                If LastClickedWindow = DevWindow Then sndBuf.Put_Byte DataCode.Dev_SetMode
            End If
        End If
        LastClickedWindow = 0
    End If
    
    'Clear the keycode
    KeyCode = 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim b As Boolean

'Update amount window

    If LastClickedWindow = AmountWindow Then
        'Backspace
        If KeyAscii = 8 Then
            If Len(AmountWindowValue) > 0 Then
                AmountWindowValue = Left$(AmountWindowValue, Len(AmountWindowValue) - 1)
            End If
        End If
        'Number
        If IsNumeric(Chr$(KeyAscii)) Then AmountWindowValue = AmountWindowValue & Chr$(KeyAscii)
        'Write mail window
    ElseIf LastClickedWindow = WriteMessageWindow Then
        If WMSelCon Then
            Select Case WMSelCon
            Case wmFrom
                If KeyAscii = 8 Then
                    If Len(WriteMailData.RecieverName) > 0 Then
                        WriteMailData.RecieverName = Left$(WriteMailData.RecieverName, Len(WriteMailData.RecieverName) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(WriteMailData.RecieverName) < 10 Then
                            WriteMailData.RecieverName = WriteMailData.RecieverName & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case wmSubject
                If KeyAscii = 8 Then
                    If Len(WriteMailData.Subject) > 0 Then
                        WriteMailData.Subject = Left$(WriteMailData.Subject, Len(WriteMailData.Subject) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(WriteMailData.Subject) < 30 Then
                            WriteMailData.Subject = WriteMailData.Subject & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case wmMessage
                If KeyAscii = 8 Then
                    If Len(WriteMailData.Message) > 0 Then
                        WriteMailData.Message = Left$(WriteMailData.Message, Len(WriteMailData.Message) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(WriteMailData.Message) < 500 Then
                            WriteMailData.Message = WriteMailData.Message & Chr$(KeyAscii)
                        End If
                    End If
                End If
            End Select
        End If
        'Dev window
    ElseIf LastClickedWindow = DevWindow Then
        If DevHasFocus Then
            Select Case DevHasFocus
            Case Lighting1
                If KeyAscii = 8 Then
                    If Len(DevValue.Lighting(1)) > 0 Then
                        DevValue.Lighting(1) = Left$(DevValue.Lighting(1), Len(DevValue.Lighting(1)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Lighting(1)) < 500 Then
                            DevValue.Lighting(1) = DevValue.Lighting(1) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Lighting2
                If KeyAscii = 8 Then
                    If Len(DevValue.Lighting(2)) > 0 Then
                        DevValue.Lighting(2) = Left$(DevValue.Lighting(2), Len(DevValue.Lighting(2)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Lighting(2)) < 500 Then
                            DevValue.Lighting(2) = DevValue.Lighting(2) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Lighting3
                If KeyAscii = 8 Then
                    If Len(DevValue.Lighting(3)) > 0 Then
                        DevValue.Lighting(3) = Left$(DevValue.Lighting(3), Len(DevValue.Lighting(3)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Lighting(3)) < 500 Then
                            DevValue.Lighting(3) = DevValue.Lighting(3) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Lighting4
                If KeyAscii = 8 Then
                    If Len(DevValue.Lighting(4)) > 0 Then
                        DevValue.Lighting(4) = Left$(DevValue.Lighting(4), Len(DevValue.Lighting(4)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Lighting(4)) < 500 Then
                            DevValue.Lighting(4) = DevValue.Lighting(4) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Grh1
                If KeyAscii = 8 Then
                    If Len(DevValue.Grh(1)) > 0 Then
                        DevValue.Grh(1) = Left$(DevValue.Grh(1), Len(DevValue.Grh(1)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Grh(1)) < 500 Then
                            DevValue.Grh(1) = DevValue.Grh(1) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Grh2
                If KeyAscii = 8 Then
                    If Len(DevValue.Grh(2)) > 0 Then
                        DevValue.Grh(2) = Left$(DevValue.Grh(2), Len(DevValue.Grh(2)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Grh(2)) < 500 Then
                            DevValue.Grh(2) = DevValue.Grh(2) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Grh3
                If KeyAscii = 8 Then
                    If Len(DevValue.Grh(3)) > 0 Then
                        DevValue.Grh(3) = Left$(DevValue.Grh(3), Len(DevValue.Grh(3)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Grh(3)) < 500 Then
                            DevValue.Grh(3) = DevValue.Grh(3) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Grh4
                If KeyAscii = 8 Then
                    If Len(DevValue.Grh(4)) > 0 Then
                        DevValue.Grh(4) = Left$(DevValue.Grh(4), Len(DevValue.Grh(4)) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Grh(4)) < 500 Then
                            DevValue.Grh(4) = DevValue.Grh(4) & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case NPC
                If KeyAscii = 8 Then
                    If Len(DevValue.NPC) > 0 Then
                        DevValue.NPC = Left$(DevValue.NPC, Len(DevValue.NPC) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.NPC) < 500 Then
                            DevValue.NPC = DevValue.NPC & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Blocked
                If KeyAscii = 8 Then
                    If Len(DevValue.Blocked) > 0 Then
                        DevValue.Blocked = Left$(DevValue.Blocked, Len(DevValue.Blocked) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Blocked) < 500 Then
                            DevValue.Blocked = DevValue.Blocked & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case ExitMap
                If KeyAscii = 8 Then
                    If Len(DevValue.ExitMap) > 0 Then
                        DevValue.ExitMap = Left$(DevValue.ExitMap, Len(DevValue.ExitMap) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.ExitMap) < 500 Then
                            DevValue.ExitMap = DevValue.ExitMap & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case ExitX
                If KeyAscii = 8 Then
                    If Len(DevValue.ExitX) > 0 Then
                        DevValue.ExitX = Left$(DevValue.ExitX, Len(DevValue.ExitX) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.ExitX) < 500 Then
                            DevValue.ExitX = DevValue.ExitX & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case ExitY
                If KeyAscii = 8 Then
                    If Len(DevValue.ExitY) > 0 Then
                        DevValue.ExitY = Left$(DevValue.ExitY, Len(DevValue.ExitY) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.ExitY) < 500 Then
                            DevValue.ExitY = DevValue.ExitY & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Mailbox
                If KeyAscii = 8 Then
                    If Len(DevValue.Mailbox) > 0 Then
                        DevValue.Mailbox = Left$(DevValue.Mailbox, Len(DevValue.Mailbox) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Mailbox) < 500 Then
                            DevValue.Mailbox = DevValue.Mailbox & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Namex
                If KeyAscii = 8 Then
                    If Len(DevValue.Name) > 0 Then
                        DevValue.Name = Left$(DevValue.Name, Len(DevValue.Name) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Name) < 500 Then
                            DevValue.Name = DevValue.Name & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Obj
                If KeyAscii = 8 Then
                    If Len(DevValue.Obj) > 0 Then
                        DevValue.Obj = Left$(DevValue.Obj, Len(DevValue.Obj) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Obj) < 500 Then
                            DevValue.Obj = DevValue.Obj & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case ObjAmount
                If KeyAscii = 8 Then
                    If Len(DevValue.ObjAmount) > 0 Then
                        DevValue.ObjAmount = Left$(DevValue.ObjAmount, Len(DevValue.ObjAmount) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.ObjAmount) < 500 Then
                            DevValue.ObjAmount = DevValue.ObjAmount & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Version
                If KeyAscii = 8 Then
                    If Len(DevValue.Version) > 0 Then
                        DevValue.Version = Left$(DevValue.Version, Len(DevValue.Version) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Version) < 500 Then
                            DevValue.Version = DevValue.Version & Chr$(KeyAscii)
                        End If
                    End If
                End If
            Case Weather
                If KeyAscii = 8 Then
                    If Len(DevValue.Weather) > 0 Then
                        DevValue.Weather = Left$(DevValue.Weather, Len(DevValue.Weather) - 1)
                    End If
                End If
                If KeyAscii >= 32 Then
                    If KeyAscii <= 126 Then
                        If Len(DevValue.Weather) < 500 Then
                            DevValue.Weather = DevValue.Weather & Chr$(KeyAscii)
                        End If
                    End If
                End If
            End Select
            Game_ClearMapTileChanged
        End If
        'Send text
    Else
        If EnterText Then
            'Backspace
            If KeyAscii = 8 Then
                If Len(EnterTextBuffer) > 0 Then EnterTextBuffer = Left$(EnterTextBuffer, Len(EnterTextBuffer) - 1)
                b = True
            End If
            'Add to text buffer
            If KeyAscii >= 32 Then
                If KeyAscii <= 126 Then
                    If Len(EnterTextBuffer) < 85 Then
                        EnterTextBuffer = EnterTextBuffer & Chr$(KeyAscii)
                        b = True
                    End If
                End If
            End If
            'Update size
            If b Then
                EnterTextBufferWidth = Engine_GetTextWidth(EnterTextBuffer)
                LastClickedWindow = 0
            End If
        End If
    End If
    
    'Clear the key
    KeyAscii = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Dim i As Integer

    'Get object off ground (alt)
    If KeyCode = 18 Then sndBuf.Put_Byte DataCode.User_Get
    
    'Reset skins (F12)
    If KeyCode = vbKeyF12 Then
        Engine_Init_GUI 0
        Game_Config_Save
    End If

    'Delete mail (Delete)
    If KeyCode = vbKeyDelete Then
        If LastClickedWindow = MailboxWindow Then
            If ShowGameWindow(MailboxWindow) Then
                If SelMessage > 0 Then
                    sndBuf.Put_Byte DataCode.Server_MailDelete
                    sndBuf.Put_Byte SelMessage
                End If
            End If
        End If
    End If

    'Use the quick bar
    If KeyCode >= vbKeyF1 Then
        If KeyCode <= vbKeyF12 Then
            Engine_UseQuickBar KeyCode - vbKeyF1 + 1
        End If
    End If

    'Send an emoticon
    If EnterText = False Then
        Select Case KeyCode
        Case vbKey1
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Dots
        Case vbKey2
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Exclimation
        Case vbKey3
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Question
        Case vbKey4
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Surprised
        Case vbKey5
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Heart
        Case vbKey6
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Hearts
        Case vbKey7
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.HeartBroken
        Case vbKey8
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Utensils
        Case vbKey9
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.Meat
        Case vbKey0
            sndBuf.Put_Byte DataCode.User_Emote
            sndBuf.Put_Byte EmoID.ExcliQuestion
        End Select
    End If
    
    'Clear the key
    KeyCode = 0

End Sub

Private Sub PingTmr_Timer()

'Ping the server

    sndBuf.Put_Byte DataCode.Server_Ping
    PingSTime = timeGetTime

    If NonRetPings < 250 Then NonRetPings = NonRetPings + 1

End Sub

Private Sub Sox_OnClose(inSox As Long)

    If SocketOpen = 1 Then IsUnloading = 1

End Sub

Private Sub Sox_OnDataArrival(inSox As Long, inData() As Byte)

'*********************************************
'Retrieve the CommandIDs and send to corresponding data handler
'*********************************************

Dim rBuf As DataBuffer
Dim CommandID As Byte
Dim BufUBound As Long

'Counts empty CommandIDs - used for debugging to remove extra packet info
'In a perfect program, this should always be 0, which means no empty data was sent and caught
Static x As Long

    'Display packet
    If DEBUG_PrintPacket_In Then
        Engine_AddToChatTextBuffer "DataIn: " & StrConv(inData, vbUnicode), -1
    End If

    'Decrypt our packet
    Select Case EncryptionType
        Case EncryptionTypeBlowfish
            Encryption_Blowfish_DecryptByte inData, EncryptionKey
        Case EncryptionTypeCryptAPI
            Encryption_CryptAPI_DecryptByte inData, EncryptionKey
        Case EncryptionTypeDES
            Encryption_DES_DecryptByte inData, EncryptionKey
        Case EncryptionTypeGost
            Encryption_Gost_DecryptByte inData, EncryptionKey
        Case EncryptionTypeRC4
            Encryption_RC4_DecryptByte inData, EncryptionKey
        Case EncryptionTypeXOR
            Encryption_XOR_DecryptByte inData, EncryptionKey
        Case EncryptionTypeSkipjack
            Encryption_Skipjack_DecryptByte inData, EncryptionKey
        Case EncryptionTypeTEA
            Encryption_TEA_DecryptByte inData, EncryptionKey
        Case EncryptionTypeTwofish
            Encryption_Twofish_DecryptByte inData, EncryptionKey
    End Select
    
    'Set up the data buffer
    Set rBuf = New DataBuffer
    rBuf.Set_Buffer inData
    BufUBound = UBound(inData)
    
    'Uncomment this to see packets going in to the client
    'Dim i As Long
    'Dim S As String
    'For i = LBound(inData) To UBound(inData)
    '    S = S & inData(i) & " "
    'Next i
    'Debug.Print S

    Do
        'Get the Command ID
        CommandID = rBuf.Get_Byte

        'Make the appropriate call based on the Command ID
        With DataCode
            Select Case CommandID
            Case 0
                If DEBUG_PrintPacketReadErrors Then
                    x = x + 1
                    Debug.Print "---Blank Byte #" & x & " - Previous CommandID was " & CommandID
                End If

            Case .Comm_Talk: Data_Comm_Talk rBuf
            Case .Comm_UMsgbox: Data_Comm_UMsgBox rBuf

            Case .Dev_SetMapInfo: Data_Dev_SetMapInfo rBuf
            Case .Dev_SetMode: Data_Dev_SetMode rBuf

            Case .Map_DoneSwitching: Data_Map_DoneSwitching
            Case .Map_EndTransfer: Data_Map_EndTransfer rBuf
            Case .Map_LoadMap: Data_Map_LoadMap rBuf
            Case .Map_SendName:  Data_Map_SendName rBuf
            Case .Map_StartTransfer: Data_Map_StartTransfer rBuf
            Case .Map_UpdateTile:  Data_Map_UpdateTile rBuf

            Case .Server_ChangeChar: Data_Server_ChangeChar rBuf
            Case .Server_CharHP: Data_Server_CharHP rBuf
            Case .Server_CharMP: Data_Server_CharMP rBuf
            Case .Server_Connect: Data_Server_Connect
            Case .Server_Disconnect: Data_Server_Disconnect
            Case .Server_EraseChar: Data_Server_EraseChar rBuf
            Case .Server_EraseObject: Data_Server_EraseObject rBuf
            Case .Server_IconBlessed: Data_Server_IconBlessed rBuf
            Case .Server_IconCursed: Data_Server_IconCursed rBuf
            Case .Server_IconIronSkin: Data_Server_IconIronSkin rBuf
            Case .Server_IconProtected: Data_Server_IconProtected rBuf
            Case .Server_IconStrengthened: Data_Server_IconStrengthened rBuf
            Case .Server_IconWarCursed:  Data_Server_IconWarCursed rBuf
            Case .Server_IconSpellExhaustion: Data_Server_IconSpellExhaustion rBuf
            Case .Server_MailBox: Data_Server_Mailbox rBuf
            Case .Server_MailItemInfo: Data_Server_MailItemInfo rBuf
            Case .Server_MailItemRemove: Data_Server_MailItemRemove rBuf
            Case .Server_MailMessage: Data_Server_MailMessage rBuf
            Case .Server_MakeChar: Data_Server_MakeChar rBuf
            Case .Server_MakeObject: Data_Server_MakeObject rBuf
            Case .Server_MoveChar: Data_Server_MoveChar rBuf
            Case .Server_Ping: Data_Server_Ping
            Case .Server_PlaySound: Data_Server_PlaySound rBuf
            Case .Server_SetCharDamage: Data_Server_SetCharDamage rBuf
            Case .Server_SetUserPosition: Data_Server_SetUserPosition rBuf
            Case .Server_UserCharIndex: Data_Server_UserCharIndex rBuf

            Case .User_AggressiveFace: Data_User_AggressiveFace rBuf
            Case .User_Attack: Data_User_Attack rBuf
            Case .User_BaseStat: Data_User_BaseStat rBuf
            Case .User_Blink: Data_User_Blink rBuf
            Case .User_CastSkill: Data_User_CastSkill rBuf
            Case .User_Emote: Data_User_Emote rBuf
            Case .User_KnownSkills: Data_User_KnownSkills rBuf
            Case .User_LookLeft: Data_User_LookLeft rBuf
            Case .User_LookRight: Data_User_LookLeft rBuf
            Case .User_ModStat: Data_User_ModStat rBuf
            Case .User_Rotate: Data_User_Rotate rBuf
            Case .User_SetInventorySlot: Data_User_SetInventorySlot rBuf
            Case .User_Target: Data_User_Target rBuf
            Case .User_Trade_StartNPCTrade: Data_User_Trade_StartNPCTrade rBuf

            Case Else
                If DEBUG_PrintPacketReadErrors Then Debug.Print "Command ID " & CommandID & " cause premature packet handling abortion!"
                Exit Do 'Something went wrong or we hit the end, either way, RUN!!!!

            End Select
        End With

        'Exit when the buffer runs out
        If rBuf.Get_ReadPos > BufUBound Then Exit Do

    Loop

End Sub

Private Function SplitCommandFromString(StringBuffer As String) As String

Dim TempSplit() As String
Dim i As Integer

    TempSplit() = Split(StringBuffer, " ")

    For i = 1 To UBound(TempSplit)
        SplitCommandFromString = SplitCommandFromString & TempSplit(i) & " "
    Next i
    SplitCommandFromString = Left$(SplitCommandFromString, Len(SplitCommandFromString) - 1)

End Function

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:36)  Decl: 4  Code: 837  Total: 841 Lines
':) CommentOnly: 51 (6.1%)  Commented: 2 (0.2%)  Empty: 77 (9.2%)  Max Logic Depth: 7
Private Sub Sox_OnState(inSox As Long, inState As SoxOCX.enmSoxState)
    
    If inState = soxConnecting Then
        If SocketOpen = 0 Then
    
            'Pre-saved character
            If SendNewChar = False Then
                sndBuf.Put_Byte DataCode.User_Login
                sndBuf.Put_String UserName
                sndBuf.Put_String UserPassword
            Else
                'New character
                sndBuf.Put_Byte DataCode.User_NewLogin
                sndBuf.Put_String UserName
                sndBuf.Put_String UserPassword
            End If
        
            'Save Game.ini
            If frmConnect.SavePassChk.Value = 0 Then UserPassword = vbNullString
            Engine_Var_Write DataPath & "Game.ini", "INIT", "Name", UserName
            Engine_Var_Write DataPath & "Game.ini", "INIT", "Password", UserPassword
            
            'Send the data
            Data_Send
            DoEvents
    
        End If
    End If

End Sub
