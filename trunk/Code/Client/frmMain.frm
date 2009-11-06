VERSION 5.00
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
   Begin VB.Timer PTDTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer ShutdownTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
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
            MousePosAdd.X = (DevData(LoopC).lData * MouseSpeed)
            MousePos.X = MousePos.X + MousePosAdd.X
            If MousePos.X < 0 Then MousePos.X = 0
            If MousePos.X > frmMain.ScaleWidth Then MousePos.X = frmMain.ScaleWidth
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
            MousePosAdd.X = 0
            MousePosAdd.Y = 0
        End If

    Next LoopC

Exit Sub

ErrOut:
    NC = 1

End Sub

Private Sub Form_Click()

    'Regain focus to Direct Input mouse
    If NC Then
        DIDevice.Acquire
        NC = 0
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim TempS() As String
Dim s As String
Dim s2 As String
Dim i As Byte
Dim j As Long

    'Disable / enable input (for debugging)
    If KeyCode = vbKeyF11 Then
        If GetAsyncKeyState(vbKeyShift) Then
            DisableInput = 1
            Engine_AddToChatTextBuffer "Input disabled", FontColor_Info
        End If
    End If
    If KeyCode = vbKeyF10 Then
        If GetAsyncKeyState(vbKeyShift) Then
            DisableInput = 0
            Engine_AddToChatTextBuffer "Input enabled", FontColor_Info
        End If
    End If
    
    '*************************
    '***** General input *****
    '*************************
    
    'Get object off ground (alt)
    If KeyCode = 18 Then
        If LastLootTime + LootDelay < timeGetTime Then
            LastLootTime = timeGetTime
            sndBuf.Put_Byte DataCode.User_Get
        End If
    End If
    
    'Use the quick bar
    If KeyCode >= vbKeyF1 Then
        If KeyCode <= vbKeyF12 Then
            Engine_UseQuickBar KeyCode - vbKeyF1 + 1
        End If
    End If
    
    'Attack key
    If KeyCode = vbKeyControl Then
        If LastAttackTime + AttackDelay < timeGetTime Then
            
            'Check for a valid attacking distance
            If UserAttackRange > 1 Then
                If TargetCharIndex > 0 Then
                    If Engine_Distance(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(TargetCharIndex).Pos.X, CharList(TargetCharIndex).Pos.Y) <= UserAttackRange Then
                        LastAttackTime = timeGetTime
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Attack
                        sndBuf.Put_Byte CharList(UserCharIndex).Heading
                    Else
                        Engine_AddToChatTextBuffer Message(91), FontColor_Fight
                    End If
                End If
            Else
                LastAttackTime = timeGetTime
                sndBuf.Allocate 2
                sndBuf.Put_Byte DataCode.User_Attack
                sndBuf.Put_Byte CharList(UserCharIndex).Heading
            End If
            
        End If
    End If
    
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
        ElseIf KeyCode = vbKeyS Then
            If ShowGameWindow(StatWindow) Then
                ShowGameWindow(StatWindow) = 0
            Else
                ShowGameWindow(StatWindow) = 1
                LastClickedWindow = StatWindow
            End If
        End If
    End If

    If KeyCode = vbKeyReturn Then
        
        '*************************
        '***** Amount window *****
        '*************************
        If LastClickedWindow = AmountWindow Then
            If AmountWindowItemIndex Then
                If AmountWindowValue <> "" Then
                    If IsNumeric(AmountWindowValue) Then
                        'Drop into mail
                        If AmountWindowUsage = AW_InvToMail Then
                            'Check for duplicate entries
                            For j = 1 To MaxMailObjs
                                If WriteMailData.ObjIndex(j) = AmountWindowItemIndex Then
                                    ShowGameWindow(AmountWindow) = 0
                                    AmountWindowUsage = 0
                                    LastClickedWindow = 0
                                    Exit Sub
                                End If
                            Next j
                            'Find the next free slot
                            j = 0
                            Do
                                j = j + 1
                                If j > MaxMailObjs Then
                                    ShowGameWindow(AmountWindow) = 0
                                    AmountWindowUsage = 0
                                    LastClickedWindow = 0
                                    Exit Sub
                                End If
                            Loop While WriteMailData.ObjIndex(j) > 0
                            WriteMailData.ObjIndex(j) = AmountWindowItemIndex
                            WriteMailData.ObjAmount(j) = CInt(AmountWindowValue)
                        'Buy from NPC
                        ElseIf AmountWindowUsage = AW_ShopToInv Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Trade_BuyFromNPC
                            sndBuf.Put_Byte AmountWindowItemIndex
                            sndBuf.Put_Integer CInt(AmountWindowValue)
                        'Sell to NPC
                        ElseIf AmountWindowUsage = AW_InvToShop Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Trade_SellToNPC
                            sndBuf.Put_Byte AmountWindowItemIndex
                            sndBuf.Put_Integer CInt(AmountWindowValue)
                        'Take from bank
                        ElseIf AmountWindowUsage = AW_BankToInv Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Bank_TakeItem
                            sndBuf.Put_Byte AmountWindowItemIndex
                            sndBuf.Put_Integer CInt(AmountWindowValue)
                        'Put in bank
                        ElseIf AmountWindowUsage = AW_InvToBank Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Bank_PutItem
                            sndBuf.Put_Byte AmountWindowItemIndex
                            sndBuf.Put_Integer CInt(AmountWindowValue)
                        'Drop on ground
                        Else
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Drop
                            sndBuf.Put_Byte AmountWindowItemIndex
                            sndBuf.Put_Integer CInt(AmountWindowValue)
                        End If
                    Else
                        AmountWindowValue = vbNullString
                    End If
                    ShowGameWindow(AmountWindow) = 0
                    AmountWindowUsage = 0
                    LastClickedWindow = 0
                End If
            End If
            
        '*****************************
        '***** Write mail window *****
        '*****************************
        ElseIf LastClickedWindow = WriteMessageWindow Then
            'Send message
            If LastMailSendTime + 4000 < timeGetTime Then   'DelayTimeMail (+1000ms for packet delay)
                If Len(WriteMailData.Subject) > 0 Then
                    If Len(WriteMailData.Message) > 0 Then
                        If Len(WriteMailData.RecieverName) > 0 Then
                            For i = 1 To MaxMailObjs
                                If WriteMailData.ObjIndex(i) = 0 Then
                                    i = i - 1
                                    Exit For
                                End If
                            Next i
                            sndBuf.Allocate 6 + Len(WriteMailData.RecieverName) + Len(WriteMailData.Subject) + Len(WriteMailData.Message)
                            sndBuf.Put_Byte DataCode.Server_MailCompose
                            sndBuf.Put_String WriteMailData.RecieverName
                            sndBuf.Put_String WriteMailData.Subject
                            sndBuf.Put_StringEX WriteMailData.Message
                            sndBuf.Put_Byte i   'Number of objects
                            If i > 0 Then
                                For j = 1 To i
                                    sndBuf.Allocate 3
                                    sndBuf.Put_Byte WriteMailData.ObjIndex(j)
                                    sndBuf.Put_Integer WriteMailData.ObjAmount(j)
                                Next j
                            End If
                            
                            WriteMailData.Message = vbNullString
                            WriteMailData.RecieverName = vbNullString
                            WriteMailData.Subject = vbNullString
                            ShowGameWindow(WriteMessageWindow) = 0
                            LastClickedWindow = 0
                            LastMailSendTime = timeGetTime
                        End If
                    End If
                End If
            End If
            
        '***********************
        '***** Chat screen *****
        '***********************
        Else
            If EnterText = True Then
                If EnterTextBuffer <> "" Then
                
                    '***** Check for commands *****
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
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/LANG" Then
                        s = LCase$(SplitCommandFromString(EnterTextBuffer))
                        If Engine_FileExist(MessagePath & s & "*.ini", vbNormal) Then
                            s = Dir$(MessagePath & s & "*.ini", vbNormal)
                            s = Left$(s, Len(s) - 4)
                            Engine_Init_Messages s
                            Engine_Var_Write DataPath & "Game.ini", "INIT", "Language", s
                            Engine_AddToChatTextBuffer Replace$(Message(90), "<lang>", s), FontColor_Info
                        Else
                            Engine_AddToChatTextBuffer Message(87), FontColor_Info
                        End If
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/SKIN" Then
                        s = LCase$(SplitCommandFromString(EnterTextBuffer))
                        If s = "" Then
                            Engine_AddToChatTextBuffer Engine_BuildSkinsList, FontColor_Info
                        Else
                            If Engine_FileExist(DataPath & "Skins\" & s & "*.ini", vbNormal) Then
                                s = Dir$(DataPath & "Skins\" & s & "*.ini", vbNormal)
                                CurrentSkin = Left$(s, Len(s) - 4)
                                Engine_Init_GUI 0
                                Engine_Var_Write DataPath & "Game.ini", "INIT", "CurrentSkin", CurrentSkin
                                Engine_AddToChatTextBuffer Replace$(Message(89), "<skin>", CurrentSkin), FontColor_Info
                            Else
                                Engine_AddToChatTextBuffer Message(88), FontColor_Info
                            End If
                        End If
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/QUEST" Then
                        If QuestInfoUBound = 0 Then
                            'No quests in place
                            Engine_AddToChatTextBuffer Message(103), FontColor_Quest
                        Else
                            j = Val(Trim$(SplitCommandFromString(EnterTextBuffer)))
                            If j < 1 Or j > QuestInfoUBound Then
                                'No valid number specified, give the list
                                Engine_AddToChatTextBuffer Message(104), FontColor_Quest
                                For i = 1 To QuestInfoUBound
                                    Engine_AddToChatTextBuffer "  " & i & ". " & QuestInfo(i).name, FontColor_Quest
                                Next i
                            Else
                                'Give the info on the specific quest
                                Engine_AddToChatTextBuffer QuestInfo(j).name & ":", FontColor_Quest
                                Engine_AddToChatTextBuffer QuestInfo(j).Desc, FontColor_Quest
                            End If
                        End If
                                
                    ElseIf UCase$(Left$(EnterTextBuffer, 4)) = "/THR" Then
                        TempS = Split(EnterTextBuffer)
                        If UBound(TempS) <> 0 Then
                            If IsNumeric(TempS(1)) Then
                                sndBuf.Put_Byte DataCode.GM_Thrall
                                sndBuf.Put_Integer Val(TempS(1))
                                If UBound(TempS) > 1 Then
                                    If IsNumeric(TempS(2)) Then
                                        sndBuf.Put_Integer Val(TempS(2))
                                    Else
                                        sndBuf.Put_Integer 1
                                    End If
                                    sndBuf.Put_Integer 1
                                End If
                            End If
                        End If
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/DETHR" Then
                        sndBuf.Put_Byte DataCode.GM_DeThrall
                        
                    ElseIf UCase$(EnterTextBuffer) = "/QUIT" Then
                        IsUnloading = 1
                        
                    ElseIf UCase$(EnterTextBuffer) = "/ACCEPT" Then
                        sndBuf.Put_Byte DataCode.User_StartQuest
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/DESC" Then
                        sndBuf.Put_Byte DataCode.User_Desc
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                        
                    ElseIf UCase$(EnterTextBuffer) = "/HELP" Then
                        sndBuf.Put_Byte DataCode.Server_Help
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/APPR" Then
                        sndBuf.Put_Byte DataCode.GM_Approach
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 4)) = "/SUM" Then
                        sndBuf.Put_Byte DataCode.GM_Summon
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/SETGM" Then
                        TempS = Split(SplitCommandFromString(EnterTextBuffer), " ")
                        If UBound(TempS) > 0 Then
                            If IsNumeric(TempS(1)) Then
                                sndBuf.Allocate 3 + Len(TempS(0))
                                sndBuf.Put_Byte DataCode.GM_SetGMLevel
                                sndBuf.Put_String TempS(0)
                                sndBuf.Put_Byte CByte(TempS(1))
                            End If
                        End If
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/BANIP" Then
                        s = SplitCommandFromString(EnterTextBuffer) 'Remove the command
                        If LenB(s) < 4 Then 'Not enough information entered
                            Engine_AddToChatTextBuffer Message(92), FontColor_Info
                            Exit Sub
                        End If
                        TempS = Split(s, " ", 2)    'Split up the IP and reason
                        If UBound(TempS) = 0 Then
                            Engine_AddToChatTextBuffer Message(93), FontColor_Info
                            Exit Sub
                        Else
                            s = TempS(0)
                            s2 = TempS(1)
                        End If
                        TempS = Split(s, ".")
                        If UBound(TempS) <> 3 Then
                            Engine_AddToChatTextBuffer Message(92), FontColor_Info
                            Exit Sub
                        End If
                        For j = 0 To 3
                            If Val(TempS(j)) < 0 Or Val(TempS(j)) > 255 Then
                                Engine_AddToChatTextBuffer Message(92), FontColor_Info
                                Exit Sub
                            End If
                        Next j
                        sndBuf.Put_Byte DataCode.GM_BanIP
                        sndBuf.Put_String Trim$(s)
                        sndBuf.Put_String Trim$(s2)
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 8)) = "/UNBANIP" Then
                        s = SplitCommandFromString(EnterTextBuffer) 'Remove the command
                        If LenB(s) < 4 Then 'Not enough information entered
                            Engine_AddToChatTextBuffer Message(92), FontColor_Info
                            Exit Sub
                        End If
                        TempS = Split(s, ".")
                        If UBound(TempS) <> 3 Then
                            Engine_AddToChatTextBuffer Message(92), FontColor_Info
                            Exit Sub
                        End If
                        For j = 0 To 3
                            If TempS(j) <> "*" Then
                                If Val(TempS(j)) < 0 Or Val(TempS(j)) > 255 Then
                                    Engine_AddToChatTextBuffer Message(92), FontColor_Info
                                    Exit Sub
                                End If
                            End If
                        Next j
                        sndBuf.Put_Byte DataCode.GM_UnBanIP
                        sndBuf.Put_String Trim$(s)
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 5)) = "/KICK" Then
                        sndBuf.Put_Byte DataCode.GM_Kick
                        sndBuf.Put_String SplitCommandFromString(EnterTextBuffer)
                        
                    ElseIf UCase$(Left$(EnterTextBuffer, 6)) = "/RAISE" Then
                        TempS() = Split(SplitCommandFromString(EnterTextBuffer), " ")
                        If UBound(TempS) > 0 Then
                            If IsNumeric(TempS(1)) Then
                                sndBuf.Allocate 6 + Len(TempS(0))
                                sndBuf.Put_Byte DataCode.GM_Raise
                                sndBuf.Put_String TempS(0)
                                sndBuf.Put_Long CLng(TempS(1))
                            End If
                        End If
                        
                    Else
                        '*** No commands sent, send as text ***
                        sndBuf.Allocate 2 + Len(EnterTextBuffer)
                        sndBuf.Put_Byte DataCode.Comm_Talk
                        sndBuf.Put_String EnterTextBuffer
                        
                        'We just sent a chat message, so check if it had triggers!
                        Engine_NPCChat_CheckForChatTriggers EnterTextBuffer
                    End If
                    EnterTextBuffer = vbNullString
                    EnterTextBufferWidth = 10
                    ShownText = vbNullString
                End If
                
                EnterText = False
            Else
                EnterText = True
            End If
        End If
    End If
    
    '*****************************
    '***** Close last screen *****
    '*****************************
    If KeyCode = vbKeyEscape Then
        If LastClickedWindow = 0 Then
            If ShowGameWindow(MenuWindow) = 1 Then
                ShowGameWindow(MenuWindow) = 0
                LastClickedWindow = 0
            Else
                If EnterText Then
                    EnterTextBuffer = vbNullString
                    EnterText = False
                Else
                    ShowGameWindow(MenuWindow) = 1
                    LastClickedWindow = MenuWindow
                End If
            End If
        Else
            If ShowGameWindow(LastClickedWindow) Then ShowGameWindow(LastClickedWindow) = 0
        End If
        LastClickedWindow = 0
    End If
    
    'Clear the keycode
    KeyCode = 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim b As Boolean

    '*************************
    '***** Amount window *****
    '*************************
    If LastClickedWindow = AmountWindow Then
        'Backspace
        If KeyAscii = 8 Then
            If Len(AmountWindowValue) > 0 Then
                AmountWindowValue = Left$(AmountWindowValue, Len(AmountWindowValue) - 1)
            End If
        End If
        'Number
        If IsNumeric(Chr$(KeyAscii)) Then
            AmountWindowValue = AmountWindowValue & Chr$(KeyAscii)
            If Val(AmountWindowValue) > MAXINT Then AmountWindowValue = Str(MAXINT)
        End If
    
    '*****************************
    '***** Write mail window *****
    '*****************************
    ElseIf LastClickedWindow = WriteMessageWindow Then
        If WMSelCon Then
            Select Case WMSelCon
            Case wmFrom
                If KeyAscii = 8 Then
                    If Len(WriteMailData.RecieverName) > 0 Then
                        WriteMailData.RecieverName = Left$(WriteMailData.RecieverName, Len(WriteMailData.RecieverName) - 1)
                    End If
                Else
                    If Len(WriteMailData.RecieverName) < 10 Then
                        If Game_ValidCharacter(KeyAscii) Then WriteMailData.RecieverName = WriteMailData.RecieverName & Chr$(KeyAscii)
                    End If
                End If
            Case wmSubject
                If KeyAscii = 8 Then
                    If Len(WriteMailData.Subject) > 0 Then
                        WriteMailData.Subject = Left$(WriteMailData.Subject, Len(WriteMailData.Subject) - 1)
                    End If
                Else
                    If Len(WriteMailData.Subject) < 30 Then
                        If Game_ValidCharacter(KeyAscii) Then WriteMailData.Subject = WriteMailData.Subject & Chr$(KeyAscii)
                    End If
                End If
            Case wmMessage
                If KeyAscii = 8 Then
                    If Len(WriteMailData.Message) > 0 Then
                        WriteMailData.Message = Left$(WriteMailData.Message, Len(WriteMailData.Message) - 1)
                    End If
                Else
                    If Len(WriteMailData.Message) < 500 Then
                        If Game_ValidCharacter(KeyAscii) Then WriteMailData.Message = WriteMailData.Message & Chr$(KeyAscii)
                    End If
                End If
            End Select
        End If
    
    '*****************************
    '***** Text input buffer *****
    '*****************************
    Else
        If EnterText Then
            'Backspace
            If KeyAscii = 8 Then
                If Len(EnterTextBuffer) > 0 Then EnterTextBuffer = Left$(EnterTextBuffer, Len(EnterTextBuffer) - 1)
                b = True
            End If
            'Add to text buffer
            If Game_ValidCharacter(KeyAscii) Then
                If Len(EnterTextBuffer) < 85 Then
                    If Game_ValidCharacter(KeyAscii) Then
                        EnterTextBuffer = EnterTextBuffer & Chr$(KeyAscii)
                        b = True
                    End If
                End If
            End If
            'Update size
            If b Then
                EnterTextBufferWidth = Engine_GetTextWidth(EnterTextBuffer)
                UpdateShownTextBuffer
                LastClickedWindow = 0
            End If
        Else
            'Auto-write a reply to the last person to whisper to us
            If KeyAscii = 114 Then  'Key R
                If LastWhisperName <> "" Then
                    EnterText = True
                    EnterTextBuffer = "/tell " & LastWhisperName & " "
                    EnterTextBufferWidth = Engine_GetTextWidth(EnterTextBuffer)
                    LastClickedWindow = 0
                End If
            End If
        End If
    End If
    
    'Clear the key
    KeyAscii = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Dim i As Integer
    
    'Reset skins (F12)
    If GetAsyncKeyState(vbKeyShift) Then
        If KeyCode = vbKeyF12 Then
            Engine_Init_GUI 0
            Game_Config_Save
        End If
    End If

    'Delete mail (Delete)
    If KeyCode = vbKeyDelete Then
        If LastClickedWindow = MailboxWindow Then
            If ShowGameWindow(MailboxWindow) Then
                If SelMessage > 0 Then
                    sndBuf.Allocate 2
                    sndBuf.Put_Byte DataCode.Server_MailDelete
                    sndBuf.Put_Byte SelMessage
                End If
            End If
        End If
    End If

    'Send an emoticon - but make sure we're not typing or entering in a mail message
    If EnterText = False Then
        If Not LastClickedWindow = WriteMessageWindow Then
            If Not LastClickedWindow = AmountWindow Then
                If Not ShowGameWindow(WriteMessageWindow) Then
                    Select Case KeyCode
                    Case vbKey1
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Dots
                    Case vbKey2
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Exclimation
                    Case vbKey3
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Question
                    Case vbKey4
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Surprised
                    Case vbKey5
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Heart
                    Case vbKey6
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Hearts
                    Case vbKey7
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.HeartBroken
                    Case vbKey8
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Utensils
                    Case vbKey9
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.Meat
                    Case vbKey0
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Emote
                        sndBuf.Put_Byte EmoID.ExcliQuestion
                    End Select
                End If
            End If
        End If
    End If
    
    'Clear the key
    KeyCode = 0

End Sub

Private Sub PTDTmr_Timer()

    sndBuf.Put_Byte DataCode.Server_PTD
    PTDSTime = timeGetTime

End Sub

Private Sub ShutdownTimer_Timer()

    On Error Resume Next    'Who cares about an error if we are closing down
    
    'Quit the client - we must user a timer since DoEvents wont work (since we're not multithreaded)

    'Close down the socket
    GOREsock_ShutDown
    GOREsock_UnHook
    If GOREsock_Loaded Then
        GOREsock_Terminate
    Else

        'Unload the engine
        Engine_Init_UnloadTileEngine
        
        'Unload the forms
        Engine_UnloadAllForms
        
        'Unload everything else
        End

    End If

End Sub

Private Function SplitCommandFromString(StringBuffer As String) As String

Dim TempSplit() As String
Dim i As Integer
    
    If StringBuffer = "" Then Exit Function
    If Len(StringBuffer) < 2 Then Exit Function
    
    TempSplit() = Split(StringBuffer, " ")
    
    If UBound(TempSplit) = 0 Then Exit Function
    
    For i = 1 To UBound(TempSplit)
        SplitCommandFromString = SplitCommandFromString & TempSplit(i) & " "
    Next i
    SplitCommandFromString = Left$(SplitCommandFromString, Len(SplitCommandFromString) - 1)

End Function

