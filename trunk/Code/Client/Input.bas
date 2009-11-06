Attribute VB_Name = "Input"
Option Explicit

Public DI As DirectInput8
Public DIDevice As DirectInputDevice8
Public MousePos As POINTAPI
Public MousePosAdd As POINTAPI
Public MouseEvent As Long
Public MouseLeftDown As Byte
Public MouseRightDown As Byte

Private Const KeyPress_Shift As Integer = 2 ^ 12
Private Const KeyPress_Control As Integer = 2 ^ 13
Private Const KeyPress_Alt As Integer = 2 ^ 14

Private Type KeyDefinitions
    MiniMap As Integer
    PickUpObj As Integer
    QuickBar(1 To 12) As Integer
    Attack As Integer
    ChatBufferUp As Integer
    ChatBufferDown As Integer
    InventoryWindow As Integer
    QuickBarWindow As Integer
    ChatWindow As Integer
    StatWindow As Integer
    MenuWindow As Integer
    ZoomIn As Integer
    ZoomOut As Integer
    MoveNorth As Integer
    MoveEast As Integer
    MoveSouth As Integer
    MoveWest As Integer
    ResetGUI As Integer
    QuickTarget As Integer
    QuickReply As Integer
End Type
Private KeyDefinitions As KeyDefinitions

Private IgnoreNextChatKey As Boolean    'Used to ignore the next keystroke going into the chat buffer (for pressing the quick-reply button)

Private Function Input_Keys_IsPressed(ByVal DefinitionValue As Integer, ByVal KeyCode As Integer) As Boolean

'*****************************************************************
'Checks if the definition requirements are met
'*****************************************************************
Dim CheckForInput As Boolean

    CheckForInput = True

    'Check for shift, alt and control requirements
    If DefinitionValue And KeyPress_Shift Then
        If GetAsyncKeyState(16) = 0 Then Exit Function
    End If
    If DefinitionValue And KeyPress_Control Then
        If GetAsyncKeyState(17) = 0 Then Exit Function
        CheckForInput = False   'No need to check for input if control is pressed
    End If
    If DefinitionValue And KeyPress_Alt Then
        If GetAsyncKeyState(18) = 0 Then Exit Function
        CheckForInput = False   'No need to check for input if alt is pressed
    End If
    
    'Remove the shift, alt and control bits, then check for the keycode requirements
    If (DefinitionValue And 2047) <> KeyCode Then Exit Function
    
    'Check for input boxes being active so we don't run commands when typing
    If CheckForInput Then
    
        'Typing in the chat buffer
        If EnterText Then Exit Function
    
        'Writing a message in the mail window
        If LastClickedWindow = WriteMessageWindow Then
            If ShowGameWindow(WriteMessageWindow) <> 0 Then Exit Function
        End If
        
        'Numeric only
        If Input_Keys_IsNumeric(KeyCode) Then
            
            'Entering a value in the amount window
            If LastClickedWindow = AmountWindow Then
                If ShowGameWindow(AmountWindow) <> 0 Then Exit Function
            End If
            
            'Entering a number on the NPC chat window
            If LastClickedWindow = NPCChatWindow Then
                If ShowGameWindow(NPCChatWindow) <> 0 Then Exit Function
            End If
            
        End If

    End If

    'Every test has been passed
    Input_Keys_IsPressed = True

End Function

Public Sub Input_Keys_LoadDefinitions()

'*****************************************************************
'Load the key definitions
'*****************************************************************
Dim i As Long

    KeyDefinitions.Attack = Val(Var_Get(DataPath & "Game.ini", "INPUT", "Attack"))
    KeyDefinitions.ChatBufferDown = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatBufferDown"))
    KeyDefinitions.ChatBufferUp = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatBufferUp"))
    KeyDefinitions.ChatWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatWindow"))
    KeyDefinitions.InventoryWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "InventoryWindow"))
    KeyDefinitions.MenuWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MenuWindow"))
    KeyDefinitions.MiniMap = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MiniMap"))
    KeyDefinitions.MoveEast = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveEast"))
    KeyDefinitions.MoveNorth = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveNorth"))
    KeyDefinitions.MoveSouth = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveSouth"))
    KeyDefinitions.MoveWest = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveWest"))
    KeyDefinitions.PickUpObj = Val(Var_Get(DataPath & "Game.ini", "INPUT", "PickUpObj"))
    KeyDefinitions.QuickBarWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickBarWindow"))
    KeyDefinitions.StatWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "StatWindow"))
    KeyDefinitions.ZoomIn = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ZoomIn"))
    KeyDefinitions.ZoomOut = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ZoomOut"))
    KeyDefinitions.ResetGUI = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ResetGUI"))
    KeyDefinitions.QuickTarget = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickTarget"))
    KeyDefinitions.QuickReply = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickReply"))
    For i = 1 To 12
        KeyDefinitions.QuickBar(i) = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickBar" & i))
    Next i

End Sub

Public Sub Input_Keys_ClearQueue()

'*****************************************************************
'Clears the GetAsyncKeyState queue to prevent key presses from a long time
' ago falling into "have been pressed"
'*****************************************************************
Dim i As Long

    For i = 1 To 145
        GetAsyncKeyState i
    Next i

End Sub

Private Function Input_GetCommand(ByVal CommandString As String) As Boolean

'*****************************************************************
'Checks if "CommandString" is the command entered in the buffer
'Partial matches return true, too, such as for example:
'Entered: /quit                CommandString: /qui
'*****************************************************************

    'Check for the command passed
    If UCase$(Left$(EnterTextBuffer, Len(CommandString))) = UCase$(CommandString) Then Input_GetCommand = True Else Input_GetCommand = False

End Function

Private Function Input_GetBufferArgs() As String

'*****************************************************************
'Returns the arguments for a command entered into the chat buffer
' (basically cuts off the command and the space after it)
'*****************************************************************
Dim s() As String

    'Split between the first space only
    s = Split(EnterTextBuffer, " ", 2)
    
    'Return the parameters if they exist
    If UBound(s) > 0 Then Input_GetBufferArgs = Trim$(s(1))

End Function

Public Sub Input_Init()

'*****************************************************************
'Init the input devices (keyboard and mouse)
'*****************************************************************
Dim diProp As DIPROPLONG

    'Create the device
    Set DI = DX.DirectInputCreate
    Set DIDevice = DI.CreateDevice("guid_SysMouse")
    
    Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
    
    'If in windowed mode, free the mouse from the screen
    If Windowed Then
        Call DIDevice.SetCooperativeLevel(frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    Else
        Call DIDevice.SetCooperativeLevel(frmMain.hwnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
    End If
    
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 50
    Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
    MouseEvent = DX.CreateEvent(frmMain)
    DIDevice.SetEventNotification MouseEvent

End Sub

Sub Input_Keys_Press(ByVal KeyAscii As Integer)

'*****************************************************************
'Handles input entering to windows (mostly just alphanumeric)
'*****************************************************************
Dim StartGold As Long
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

    '*************************
    '***** Trade window ******
    '*************************
    ElseIf LastClickedWindow = TradeWindow Then
        StartGold = TradeTable.Gold1
        'Backspace
        If KeyAscii = 8 Then
            If Len(Str$(TradeTable.Gold1)) > 0 Then
                If Len(Str$(TradeTable.Gold1)) - 1 <= 1 Then
                    TradeTable.Gold1 = 0
                Else
                    TradeTable.Gold1 = Left$(Str$(TradeTable.Gold1), Len(Str$(TradeTable.Gold1)) - 1)
                End If
            End If
        End If
        'Number
        If IsNumeric(Chr$(KeyAscii)) Then
            If Len(Str$(TradeTable.Gold1) & Chr$(KeyAscii)) < Len(Str$(MAXLONG)) Then
                TradeTable.Gold1 = Val(Str$(TradeTable.Gold1) & Chr$(KeyAscii))
                If TradeTable.Gold1 > MAXLONG Then TradeTable.Gold1 = MAXLONG
            Else
                TradeTable.Gold1 = MAXLONG
            End If
            If TradeTable.Gold1 > BaseStats(SID.Gold) Then TradeTable.Gold1 = BaseStats(SID.Gold)

        End If
        'Check if the gold has changed, if so update it on the server
        If TradeTable.Gold1 <> StartGold Then
            sndBuf.Put_Byte DataCode.User_Trade_UpdateTrade
            sndBuf.Put_Byte 0
            sndBuf.Put_Long TradeTable.Gold1
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
            
            'Check if to ignore this keystroke
            If IgnoreNextChatKey Then
                IgnoreNextChatKey = False
            Else
                
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
                
            End If
            
        End If
    End If

End Sub

Private Sub Input_Keys_Down_Return()

'*****************************************************************
'Return was pressed down
'*****************************************************************
Dim j As Long
Dim i As Long

    '*************************
    '***** Amount window *****
    '*************************
    If LastClickedWindow = AmountWindow Then
        If AmountWindowItemIndex Then
            If AmountWindowValue <> vbNullString Then
                If IsNumeric(AmountWindowValue) Then
                    'Drop into mail
                    If AmountWindowUsage = AW_InvToMail Then
                        'Check for duplicate entries
                        For j = 1 To MaxMailObjs
                            If WriteMailData.ObjIndex(j) = AmountWindowItemIndex Then
                                ShowGameWindow(AmountWindow) = 0
                                AmountWindowUsage = 0
                                If LastClickedWindow = AmountWindow Then LastClickedWindow = 0
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
                                If LastClickedWindow = AmountWindow Then LastClickedWindow = 0
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
                    'Put in trade
                    ElseIf AmountWindowUsage = AW_InvToTrade Then
                        sndBuf.Put_Byte DataCode.User_Trade_UpdateTrade
                        sndBuf.Put_Byte AmountWindowItemIndex
                        sndBuf.Put_Long CInt(AmountWindowValue)
                             
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
                If LastClickedWindow = AmountWindow Then LastClickedWindow = 0
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
                        If LastClickedWindow = WriteMessageWindow Then LastClickedWindow = 0
                        LastMailSendTime = timeGetTime
                    End If
                End If
            End If
        End If
        
    End If
    
    '***********************
    '***** Chat screen *****
    '***********************
    If LastClickedWindow <> WriteMessageWindow Then
        If LastClickedWindow <> ViewMessageWindow Then
            If LastClickedWindow <> AmountWindow Then
                If EnterText = True Then
                    If EnterTextBuffer <> vbNullString Then Input_HandleCommands
                    EnterText = False
                Else
                    EnterText = True
                End If
            End If
        End If
    End If

End Sub

Private Function Input_Keys_IsNumeric(ByVal KeyCode As Integer) As Boolean

'*****************************************************************
'Check if a numeric key (0 to 9) was pressed
'*****************************************************************

    '0 = 48
    '9 = 57
    If KeyCode > 47 Then
        If KeyCode < 58 Then
            Input_Keys_IsNumeric = True
        End If
    End If

End Function

Private Function Input_Keys_IsAlpha(ByVal KeyCode As Integer) As Boolean

'*****************************************************************
'Check if an alphabet key (A to Z) was pressed
'*****************************************************************

    'a = 65
    'z = 90
    If KeyCode > 64 Then
        If KeyCode < 91 Then
            Input_Keys_IsAlpha = True
        End If
    End If
    
End Function

Private Function Input_Keys_IsAlphaNumeric(ByVal KeyCode As Integer) As Boolean

'*****************************************************************
'Check if an alphanumeric key (A to Z, 0 to 9) was pressed
'*****************************************************************

    Input_Keys_IsAlphaNumeric = (Input_Keys_IsNumeric(KeyCode) And Input_Keys_IsAlpha(KeyCode))

End Function

Sub Input_Keys_Down(ByVal KeyCode As Integer)

'*****************************************************************
'Checks keys and respond
'*****************************************************************
Dim i As Long

    'Return was pressed
    If KeyCode = vbKeyReturn Then
        Input_Keys_Down_Return
        Exit Sub
    End If
    
    'Escape was pressed
    If KeyCode = vbKeyEscape Then
        If LastClickedWindow = 0 Then
            If ShowGameWindow(MenuWindow) = 0 Then
                If EnterText Then
                    EnterTextBuffer = vbNullString
                    EnterTextBufferWidth = 10
                    UpdateShownTextBuffer
                    EnterText = False
                End If
            End If
        Else
            ShowGameWindow(LastClickedWindow) = 0
            LastClickedWindow = 0
            Exit Sub
        End If
    End If
    
    'Hide/show the mini-map
    If Input_Keys_IsPressed(KeyDefinitions.MiniMap, KeyCode) Then
        If ShowMiniMap = 0 Then ShowMiniMap = 1 Else ShowMiniMap = 0
    End If
    
    'Get object off ground (alt)
    If Input_Keys_IsPressed(KeyDefinitions.PickUpObj, KeyCode) Then
        If Engine_OBJ_AtTile(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y) Then
            If LastLootTime < timeGetTime Then
                LastLootTime = timeGetTime + LootDelay
                sndBuf.Put_Byte DataCode.User_Get
            End If
        End If
    End If
    
    'Use the quick bar
    For i = 1 To 12
        If Input_Keys_IsPressed(KeyDefinitions.QuickBar(i), KeyCode) Then
            Engine_UseQuickBar KeyCode - vbKeyF1 + 1
        End If
    Next i
    
    'Attack key
    If Input_Keys_IsPressed(KeyDefinitions.Attack, KeyCode) Then
        If UserCharIndex > 0 Then
            If LastAttackTime < timeGetTime Then
                LastAttackTime = timeGetTime + AttackDelay
                
                'Check for a valid attacking distance
                If UserAttackRange > 1 Then
                    If TargetCharIndex > 0 Then
                        If TargetCharIndex <> UserCharIndex Then
                            If Engine_Distance(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(TargetCharIndex).Pos.X, CharList(TargetCharIndex).Pos.Y) <= UserAttackRange Then
                                LastAttackTime = timeGetTime
                                sndBuf.Allocate 2
                                sndBuf.Put_Byte DataCode.User_Attack
                                sndBuf.Put_Byte CharList(UserCharIndex).Heading
                            Else
                                Engine_AddToChatTextBuffer Message(91), FontColor_Fight
                            End If
                        End If
                    End If
                Else
                    If Engine_UserIsFacingChar Then
                        sndBuf.Allocate 2
                        sndBuf.Put_Byte DataCode.User_Attack
                        sndBuf.Put_Byte CharList(UserCharIndex).Heading
                    End If
                End If
                
            End If
        End If
    End If
    
    'Chat buffer scrolling
    If Input_Keys_IsPressed(KeyDefinitions.ChatBufferUp, KeyCode) Then
        If ShowGameWindow(ChatWindow) Then
            ChatBufferChunk = ChatBufferChunk + 0.5
            Engine_UpdateChatArray
        End If
    End If
    If Input_Keys_IsPressed(KeyDefinitions.ChatBufferDown, KeyCode) Then
        If ShowGameWindow(ChatWindow) Then
            If ChatBufferChunk > 1 Then
                ChatBufferChunk = ChatBufferChunk - 0.5
                Engine_UpdateChatArray
            End If
        End If
    End If
    
    'Hide/show windows
    If Input_Keys_IsPressed(KeyDefinitions.InventoryWindow, KeyCode) Then
        If ShowGameWindow(InventoryWindow) Then
            ShowGameWindow(InventoryWindow) = 0
            If LastClickedWindow = InventoryWindow Then LastClickedWindow = 0
        Else
            ShowGameWindow(InventoryWindow) = 1
            LastClickedWindow = InventoryWindow
        End If
    End If
    If Input_Keys_IsPressed(KeyDefinitions.QuickBarWindow, KeyCode) Then
        If ShowGameWindow(QuickBarWindow) Then
            ShowGameWindow(QuickBarWindow) = 0
            If LastClickedWindow = QuickBarWindow Then LastClickedWindow = 0
        Else
            ShowGameWindow(QuickBarWindow) = 1
            LastClickedWindow = QuickBarWindow
        End If
    End If
    If Input_Keys_IsPressed(KeyDefinitions.ChatWindow, KeyCode) Then
        If ShowGameWindow(ChatWindow) Then
            ShowGameWindow(ChatWindow) = 0
            If LastClickedWindow = ChatWindow Then LastClickedWindow = 0
        Else
            ShowGameWindow(ChatWindow) = 1
            LastClickedWindow = ChatWindow
        End If
    End If
    If Input_Keys_IsPressed(KeyDefinitions.StatWindow, KeyCode) Then
        If ShowGameWindow(StatWindow) Then
            ShowGameWindow(StatWindow) = 0
            If LastClickedWindow = StatWindow Then LastClickedWindow = 0
        Else
            ShowGameWindow(StatWindow) = 1
            LastClickedWindow = StatWindow
        End If
    End If
    If Input_Keys_IsPressed(KeyDefinitions.MenuWindow, KeyCode) Then
        If ShowGameWindow(MenuWindow) Then
            ShowGameWindow(MenuWindow) = 0
            If LastClickedWindow = MenuWindow Then LastClickedWindow = 0
        Else
            ShowGameWindow(MenuWindow) = 1
            LastClickedWindow = MenuWindow
        End If
    End If
    
    'Reset skin positions
    If Input_Keys_IsPressed(KeyDefinitions.ResetGUI, KeyCode) Then
        Engine_Init_GUI 0
        Game_Config_Save
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
    
    'Auto-write a reply to the last person to whisper to us
    If Input_Keys_IsPressed(KeyDefinitions.QuickReply, KeyCode) Then
        If LenB(LastWhisperName) <> 0 Then
            EnterText = True
            EnterTextBuffer = "/tell " & LastWhisperName & " "
            EnterTextBufferWidth = Engine_GetTextWidth(EnterTextBuffer)
            IgnoreNextChatKey = True
            UpdateShownTextBuffer
            LastClickedWindow = 0
        End If
    End If
    
    'Target the closest character
    If Input_Keys_IsPressed(KeyDefinitions.QuickTarget, KeyCode) Then
        i = Game_ClosestTargetNPC
        If i > 0 Then
            sndBuf.Allocate 3
            sndBuf.Put_Byte DataCode.User_Target
            sndBuf.Put_Integer i
        End If
    End If

    'Send an emoticon - but make sure we're not typing or entering in a mail message
    If EnterText = False Then
        If Not LastClickedWindow = WriteMessageWindow Then
            If Not LastClickedWindow = AmountWindow Then
                If ShowGameWindow(WriteMessageWindow) = 0 Then
                    If ShowGameWindow(NPCChatWindow) = 0 Then
                        If EmoticonDelay < timeGetTime Then
                            EmoticonDelay = timeGetTime + 2000  'Wait 2000ms (two seconds) between emoticon usages
                            
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
                        
                    Else
                        
                        If KeyCode >= 49 Then
                            If KeyCode - 48 <= GameWindow.NPCChat.NumAnswers Then
                                i = NPCChat(ActiveAsk.ChatIndex).Ask.Ask(ActiveAsk.AskIndex).Answer(KeyCode - 48).GotoID
                                If i > 0 Then
                                    Engine_ShowNPCChatWindow ActiveAsk.AskName, ActiveAsk.ChatIndex, i
                                Else
                                    ShowGameWindow(NPCChatWindow) = 0
                                    If LastClickedWindow = NPCChatWindow Then LastClickedWindow = 0
                                End If
                            End If
                        End If
                    
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub Input_HandleCommands()

'*****************************************************************
'Handles all the chat commands - when aborting, use either GoTo CleanUp
' to ignore the keystroke (buffer is not cleared) or GoTo CleanUp to
' clear the buffer, too (its all just about preference)
'*****************************************************************
Dim TempS() As String
Dim s As String
Dim s2 As String
Dim i As Long
Dim j As Long

    '***** Check for commands *****
    If Input_GetCommand("/BLI") Then
        sndBuf.Put_Byte DataCode.User_Blink
        
    ElseIf Input_GetCommand("/LOOKL") Then
        sndBuf.Put_Byte DataCode.User_LookLeft
        
    ElseIf Input_GetCommand("/LOOKR") Then
        sndBuf.Put_Byte DataCode.User_LookRight
        
    ElseIf Input_GetCommand("/WHO") Then
        sndBuf.Put_Byte DataCode.Server_Who
        
    ElseIf Input_GetCommand("/SH") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.Comm_Shout
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/GINFO") Or Input_GetCommand("/GROUPI") Then
        sndBuf.Put_Byte DataCode.User_Group_Info
        
    ElseIf Input_GetCommand("/TELL") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        TempS() = Split(s, " ", 2)
        If UBound(TempS) < 1 Then GoTo CleanUp
        If LenB(Trim$(TempS(0))) = 0 Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.Comm_Whisper
        sndBuf.Put_String Trim$(TempS(0))
        sndBuf.Put_String Trim$(TempS(1))
        
    ElseIf Input_GetCommand("/DEP") Then
        j = Val(Input_GetBufferArgs)
        If j <= 0 Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.User_Bank_Deposit
        sndBuf.Put_Long j
        'We will assume that the deposit was successful
        Engine_AddToChatTextBuffer Replace$(Message(118), "<amount>", Str(j)), FontColor_Info
        
    ElseIf Input_GetCommand("/WITH") Then
        j = Val(Input_GetBufferArgs)
        If j <= 0 Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.User_Bank_Withdraw
        sndBuf.Put_Long j
    ElseIf Input_GetCommand("/TRADE") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then
            Engine_AddToChatTextBuffer Message(136), FontColor_Info
            GoTo CleanUp
        End If
        If UCase$(s) = UCase$(CharList(UserCharIndex).Name) Then
            Engine_AddToChatTextBuffer Message(133), FontColor_Info
            GoTo CleanUp
        End If
        sndBuf.Put_Byte DataCode.User_Trade_Trade
        sndBuf.Put_String s
    ElseIf Input_GetCommand("/BALAN") Then
        sndBuf.Put_Byte DataCode.User_Bank_Balance
        
    ElseIf Input_GetCommand("/G ") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.Comm_GroupTalk
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/CREATEG") Or Input_GetCommand("/MAKEG") Or Input_GetCommand("/NEWG") Then
        sndBuf.Put_Byte DataCode.User_Group_Make
    
    ElseIf Input_GetCommand("/INVITE") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.User_Group_Invite
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/LEAVEG") Or Input_GetCommand("/EXITG") Then
        sndBuf.Put_Byte DataCode.User_Group_Leave
        
    ElseIf Input_GetCommand("/JOING") Then
        sndBuf.Put_Byte DataCode.User_Group_Join
        
    ElseIf Input_GetCommand("/ME") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.Comm_Emote
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/EM") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.Comm_Emote
        sndBuf.Put_String s

    ElseIf Input_GetCommand("/LANG") Then
        s = LCase$(Input_GetBufferArgs)
        If s = vbNullString Then GoTo CleanUp
        If Engine_FileExist(MessagePath & s & "*.ini", vbNormal) Then
            s = Dir$(MessagePath & s & "*.ini", vbNormal)
            s = Left$(s, Len(s) - 4)
            s = Engine_Init_Messages(s)
            Engine_Init_Signs s
            Var_Write DataPath & "Game.ini", "INIT", "Language", s
            Engine_AddToChatTextBuffer Replace$(Message(90), "<lang>", s), FontColor_Info
        Else
            Engine_AddToChatTextBuffer Message(87), FontColor_Info
        End If
        
    ElseIf Input_GetCommand("/SKIN") Then
        s = LCase$(Input_GetBufferArgs)
        If s = vbNullString Then
            Engine_AddToChatTextBuffer Engine_BuildSkinsList, FontColor_Info
            GoTo CleanUp
        End If
        If Engine_FileExist(DataPath & "Skins\" & s & "*.ini", vbNormal) Then
            s = Dir$(DataPath & "Skins\" & s & "*.ini", vbNormal)
            CurrentSkin = Left$(s, Len(s) - 4)
            Engine_Init_GUI 0
            Var_Write DataPath & "Game.ini", "INIT", "CurrentSkin", CurrentSkin
            Engine_AddToChatTextBuffer Replace$(Message(89), "<skin>", CurrentSkin), FontColor_Info
        Else
            Engine_AddToChatTextBuffer Message(88), FontColor_Info
        End If
        
    ElseIf Input_GetCommand("/QUEST") Then
        If QuestInfoUBound = 0 Then
            'No quests in place
            Engine_AddToChatTextBuffer Message(103), FontColor_Quest
        Else
            j = Val(Input_GetBufferArgs)
            If j < 1 Or j > QuestInfoUBound Then
                'No valid number specified, give the list
                Engine_AddToChatTextBuffer Message(104), FontColor_Quest
                For i = 1 To QuestInfoUBound
                    Engine_AddToChatTextBuffer "  " & i & ". " & QuestInfo(i).Name, FontColor_Quest
                Next i
            Else
                'Give the info on the specific quest
                Engine_AddToChatTextBuffer QuestInfo(j).Name & ":", FontColor_Quest
                Engine_AddToChatTextBuffer QuestInfo(j).Desc, FontColor_Quest
            End If
        End If
        
    ElseIf Input_GetCommand("/CANCELQUEST") Or Input_GetCommand("/ENDQUEST") Then
        If QuestInfoUBound = 0 Then GoTo CleanUp
        j = Val(Input_GetBufferArgs)
        If j < 1 Or j > QuestInfoUBound Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.User_CancelQuest
        sndBuf.Put_Byte CByte(j)
                
    ElseIf Input_GetCommand("/THR") Then
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
        
    ElseIf Input_GetCommand("/DETHR") Then
        sndBuf.Put_Byte DataCode.GM_DeThrall
        
    ElseIf Input_GetCommand("/QUIT") Then
        IsUnloading = 1
        
    ElseIf Input_GetCommand("/ACCEPT") Then
        sndBuf.Put_Byte DataCode.User_StartQuest
        
    ElseIf Input_GetCommand("/DESC") Then
        s = Input_GetBufferArgs
        sndBuf.Put_Byte DataCode.User_Desc
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/HELP") Then
        sndBuf.Put_Byte DataCode.Server_Help
        
    ElseIf Input_GetCommand("/APPR") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.GM_Approach
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/SUM") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.GM_Summon
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/SETGM") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        TempS = Split(s, " ")
        If UBound(TempS) > 0 Then
            If IsNumeric(TempS(1)) Then
                sndBuf.Allocate 3 + Len(TempS(0))
                sndBuf.Put_Byte DataCode.GM_SetGMLevel
                sndBuf.Put_String TempS(0)
                sndBuf.Put_Byte CByte(TempS(1))
            End If
        End If
        
    ElseIf Input_GetCommand("/CLICKWARP") Then
        If UseClickWarp = 1 Then UseClickWarp = 0 Else UseClickWarp = 1
        Engine_AddToChatTextBuffer Replace$(Message(124), "<value>", UseClickWarp), FontColor_Info
        
    ElseIf Input_GetCommand("/BANIP") Then
        s = Input_GetBufferArgs 'Remove the command
        If LenB(s) < 4 Then 'Not enough information entered
            Engine_AddToChatTextBuffer Message(92), FontColor_Info
            GoTo CleanUp
        End If
        TempS = Split(s, " ", 2)    'Split up the IP and reason
        If UBound(TempS) = 0 Then
            Engine_AddToChatTextBuffer Message(93), FontColor_Info
            GoTo CleanUp
        Else
            s = TempS(0)
            s2 = TempS(1)
        End If
        TempS = Split(s, ".")
        If UBound(TempS) <> 3 Then
            Engine_AddToChatTextBuffer Message(92), FontColor_Info
            GoTo CleanUp
        End If
        For j = 0 To 3
            If Val(TempS(j)) < 0 Or Val(TempS(j)) > 255 Then
                Engine_AddToChatTextBuffer Message(92), FontColor_Info
                GoTo CleanUp
            End If
        Next j
        sndBuf.Put_Byte DataCode.GM_BanIP
        sndBuf.Put_String Trim$(s)
        sndBuf.Put_String Trim$(s2)
        
    ElseIf Input_GetCommand("/UNBANIP") Then
        s = Input_GetBufferArgs 'Remove the command
        If LenB(s) < 4 Then 'Not enough information entered
            Engine_AddToChatTextBuffer Message(92), FontColor_Info
            GoTo CleanUp
        End If
        TempS = Split(s, ".")
        If UBound(TempS) <> 3 Then
            Engine_AddToChatTextBuffer Message(92), FontColor_Info
            GoTo CleanUp
        End If
        For j = 0 To 3
            If TempS(j) <> "*" Then
                If Val(TempS(j)) < 0 Or Val(TempS(j)) > 255 Then
                    Engine_AddToChatTextBuffer Message(92), FontColor_Info
                    GoTo CleanUp
                End If
            End If
        Next j
        sndBuf.Put_Byte DataCode.GM_UnBanIP
        sndBuf.Put_String Trim$(s)
        
    ElseIf Input_GetCommand("/KICK") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.GM_Kick
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/SEARCHI") Or Input_GetCommand("/FINDI") Or Input_GetCommand("/FINDO") Or Input_GetCommand("/SEARCHO") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.GM_FindItem
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/GIVESK") Or Input_GetCommand("/GIVESP") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        TempS = Split(s, " ")
        If UBound(TempS) <> 1 Then GoTo CleanUp
        If Val(TempS(1)) <= 0 Or Val(TempS(1)) > 255 Then Exit Sub
        sndBuf.Put_Byte DataCode.GM_GiveSkill
        sndBuf.Put_String TempS(0)
        sndBuf.Put_Long Val(TempS(1))
        
    ElseIf Input_GetCommand("/SQL") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        sndBuf.Put_Byte DataCode.GM_SQL
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/KILLMAP") Then
        sndBuf.Put_Byte DataCode.GM_KillMap

    ElseIf Input_GetCommand("/KILL") Then
        If TargetCharIndex = UserCharIndex Or TargetCharIndex = 0 Then
            Engine_AddToChatTextBuffer "Suicide is not the answer...", FontColor_Info
        Else
            sndBuf.Put_Byte DataCode.GM_Kill
        End If
        
    ElseIf Input_GetCommand("/GIVEGO") Then
        s = Input_GetBufferArgs
        If Val(s) <= 0 Or Val(s) > MAXLONG Then
            Engine_AddToChatTextBuffer "Please enter an amount greater than 0.", FontColor_Info
            GoTo CleanUp
        End If
        sndBuf.Put_Byte DataCode.GM_GiveGold
        sndBuf.Put_Long Val(s)
        
    ElseIf Input_GetCommand("/GIVEOBJ") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        TempS = Split(s, " ")
        If UBound(TempS) <> 1 Then
            Engine_AddToChatTextBuffer "Please use the format: <ObjIndex> <Amount>", FontColor_Info
            GoTo CleanUp
        End If
        If Val(TempS(0)) <= 0 Or Val(TempS(0)) > MAXINT Then
            Engine_AddToChatTextBuffer "Invalid ObjIndex parameter - enter a value between 1 and " & MAXINT & ".", FontColor_Info
            GoTo CleanUp
        End If
        If Val(TempS(1)) <= 0 Or Val(TempS(1)) > MAXINT Then
            Engine_AddToChatTextBuffer "Invalid Amount parameter - enter a value between 1 and " & MAXINT & ".", FontColor_Info
            GoTo CleanUp
        End If
        sndBuf.Put_Byte DataCode.GM_GiveObject
        sndBuf.Put_Integer Val(TempS(0))
        sndBuf.Put_Integer Val(TempS(1))
        
    ElseIf Input_GetCommand("/WARP") Then
        i = Val(Input_GetBufferArgs)
        If Not Engine_FileExist(MapPath & i & ".map", vbNormal) Then
            Engine_AddToChatTextBuffer "Please enter a valid map number.", FontColor_Info
            GoTo CleanUp
        End If
        sndBuf.Put_Byte DataCode.GM_WarpToMap
        sndBuf.Put_Integer i
    
    ElseIf Input_GetCommand("/IPINFO") Then
        s = Input_GetBufferArgs
        If s = vbNullString Then GoTo CleanUp
        TempS = Split(s, ".")   'All of this is just a check for a valid IP
        If UBound(TempS) <> 3 Then  'Check for 3 periods
            Engine_AddToChatTextBuffer Message(92), FontColor_Info
            GoTo CleanUp
        End If
        For j = 0 To 3  'Check for values between 0 and 255
            If Val(TempS(j)) < 0 Or Val(TempS(j)) > 255 Then
                Engine_AddToChatTextBuffer Message(92), FontColor_Info
                GoTo CleanUp
            End If
        Next j
        sndBuf.Put_Byte DataCode.GM_IPInfo
        sndBuf.Put_String s
        
    ElseIf Input_GetCommand("/BANLIST") Then
        sndBuf.Put_Byte DataCode.GM_BanList
        
    ElseIf Input_GetCommand("/RAISE") Then
        TempS() = Split(Input_GetBufferArgs, " ")
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
        EnterTextBuffer = Trim$(EnterTextBuffer)
        sndBuf.Allocate 2 + Len(EnterTextBuffer)
        sndBuf.Put_Byte DataCode.Comm_Talk
        sndBuf.Put_String EnterTextBuffer
        
        'We just sent a chat message, so check if it had triggers!
        Engine_NPCChat_CheckForChatTriggers EnterTextBuffer
        
    End If
    
CleanUp:
    
    'Cleans up the buffer
    EnterTextBuffer = vbNullString
    EnterTextBufferWidth = 10
    ShownText = vbNullString

End Sub

Sub Input_Keys_General()

'*****************************************************************
'Checks keys and respond
'*****************************************************************

    If GetActiveWindow = 0 Then Exit Sub
    
    'Dont move when Control is pressed
    If GetAsyncKeyState(vbKeyControl) Then Exit Sub

    'Check if certain screens are open that require ASDW keys
    If ShowGameWindow(WriteMessageWindow) Then
        If WMSelCon <> 0 Then Exit Sub
    End If

    'Zoom in / out
    If LastClickedWindow <> TradeWindow Then
        If LastClickedWindow <> ChatWindow Then
            If GetAsyncKeyState(KeyDefinitions.ZoomIn) Then       'In
                ZoomLevel = ZoomLevel + (ElapsedTime * 0.0003)
                If ZoomLevel > MaxZoomLevel Then ZoomLevel = MaxZoomLevel
            ElseIf GetAsyncKeyState(KeyDefinitions.ZoomOut) Then  'Out
                ZoomLevel = ZoomLevel - (ElapsedTime * 0.0003)
                If ZoomLevel < 0 Then ZoomLevel = 0
            End If
        End If
    End If

    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If GetAsyncKeyState(vbKeyTab) Then
            'Move Up-Right
            If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyRight) < 0 Then
                Engine_ChangeHeading NORTHEAST
                Exit Sub
            End If
            'Move Up-Left
            If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyLeft) < 0 Then
                Engine_ChangeHeading NORTHWEST
                Exit Sub
            End If
            'Move Down-Right
            If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyRight) < 0 Then
                Engine_ChangeHeading SOUTHEAST
                Exit Sub
            End If
            'Move Down-Left
            If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyLeft) < 0 Then
                Engine_ChangeHeading SOUTHWEST
                Exit Sub
            End If
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                Engine_ChangeHeading NORTH
                Exit Sub
            End If
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
                Engine_ChangeHeading EAST
                Exit Sub
            End If
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                Engine_ChangeHeading SOUTH
                Exit Sub
            End If
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                Engine_ChangeHeading WEST
                Exit Sub
            End If
            If EnterText = False Then
                If GetKeyState(KeyDefinitions.MoveNorth) < 0 And GetKeyState(KeyDefinitions.MoveEast) < 0 Then
                    Engine_ChangeHeading NORTHEAST
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveNorth) < 0 And GetKeyState(KeyDefinitions.MoveWest) < 0 Then
                    Engine_ChangeHeading NORTHWEST
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveSouth) < 0 And GetKeyState(KeyDefinitions.MoveEast) < 0 Then
                    Engine_ChangeHeading SOUTHEAST
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveSouth) < 0 And GetKeyState(KeyDefinitions.MoveWest) < 0 Then
                    Engine_ChangeHeading SOUTHWEST
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveNorth) < 0 Then
                    Engine_ChangeHeading NORTH
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveEast) < 0 Then
                    Engine_ChangeHeading EAST
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveSouth) < 0 Then
                    Engine_ChangeHeading SOUTH
                    Exit Sub
                End If
                If GetKeyState(KeyDefinitions.MoveWest) < 0 Then
                    Engine_ChangeHeading WEST
                    Exit Sub
                End If
            End If
        Else
            'Move Up-Right
            If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyRight) < 0 Then
                Engine_MoveUser NORTHEAST
                Exit Sub
            End If
            'Move Up-Left
            If GetKeyState(vbKeyUp) < 0 And GetKeyState(vbKeyLeft) < 0 Then
                Engine_MoveUser NORTHWEST
                Exit Sub
            End If
            'Move Down-Right
            If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyRight) < 0 Then
                Engine_MoveUser SOUTHEAST
                Exit Sub
            End If
            'Move Down-Left
            If GetKeyState(vbKeyDown) < 0 And GetKeyState(vbKeyLeft) < 0 Then
                Engine_MoveUser SOUTHWEST
                Exit Sub
            End If
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
                Engine_MoveUser NORTH
                Exit Sub
            End If
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
                Engine_MoveUser EAST
                Exit Sub
            End If
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                Engine_MoveUser SOUTH
                Exit Sub
            End If
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                Engine_MoveUser WEST
                Exit Sub
            End If
            If EnterText = False Then
                If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyD) < 0 Then
                    Engine_MoveUser NORTHEAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyW) < 0 And GetKeyState(vbKeyA) < 0 Then
                    Engine_MoveUser NORTHWEST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyD) < 0 Then
                    Engine_MoveUser SOUTHEAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 And GetKeyState(vbKeyA) < 0 Then
                    Engine_MoveUser SOUTHWEST
                    Exit Sub
                End If
                If GetKeyState(vbKeyW) < 0 Then
                    Engine_MoveUser NORTH
                    Exit Sub
                End If
                If GetKeyState(vbKeyD) < 0 Then
                    Engine_MoveUser EAST
                    Exit Sub
                End If
                If GetKeyState(vbKeyS) < 0 Then
                    Engine_MoveUser SOUTH
                    Exit Sub
                End If
                If GetKeyState(vbKeyA) < 0 Then
                    Engine_MoveUser WEST
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub

Sub Input_Mouse_LeftClick()

'******************************************
'Left click mouse
'******************************************
Dim tX As Integer
Dim tY As Integer
Dim i As Long

    'Make sure engine is running
    If Not EngineRun Then Exit Sub

    '***Check for skill list click***
    'Skill lists, because it is not actually a window, must be handled differently
    If QuickBarSetSlot <= 0 Then DrawSkillList = 0
    If DrawSkillList Then
        If SkillListSize Then
            For tX = 1 To SkillListSize
                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, SkillList(tX).X, SkillList(tX).Y, 32, 32) Then
                    QuickBarID(QuickBarSetSlot).ID = SkillList(tX).SkillID
                    QuickBarID(QuickBarSetSlot).Type = QuickBarType_Skill
                    DrawSkillList = 0
                    QuickBarSetSlot = 0
                    Exit Sub
                End If
            Next tX
        End If
    End If

    '***Check for a window click***
    WMSelCon = 0

    'Start with the last clicked window, then move in order of importance
    If LastClickedWindow > 0 Then
        If Input_Mouse_LeftClick_Window(LastClickedWindow) = 1 Then Exit Sub
    End If
    For i = 1 To NumGameWindows
        If LastClickedWindow <> i Then
            If Input_Mouse_LeftClick_Window(i) = 1 Then Exit Sub
        End If
    Next i

    'No windows clicked, so a tile click will take place
    'Get the tile positions
    Engine_ConvertCPtoTP MousePos.X, MousePos.Y, tX, tY

    'Send left click
    sndBuf.Allocate 3
    sndBuf.Put_Byte DataCode.User_LeftClick
    sndBuf.Put_Byte CByte(tX)
    sndBuf.Put_Byte CByte(tY)

    'If there was a click on the game screen and the
    ' skill list is up, but no window clicked, set to 0
    If DrawSkillList Then
        If QuickBarSetSlot Then
            QuickBarID(QuickBarSetSlot).ID = 0
            QuickBarID(QuickBarSetSlot).Type = 0
            DrawSkillList = 0
            QuickBarSetSlot = 0
        End If
    End If
    
    'Last clicked window was nothing, so set to nothing :)
    LastClickedWindow = 0

End Sub

Function Input_Mouse_LeftClick_Window(ByVal WindowIndex As Byte) As Byte

'******************************************
'Left click a game window
'******************************************
Dim i As Byte
Dim j As Byte

    Select Case WindowIndex

        Case TradeWindow
            If ShowGameWindow(TradeWindow) Then
                With GameWindow.Trade
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = TradeWindow
                        'Item window
                        For i = 1 To 9
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Trade1(i).X, .Screen.Y + .Trade1(i).Y, 32, 32) Then
                                sndBuf.Allocate 2
                                sndBuf.Put_Byte DataCode.User_Trade_RemoveItem
                                sndBuf.Put_Byte i
                                Exit Function
                            End If
                        Next i
                        'Accept button
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Accept.X, .Screen.Y + .Accept.Y, .Accept.Width, .Accept.Height) Then
                            sndBuf.Put_Byte DataCode.User_Trade_Accept
                            Exit Function
                        End If
                        'Finish button
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Trade.X, .Screen.Y + .Trade.Y, .Trade.Width, .Trade.Height) Then
                            sndBuf.Put_Byte DataCode.User_Trade_Finish
                            Exit Function
                        End If
                        'Cancel button
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Cancel.X, .Screen.Y + .Cancel.Y, .Cancel.Width, .Cancel.Height) Then
                            sndBuf.Put_Byte DataCode.User_Trade_Cancel
                            Exit Function
                        End If
                        SelGameWindow = TradeWindow
                    End If
                End With
            End If
        
        Case NPCChatWindow
            If ShowGameWindow(NPCChatWindow) Then
                With GameWindow.NPCChat
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = NPCChatWindow
                        For i = 1 To .NumAnswers
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Answer(i).X, .Screen.Y + .Answer(i).Y, .Answer(i).Width, .Answer(i).Height) Then
                                j = NPCChat(ActiveAsk.ChatIndex).Ask.Ask(ActiveAsk.AskIndex).Answer(i).GotoID
                                If j > 0 Then
                                    Engine_ShowNPCChatWindow ActiveAsk.AskName, ActiveAsk.ChatIndex, j
                                Else
                                    ShowGameWindow(NPCChatWindow) = 0
                                    If LastClickedWindow = NPCChatWindow Then LastClickedWindow = 0
                                End If
                                Exit For
                            End If
                        Next i
                        SelGameWindow = NPCChatWindow
                    End If
                End With
            End If
    
        Case MenuWindow
            If ShowGameWindow(MenuWindow) Then
                With GameWindow.Menu
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = MenuWindow
                        'Quit button
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .QuitLbl.X, .Screen.Y + .QuitLbl.Y, .QuitLbl.Width, .QuitLbl.Height) Then
                            IsUnloading = 1
                            Exit Function
                        End If
                        SelGameWindow = MenuWindow
                    End If
                End With
            End If
            
        Case StatWindow
            If ShowGameWindow(StatWindow) Then
                With GameWindow.StatWindow
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = StatWindow
                        'Raise str
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .AddStr.X, .Screen.Y + .AddStr.Y, .AddStr.Width, .AddStr.Height) Then
                            sndBuf.Allocate 2
                            sndBuf.Put_Byte DataCode.User_BaseStat
                            sndBuf.Put_Byte SID.Str
                        End If
                        'Raise agi
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .AddAgi.X, .Screen.Y + .AddAgi.Y, .AddAgi.Width, .AddAgi.Height) Then
                            sndBuf.Allocate 2
                            sndBuf.Put_Byte DataCode.User_BaseStat
                            sndBuf.Put_Byte SID.Agi
                        End If
                        'Raise mag
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .AddMag.X, .Screen.Y + .AddMag.Y, .AddMag.Width, .AddMag.Height) Then
                            sndBuf.Allocate 2
                            sndBuf.Put_Byte DataCode.User_BaseStat
                            sndBuf.Put_Byte SID.Mag
                        End If
                        SelGameWindow = StatWindow
                    End If
                End With
            End If
            
        Case ChatWindow
            If ShowGameWindow(ChatWindow) Then
                With GameWindow.ChatWindow
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Text.X, .Screen.Y + .Text.Y, .Text.Width, .Text.Height) Then
                            EnterText = True
                        End If
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = ChatWindow
                        SelGameWindow = ChatWindow
                        Exit Function
                    End If
                End With
            End If
        
        Case QuickBarWindow
            If ShowGameWindow(QuickBarWindow) Then
                With GameWindow.QuickBar
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = QuickBarWindow
                        'Cancel changes to quick bar items
                        DrawSkillList = 0
                        QuickBarSetSlot = 0
                        'Check if an item was clicked
                        For i = 1 To 12
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If GetAsyncKeyState(vbKeyShift) Then
                                    QuickBarSetSlot = i
                                    DrawSkillList = 1
                                Else
                                    Engine_UseQuickBar i
                                End If
                                Exit Function
                            End If
                        Next i
                        'Item was not clicked
                        SelGameWindow = QuickBarWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case InventoryWindow
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = InventoryWindow
                        'Check if an item was clicked
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If GetAsyncKeyState(vbKeyShift) Then
                                    If Game_ClickItem(i) Then
                                        If UserInventory(i).Amount = 1 Then
                                            'Drop item into mailbox
                                            If ShowGameWindow(WriteMessageWindow) Then
                                                'Check for duplicate entries
                                                For j = 1 To MaxMailObjs
                                                    If WriteMailData.ObjIndex(j) = i Then Exit Function
                                                Next j
                                                'Place item in next free slot (if any)
                                                j = 0
                                                Do
                                                    j = j + 1
                                                    If j > MaxMailObjs Then Exit Function
                                                Loop While WriteMailData.ObjIndex(j) > 0
                                                WriteMailData.ObjIndex(j) = i
                                                WriteMailData.ObjAmount(j) = 1
                                            'Sell item to shopkeeper
                                            ElseIf ShowGameWindow(ShopWindow) Then
                                                sndBuf.Allocate 4
                                                sndBuf.Put_Byte DataCode.User_Trade_SellToNPC
                                                sndBuf.Put_Byte i
                                                sndBuf.Put_Integer 1
                                            'Put item in the bank
                                            ElseIf ShowGameWindow(BankWindow) Then
                                                sndBuf.Allocate 4
                                                sndBuf.Put_Byte DataCode.User_Bank_PutItem
                                                sndBuf.Put_Byte i
                                                sndBuf.Put_Integer 1
                                            'Drop item on ground
                                            Else
                                                sndBuf.Allocate 4
                                                sndBuf.Put_Byte DataCode.User_Drop
                                                sndBuf.Put_Byte i
                                                sndBuf.Put_Integer 1
                                            End If
                                        Else
                                            'Drop item into mailbox
                                            If ShowGameWindow(WriteMessageWindow) Then
                                                'Check for duplicate entries
                                                For j = 1 To MaxMailObjs
                                                    If WriteMailData.ObjIndex(j) = i Then Exit Function
                                                Next j
                                                'Check for free slots
                                                j = 0
                                                Do
                                                    j = j + 1
                                                    If j > MaxMailObjs Then Exit Function
                                                Loop While WriteMailData.ObjIndex(j) > 0
                                                'Open the amount window
                                                ShowGameWindow(AmountWindow) = 1
                                                LastClickedWindow = AmountWindow
                                                AmountWindowValue = vbNullString
                                                AmountWindowItemIndex = i
                                                AmountWindowUsage = AW_InvToMail
                                            'Sell item to shopkeeper
                                            ElseIf ShowGameWindow(ShopWindow) Then
                                                ShowGameWindow(AmountWindow) = 1
                                                LastClickedWindow = AmountWindow
                                                AmountWindowValue = vbNullString
                                                AmountWindowItemIndex = i
                                                AmountWindowUsage = AW_InvToShop
                                            'Put item in the bank
                                            ElseIf ShowGameWindow(BankWindow) Then
                                                ShowGameWindow(AmountWindow) = 1
                                                LastClickedWindow = AmountWindow
                                                AmountWindowValue = vbNullString
                                                AmountWindowItemIndex = i
                                                AmountWindowUsage = AW_InvToBank
                                            'Drop item on ground
                                            Else
                                                ShowGameWindow(AmountWindow) = 1
                                                LastClickedWindow = AmountWindow
                                                AmountWindowValue = vbNullString
                                                AmountWindowItemIndex = i
                                                AmountWindowUsage = AW_Drop
                                            End If
                                        End If
                                    End If
                                Else
                                    If Game_ClickItem(i) Then
                                        sndBuf.Allocate 2
                                        sndBuf.Put_Byte DataCode.User_Use
                                        sndBuf.Put_Byte i
                                    End If
                                End If
                                Exit Function
                            End If
                        Next i
                        'Item was not clicked
                        SelGameWindow = InventoryWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case ShopWindow
            If ShowGameWindow(ShopWindow) Then
                With GameWindow.Shop
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = ShopWindow
                        'Check if an item was clicked
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If Game_ClickItem(i, 2) > 0 Then
                                    sndBuf.Allocate 4
                                    sndBuf.Put_Byte DataCode.User_Trade_BuyFromNPC
                                    sndBuf.Put_Byte i
                                    sndBuf.Put_Integer 1
                                End If
                                Exit Function
                            End If
                        Next i
                        'Item was not clicked
                        SelGameWindow = ShopWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case BankWindow
            If ShowGameWindow(BankWindow) Then
                With GameWindow.Bank
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = BankWindow
                        'Check if an item was clicked
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If Game_ClickItem(i, 3) > 0 Then
                                    sndBuf.Allocate 4
                                    sndBuf.Put_Byte DataCode.User_Bank_TakeItem
                                    sndBuf.Put_Byte i
                                    sndBuf.Put_Integer 1
                                End If
                                Exit Function
                            End If
                        Next i
                        'Item was not clicked
                        SelGameWindow = BankWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case MailboxWindow
            If ShowGameWindow(MailboxWindow) Then
                With GameWindow.Mailbox
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = MailboxWindow
                        'Check if Write was clicked
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .WriteLbl.X, .Screen.Y + .WriteLbl.Y, .WriteLbl.Width, .WriteLbl.Height) Then
                            For i = 1 To MaxMailObjs
                                WriteMailData.ObjIndex(i) = 0
                                WriteMailData.ObjAmount(i) = 0
                            Next i
                            WriteMailData.Message = vbNullString
                            WriteMailData.Subject = vbNullString
                            WriteMailData.RecieverName = vbNullString
                            ShowGameWindow(MailboxWindow) = 0
                            ShowGameWindow(WriteMessageWindow) = 1
                            LastClickedWindow = WriteMessageWindow
                            Exit Function
                        End If
                        If SelMessage > 0 Then
                            'Check if Delete was clicked
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .DeleteLbl.X, .Screen.Y + .DeleteLbl.Y, .DeleteLbl.Width, .DeleteLbl.Height) Then
                                sndBuf.Allocate 2
                                sndBuf.Put_Byte DataCode.Server_MailDelete
                                sndBuf.Put_Byte SelMessage
                                Exit Function
                            End If
                            'Check if Read was clicked
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .ReadLbl.X, .Screen.Y + .ReadLbl.Y, .ReadLbl.Width, .ReadLbl.Height) Then
                                sndBuf.Allocate 2
                                sndBuf.Put_Byte DataCode.Server_MailMessage
                                sndBuf.Put_Byte SelMessage
                                Exit Function
                            End If
                        End If
                        'Check if List was clicked
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .List.X + .List.X, .Screen.Y + .List.Y, .List.Width, .List.Height) Then
                            For i = 1 To (.List.Height \ Font_Default.CharHeight)
                                If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .List.X + .List.X, .Screen.Y + .List.Y + ((i - 1) * Font_Default.CharHeight), .List.Width, Font_Default.CharHeight) Then
                                    If SelMessage = i Then
                                        sndBuf.Allocate 2
                                        sndBuf.Put_Byte DataCode.Server_MailMessage
                                        sndBuf.Put_Byte i
                                    Else
                                        SelMessage = i
                                    End If
                                    Exit Function
                                End If
                            Next i
                            Exit Function
                        End If
                        SelGameWindow = MailboxWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case ViewMessageWindow
            If ShowGameWindow(ViewMessageWindow) Then
                With GameWindow.ViewMessage
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = ViewMessageWindow
                        'Click an item
                        For i = 1 To MaxMailObjs
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                                sndBuf.Allocate 2
                                sndBuf.Put_Byte DataCode.Server_MailItemTake
                                sndBuf.Put_Byte i
                                Exit Function
                            End If
                        Next i
                        SelGameWindow = ViewMessageWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case WriteMessageWindow
            If ShowGameWindow(WriteMessageWindow) Then
                With GameWindow.WriteMessage
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = WriteMessageWindow
                        'Click From
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .From.X + .Screen.X, .From.Y + .Screen.Y, .From.Width, .From.Height) Then
                            WMSelCon = wmFrom
                            Exit Function
                        End If
                        'Click Subject
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Subject.X + .Screen.X, .Subject.Y + .Screen.Y, .Subject.Width, .Subject.Height) Then
                            WMSelCon = wmSubject
                            Exit Function
                        End If
                        'Click Message
                        If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Message.X + .Screen.X, .Message.Y + .Screen.Y, .Message.Width, .Message.Height) Then
                            WMSelCon = wmMessage
                            Exit Function
                        End If
                        'Click an item
                        For i = 1 To MaxMailObjs
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                                WriteMailData.ObjIndex(i) = 0
                                WriteMailData.ObjAmount(i) = 0
                                Exit Function
                            End If
                        Next i
                        SelGameWindow = WriteMessageWindow
                        Exit Function
                    End If
                End With
            End If
            
        Case AmountWindow
            If ShowGameWindow(AmountWindow) Then
                With GameWindow.Amount
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_LeftClick_Window = 1
                        LastClickedWindow = AmountWindow
                    End If
                    SelGameWindow = AmountWindow
                    Exit Function
                End With
            End If
        
    End Select

End Function

Sub Input_Mouse_Move()

'******************************************
'Move mouse
'******************************************

    'Make sure engine is running
    If Not EngineRun Then Exit Sub

    'Clear item info display
    ItemDescLines = 0

    'Check if left mouse is pressed
    If MouseLeftDown Then

        Select Case SelGameWindow

                'Move QuickBar
            Case QuickBarWindow
                With GameWindow.QuickBar.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move ChatWindow
            Case ChatWindow
                With GameWindow.ChatWindow.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                    Engine_UpdateChatArray
                End With
                
                'Move Stat Window
            Case StatWindow
                With GameWindow.StatWindow.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move Inventory
            Case InventoryWindow
                With GameWindow.Inventory.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move Shop
            Case ShopWindow
                With GameWindow.Shop.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move Bank
            Case BankWindow
                With GameWindow.Bank.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move Mailbox
            Case MailboxWindow
                With GameWindow.Mailbox.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move View Message
            Case ViewMessageWindow
                With GameWindow.ViewMessage.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move write message
            Case WriteMessageWindow
                With GameWindow.WriteMessage.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move Amount
            Case AmountWindow
                With GameWindow.Amount.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move Chat window
            Case NPCChatWindow
                With GameWindow.NPCChat.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
                
                'Move the trade window
            Case TradeWindow
                With GameWindow.Trade.Screen
                    .X = .X + MousePosAdd.X
                    .Y = .Y + MousePosAdd.Y
                    If WindowsInScreen Then
                        If .X < 0 Then .X = 0
                        If .Y < 0 Then .Y = 0
                        If .X > ScreenWidth - .Width Then .X = ScreenWidth - .Width
                        If .Y > ScreenHeight - .Height Then .Y = ScreenHeight - .Height
                    End If
                End With
            
        End Select
        
    End If

End Sub

Sub Input_Mouse_RightClick()

'******************************************
'Right click mouse
'******************************************
Dim tX As Integer
Dim tY As Integer
Dim i As Long

    'Make sure engine is running
    If Not EngineRun Then Exit Sub

    '***Check for a window click***
    'Start with the last clicked window, then move in order of importance
    If Input_Mouse_RightClick_Window(LastClickedWindow) = 1 Then Exit Sub
    For i = 1 To NumGameWindows
        If Input_Mouse_RightClick_Window(i) = 1 Then Exit Sub
    Next i
                                                                
    'No windows clicked, so a tile click will take place
    'Get the tile positions
    Engine_ConvertCPtoTP MousePos.X, MousePos.Y, tX, tY
    
    'Check if a NPC was clicked that has ASK responses
    For i = 1 To LastChar
        If CharList(i).Pos.X = tX Then
            If CharList(i).Pos.Y = tY Then
                If CharList(i).NPCChatIndex > 0 Then
                    If NPCChat(CharList(i).NPCChatIndex).Ask.StartAsk > 0 Then
                        Engine_ShowNPCChatWindow CharList(i).Name, CharList(i).NPCChatIndex, NPCChat(CharList(i).NPCChatIndex).Ask.StartAsk
                    End If
                End If
                Exit For
            End If
        End If
    Next i

    'Normal click
    If UseClickWarp = 0 Then
        
        'Check if a sign was clicked
        If MapData(tX, tY).Sign Then Engine_AddToChatTextBuffer Replace$(Message(126), "<text>", Signs(MapData(tX, tY).Sign)), FontColor_Info
        
        'Send left click
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.User_RightClick
        sndBuf.Put_Byte CByte(tX)
        sndBuf.Put_Byte CByte(tY)
        
    'Warp click
    Else
    
        sndBuf.Allocate 3
        sndBuf.Put_Byte DataCode.GM_Warp
        sndBuf.Put_Byte CByte(tX)
        sndBuf.Put_Byte CByte(tY)
        
    End If

End Sub

Function Input_Mouse_RightClick_Window(ByVal WindowIndex As Byte) As Byte

'******************************************
'Left click a game window
'******************************************
Dim i As Integer

    Select Case WindowIndex
    
        Case TradeWindow
            If ShowGameWindow(TradeWindow) Then
                With GameWindow.Trade
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = TradeWindow
                        'Item window for user 1
                        For i = 1 To 9
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Trade1(i).X, .Screen.Y + .Trade1(i).Y, 32, 32) Then
                                If TradeTable.Trade1(i).Grh > 0 Then
                                    Engine_SetItemDesc TradeTable.Trade1(i).Name, TradeTable.Trade1(i).Amount, TradeTable.Trade1(i).Value
                                    Exit Function
                                End If
                            End If
                        Next i
                        'Item window for user 2
                        For i = 1 To 9
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Trade2(i).X, .Screen.Y + .Trade2(i).Y, 32, 32) Then
                                If TradeTable.Trade2(i).Grh > 0 Then
                                    Engine_SetItemDesc TradeTable.Trade2(i).Name, TradeTable.Trade2(i).Amount, TradeTable.Trade2(i).Value
                                    Exit Function
                                End If
                            End If
                        Next i
                    End If
                End With
            End If
        
        Case QuickBarWindow
            If ShowGameWindow(QuickBarWindow) Then
                With GameWindow.QuickBar
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = QuickBarWindow
                        'Check if an item was clicked
                        For i = 1 To 12
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                'An item in the quickbar was clicked - get description
                                If QuickBarID(i).Type = QuickBarType_Item Then
                                    Engine_SetItemDesc UserInventory(QuickBarID(i).ID).Name, UserInventory(QuickBarID(i).ID).Amount
                                    'A skill in the quickbar was clicked - get the name
                                ElseIf QuickBarID(i).Type = QuickBarType_Skill Then
                                    Engine_SetItemDesc Engine_SkillIDtoSkillName(QuickBarID(i).ID)
                                End If
                                Exit Function
                            End If
                        Next i
                    End If
                End With
            End If
            
        Case InventoryWindow
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = InventoryWindow
                        'Check if an item was clicked
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If UserInventory(i).GrhIndex > 0 Then
                                    Engine_SetItemDesc UserInventory(i).Name, UserInventory(i).Amount, UserInventory(i).Value
                                    DragSourceWindow = InventoryWindow
                                    DragItemSlot = i
                                End If
                                Exit Function
                            End If
                        Next i
                    End If
                End With
            End If
            
        Case ShopWindow
            If ShowGameWindow(ShopWindow) Then
                With GameWindow.Shop
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = ShopWindow
                        'Check if an item was clicked
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If i <= NPCTradeItemArraySize Then
                                    If NPCTradeItems(i).GrhIndex > 0 Then
                                        Engine_SetItemDesc NPCTradeItems(i).Name, 0, NPCTradeItems(i).Value
                                        DragSourceWindow = ShopWindow
                                        DragItemSlot = i
                                    End If
                                End If
                                Exit Function
                            End If
                        Next i
                    End If
                End With
            End If
            
        Case BankWindow
            If ShowGameWindow(BankWindow) Then
                With GameWindow.Bank
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = BankWindow
                        'Check if an item was clicked
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If UserBank(i).GrhIndex > 0 Then Engine_SetItemDesc UserBank(i).Name, UserBank(i).Amount
                                DragSourceWindow = BankWindow
                                DragItemSlot = i
                                Exit Function
                            End If
                        Next i
                    End If
                End With
            End If
            
        Case ViewMessageWindow
            If ShowGameWindow(ViewMessageWindow) Then
                With GameWindow.ViewMessage
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = ViewMessageWindow
                        'Click an item
                        For i = 1 To MaxMailObjs
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                                Engine_SetItemDesc ReadMailData.ObjName(i), ReadMailData.ObjAmount(i)
                                Exit Function
                            End If
                        Next i
                    End If
                End With
            End If
            
        Case WriteMessageWindow
            If ShowGameWindow(WriteMessageWindow) Then
                With GameWindow.WriteMessage
                    'Check if the screen was clicked
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = WriteMessageWindow
                        'Click an item
                        For i = 1 To MaxMailObjs
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X + .Image(i).X, .Screen.Y + .Image(i).Y, .Image(i).Width, .Image(i).Height) Then
                                Engine_SetItemDesc UserInventory(WriteMailData.ObjIndex(i)).Name, WriteMailData.ObjAmount(i)
                                Exit Function
                            End If
                        Next i
                    End If
                End With
            End If
            
            
        Case ChatWindow
            If ShowGameWindow(ChatWindow) Then
                With GameWindow.ChatWindow
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = ChatWindow
                    End If
                End With
            End If
        
        Case MenuWindow
            If ShowGameWindow(MenuWindow) Then
                With GameWindow.Menu
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = MenuWindow
                    End If
                End With
            End If
            
        Case StatWindow
            If ShowGameWindow(StatWindow) Then
                With GameWindow.StatWindow
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = StatWindow
                    End If
                End With
            End If
            
        Case ViewMessageWindow
            If ShowGameWindow(ViewMessageWindow) Then
                With GameWindow.ViewMessage
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = ViewMessageWindow
                    End If
                End With
            End If
            
        Case AmountWindow
            If ShowGameWindow(AmountWindow) Then
                With GameWindow.Amount
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = AmountWindow
                    End If
                End With
            End If
            
        Case NPCChatWindow
            If ShowGameWindow(NPCChatWindow) Then
                With GameWindow.NPCChat
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        Input_Mouse_RightClick_Window = 1
                        LastClickedWindow = NPCChatWindow
                    End If
                End With
            End If
    
    End Select

End Function

Sub Input_Mouse_RightRelease()

'******************************************
'Right mouse button released
'******************************************
Dim i As Byte

    'Check if we released mouse and have an item in being dragged
    If DragItemSlot Then
    
        'Inventory -> Trade Window
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(TradeWindow) Then
                With GameWindow.Trade
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        For i = 1 To 9
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Trade1(i).X + .Screen.X, .Trade1(i).Y + .Screen.Y, 32, 32) Then

                                If UserInventory(DragItemSlot).Amount = 1 Then
                                    sndBuf.Put_Byte DataCode.User_Trade_UpdateTrade
                                    sndBuf.Put_Byte DragItemSlot
                                    sndBuf.Put_Long 1
                                Else
                                    ShowGameWindow(AmountWindow) = 1
                                    LastClickedWindow = AmountWindow
                                    AmountWindowItemIndex = DragItemSlot
                                    AmountWindowValue = vbNullString
                                    AmountWindowUsage = AW_InvToTrade
                                End If
                                
                                'Clear and leave
                                DragSourceWindow = 0
                                DragItemSlot = 0
                                
                                Exit Sub
                                
                            End If
                            
                        Next i
                    End If
                End With
            End If
        End If
        
        'Inventory -> Inventory (change slot)
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        For i = 1 To 49
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                If DragItemSlot <> i Then
                                    'Switch slots
                                    sndBuf.Allocate 3
                                    sndBuf.Put_Byte DataCode.User_ChangeInvSlot
                                    sndBuf.Put_Byte DragItemSlot
                                    sndBuf.Put_Byte i
                                    'Clear and leave
                                    DragSourceWindow = 0
                                    DragItemSlot = 0
                                    Exit Sub
                                End If
                            End If
                        Next i
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If

        'Inventory -> Quick Bar
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(QuickBarWindow) Then
                With GameWindow.QuickBar
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        For i = 1 To 12
                            If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Image(i).X + .Screen.X, .Image(i).Y + .Screen.Y, .Image(i).Width, .Image(i).Height) Then
                                'Drop into quick use slot
                                QuickBarID(i).Type = QuickBarType_Item
                                QuickBarID(i).ID = DragItemSlot
                                'Clear and leave
                                DragSourceWindow = 0
                                DragItemSlot = 0
                                Exit Sub
                            End If
                        Next i
                    End If
                End With
            End If
        End If
        
        'Inventory -> Depot
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(BankWindow) Then
                With GameWindow.Bank
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Single item
                        If UserInventory(DragItemSlot).Amount = 1 Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Bank_PutItem
                            sndBuf.Put_Byte DragItemSlot
                            sndBuf.Put_Integer 1
                        'Multiple items
                        Else
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_InvToBank
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                    End If
                End With
            End If
        End If
        
        'Inventory -> Shop
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(ShopWindow) Then
                With GameWindow.Shop
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Single item
                        If UserInventory(DragItemSlot).Amount = 1 Then
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Trade_SellToNPC
                            sndBuf.Put_Byte DragItemSlot
                            sndBuf.Put_Integer 1
                        'Multiple items
                        Else
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_InvToShop
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
        
        'Shop -> Inventory
        If DragSourceWindow = ShopWindow Then
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Bring up amount window for bulk buying
                        ShowGameWindow(AmountWindow) = 1
                        LastClickedWindow = AmountWindow
                        AmountWindowValue = vbNullString
                        AmountWindowItemIndex = DragItemSlot
                        AmountWindowUsage = AW_ShopToInv
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
        
        'Bank -> Inventory
        If DragSourceWindow = BankWindow Then
            If ShowGameWindow(InventoryWindow) Then
                With GameWindow.Inventory
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        If UserBank(DragItemSlot).Amount > 1 Then
                            'Bring up amount window for bulk withdrawing
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_BankToInv
                        Else
                            sndBuf.Allocate 4
                            sndBuf.Put_Byte DataCode.User_Bank_TakeItem
                            sndBuf.Put_Byte DragItemSlot
                            sndBuf.Put_Integer 1
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
                                
        'Inventory -> Mail
        If DragSourceWindow = InventoryWindow Then
            If ShowGameWindow(WriteMessageWindow) Then
                With GameWindow.WriteMessage
                    If Engine_Collision_Rect(MousePos.X, MousePos.Y, 1, 1, .Screen.X, .Screen.Y, .Screen.Width, .Screen.Height) Then
                        'Single item
                        If UserInventory(DragItemSlot).Amount = 1 Then
                            'Check for duplicate entries
                            For i = 1 To MaxMailObjs
                                If WriteMailData.ObjIndex(i) = DragItemSlot Then
                                    DragSourceWindow = 0
                                    DragItemSlot = 0
                                    Exit Sub
                                End If
                            Next i
                            'Place item in next free slot (if any)
                            i = 0
                            Do
                                i = i + 1
                                If i > MaxMailObjs Then
                                    DragSourceWindow = 0
                                    DragItemSlot = 0
                                    Exit Sub
                                End If
                            Loop While WriteMailData.ObjIndex(i) > 0
                            WriteMailData.ObjIndex(i) = DragItemSlot
                            WriteMailData.ObjAmount(i) = 1
                        'Multiple items
                        Else
                            ShowGameWindow(AmountWindow) = 1
                            LastClickedWindow = AmountWindow
                            AmountWindowValue = vbNullString
                            AmountWindowItemIndex = DragItemSlot
                            AmountWindowUsage = AW_InvToMail
                        End If
                        'Clear and leave
                        DragSourceWindow = 0
                        DragItemSlot = 0
                        Exit Sub
                    End If
                End With
            End If
        End If
        
        'Inventory -> Ground
        If DragSourceWindow = InventoryWindow Then
            'Single item
            If UserInventory(DragItemSlot).Amount = 1 Then
                sndBuf.Allocate 4
                sndBuf.Put_Byte DataCode.User_Drop
                sndBuf.Put_Byte DragItemSlot
                sndBuf.Put_Integer 1
            'Multiple items
            Else
                ShowGameWindow(AmountWindow) = 1
                LastClickedWindow = AmountWindow
                AmountWindowValue = vbNullString
                AmountWindowItemIndex = DragItemSlot
                AmountWindowUsage = AW_Drop
            End If
            'Clear and leave
            DragSourceWindow = 0
            DragItemSlot = 0
            Exit Sub
        End If

        'Didn't release over a valid area
        DragSourceWindow = 0
        DragItemSlot = 0

    End If

End Sub
