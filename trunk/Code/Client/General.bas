Attribute VB_Name = "General"
Option Explicit

Public Enum LogType
    General = 0
    CodeTracker = 1
    PacketIn = 2
    PacketOut = 3
    CriticalError = 4
    InvalidPacketData = 5
End Enum

Public Type NPCTradeItems
    Name As String
    Value As Long
    GrhIndex As Long
End Type

Public NumBytesForSkills As Long

Public NPCTradeItems() As NPCTradeItems
Public NPCTradeItemArraySize As Byte

Public FPSCap As Long   'The FPS cap the user defined to use (in milliseconds, not FPS)

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)

'Sleep API - used to put a process into "idle" for X milliseconds
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Like the Shell function, but more powerful - used to call another application to load it
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Log(ByVal DummyT As String, ByVal DummyB As LogType)

'***************************************************
'Dummy routine for logs from the server since some files are shared between multiple projects
'***************************************************

End Sub

Public Function Engine_ValidChar(ByVal CharIndex As Integer) As Boolean

'***************************************************
'Checks for a valid char index
'***************************************************

    If CharIndex <= 0 Then GoTo InvalidChar
    If CharIndex > LastChar Then GoTo InvalidChar
    If CharList(CharIndex).Active = 0 Then GoTo InvalidChar
    
    Engine_ValidChar = True
    Exit Function
    
InvalidChar:

    sndBuf.Allocate 3
    sndBuf.Put_Byte DataCode.User_RequestMakeChar
    sndBuf.Put_Integer CharIndex
    Engine_ValidChar = False
    
End Function

Public Function Engine_BuildSkinsList() As String

'***************************************************
'Returns the list of all the skins
'***************************************************
Dim TempSplit() As String
Dim Files() As String
Dim i As Long

    'Get the list of files
    Files() = AllFilesInFolders(DataPath & "Skins\", False)
    
    'Show the header message
    Engine_AddToChatTextBuffer "The following skins are available:", FontColor_Info
    
    'Look for files ending with ".ini" only
    For i = LBound(Files) To UBound(Files)
        If Right$(Files(i), 4) = ".ini" Then
            
            'Crop out the skin name and add it to the function
            TempSplit() = Split(Files(i), "\")
            If LenB(Engine_BuildSkinsList) <> 0 Then Engine_BuildSkinsList = Engine_BuildSkinsList & vbCrLf
            Engine_BuildSkinsList = Engine_BuildSkinsList & " * |" & Left$(TempSplit(UBound(TempSplit)), Len(TempSplit(UBound(TempSplit))) - 4) & "|"

        End If
    Next i
    
End Function

Sub Game_BuildFilter()

'*****************************************************************
'Creates the filtering strings
'*****************************************************************
Dim sGroup() As String
Dim sSplit() As String
Dim i As Long

    'Check if we even have filtered words
    If LenB(FilterString) = 0 Then Exit Sub

    'Split up the word groups
    sGroup() = Split(FilterString, ",")
    ReDim FilterFind(0 To UBound(sGroup()))
    ReDim FilterReplace(0 To UBound(sGroup()))
    For i = 0 To UBound(sGroup())
        
        'Split up the group to get the word to search for, and the word to replace it with
        sSplit() = Split(sGroup(i), "-")
        
        'Store the values
        FilterFind(i) = Trim$(sSplit(0))
        FilterReplace(i) = Trim$(sSplit(1))
        
    Next i
    
End Sub

Function Game_FilterString(ByVal s As String) As String

'*****************************************************************
'Filters a string from all illegal characters and swear words
'*****************************************************************
Dim i As Long
Dim a As Integer
Dim t As String

    'Check for a legal string
    If LenB(s) = 0 Then
        Game_FilterString = s
        Exit Function
    End If

    'Filter illegal character
    For i = 1 To Len(s) - 1
        a = Asc(Mid$(s, i, 1))
        If Not Game_ValidCharacter(a) Then
            t = vbNullString
            If i > 1 Then t = t & Left$(s, i - 1)
            t = t & "X"
            If i < Len(s) - 1 Then t = t & Right$(s, Len(s) - i)
            s = t
        End If
    Next i

    'Call the swear filter
    s = Game_SwearFilterString(s)
    
    'Return the string
    Game_FilterString = s

End Function

Function Game_ClosestTargetNPC() As Integer

'*****************************************************************
'Find the closest NPC to target
'*****************************************************************
Dim LowestValue As Long
Dim LowestValueChar As Long
Dim UserAngleMod As Long
Dim TempAngle As Long
Dim TempValue As Long
Dim j As Long

    'Check for characters
    If LastChar = 1 Then Exit Function  'If theres only one character, its probably the user
    
    'Get the initial size of the chars array
    ReDim CharValue(1 To LastChar)
    
    'Calculate the modifier of the user's heading
    Select Case CharList(UserCharIndex).Heading
        Case NORTH: UserAngleMod = 0 * 45
        Case NORTHEAST: UserAngleMod = 1 * 45
        Case EAST: UserAngleMod = 2 * 45
        Case SOUTHEAST: UserAngleMod = 3 * 45
        Case SOUTH: UserAngleMod = 4 * 45
        Case SOUTHWEST: UserAngleMod = 5 * 45
        Case WEST: UserAngleMod = 6 * 45
        Case NORTHWEST: UserAngleMod = 7 * 45
    End Select
    
    'Loop through all the characters
    For j = 1 To LastChar
    
        'Make sure the character is used
        If CharList(j).Active Then
            If j <> UserCharIndex Then
                If j <> TargetCharIndex Then
                    
                    'Check that the character is in the screen
                    If CharList(j).Pos.X > ScreenMinX Then
                        If CharList(j).Pos.X < ScreenMaxX Then
                            If CharList(j).Pos.Y > ScreenMinY Then
                                If CharList(j).Pos.Y < ScreenMaxY Then
                                    
                                    'Get the angle between the user and the NPC
                                    TempAngle = -UserAngleMod + Engine_GetAngle(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(j).Pos.X, CharList(j).Pos.Y)
                                    
                                    'Make sure the angle is between 0 and 360
                                    Do While TempAngle >= 360
                                        TempAngle = TempAngle - 360
                                    Loop
                                    Do While TempAngle < 0
                                        TempAngle = TempAngle + 360
                                    Loop
                                    
                                    'Check that the angle is less between -95 and 95 (not behind them)
                                    If TempAngle < 95 Or TempAngle > 265 Then
                                        
                                        'Convert the angle to the distance from 0 degrees
                                        If TempAngle > 180 Then TempAngle = Abs(360 - TempAngle)
                                        If TempAngle = 360 Then TempAngle = 0

                                        'Calculate the value of the character
                                        'Value = Angle * 2 + Distance
                                        TempValue = (TempAngle * 0.5) + Engine_Distance(CharList(UserCharIndex).Pos.X, CharList(UserCharIndex).Pos.Y, CharList(j).Pos.X, CharList(j).Pos.Y)
                                        
                                        'Check if this value is lower then the first value
                                        If LowestValue = 0 Then
                                            LowestValue = TempValue
                                            LowestValueChar = j
                                        Else
                                            If LowestValue > TempValue Then
                                                LowestValue = TempValue
                                                LowestValueChar = j
                                            End If
                                        End If
                                        
                                    End If
                                
                                End If
                            End If
                        End If
                    End If
                
                End If
            End If
        End If
    
    Next j
    
    'Return the index of the character with the lowest value (best target)
    Game_ClosestTargetNPC = LowestValueChar

End Function

Function Game_SwearFilterString(ByVal s As String) As String

'*****************************************************************
'Checks the passed string for any swear words to filter out
'*****************************************************************
Dim i As Long

    'Check if we even have filtered words
    If LenB(FilterString) = 0 Then
        Game_SwearFilterString = s
        Exit Function
    End If

    'Loop through all the filters
    For i = 0 To UBound(FilterFind())
        s = Replace$(s, FilterFind(i), FilterReplace(i))
    Next i
    
    'Return the string
    Game_SwearFilterString = s

End Function

Function Game_CheckUserData() As Boolean

'*****************************************************************
'Checks all user data for mistakes and reports them.
'*****************************************************************

    'Password
    If Len(UserPassword) < 3 Then
        MsgBox ("Password box is empty.")
        Exit Function
    End If
    If Len(UserPassword) > 10 Then
        MsgBox ("Password must be 10 characters or less.")
        Exit Function
    End If
    If Game_LegalString(UserPassword) = False Then
        MsgBox ("Invalid Password.")
        Exit Function
    End If
    
    'Name
    If Len(UserName) < 3 Then
        MsgBox ("Name box is empty.")
        Exit Function
    End If
    If Len(UserName) > 10 Then
        MsgBox ("Name must be 10 characters or less.")
        Exit Function
    End If
    If Game_LegalString(UserName) = False Then
        MsgBox ("Invalid Name.")
        Exit Function
    End If
    
    'If all good send true
    Game_CheckUserData = True

End Function

Function Game_ClickItem(ByVal ItemIndex As Byte, Optional ByVal InventoryType As Long = 1) As Long

'***************************************************
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    
    'Make sure item index is within limits
    If ItemIndex <= 0 Then Exit Function
    If ItemIndex > MAX_INVENTORY_SLOTS Then Exit Function
    
    'Check by the appropriate window
    Select Case InventoryType
        
        'User inventory
        Case 1
            If UserInventory(ItemIndex).GrhIndex > 0 Then Game_ClickItem = 1
            
        'Shop inventory
        Case 2
            If NPCTradeItems(ItemIndex).GrhIndex > 0 Then Game_ClickItem = 1
        
        'Bank depot
        Case 3
            If UserBank(ItemIndex).GrhIndex > 0 Then Game_ClickItem = 1
            
    End Select

End Function

Function Game_ValidCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is used for chat/etc)
'Make sure you update the server's Server_ValidCharacter, too!
'*****************************************************************

    Log "Call Game_ValidCharacter(" & KeyAscii & ")", CodeTracker '//\\LOGLINE//\\

    If KeyAscii >= 32 Then Game_ValidCharacter = True

End Function

Function Game_LegalCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is for username/pass)
'Make sure you update the server's Server_LegalCharacter, too!
'*****************************************************************

    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 Then
        If KeyAscii <= 57 Then
            Game_LegalCharacter = True
            Exit Function
        End If
    End If
    
    'Allow characters A to Z
    If KeyAscii >= 65 Then
        If KeyAscii <= 90 Then
            Game_LegalCharacter = True
            Exit Function
        End If
    End If
    
    'Allow characters a to z
    If KeyAscii >= 97 Then
        If KeyAscii <= 122 Then
            Game_LegalCharacter = True
            Exit Function
        End If
    End If
    
    'Allow foreign characters
    If KeyAscii >= 128 Then
        If KeyAscii <= 168 Then
            Game_LegalCharacter = True
            Exit Function
        End If
    End If
    
Exit Function

ErrOut:

    'Something bad happened, so the character must be invalid
    Game_LegalCharacter = False
    
End Function

Function Game_ValidString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for illegal characters in the string (wrapper for Game_ValidCharacter)
'*****************************************************************
Dim i As Long

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then Exit Function
    If LenB(CheckString) < 1 Then Exit Function

    'Loop through the string
    For i = 1 To Len(CheckString)
        
        'Check the values
        If Game_ValidCharacter(AscB(Mid$(CheckString, i, 1))) = False Then Exit Function
        
    Next i
    
    'If we have made it this far, then all is good
    Game_ValidString = True

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Game_ValidString = False

End Function

Function Game_LegalString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for illegal characters in the string (wrapper for Server_LegalCharacter)
'*****************************************************************
Dim i As Long

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then Exit Function
    If LenB(CheckString) < 1 Then Exit Function

    'Loop through the string
    For i = 1 To Len(CheckString)
        
        'Check the values
        If Game_LegalCharacter(AscB(Mid$(CheckString, i, 1))) = False Then Exit Function
        
    Next i
    
    'If we have made it this far, then all is good
    Game_LegalString = True

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Game_LegalString = False

End Function

Public Sub Game_Config_Load()

'***************************************************
'Load the user configuration
'***************************************************

Dim i As Byte

    'Quickbar
    For i = 1 To 12
        QuickBarID(i).ID = Val(Var_Get(DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "ID"))
        QuickBarID(i).Type = Val(Var_Get(DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "Type"))
    Next i
    
    'Skin
    CurrentSkin = Var_Get(DataPath & "Game.ini", "INIT", "CurrentSkin")

End Sub

Sub Game_Map_Switch(Map As Integer)

'*****************************************************************
'Loads and switches to a new map
'*****************************************************************
Dim LargestTileSize As Long
Dim MapBuf As DataBuffer
Dim GetParticleCount As Integer
Dim GetEffectNum As Byte
Dim GetDirection As Integer
Dim GetGfx As Byte
Dim GetX As Integer
Dim GetY As Integer
Dim ByFlags As Long
Dim MapNum As Byte
Dim i As Integer
Dim Y As Byte
Dim X As Byte
Dim b() As Byte
Dim TempInt As Integer

    'Check if there was a map before this one - if so, clear it up
    If MapInfo.Width > 0 Then

        'Clear the offset values for the particle engine
        ParticleOffsetX = 0
        ParticleOffsetY = 0
        LastOffsetX = 0
        LastOffsetY = 0
    
        'Reset the user's position (it won't be drawn at 0,0 since it is an invalid position anyways)
        UserPos.X = 0
        UserPos.Y = 0
    
        'Erase characters
        LastChar = 0
        Erase CharList
    
        'Erase objects
        LastObj = 0
        Erase OBJList
        
        'Erase particle effects
        LastEffect = 0
        ReDim Effect(1 To NumEffects)

    End If

    'Open map file
    MapNum = FreeFile
    Open MapPath & Map & ".map" For Binary As #MapNum
        Seek #MapNum, 1
        
        'Store the data in the buffer
        ReDim b(0 To LOF(MapNum) - 1)
        Get #MapNum, , b()
        
    'Close the map file
    Close #MapNum
    
    'Assign the buffer data
    Set MapBuf = New DataBuffer
    MapBuf.Set_Buffer b()
    
    'Clear the data array (since its now in the buffer)
    Erase b()

    'Map Header
    TempInt = MapBuf.Get_Integer    'Not stored in memory
    MapInfo.Width = MapBuf.Get_Byte
    MapInfo.Height = MapBuf.Get_Byte
    
    'Resize mapdata array
    ReDim MapData(1 To MapInfo.Width, 1 To MapInfo.Height) As MapBlock

    'Resize the save light buffer
    ReDim SaveLightBuffer(1 To MapInfo.Width, 1 To MapInfo.Height)
    
    'Load arrays
    For Y = 1 To MapInfo.Height
        For X = 1 To MapInfo.Width
        
            'Clear the graphic layers
            For i = 1 To 6
                MapData(X, Y).Graphic(i).GrhIndex = 0
            Next i

            'Get flag's byte
            ByFlags = MapBuf.Get_Long

            'Blocked
            If ByFlags And 1 Then MapData(X, Y).Blocked = MapBuf.Get_Byte Else MapData(X, Y).Blocked = 0

            'Graphic layers
            If ByFlags And 2 Then
                MapData(X, Y).Graphic(1).GrhIndex = MapBuf.Get_Long
                Engine_Init_Grh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
                
                'Find the size of the largest tile used
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(1).GrhIndex).pixelHeight
                End If
                
            End If
            If ByFlags And 4 Then
                MapData(X, Y).Graphic(2).GrhIndex = MapBuf.Get_Long
                Engine_Init_Grh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(2).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 8 Then
                MapData(X, Y).Graphic(3).GrhIndex = MapBuf.Get_Long
                Engine_Init_Grh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(3).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 16 Then
                MapData(X, Y).Graphic(4).GrhIndex = MapBuf.Get_Long
                Engine_Init_Grh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(4).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 32 Then
                MapData(X, Y).Graphic(5).GrhIndex = MapBuf.Get_Long
                Engine_Init_Grh MapData(X, Y).Graphic(5), MapData(X, Y).Graphic(5).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(5).GrhIndex).pixelHeight
                End If
            End If
            If ByFlags And 64 Then
                MapData(X, Y).Graphic(6).GrhIndex = MapBuf.Get_Long
                Engine_Init_Grh MapData(X, Y).Graphic(6), MapData(X, Y).Graphic(6).GrhIndex
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelWidth Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelWidth
                End If
                If LargestTileSize < GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelHeight Then
                    LargestTileSize = GrhData(MapData(X, Y).Graphic(6).GrhIndex).pixelHeight
                End If
            End If
            
            'Set light to default (-1) - it will be set again if it is not -1 from the code below
            For i = 1 To 24
                MapData(X, Y).Light(i) = -1
            Next i
            
            'Get lighting values
            If ByFlags And 128 Then
                For i = 1 To 4
                    MapData(X, Y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 256 Then
                For i = 5 To 8
                    MapData(X, Y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 512 Then
                For i = 9 To 12
                    MapData(X, Y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 1024 Then
                For i = 13 To 16
                    MapData(X, Y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 2048 Then
                For i = 17 To 20
                    MapData(X, Y).Light(i) = MapBuf.Get_Long
                Next i
            End If
            If ByFlags And 4096 Then
                For i = 21 To 24
                    MapData(X, Y).Light(i) = MapBuf.Get_Long
                Next i
            End If

            'Store the lighting in the SaveLightBuffer
            For i = 1 To 24
                SaveLightBuffer(X, Y).Light(i) = MapData(X, Y).Light(i)
            Next i

            'Mailbox - Not used by the client
            'If ByFlags And 8192 Then MapData(X, Y).Mailbox = 1 Else MapData(X, Y).Mailbox = 0

            'Shadows
            If ByFlags And 16384 Then MapData(X, Y).Shadow(1) = 1 Else MapData(X, Y).Shadow(1) = 0
            If ByFlags And 32768 Then MapData(X, Y).Shadow(2) = 1 Else MapData(X, Y).Shadow(2) = 0
            If ByFlags And 65536 Then MapData(X, Y).Shadow(3) = 1 Else MapData(X, Y).Shadow(3) = 0
            If ByFlags And 131072 Then MapData(X, Y).Shadow(4) = 1 Else MapData(X, Y).Shadow(4) = 0
            If ByFlags And 262144 Then MapData(X, Y).Shadow(5) = 1 Else MapData(X, Y).Shadow(5) = 0
            If ByFlags And 524288 Then MapData(X, Y).Shadow(6) = 1 Else MapData(X, Y).Shadow(6) = 0
            
            'Clear any old sfx
            If Not MapData(X, Y).Sfx Is Nothing Then
                MapData(X, Y).Sfx.Stop
                Set MapData(X, Y).Sfx = Nothing
            End If
            
            'Set the sfx
            If ByFlags And 1048576 Then
                i = MapBuf.Get_Integer
                Sound_SetToMap i, X, Y
            End If
            
            'Blocked attack
            If ByFlags And 2097152 Then MapData(X, Y).BlockedAttack = 1 Else MapData(X, Y).BlockedAttack = 0
            
            'Sign
            If ByFlags And 4194304 Then MapData(X, Y).Sign = MapBuf.Get_Integer Else MapData(X, Y).Sign = 0
            
            'If there is a warp
            If ByFlags And 8388608 Then MapData(X, Y).Warp = 1 Else MapData(X, Y).Warp = 0

        Next X
    Next Y
    
    'Get the number of effects
    Y = MapBuf.Get_Byte

    'Store the individual particle effect types
    If Y > 0 Then
        For X = 1 To Y
            GetEffectNum = MapBuf.Get_Byte
            GetX = MapBuf.Get_Integer
            GetY = MapBuf.Get_Integer
            GetParticleCount = MapBuf.Get_Integer
            GetGfx = MapBuf.Get_Byte
            GetDirection = MapBuf.Get_Integer
            Effect_Begin GetEffectNum, GetX, GetY, GetGfx, GetParticleCount, GetDirection
        Next X
    End If
    
    'Clear the map data
    Set MapBuf = Nothing
    
    'Create the minimap
    Engine_BuildMiniMap

    'Clear out old mapinfo variables
    MapInfo.Name = vbNullString

    'Set current map
    CurMap = Map
    
    'Auto-calculate the maximum size to set the tile buffer
    LargestTileSize = LargestTileSize + (32 - (LargestTileSize Mod 32)) 'Round to the next highest factor of 32
    TileBufferSize = (LargestTileSize \ 32) 'Divide into tiles
    
    'Force to 2 to draw characters since they are 2 tiles tall
    'If you have characters or paperdoll parts > 64 pixels in width or high, you need to increase this
    If TileBufferSize < 2 Then TileBufferSize = 2

End Sub

Public Sub Game_Config_Save()

'***************************************************
'Load the user configuration
'***************************************************
Dim t As String
Dim i As Byte

    'Quickbar
    For i = 1 To 12
        Var_Write DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "ID", Str$(QuickBarID(i).ID)
        Var_Write DataPath & "Game.ini", "QUICKBARVALUES", "Slot" & i & "Type", Str$(QuickBarID(i).Type)
    Next i
    
    'Skin
    Var_Write DataPath & "Game.ini", "INIT", "CurrentSkin", CurrentSkin
    
    'Skin positions
    t = DataPath & "Skins\" & CurrentSkin & ".dat"   'Set the custom positions file for the skin
    With GameWindow
        Var_Write t, "QUICKBAR", "ScreenX", Str$(.QuickBar.Screen.X)
        Var_Write t, "QUICKBAR", "ScreenY", Str$(.QuickBar.Screen.Y)
        Var_Write t, "CHATWINDOW", "ScreenX", Str$(.ChatWindow.Screen.X)
        Var_Write t, "CHATWINDOW", "ScreenY", Str$(.ChatWindow.Screen.Y)
        Var_Write t, "INVENTORY", "ScreenX", Str$(.Inventory.Screen.X)
        Var_Write t, "INVENTORY", "ScreenY", Str$(.Inventory.Screen.Y)
        Var_Write t, "SHOP", "ScreenX", Str$(.Shop.Screen.X)
        Var_Write t, "SHOP", "ScreenY", Str$(.Shop.Screen.Y)
        Var_Write t, "MAILBOX", "ScreenX", Str$(.Mailbox.Screen.X)
        Var_Write t, "MAILBOX", "ScreenY", Str$(.Mailbox.Screen.Y)
        Var_Write t, "VIEWMESSAGE", "ScreenX", Str$(.ViewMessage.Screen.X)
        Var_Write t, "VIEWMESSAGE", "ScreenY", Str$(.ViewMessage.Screen.Y)
        Var_Write t, "WRITEMESSAGE", "ScreenX", Str$(.WriteMessage.Screen.X)
        Var_Write t, "WRITEMESSAGE", "ScreenY", Str$(.WriteMessage.Screen.Y)
        Var_Write t, "AMOUNT", "ScreenX", Str$(.Amount.Screen.X)
        Var_Write t, "AMOUNT", "ScreenY", Str$(.Amount.Screen.Y)
        Var_Write t, "MENU", "ScreenX", Str$(.Menu.Screen.X)
        Var_Write t, "MENU", "ScreenY", Str$(.Menu.Screen.Y)
        Var_Write t, "BANK", "ScreenX", Str$(.Bank.Screen.X)
        Var_Write t, "BANK", "ScreenY", Str$(.Bank.Screen.Y)
        Var_Write t, "NPCCHAT", "ScreenX", Str$(.NPCChat.Screen.X)
        Var_Write t, "NPCCHAT", "ScreenY", Str$(.NPCChat.Screen.Y)
    End With

End Sub

Sub UpdateShownTextBuffer()

'*****************************************************************
'Updates the ShownTextBuffer
'*****************************************************************
Dim X As Long
Dim j As Long
    
    'Check if the width is larger then the screen
    If EnterTextBufferWidth > GameWindow.ChatWindow.Text.Width - 24 Then
        
        'Loop through the characters backwards
        For X = Len(EnterTextBuffer) To 1 Step -1
            
            'Add up the size
            j = j + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(EnterTextBuffer, X, 1)))
            
            'Check if the size has become too large
            If j > GameWindow.ChatWindow.Text.Width - 24 Then
            
                'If the size has become too large, the character before (since we are looping backwards, it is + 1) is the limit
                ShownText = Right$(EnterTextBuffer, Len(EnterTextBuffer) - X + 1)
                Exit For
                
            End If
        Next X
    Else
    
        'Set the shown text buffer to the full buffer
        ShownText = EnterTextBuffer
    
    End If

End Sub

Sub Main()

'*****************************************************************
'Main
'*****************************************************************
Dim KeyClearTime As Long
Dim PacketKeys() As String
Dim LastUnloadTime As Long
Dim StartTime As Long
Dim i As Integer

    'Set the high-resolution timer
    timeBeginPeriod 1

    'Init file paths
    InitFilePaths
    
    'Load frmMain
    Load frmMain
    frmMain.Hide
    DoEvents

    'Check if we need to run the updater
    If ForceUpdateCheck Then
    
        'Check for the right parameter
        If Command$ <> "-sdf@041jkdf0)21`~" Then

            'Force the creation of frmConnect, thus forcing the creation of its hWnd
            Load frmConnect
            frmConnect.Show
            frmConnect.Hide
            
            'Load the updater
            ShellExecute frmConnect.hwnd, vbNullString, App.Path & "\UpdateClient.exe", vbNullString, vbNullString, 1   'The 1 means "show normal"
    
            'Unload the client
            Engine_UnloadAllForms
            End
        
        End If
    End If
    
    'Generate the packet keys
    GenerateEncryptionKeys PacketKeys
    frmMain.GOREsock.ClearPicture
    frmMain.GOREsock.SetEncryption PacketEncTypeServerIn, PacketEncTypeServerOut, PacketKeys()
    Erase PacketKeys
    
    'Number of bytes required to fill the skills
    NumBytesForSkills = Int((NumSkills - 1) / 8) + 1
    
    'Load the font information
    Engine_Init_FontSettings
    
    'Load the messages
    Engine_Init_Messages LCase$(Var_Get(DataPath & "Game.ini", "INIT", "Language"))
    Engine_Init_Signs LCase$(Var_Get(DataPath & "Game.ini", "INIT", "Language"))
    
    'Fill startup variables for the tile engine
    EnterTextBufferWidth = 1
    ReDim SkillListIDs(1 To NumSkills)

    'Set intial user position
    UserPos.X = 1
    UserPos.Y = 1

    'Set scroll pixels per frame
    ShowGameWindow(QuickBarWindow) = 1
    ShowGameWindow(ChatWindow) = 1

    'Set the array sizes by the number of graphic files
    NumGrhFiles = CLng(Var_Get(DataPath & "Grh.ini", "INIT", "NumGrhFiles"))
    ReDim SurfaceDB(1 To NumGrhFiles)
    ReDim SurfaceSize(1 To NumGrhFiles)
    ReDim SurfaceTimer(1 To NumGrhFiles)
    
    'Load graphic data into memory
    Engine_Init_GrhData
    Engine_Init_BodyData
    Engine_Init_WeaponData
    Engine_Init_WingData
    Engine_Init_HeadData
    Engine_Init_HairData
    
    'Load the config
    Game_Config_Load
    Engine_Init_GUI

    'Create the buffer
    Set sndBuf = New DataBuffer
    sndBuf.Clear

    'Set the form starting positions
    DoEvents

    'Load the data commands
    InitDataCommands
    
    'Build the word filters
    Game_BuildFilter

    'Display connect window
    frmConnect.Visible = True

    'Main Loop
    Do
    
        'Calculate the starttime - this is the absolute time it takes from start to finish, disincluding DoEvents
        ' The idea is that it works just like the ElapsedTime, but in slightly different placing
        StartTime = timeGetTime
    
        'Check if unloading
        If IsUnloading = 1 Then Exit Do
        
        'Clear the key cache
        If KeyClearTime < timeGetTime Then
            Input_Keys_ClearQueue
            KeyClearTime = timeGetTime + 200
        End If
        
        'Don't draw frame is window is minimized or there is no map loaded
        If frmMain.WindowState <> 1 Then
            If CurMap > 0 Then

                'Show the next frame
                Engine_ShowNextFrame

                'Check for key inputs
                Input_Keys_General
                
                'Keep the music looping
                If MapInfo.Music > 0 Then Music_Loop 1

            End If
        End If
        
        'Send the data buffer
        If SocketOpen Then Data_Send

        'Check to unload stuff from memory (only check every 5 seconds)
        If LastUnloadTime < timeGetTime Then
            For i = 1 To NumGrhFiles    'Check to unload surfaces
                If SurfaceTimer(i) > 0 Then 'Only update surfaces in use
                    If SurfaceTimer(i) < timeGetTime Then   'Unload the surface
                        Set SurfaceDB(i) = Nothing
                        SurfaceTimer(i) = 0
                    End If
                End If
            Next i
            For i = 1 To NumSfx 'Check to unload sound buffers
                If SoundBufferTimer(i) > 0 Then 'Only update sound buffers in use
                    If SoundBufferTimer(i) < timeGetTime Then   'Unload the sound buffer
                        Set DSBuffer(i) = Nothing
                        SoundBufferTimer(i) = 0
                    End If
                End If
            Next i
            LastUnloadTime = timeGetTime + 5000 'States we will check the unload routine again in 5000 milliseconds
        End If
        
        'Check to change servers
        If SocketMoveToPort > 0 Then
            If frmMain.GOREsock.ShutDown <> soxERROR Then
                
                'Set up the socket
                'Leave the GetIPFromHost() wrapper there, this will convert a host name to IP if needed, or leave it as an IP if you pass an IP
                SoxID = frmMain.GOREsock.Connect(GetIPFromHost(SocketMoveToIP), SocketMoveToPort)
                SocketOpen = 1
                
                'If the SoxID = -1, then the connection failed, elsewise, we're good to go! W00t! ^_^
                If SoxID = -1 Then
                    MsgBox "Unable to connect to the game server!" & vbCrLf & "Either the server is down or you are not connected to the internet.", vbOKOnly Or vbCritical
                    IsUnloading = 1
                Else
                    frmMain.GOREsock.SetOption SoxID, soxSO_TCP_NODELAY, True
                End If
                
                'Clear the temp values
                SocketMoveToPort = 0
                SocketMoveToIP = vbNullString
                
            End If
        End If

        'Do other events
        DoEvents
        
        'Do sleep event - force FPS at the FPS cap
        If FPSCap > 0 Then
            If (timeGetTime - StartTime) < FPSCap Then  'If Elapsed Time < Time required for requested highest fps
                Sleep FPSCap - (timeGetTime - StartTime)
            End If
        End If

    Loop
    
    'Save the config
    Game_Config_Save
    
    'Close down
    frmMain.ShutdownTimer.Enabled = True

End Sub

Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

    Var_Get = Space$(1000)
    getprivateprofilestring Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub

Public Function Engine_WordWrap(ByVal Text As String, ByVal MaxLineLen As Integer) As String

'************************************************************
'Wrap a long string to multiple lines by vbNewLine
'************************************************************
Dim TempSplit() As String
Dim TSLoop As Long
Dim LastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long

    'Too small of text
    If Len(Text) < 2 Then
        Engine_WordWrap = Text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        b = 1
        LastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": LastSpace = i
                    Case "_": LastSpace = i
                    Case "-": LastSpace = i
                End Select
    
                'Add up the size - Do not count the "|" character (high-lighter)!
                If Not Mid$(TempSplit(TSLoop), i, 1) = "|" Then
                    Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
                End If
                
                'Check for too large of a size
                If Size > MaxLineLen Then
                    
                    'Check if the last space was too far back
                    If i - LastSpace > 4 Then
                        
                        'Too far away to the last space, so break at the last character
                        Engine_WordWrap = Engine_WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                        b = i - 1
                        Size = 0
                        
                    Else
                    
                        'Break at the last space to preserve the word
                        Engine_WordWrap = Engine_WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, LastSpace - b)) & vbNewLine
                        b = LastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = Engine_GetTextWidth(Mid$(TempSplit(TSLoop), LastSpace, i - LastSpace))
                        
                    End If
                    
                End If
                
                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If b <> i Then
                        Engine_WordWrap = Engine_WordWrap & Mid$(TempSplit(TSLoop), b, i)
                    End If
                End If
                
            Next i
            
        Else
        
            Engine_WordWrap = Engine_WordWrap & TempSplit(TSLoop)
        
        End If
        
    Next TSLoop

End Function
