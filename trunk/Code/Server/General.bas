Attribute VB_Name = "General"
Option Explicit

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single

'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function

Private Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte

'*****************************************************************
'Check if two lines intersect (return 1 if true)
'*****************************************************************

Dim m1 As Single
Dim M2 As Single
Dim B1 As Single
Dim B2 As Single
Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    B1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    B2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If B2 = B1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((B2 - B1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
        
    End If
    
End Function

Public Function Engine_ClearPath(ByVal Map As Integer, ByVal UserX As Long, ByVal UserY As Long, ByVal TargetX As Long, ByVal TargetY As Long) As Byte

'***************************************************
'Check if the path is clear from the user to the target of blocked tiles
'For the line-rect collision, we pretend that each tile is 2 units wide so we can give them a width of 1 to center things
'***************************************************
Dim X As Long
Dim Y As Long

    '****************************************
    '***** Target is on top of the user *****
    '****************************************
    
    'If the target position = user position, we must be targeting ourself, so nothing can be blocking us from us (I hope o.O )
    If UserX = TargetX Then
        If UserY = TargetY Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If

    '********************************************
    '***** Target is right next to the user *****
    '********************************************
    
    'Target is at one of the 4 diagonals of the user
    If Abs(UserX - TargetX) = 1 Then
        If Abs(UserY - TargetY) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Target is above or below the user
    If UserX = TargetX Then
        If Abs(UserY - TargetY) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Target is to the left or right of the user
    If UserY = TargetY Then
        If Abs(UserX - TargetX) = 1 Then
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    '********************************************
    '***** Target is diagonal from the user *****
    '********************************************
    
    'Check if the target is diagonal from the user - only do the following checks if diagonal from the target
    If Abs(UserX - TargetX) = Abs(UserY - TargetY) Then

        If UserX > TargetX Then
                        
            'Diagonal to the top-left
            If UserY > TargetY Then
                For X = TargetX To UserX - 1
                    For Y = TargetY To UserY - 1
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            
            'Diagonal to the bottom-left
            Else
                For X = TargetX To UserX - 1
                    For Y = UserY + 1 To TargetY
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            End If

        End If
        
        If UserX < TargetX Then
        
            'Diagonal to the top-right
            If UserY > TargetY Then
                For X = UserX + 1 To TargetX
                    For Y = TargetY To UserY - 1
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
                
            'Diagonal to the bottom-right
            Else
                For X = UserX + 1 To TargetX
                    For Y = UserY + 1 To TargetY
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    Next Y
                Next X
            End If
        
        End If
    
        Engine_ClearPath = 1
        Exit Function
    
    End If

    '*******************************************************************
    '***** Target is directly vertical or horizontal from the user *****
    '*******************************************************************
    
    'Check if target is directly above the user
    If UserX = TargetX Then 'Check if x values are the same (straight line between the two)
        If UserY > TargetY Then
            For Y = TargetY + 1 To UserY - 1
                If MapData(Map, UserX, Y).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next Y
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly below the user
    If UserX = TargetX Then
        If UserY < TargetY Then
            For Y = UserY + 1 To TargetY - 1
                If MapData(Map, UserX, Y).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next Y
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly to the left of the user
    If UserY = TargetY Then
        If UserX > TargetX Then
            For X = TargetX + 1 To UserX - 1
                If MapData(Map, X, UserY).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next X
            Engine_ClearPath = 1
            Exit Function
        End If
    End If
    
    'Check if the target is directly to the right of the user
    If UserY = TargetY Then
        If UserX < TargetX Then
            For X = UserX + 1 To TargetX - 1
                If MapData(Map, X, UserY).Blocked And 128 Then
                    Engine_ClearPath = 0
                    Exit Function
                End If
            Next X
            Engine_ClearPath = 1
            Exit Function
        End If
    End If

    '*******************************************************************
    '***** Target is directly vertical or horizontal from the user *****
    '*******************************************************************
    
    
    If UserY > TargetY Then
    
        'Check if the target is to the top-left of the user
        If UserX > TargetX Then
            For X = TargetX To UserX - 1
                For Y = TargetY To UserY - 1
                    'We must do * 2 on the tiles so we can use +1 to get the center (its like * 32 and + 16 - this does the same affect)
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
            Engine_ClearPath = 1
            Exit Function
    
        'Check if the target is to the top-right of the user
        Else
            For X = UserX + 1 To TargetX
                For Y = TargetY To UserY - 1
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        End If
        
    Else
    
        'Check if the target is to the bottom-left of the user
        If UserX > TargetX Then
            For X = TargetX To UserX - 1
                For Y = UserY + 1 To TargetY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        
        'Check if the target is to the bottom-right of the user
        Else
            For X = UserX + 1 To TargetX
                For Y = UserY + 1 To TargetY
                    If Engine_Collision_LineRect(X * 2, Y * 2, 2, 2, UserX * 2 + 1, UserY * 2 + 1, TargetX * 2 + 1, TargetY * 2 + 1) Then
                        If MapData(Map, X, Y).Blocked And 128 Then
                            Engine_ClearPath = 0
                            Exit Function
                        End If
                    End If
                Next Y
            Next X
        End If
    
    End If
    
    Engine_ClearPath = 1

End Function

Private Function Engine_Collision_LineRect(ByVal SX As Long, ByVal SY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Byte

'*****************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'*****************************************************************

    'Top line
    If Engine_Collision_Line(SX, SY, SX + SW, SY, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(SX + SW, SY, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(SX, SY + SH, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(SX, SY, SX, SY + SW, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

End Function

Private Function Engine_Collision_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte

'*****************************************************************
'Find if a value is between two other values (used for line collision)
'*****************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If Value >= Bound1 Then
            If Value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
End Function

Public Function ByteArrayToStr(ByRef ByteArray() As Byte) As String

'*****************************************************************
'Take a byte array and print it out in a readable string
'Example output: 084[T] 086[V] 088[X] 090[Z] 092[\] 094[^]
'*****************************************************************

On Error GoTo ErrOut

Dim i As Long
    
    Log "ByteArrayToStr: ByteArray LBound() = " & LBound(ByteArray) & " UBound() = " & UBound(ByteArray), CodeTracker '//\\LOGLINE//\\
    For i = LBound(ByteArray) To UBound(ByteArray)
        If ByteArray(i) >= 100 Then
            ByteArrayToStr = ByteArrayToStr & ByteArray(i) & "[" & Chr$(ByteArray(i)) & "] "
        ElseIf ByteArray(i) >= 10 Then
            ByteArrayToStr = ByteArrayToStr & "0" & ByteArray(i) & "[" & Chr$(ByteArray(i)) & "] "
        Else
            ByteArrayToStr = ByteArrayToStr & "00" & ByteArray(i) & "[" & Chr$(ByteArray(i)) & "] "
        End If
    Next i
    ByteArrayToStr = Left$(ByteArrayToStr, Len(ByteArrayToStr) - 1)
    
'If there was an error, we were probably passed an erased ByteArray
ErrOut:

    Log "ByteArrayToStr: Unknown error in routine!", CriticalError '//\\LOGLINE//\\
    
End Function

Function Server_WalkTimePerTile(ByVal Speed As Long, Optional ByVal LagBuffer As Byte = 150) As Long
'*****************************************************************
'Takes a speed value and returns the time it takes to walk a tile
'To fine the value:
'(Speed + 4) * BaseWalkSpeed = Pixels/second
'Pixels/sec / 32 = Tiles/sec
'1000 / Tiles/sec = Seconds per tile - how long it takes to walk by one tile
'*****************************************************************

    Log "Call Server_WalkTimePerTile(" & Speed & ")", CodeTracker '//\\LOGLINE//\\

    '4 = The client works off a base value of 4 for speed, so the speed is calculated as 4 + Speed in the client
    '11 = BaseWalkSpeed - how fast we move in pixels/sec
    '32 = The size of a tile
    '150 = We have to give some slack for network lag and client lag - raise this value if people skip too much
    '     and lower it if people are speedhacking and getting too much extra speed
    '1000 = Miliseconds in a second
    Server_WalkTimePerTile = 1000 / (((Speed + 4) * 11) / 32) - LagBuffer
    
    Log "Rtrn Server_WalkTimePerSecond = " & Server_WalkTimePerTile, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_UserExist(ByVal UserName As String) As Boolean
'*****************************************************************
'Checks the database for if a user exists by the specified name
'*****************************************************************

    Log "Call Server_UserExist(" & UserName & ")", CodeTracker '//\\LOGLINE//\\

    'Make the query
    DB_RS.Open "SELECT name FROM users WHERE `name`='" & UserName & "'", DB_Conn, adOpenStatic, adLockOptimistic

    'If End Of File = true, then the user doesn't exist
    If DB_RS.EOF = True Then Server_UserExist = False Else Server_UserExist = True

    'Close the recordset
    DB_RS.Close
    
    Log "Rtrn Server_UserExist = " & Server_UserExist, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_LegalString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for illegal characters in the string (string wrapper for Server_LegalCharacter)
'*****************************************************************
Dim b() As Byte
Dim i As Long

    Log "Call Server_LegalString(" & CheckString & ")", CodeTracker '//\\LOGLINE//\\

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then
        Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If LenB(CheckString) < 1 Then
        Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Copy the string to a byte array
    b() = StrConv(CheckString, vbFromUnicode)

    'Loop through the string
    For i = 0 To UBound(b)
        
        'Check the values
        If Server_LegalCharacter(b(i)) = False Then
            Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
        
    Next i
    
    'If we have made it this far, then all is good
    Server_LegalString = True
    
    Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Server_LegalString = False
    
    Log "Rtrn Server_LegalString = " & Server_LegalString, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_ValidString(ByVal CheckString As String) As Boolean

'*****************************************************************
'Check for valid characters in the string (string wrapper for Server_ValidCharacter)
'Make sure to update on the client, too!
'*****************************************************************
Dim b() As Byte
Dim i As Long

    Log "Call Server_ValidString(" & CheckString & ")", CodeTracker '//\\LOGLINE//\\

    On Error GoTo ErrOut

    'Check for invalid string
    If CheckString = vbNullChar Then
        Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    If LenB(CheckString) < 1 Then
        Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Copy the string to a byte array
    b() = StrConv(CheckString, vbFromUnicode)

    'Loop through the string
    For i = 0 To UBound(b)
        
        'Check the values
        If Server_ValidCharacter(b(i)) = False Then
            Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\
            Exit Function
        End If
        
    Next i
    
    'If we have made it this far, then all is good
    Server_ValidString = True
    
    Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Server_ValidString = False
    
    Log "Rtrn Server_ValidString = " & Server_ValidString, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_ValidCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is used for chat/etc)
'Make sure you update the client's Game_ValidCharacter, too!
'*****************************************************************

    Log "Call Server_ValidCharacter(" & KeyAscii & ")", CodeTracker '//\\LOGLINE//\\

    If KeyAscii >= 32 Then Server_ValidCharacter = True

End Function

Function Server_LegalCharacter(ByVal KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters (this is for username/pass)
'Make sure you update the client's Game_LegalCharacter, too!
'*****************************************************************

    Log "Call Server_LegalCharacter(" & KeyAscii & ")", CodeTracker '//\\LOGLINE//\\
    
    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Server_LegalCharacter = True
        Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Allow letters A to Z
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        Server_LegalCharacter = True
        Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Allow letters a to z
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        Server_LegalCharacter = True
        Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    'Allow foreign characters
    If KeyAscii >= 128 And KeyAscii <= 168 Then
        Server_LegalCharacter = True
        Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
        Exit Function
    End If
    
    Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
    
Exit Function

ErrOut:

    'Something bad happened, so the character must be invalid
    Server_LegalCharacter = False
    Log "Rtrn Server_LegalCharacter = " & Server_LegalCharacter, CodeTracker '//\\LOGLINE//\\
    
End Function

Function Server_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single

'*****************************************************************
'Finds the distance between two points
'*****************************************************************

    Log "Call Server_Distance(" & x1 & "," & Y1 & "," & x2 & "," & Y2 & ")", CodeTracker '//\\LOGLINE//\\

    Server_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))
    
    Log "Rtrn Server_Distance = " & Server_Distance, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_RectDistance(ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal MaxXDist As Long, ByVal MaxYDist As Long) As Byte

'*****************************************************************
'Check if two tile points are in the same screen
'*****************************************************************

    Log "Call Server_RectDistance(" & x1 & "," & Y1 & "," & x2 & "," & Y2 & "," & MaxXDist & "," & MaxYDist & ")", CodeTracker '//\\LOGLINE//\\

    If Abs(x1 - x2) < MaxXDist + 1 Then
        If Abs(Y1 - Y2) < MaxYDist + 1 Then
            Server_RectDistance = True
        End If
    End If
    
    Log "Rtrn Server_RectDistance = " & Server_RectDistance, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
On Error GoTo ErrOut
    
    Log "Call Server_FileExist(" & File & "," & FileType & ")", CodeTracker '//\\LOGLINE//\\

    If Dir$(File, FileType) <> "" Then Server_FileExist = True
    
    Log "Rtrn Server_FileExist = " & Server_FileExist, CodeTracker '//\\LOGLINE//\\

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Server_FileExist = False
    Log "Rtrn Server_FileExist = " & Server_FileExist, CodeTracker '//\\LOGLINE//\\

End Function

Function Server_RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Integer

'*****************************************************************
'Find a Random number between a range
'*****************************************************************

    Log "Call Server_RandomNumber(" & LowerBound & "," & UpperBound & ")", CodeTracker '//\\LOGLINE//\\

    Server_RandomNumber = Fix((UpperBound - LowerBound + 1) * Rnd) + LowerBound
    
    Log "Rtrn Server_RandomNumber = " & Server_RandomNumber, CodeTracker '//\\LOGLINE//\\

End Function

Sub Server_RefreshUserListBox()

'*****************************************************************
'Refreshes the User list box
'*****************************************************************

Dim LoopC As Long

    Log "Call Server_RefreshUserListBox", CodeTracker '//\\LOGLINE//\\

    If LastUser < 0 Then
        Log "Server_RefreshUserListBox: No users to update", CodeTracker '//\\LOGLINE//\\
        frmMain.Userslst.Clear
        Exit Sub
    End If

    frmMain.Userslst.Clear
    CurrConnections = 0
    Log "Server_RefreshUserListBox: Updating " & LastUser & " users", CodeTracker '//\\LOGLINE//\\
    For LoopC = 1 To LastUser
        If UserList(LoopC).Name <> "" Then
            frmMain.Userslst.AddItem UserList(LoopC).Name
            CurrConnections = CurrConnections + 1
        End If
    Next LoopC
    TrayModify ToolTip, "Game Server: " & CurrConnections & " connections"

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:47)  Decl: 1  Code: 368  Total: 369 Lines
':) CommentOnly: 42 (11.4%)  Commented: 0 (0%)  Empty: 46 (12.5%)  Max Logic Depth: 4
