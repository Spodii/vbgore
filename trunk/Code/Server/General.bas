Attribute VB_Name = "General"
Option Explicit

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

Function Server_WalkTimePerTile(ByVal Speed As Long) As Long
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
    Server_WalkTimePerTile = 1000 / (((Speed + 4) * 11) / 32) - 150
    
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

    If KeyAscii > 32 Then Server_ValidCharacter = True

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
