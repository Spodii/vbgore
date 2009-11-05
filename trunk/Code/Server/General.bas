Attribute VB_Name = "General"
Option Explicit

Function Server_LegalString(ByVal CheckString As String) As Boolean

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
        If Server_LegalCharacter(AscB(Mid$(CheckString, i, 1))) = False Then Exit Function
        
    Next i
    
    'If we have made it this far, then all is good
    Server_LegalString = True

Exit Function

ErrOut:

    'Something bad happened, so the string must be invalid
    Server_LegalString = False

End Function

Function Server_LegalCharacter(KeyAscii As Byte) As Boolean

'*****************************************************************
'Only allow certain specified characters
'*****************************************************************

    On Error GoTo ErrOut

    'Allow numbers between 0 and 9
    If KeyAscii >= 48 Or KeyAscii <= 57 Then
        Server_LegalCharacter = True
        Exit Function
    End If
    
    'Allow letters A to Z
    If KeyAscii >= 65 Or KeyAscii <= 90 Then
        Server_LegalCharacter = True
        Exit Function
    End If
    
    'Allow letters a to z
    If KeyAscii >= 97 Or KeyAscii <= 122 Then
        Server_LegalCharacter = True
        Exit Function
    End If
    
Exit Function

ErrOut:

    'Something bad happened, so the character must be invalid
    Server_LegalCharacter = False
    
End Function

Function Server_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single

'*****************************************************************
'Finds the distance between two points
'*****************************************************************

    Server_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))

End Function

Function Server_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
On Error GoTo ErrOut

    If Dir$(File, FileType) <> "" Then Server_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Server_FileExist = False

End Function

Function Server_RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Integer

'*****************************************************************
'Find a Random number between a range
'*****************************************************************

    Server_RandomNumber = Fix((UpperBound - LowerBound + 1) * Rnd) + LowerBound

End Function

Sub Server_RefreshUserListBox()

'*****************************************************************
'Refreshes the User list box
'*****************************************************************

Dim LoopC As Long

    If LastUser < 0 Then
        frmMain.Userslst.Clear
        Exit Sub
    End If

    frmMain.Userslst.Clear
    CurrConnections = 0
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
