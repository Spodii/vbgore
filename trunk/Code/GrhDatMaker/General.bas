Attribute VB_Name = "General"
Option Explicit

Private Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    FileExist = (Dir$(File, FileType) <> "")

End Function

Private Function GetVar(File As String, Main As String, Var As String) As String

'*****************************************************************
'Get a variable from a a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space$(150) ' This tells the computer how long the longest string can be

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Private Sub Main()

'*****************************************************************
'Loads Grh.raw, parses and outputs Grh.dat
'*****************************************************************

Dim GrhBuffer() As Integer  'Holds all the entered Grh values

Dim sX As Integer
Dim sY As Integer
Dim PixelWidth As Integer
Dim PixelHeight As Integer
Dim FileNum As Long
Dim NumFrames As Byte
Dim Frames() As Long
Dim Speed As Single
Dim LastGrh As Long
Dim Grh As Long
Dim Frame As Long
Dim ln As String
Dim TempLine As String
Dim RawFile As String
Dim NumRawFile As Long
Dim LastFile As Long
Dim Lines As Long

    InitFilePaths

    'Delete any old file
    If FileExist(DataPath & "grh.dat", vbNormal) = True Then Kill DataPath & "grh.dat"

    'Open new file
    Open DataPath & "grh.dat" For Binary As #1
    Seek #1, 1

    RawFile = Dir$(Data2Path & "grh*.raw", vbArchive)

    Do While RawFile <> ""
        Open Data2Path & RawFile For Input As #2

        'Set the buffer's initial size
        ReDim GrhBuffer(1 To 2000000)

        'Do a loop to check for repeat numbers
        While Not EOF(2)
            DoEvents
            Line Input #2, TempLine
            If LCase$(Left$(TempLine, 3)) = "grh" Then
                TempLine = Right$(TempLine, Len(TempLine) - 3)
                If InStr(1, TempLine, "=") <= 0 Then GoTo ErrorHandler
                Grh = CLng(Left$(TempLine, InStr(1, TempLine, "=", vbTextCompare) - 1))
                Lines = Lines + 1
                ReDim Preserve GrhBuffer(1 To Lines)
                GrhBuffer(Lines) = Grh
            End If

        Wend

        'Check for duplicate entries (slow, but whatever - this tool doesn't need to be fast)
        For sX = 1 To Lines
            For sY = sX + 1 To Lines
                If GrhBuffer(sX) = GrhBuffer(sY) Then

                    'Notify of duplicate
                    If MsgBox("Duplcates entries of Grh" & GrhBuffer(sX) & " found. Do you wish to continue compiling Grh.dat?" & vbCrLf _
                              & "Duplicate Grh numbers can lead to graphical display failures and artifacts.", vbYesNo) = vbNo Then
                        Exit Sub
                    End If

                End If
            Next sY
        Next sX

        Close #2

        Open Data2Path & RawFile For Input As #2

        'Clear variables
        sX = 0
        sY = 0
        Lines = 0
        Grh = 0

        While Not EOF(2)
            DoEvents
            Line Input #2, TempLine
            If LCase$(Left$(TempLine, 3)) = "grh" Then
                TempLine = Right$(TempLine, Len(TempLine) - 3)

                If InStr(1, TempLine, "=") <= 0 Then GoTo ErrorHandler

                Grh = CLng(Left$(TempLine, InStr(1, TempLine, "=", vbTextCompare) - 1))
                If Grh > LastGrh Then LastGrh = Grh

                ln = Right$(TempLine, Len(TempLine) - Len(CStr(Grh)) - 1)

                If ln <> "" Then
                    'Get number of frames and check
                    NumFrames = Val(ReadField(1, ln, "-"))
                    If NumFrames <= 0 Then GoTo ErrorHandler

                    'Put grh number
                    Put #1, , Grh
                    
                    'Put number of frames
                    Put #1, , NumFrames

                    If NumFrames > 1 Then
                        ReDim Frames(1 To NumFrames)
                        
                        'Read a animation GRH set
                        For Frame = 1 To NumFrames
                        
                            'Check and put each frame
                            Frames(Frame) = Val(ReadField(Frame + 1, ln, "-"))
                            If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then GoTo ErrorHandler
                            Put #1, , Frames(Frame)
                
                        Next Frame
                        
                        'Check and put speed
                        Speed = CSng(ReadField(NumFrames + 2, ln, "-"))
                        If Speed = 0 Then GoTo ErrorHandler
                        Put #1, , Speed
                        
                    Else
                        'check and put normal GRH data
                        FileNum = Val(ReadField(2, ln, "-"))
                        If FileNum <= 0 Then GoTo ErrorHandler
                        If FileNum > LastFile Then LastFile = FileNum

                        Put #1, , FileNum

                        sX = Val(ReadField(3, ln, "-"))
                        If sX < 0 Then GoTo ErrorHandler
                        Put #1, , sX

                        sY = Val(ReadField(4, ln, "-"))
                        If sY < 0 Then GoTo ErrorHandler
                        Put #1, , sY

                        PixelWidth = Val(ReadField(5, ln, "-"))
                        If PixelWidth <= 0 Then GoTo ErrorHandler
                        Put #1, , PixelWidth

                        PixelHeight = Val(ReadField(6, ln, "-"))
                        If PixelHeight <= 0 Then GoTo ErrorHandler
                        Put #1, , PixelHeight
                    End If
                End If
            End If
        Wend

        Close #2
        RawFile = Dir
    Loop
    Close #1

    WriteVar DataPath & "grh.ini", "INIT", "NumGrhFiles", CStr(LastFile)
    WriteVar DataPath & "grh.ini", "INIT", "NumGrhs", CStr(LastGrh)

    'Display finish box
    MsgBox "Finished.", vbOKOnly

    'Unload
    Erase GrhBuffer
    End
    
Exit Sub
ErrorHandler:

    Close #2
    Close #1
    MsgBox "Error on Grh" & Grh & "!", vbOKOnly Or vbCritical

End Sub

Private Function ReadField(ByVal field_pos As Long, ByVal text As String, ByVal delimiter As String) As String

'*****************************************************************
'Gets a field from a delimited string
'*****************************************************************

Dim i As Long
Dim LastPos As Long
Dim CurrentPos As Long

    LastPos = 0
    CurrentPos = 0

    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, text, delimiter, vbBinaryCompare)
    Next i

    If CurrentPos = 0 Then
        ReadField = Mid$(text, LastPos + 1, Len(text) - LastPos)
    Else
        ReadField = Mid$(text, LastPos + 1, CurrentPos - LastPos - 1)
    End If

End Function

Private Sub WriteVar(File As String, Main As String, Var As String, value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, value, File

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:56)  Decl: 62  Code: 250  Total: 312 Lines
':) CommentOnly: 88 (28.2%)  Commented: 4 (1.3%)  Empty: 67 (21.5%)  Max Logic Depth: 7
