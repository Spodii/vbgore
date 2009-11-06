Attribute VB_Name = "General"
Option Explicit

Private GrhBuffer() As Long  'Holds all the entered Grh values in order they appear in GrhRaw.txt
Private GrhBufUBound As Long

Private GrhExist() As Byte   'Holds if the Grh (specified by index) exists or not - created with
Private GrhExistUBound As Long  'the GrhBuffer() array after the duplication check

Private Type BadPaperdollGrh
    Grh As Long
    RetStr As String
    RetStrNum As Long
    File As String
End Type
Private BadPaperdollGrh() As BadPaperdollGrh
Private BadPaperdollGrhUBound As Long

Private Powersof2(1 To 14) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Private Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function

Private Sub MakePowersof2()
Dim i As Long

    For i = 1 To 14
        Powersof2(i) = 2 ^ i
    Next i

End Sub

Private Function IsPowerof2(ByVal Number As Long) As Boolean
Dim i As Long

    For i = 1 To 14
        If Number = Powersof2(i) Then
            IsPowerof2 = True
        Else
            If Powersof2(i) > Number Then Exit Function
        End If
    Next i

End Function

Private Sub CreateGrhExist()

'*****************************************************************
'Turns GrhBuffer() into GrhExist()
'GrhBuffer() holds all of the Grh indexes in order they appear
'GrhExist() holds if a grh exists or not, where the index of GrhExist() is the Grh index
'*****************************************************************
Dim i As Long

    'Find the highest index
    For i = 1 To GrhBufUBound
        If GrhBuffer(i) > GrhExistUBound Then GrhExistUBound = GrhBuffer(i)
    Next i
    
    'Create the GrhExist() array
    ReDim GrhExist(1 To GrhExistUBound)
    
    'Fill in the values for Grhs that exist
    For i = 1 To GrhBufUBound
        GrhExist(GrhBuffer(i)) = 1
    Next i
    
    'Now, every value that wasn't in GrhBuffer() is set to 0 in GrhExist(), and those that are in there are set to 1
    'Erase the GrhBuffer, we're done with it
    Erase GrhBuffer

End Sub

Private Sub Main()

'*****************************************************************
'Loads GrhRaw.txt, parses and outputs Grh.dat
'*****************************************************************
Dim ImageSize As CImageInfo
Dim FileList() As String
Dim TempSplit() As String
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
Dim NumRawFile As Long
Dim LastFile As Long
Dim Lines As Long

    MakePowersof2
    InitFilePaths
    
    '*** Check for valid file sizes ***
    Set ImageSize = New CImageInfo
    FileList() = AllFilesInFolders(GrhPath, False)
    For FileNum = 0 To UBound(FileList)
        If LCase$(Right$(FileList(FileNum), 4)) = ".png" Then
            TempSplit = Split(FileList(FileNum), "\")
            If IsNumeric(Left$(TempSplit(UBound(TempSplit)), Len(TempSplit(UBound(TempSplit))) - 4)) Then
                If Val(Left$(TempSplit(UBound(TempSplit)), Len(TempSplit(UBound(TempSplit))) - 4)) > 32767 Then
                    MsgBox "The following texture file was found with a number higher than 32767:" & vbNewLine & _
                        FileList(FileNum) & vbNewLine & vbNewLine & "You may not use file numbers higher than 32767." & vbNewLine & _
                        "This is for performance reasons, and highly recommended not to try and add support for!", vbOKOnly
                        End
                End If
                ImageSize.ReadImageInfo FileList(FileNum)
                If IsPowerof2(ImageSize.Width) = False Or IsPowerof2(ImageSize.Height) = False Then
                    TempLine = TempLine & TempSplit(UBound(TempSplit))
                End If
            End If
        End If
    Next FileNum
    If LenB(TempLine) Then
        MsgBox "The following image files were found to not have sizes by powers of 2!" & vbNewLine & _
            "Leaving images not in powers of 2 will cause lots of graphical errors!" & vbNewLine & vbNewLine & TempLine & _
            vbNewLine & vbNewLine & "Not to be confused with divisible by two, powers of two goes by 2^X, and contains the values:" & vbNewLine & _
            "2,4,8,16,32,64,128,256,512,1024,2048..." & _
            vbNewLine & vbNewLine & "Grh.dat will still be made, but graphics may not appear correctly in-game.", vbOKOnly Or vbCritical
    End If
    Set ImageSize = Nothing
    TempLine = vbNullString
    FileNum = 0
    sX = 0
    sY = 0
    Erase FileList

    'Delete any old file
    If FileExist(DataPath & "grh.dat", vbNormal) = True Then Kill DataPath & "grh.dat"

    'Open new file
    Open DataPath & "grh.dat" For Binary As #1
    Seek #1, 1

    Open Data2Path & "GrhRaw.txt" For Input As #2

    '*** Do a loop to check for repeat numbers ***
    GrhBufUBound = 1000
    ReDim GrhBuffer(1 To GrhBufUBound)
    
    'Grab all the numbers
    While Not EOF(2)
        Line Input #2, TempLine

        If LCase$(Left$(TempLine, 3)) = "grh" Then
            TempLine = Right$(TempLine, Len(TempLine) - 3)
            If InStr(1, TempLine, "=") <= 0 Then GoTo ErrorHandler
            Grh = CLng(Left$(TempLine, InStr(1, TempLine, "=", vbTextCompare) - 1))
            Lines = Lines + 1
            If Lines > GrhBufUBound Then
                GrhBufUBound = GrhBufUBound + 250
                ReDim Preserve GrhBuffer(1 To GrhBufUBound)
            End If
            GrhBuffer(Lines) = Grh
        End If
        
    Wend
    
    'Trim down to the smallest buffer size
    GrhBufUBound = Lines
    ReDim Preserve GrhBuffer(1 To GrhBufUBound)

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
    
    'Create the GrhExist() array
    CreateGrhExist
    
    '*** Build Grh.dat ***

    Open Data2Path & "GrhRaw.txt" For Input As #2

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

            If ln <> vbNullString Then
            
                'Split the string
                TempSplit() = Split(ln, "-")
            
                'Get number of frames and check
                NumFrames = Val(TempSplit(0))
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
                        Frames(Frame) = Val(TempSplit(Frame))
                        If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then GoTo ErrorHandler
                        Put #1, , Frames(Frame)
            
                    Next Frame
                    
                    'Check and put speed
                    Speed = CSng(TempSplit(NumFrames + 1))
                    If Speed = 0 Then GoTo ErrorHandler
                    Put #1, , Speed
                    
                Else
                    'check and put normal GRH data
                    FileNum = Val(TempSplit(1))
                    If FileNum <= 0 Then GoTo ErrorHandler
                    If FileNum > LastFile Then LastFile = FileNum

                    Put #1, , FileNum

                    sX = Val(TempSplit(2))
                    If sX < 0 Then GoTo ErrorHandler
                    Put #1, , sX

                    sY = Val(TempSplit(3))
                    If sY < 0 Then GoTo ErrorHandler
                    Put #1, , sY

                    PixelWidth = Val(TempSplit(4))
                    If PixelWidth <= 0 Then GoTo ErrorHandler
                    Put #1, , PixelWidth

                    PixelHeight = Val(TempSplit(5))
                    If PixelHeight <= 0 Then GoTo ErrorHandler
                    Put #1, , PixelHeight
                End If
            End If
        End If
    Wend

    Close #2
    Close #1

    WriteVar DataPath & "grh.ini", "INIT", "NumGrhFiles", CStr(LastFile)
    WriteVar DataPath & "grh.ini", "INIT", "NumGrhs", CStr(LastGrh)
    
    'Check if the entries in the paperdolling .dat files
    CheckPaperdollGrhs DataPath & "Body.dat"
    CheckPaperdollGrhs DataPath & "Hair.dat"
    CheckPaperdollGrhs DataPath & "Wing.dat"
    CheckPaperdollGrhs DataPath & "Head.dat"
    CheckPaperdollGrhs DataPath & "Weapon.dat"

    'Display bad paperdoll grhs
    If BadPaperdollGrhUBound > 0 Then
        If MsgBox(BadPaperdollGrhUBound & " bad Grh entries have been found in your paper-dolling files. Do you wish to view them?", vbYesNo) = vbYes Then
            For sX = 1 To BadPaperdollGrhUBound
                With BadPaperdollGrh(BadPaperdollGrhUBound)
                    MsgBox "File: " & .File & vbNewLine & _
                        "Line (" & .RetStrNum & "): " & .RetStr & vbNewLine & _
                        "Grh value used: " & .Grh, vbOKOnly
                End With
            Next sX
        End If
    End If

    'Display finish box
    MsgBox "Finished.", vbOKOnly

    'Unload
    End
    
Exit Sub
ErrorHandler:
Dim Loc1 As Long
Dim Loc2 As Long

    Loc1 = Loc(1)
    Loc2 = Loc(2)
    Close #2
    Close #1
    MsgBox "Error on Grh" & Grh & "!" & vbNewLine & vbNewLine & "Last GrhRaw.txt line: " & Loc2 & vbNewLine & "Last Grh.Dat line: " & Loc1, vbOKOnly Or vbCritical

End Sub

Private Sub AddBadPaperdollGrh(ByVal GrhIndex As Long, ByVal Line As String, ByVal LineNum As Long, ByVal File As String)

'*****************************************************************
'Adds an entry to the "Bad Paperdoll Grh" list
'*****************************************************************

    BadPaperdollGrhUBound = BadPaperdollGrhUBound + 1
    ReDim Preserve BadPaperdollGrh(1 To BadPaperdollGrhUBound)
    With BadPaperdollGrh(BadPaperdollGrhUBound)
        .File = File
        .RetStr = Line
        .Grh = GrhIndex
        .RetStrNum = LineNum
    End With

End Sub

Private Sub CheckPaperdollGrhs(ByVal FilePath As String)

'*****************************************************************
'Checks that the Grh values in the paperdolling files are valid
'*****************************************************************
Dim FileNum As Byte
Dim s() As String
Dim ln As String
Dim Origln As String
Dim v As Long
Dim FileName As String

    'Check the file exists
    If Dir$(FilePath, vbNormal) = vbNullString Then Exit Sub

    'Get the file name
    s = Split(FilePath, "\")
    FileName = UBound(s)

    'Open the file
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    
        'Loop through the whole file
        Do While Not EOF(FileNum)
        
            'Grab the line
            Line Input #FileNum, ln
            ln = Trim$(ln)
            Origln = ln
            
            'Check for a valid line
            If ln <> vbNullString Then
                If Len(ln) > 2 Then
                    If Left$(ln, 1) <> "'" Then
                        If Left$(ln, 1) <> "[" Then
                            ln = UCase$(ln)
                            If InStr(1, ln, "=") Then
                                If Left$(ln, 3) <> "NUM" Then
                                    If Left$(ln, 10) <> "HEADOFFSET" Then
                                        
                                        'Split the string by the equal sign
                                        s() = Split(ln, "=")
                                        
                                        'If there is more than 1 = sign, we don't want it
                                        If UBound(s) <> 1 Then Exit Sub
                                        
                                        'Check the value
                                        s(1) = Trim$(s(1))
                                        If Not IsNumeric(s(1)) Then Exit Sub
                                        v = Val(s(1))
                                        
                                        'Finally, we can assume we have a Grh value, so check if it is valid
                                        If v <= 0 Then
                                            AddBadPaperdollGrh v, Origln, Loc(FileNum), FileName
                                            GoTo NextLoop
                                        End If
                                        If v > GrhExistUBound Then
                                            AddBadPaperdollGrh v, Origln, Loc(FileNum), FileName
                                            GoTo NextLoop
                                        End If
                                        If GrhExist(v) = 0 Then
                                            AddBadPaperdollGrh v, Origln, Loc(FileNum), FileName
                                            GoTo NextLoop
                                        End If
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
                                            
NextLoop:
            
        Loop
        
    Close #FileNum

End Sub

Private Sub WriteVar(File As String, Main As String, Var As String, value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    WritePrivateProfileString Main, Var, value, File

End Sub

