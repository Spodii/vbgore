Attribute VB_Name = "Compressions"
Option Explicit

Public Enum CompressMethods
    RLE = 1
    RLE_Loop = 2
    LZMA = 3
    PAQ8l = 4
    Deflate64 = 5
    MonkeyAudio = 6     '*.wav only
End Enum
#If False Then
Private RLE, RLE_Loop, LZMA, PAQ8l, Deflate64, MonkeyAudio
#End If

'Value between 0 and 9, higher requiring more RAM/CPU but better compression
'Keep in mind decompressing requires a lot of RAM, too, so don't go higher than 7
Private Const PAQ8l_Level As Byte = 6

Private CompressArray() As Byte
Private PosStream() As Byte
Private DistStream() As Byte
Private ContStream() As Byte
Private LengthStream() As Byte
Private OutStream() As Byte
Private OutPos As Long
Private PosPos As Long
Private DistPos As Long
Private ReadBitPos As Integer
Private CntPos As Long
Private CntByteBuf As Integer
Private CntBitCount As Integer
Private LengthPos As Long

Private Dict() As String
Private AddDict As Integer
Private addDictPos As Integer
Private MaxDictBitPos As Integer
Private MaxDict As Integer
Private NowBitLength As Integer
Private UsedDicts As Integer
Private Const DictionarySize As Long = 3

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationname As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub CommandLine(ByVal CommandLineString As String)
Dim Start As STARTUPINFO
Dim Proc As PROCESS_INFORMATION

    Start.dwFlags = &H1
    Start.wShowWindow = 0
    CreateProcessA 0&, CommandLineString, 0&, 0&, 1&, &H20&, 0&, 0&, Start, Proc
    Do While WaitForSingleObject(Proc.hProcess, 0) = 258
        DoEvents
        Sleep 1
    Loop

End Sub

Private Sub SaveBytes(ByRef Bytes() As Byte, ByVal File As String)
Dim FileNum As Byte

    If Dir$(File, vbNormal) <> vbNullString Then Kill File
    FileNum = FreeFile
    Open File For Binary Access Write As #FileNum
        Put #FileNum, , Bytes()
    Close #FileNum

End Sub

Private Sub LoadBytes(ByRef Bytes() As Byte, ByVal File As String)
Dim FileNum As Byte

    FileNum = FreeFile
    Open File For Binary Access Read As #FileNum
        If LOF(FileNum) > 0 Then
            ReDim Bytes(0 To LOF(FileNum) - 1)
            Get #FileNum, , Bytes()
        End If
    Close #FileNum

End Sub

Public Sub Compression_Compress_PAQ8l(ByteArray() As Byte)

    SaveBytes ByteArray(), App.Path & "\contemp.bin"
    CommandLine DataPath & "paq8l.exe -" & PAQ8l_Level & " " & Chr$(34) & App.Path & "\contemp.bin" & Chr$(34)
    If Dir$(App.Path & "\contemp.bin.paq8l") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.bin.paq8l"
        Kill App.Path & "\contemp.bin.paq8l"
    End If
    Kill App.Path & "\contemp.bin"

End Sub

Public Sub Compression_DeCompress_PAQ8l(ByteArray() As Byte)

    SaveBytes ByteArray(), App.Path & "\contemp.bin.paq8l"
    CommandLine DataPath & "paq8l.exe -d " & Chr$(34) & App.Path & "\contemp.bin.paq8l" & Chr$(34)
    If Dir$(App.Path & "\contemp.bin") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.bin"
        Kill App.Path & "\contemp.bin"
    End If
    Kill App.Path & "\contemp.bin.paq8l"

End Sub

Private Sub Compression_Add_ASCtoArray(WhichArray() As Byte, ToPos As Long, ByVal Text As String)

Dim X As Long

    If ToPos + Len(Text) > UBound(WhichArray) Then ReDim Preserve WhichArray(ToPos + Len(Text) + 500)
    For X = 1 To Len(Text)
        WhichArray(ToPos) = Asc(Mid$(Text, X, 1))
        ToPos = ToPos + 1
    Next X

End Sub

Private Sub Compression_Add_BitsToContStream(ByVal Number As Long, ByVal NumBits As Integer)

Dim X As Long

    For X = NumBits - 1 To 0 Step -1
        CntByteBuf = CntByteBuf * 2 + (-1 * ((Number And CDbl(2 ^ X)) > 0))
        CntBitCount = CntBitCount + 1
        If CntBitCount = 8 Then
            ContStream(CntPos) = CntByteBuf
            CntBitCount = 0
            CntByteBuf = 0
            CntPos = CntPos + 1
            If CntPos > UBound(ContStream) Then
                ReDim Preserve ContStream(CntPos + 500)
            End If
        End If
    Next X

End Sub

Private Sub Compression_Add_CharToArray(ToArray() As Byte, ToPos As Long, ByVal Char As Byte)

    If ToPos > UBound(ToArray) Then
        ReDim Preserve ToArray(ToPos + 500)
    End If
    ToArray(ToPos) = Char
    ToPos = ToPos + 1

End Sub

Private Sub Compression_Add_CharToDict4(Char As String)

    Do
        If LenB(Dict(AddDict)) = 0 Then Dict(AddDict) = String$(255, Asc(" "))
        If addDictPos + Len(Char) < 255 Then
            Mid$(Dict(AddDict), addDictPos, Len(Char)) = Char
            addDictPos = addDictPos + Len(Char)
            Char = vbNullString
        Else
            If addDictPos < 256 Then
                Mid$(Dict(AddDict), addDictPos, 256 - addDictPos) = Left$(Char, 256 - addDictPos)
                Char = Mid$(Char, 256 - addDictPos)
            End If
            addDictPos = 1
            AddDict = AddDict + 1
            If AddDict > MaxDict Then AddDict = 1
            If AddDict > UsedDicts Then UsedDicts = AddDict
        End If
    Loop While LenB(Char)

End Sub

Public Sub Compression_Compress(SrcFile As String, DestFile As String, Compression As CompressMethods)
Dim Dummy As Boolean

    If Compression_File_Load(SrcFile) = 0 Then Exit Sub
    Select Case Compression
        Case RLE
            Compression_Compress_RLE CompressArray(), Dummy
        Case RLE_Loop
            Compression_Compress_RLELoop CompressArray()
        Case LZMA
            Compression_Compress_LZMA CompressArray()
        Case PAQ8l
            Compression_Compress_PAQ8l CompressArray()
        Case Deflate64
            Compression_Compress_Deflate64 CompressArray()
    End Select
    Compression_File_Save DestFile

End Sub

Public Sub Compression_Compress_LZMA(ByteArray() As Byte)

    SaveBytes ByteArray(), App.Path & "\contemp.bin"
    CommandLine DataPath & "7za.exe a -t7z " & Chr$(34) & App.Path & "\contemp.bin.7z" & Chr$(34) & " -aoa " & Chr$(34) & App.Path & "\contemp.bin" & Chr$(34) & " -mx9 -m0=LZMA:d80m:fb273:lc5:pb1:mc10000"
    If Dir$(App.Path & "\contemp.bin.7z") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.bin.7z"
        Kill App.Path & "\contemp.bin.7z"
    End If
    Kill App.Path & "\contemp.bin"

End Sub

Public Sub Compression_Compress_MonkeyAudio(ByteArray() As Byte)

    '*.wav only
    SaveBytes ByteArray(), App.Path & "\contemp.wav"
    CommandLine DataPath & "mac.exe " & Chr$(34) & App.Path & "\contemp.wav" & Chr$(34) & " " & Chr$(34) & App.Path & "\contemp.wav.ape" & Chr$(34) & " -c5000"
    If Dir$(App.Path & "\contemp.wav.ape") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.wav.ape"
        Kill App.Path & "\contemp.wav.ape"
    End If
    Kill App.Path & "\contemp.wav"

End Sub

Public Sub Compression_DeCompress_MonkeyAudio(ByteArray() As Byte)

    '*.wav only
    SaveBytes ByteArray(), App.Path & "\contemp.wav.ape"
    CommandLine DataPath & "mac.exe " & Chr$(34) & App.Path & "\contemp.wav.ape" & Chr$(34) & " " & Chr$(34) & App.Path & "\contemp.wav" & Chr$(34) & " -d"
    If Dir$(App.Path & "\contemp.wav") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.wav"
        Kill App.Path & "\contemp.wav"
    End If
    Kill App.Path & "\contemp.wav.ape"

End Sub

Public Sub Compression_Compress_Deflate64(ByteArray() As Byte)

    SaveBytes ByteArray(), App.Path & "\contemp.bin"
    CommandLine DataPath & "7za.exe a -tzip " & Chr$(34) & App.Path & "\contemp.bin.7z" & Chr$(34) & " -aoa " & Chr$(34) & App.Path & "\contemp.bin" & Chr$(34) & " -mx9 -mm=Deflate64 -mfb=257 -mpass=15 -mmc=1000"
    If Dir$(App.Path & "\contemp.bin.7z") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.bin.7z"
        Kill App.Path & "\contemp.bin.7z"
    End If
    Kill App.Path & "\contemp.bin"

End Sub

Public Sub Compression_DeCompress_Deflate64(ByteArray() As Byte)

    SaveBytes ByteArray(), App.Path & "\contemp.bin.7z"
    CommandLine DataPath & "7za.exe e " & Chr$(34) & App.Path & "\contemp.bin.7z" & Chr$(34)
    If Dir$(App.Path & "\contemp.bin") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.bin"
        Kill App.Path & "\contemp.bin"
    End If
    Kill App.Path & "\contemp.bin.7z"

End Sub

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed by the decompressor

Public Sub Compression_Compress_RLE(ByteArray() As Byte, IsCompressed As Boolean)

Dim X As Long
Dim Y As Long
Dim ByteCount As Long
Dim LastAsc As Integer
Dim TelSame As Long
Dim IsRun As Boolean
Dim ZeroCount As Integer
Dim LengthPos As Long
Dim NoLength As Boolean

    If UBound(ByteArray) = 0 Then Exit Sub

    ReDim ContStream(200)
    ReDim LengthStream(200)
    ReDim OutStream(500)
    IsCompressed = False
    CntPos = 1
    OutPos = 0

    For X = 0 To UBound(ByteArray)
        IsRun = LastAsc = ByteArray(X) And X <> 0
        If Not IsRun Then
            If TelSame = 1 Then
                TelSame = 0
                Compression_Add_CharToArray OutStream, OutPos, CByte(LastAsc)
                ByteCount = ByteCount + 1
            ElseIf TelSame > 1 Then
                For Y = 1 To Int(ByteCount / 255)
                    Compression_Add_CharToArray ContStream, CntPos, 255
                Next Y
                ByteCount = ByteCount Mod 255
                If ByteCount = 0 Then ZeroCount = ZeroCount + 1
                Compression_Add_CharToArray ContStream, CntPos, CByte(ByteCount)
                ByteCount = 0
                For Y = 1 To Int(TelSame / 255)
                    Compression_Add_CharToArray LengthStream, LengthPos, 255
                Next Y
                TelSame = TelSame Mod 255
                Compression_Add_CharToArray LengthStream, LengthPos, CByte(TelSame)
                TelSame = 0
            End If
            Compression_Add_CharToArray OutStream, OutPos, ByteArray(X)
            ByteCount = ByteCount + 1
        Else
            TelSame = TelSame + 1
        End If
        LastAsc = ByteArray(X)
    Next X
    
    If IsRun Then
        If TelSame < 2 Then
            Compression_Add_CharToArray OutStream, OutPos, CByte(LastAsc)
        Else
            For Y = 1 To Int(ByteCount / 255)
                Compression_Add_CharToArray ContStream, CntPos, 255
            Next Y
            ByteCount = ByteCount Mod 255
            Compression_Add_CharToArray ContStream, CntPos, CByte(ByteCount)
            For Y = 1 To Int(TelSame / 255)
                Compression_Add_CharToArray LengthStream, LengthPos, 255
            Next Y
            TelSame = TelSame Mod 255
            Compression_Add_CharToArray LengthStream, LengthPos, CByte(TelSame)
        End If
    End If
    
    ContStream(0) = CByte(ZeroCount)
    If CntPos > 1 Then IsCompressed = True
    Call Compression_Add_CharToArray(ContStream, CntPos, 0)  'No Run Till EOF
    ReDim Preserve ContStream(CntPos - 1) As Byte
    
    If LengthPos > 0 Then
        ReDim Preserve LengthStream(LengthPos - 1)
    Else
        NoLength = True
    End If
    
    ReDim Preserve OutStream(OutPos - 1) As Byte
    CntPos = UBound(ContStream) + 1
    LengthPos = 0
    If Not NoLength Then LengthPos = UBound(LengthStream) + 1
    OutPos = UBound(OutStream) + 1
    ReDim ByteArray(CntPos + LengthPos + OutPos - 1)
    CopyMem ByteArray(0), ContStream(0), CntPos
    If LengthPos > 0 Then CopyMem ByteArray(CntPos), LengthStream(0), LengthPos
    CopyMem ByteArray(CntPos + LengthPos), OutStream(0), OutPos

End Sub

Public Sub Compression_Compress_RLELoop(ByteArray() As Byte)
Dim TimesRLE As Integer
Dim IsCompressed As Boolean

    Do
        Compression_Compress_RLE ByteArray, IsCompressed
        TimesRLE = TimesRLE + 1
    Loop While IsCompressed
    ReDim Preserve ByteArray(UBound(ByteArray) + 1)
    ByteArray(UBound(ByteArray)) = TimesRLE

End Sub

Public Sub Compression_DeCompress(SrcFile As String, DestFile As String, Compression As CompressMethods)

    If Compression_File_Load(SrcFile) = 0 Then Exit Sub
    Select Case Compression
        Case RLE
            Compression_DeCompress_RLE CompressArray()
        Case RLE_Loop
            Compression_DeCompress_RLELoop CompressArray()
        Case LZMA
            Compression_DeCompress_LZMA CompressArray()
        Case PAQ8l
            Compression_DeCompress_PAQ8l CompressArray()
        Case Deflate64
            Compression_DeCompress_Deflate64 CompressArray()
    End Select
    Compression_File_Save DestFile

End Sub

Public Sub Compression_DeCompress_LZMA(ByteArray() As Byte)

    SaveBytes ByteArray(), App.Path & "\contemp.bin.7z"
    CommandLine DataPath & "7za.exe e " & Chr$(34) & App.Path & "\contemp.bin.7z" & Chr$(34)
    If Dir$(App.Path & "\contemp.bin") <> vbNullString Then
        LoadBytes ByteArray(), App.Path & "\contemp.bin"
        Kill App.Path & "\contemp.bin"
    End If
    Kill App.Path & "\contemp.bin.7z"

End Sub

Public Sub Compression_DeCompress_RLE(ByteArray() As Byte)

Dim X As Long
Dim CntCount As Long
Dim bytLastChar As Byte
Dim ByteCount As Long
Dim InpPos As Long
Dim ZeroCount As Integer
Dim LengthPos As Long

    ZeroCount = 0
    For X = 1 To UBound(ByteArray)
        If ByteArray(X) = 0 Then
            If ZeroCount = ByteArray(0) Then Exit For
            ZeroCount = ZeroCount + 1
        End If
        If ByteArray(X) <> 255 Then
            CntCount = CntCount + 1
        End If
    Next X
    
    OutPos = 0
    CntPos = 1
    LengthPos = X + 1
    InpPos = LengthPos
    
    Do While CntCount > 0
        If ByteArray(InpPos) <> 255 Then
            CntCount = CntCount - 1
        End If
        InpPos = InpPos + 1
    Loop
    ReDim OutStream(UBound(ByteArray) - InpPos + 1)
    ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
    CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
    Do
        If ByteCount = 0 Then
            For X = 1 To UBound(ByteArray) - InpPos + 1
                bytLastChar = Compression_ReadFromArray_Char(ByteArray, InpPos)
                Compression_Add_CharToArray OutStream, OutPos, bytLastChar
            Next X
        Else
            For X = 1 To ByteCount
                bytLastChar = Compression_ReadFromArray_Char(ByteArray, InpPos)
                Compression_Add_CharToArray OutStream, OutPos, bytLastChar
            Next X
            If ByteCount = 255 Then
                Do
                    ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
                    For X = 1 To ByteCount
                        bytLastChar = Compression_ReadFromArray_Char(ByteArray, InpPos)
                        Compression_Add_CharToArray OutStream, OutPos, bytLastChar
                    Next X
                Loop While ByteCount = 255
                ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
            Else
                ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
            End If
            For X = 1 To CntCount
                Compression_Add_CharToArray OutStream, OutPos, bytLastChar
            Next X
            If CntCount = 255 Then
                Do
                    CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
                    For X = 1 To CntCount
                        Compression_Add_CharToArray OutStream, OutPos, bytLastChar
                    Next X
                Loop While CntCount = 255
                CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
            Else
                CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
            End If
        End If
    Loop While InpPos <= UBound(ByteArray)
    ReDim ByteArray(OutPos - 1) As Byte
    CopyMem ByteArray(0), OutStream(0), OutPos

End Sub

Public Sub Compression_DeCompress_RLELoop(ByteArray() As Byte)
Dim X As Integer
Dim TimesRLE As Integer

    TimesRLE = ByteArray(UBound(ByteArray))
    ReDim Preserve ByteArray(UBound(ByteArray) - 1)
    
    For X = 1 To TimesRLE
        Compression_DeCompress_RLE ByteArray
    Next X

End Sub

Private Function Compression_File_Load(FilePath As String) As Byte

Dim FreeNum As Integer

    If Not Len(FilePath) = 0 Then
        FreeNum = FreeFile
        Open FilePath For Binary As #FreeNum
        If LOF(FreeNum) = 0 Then
            Close #FreeNum
            Compression_File_Load = 0
            Exit Function
        End If
        ReDim CompressArray(0 To LOF(FreeNum) - 1)
        Get #FreeNum, , CompressArray()
        Close #FreeNum
    End If
    Compression_File_Load = 1

End Function

Private Sub Compression_File_Save(FilePath As String)

Dim FreeNum As Integer

    If LenB(FilePath) Then
        If LenB(Dir$(FilePath, vbNormal)) Then Kill FilePath
        FreeNum = FreeFile
        Open FilePath For Binary As #FreeNum
        Put #FreeNum, , CompressArray()
        Close #FreeNum
    End If
    
End Sub

Private Sub Compression_MultiDictionary4_Init()

Dim X As Integer

    MaxDict = (2 ^ MaxDictBitPos) - 1
    ReDim Dict(MaxDict)
    For X = 0 To 255
        Dict(0) = Dict(0) & Chr$(X)
    Next X
    For X = 1 To MaxDict
        Dict(X) = vbNullString
    Next X
    AddDict = 1
    UsedDicts = AddDict
    addDictPos = 1
    NowBitLength = 1
    PosPos = 0
    DistPos = 0
    CntPos = 0
    LengthPos = 0
    CntBitCount = 0
    CntByteBuf = 0
    ReadBitPos = 0

End Sub

Private Sub Compression_MultiDictionary4_Search(ByVal Char As String, DictNum As Integer, DictPos As Integer)

    If Len(Char) = 1 Then
        DictNum = 0
        DictPos = Asc(Char)
    Else
        DictNum = 1
        Do While DictNum <= UsedDicts
            DictPos = InStr(Dict(DictNum), Char)
            If DictPos <> 0 Then Exit Sub
            DictNum = DictNum + 1
        Loop
    End If

End Sub

Private Function Compression_ReadFromArray_ASC(WhichArray() As Byte, FromPos As Long) As Integer

    Compression_ReadFromArray_ASC = WhichArray(FromPos)
    FromPos = FromPos + 1

End Function

Private Function Compression_ReadFromArray_Bits(FromArray() As Byte, FromPos As Long, ByVal NumBits As Integer) As Long

Dim X As Integer
Dim Temp As Long

    For X = 1 To NumBits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - ReadBitPos)) > 0))
        ReadBitPos = ReadBitPos + 1
        If ReadBitPos = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While X < NumBits
                    Temp = Temp * 2
                    X = X + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            ReadBitPos = 0
        End If
    Next X
    Compression_ReadFromArray_Bits = Temp

End Function

Private Function Compression_ReadFromArray_Char(FromArray() As Byte, FromPos As Long) As Byte

    Compression_ReadFromArray_Char = FromArray(FromPos)
    FromPos = FromPos + 1

End Function
