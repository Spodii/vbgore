Attribute VB_Name = "Compressions"
Option Explicit

Public Enum CompressMethods
    RLE = 1
    RLE_Loop = 2
    LZW = 3
End Enum
#If False Then
Private RLE, RLE_Loop, LZW
#End If

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
Private Const DictionarySize = 3

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Compression_Add_ASCtoArray(WhichArray() As Byte, ToPos As Long, Text As String)

Dim x As Long

    If ToPos + Len(Text) > UBound(WhichArray) Then ReDim Preserve WhichArray(ToPos + Len(Text) + 500)
    For x = 1 To Len(Text)
        WhichArray(ToPos) = Asc(Mid$(Text, x, 1))
        ToPos = ToPos + 1
    Next

End Sub

Private Sub Compression_Add_BitsToContStream(Number As Long, NumBits As Integer)

Dim x As Long

    For x = NumBits - 1 To 0 Step -1
        CntByteBuf = CntByteBuf * 2 + (-1 * ((Number And CDbl(2 ^ x)) > 0))
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
    Next

End Sub

Private Sub Compression_Add_CharToArray(ToArray() As Byte, ToPos As Long, Char As Byte)

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
    Loop While Char <> ""

End Sub

Public Sub Compression_Compress(SrcFile As String, DestFile As String, Compression As CompressMethods)

Dim Dummy As Boolean

    Compression_File_Load SrcFile
    Select Case Compression
    Case RLE
        Compression_Compress_RLE CompressArray(), Dummy
    Case RLE_Loop
        Compression_Compress_RLELoop CompressArray()
    Case LZW
        Compression_Compress_LZW CompressArray()
    End Select
    Compression_File_Save DestFile

End Sub

Public Sub Compression_Compress_LZW(ByteArray() As Byte)

Dim ByteValue As Byte
Dim TempByte As Long
Dim ExtraBits As Integer
Dim DictStr As String
Dim NewStr As String
Dim CompPos As Long
Dim DictVal As Long
Dim DictPosit As Long
Dim DictPositOld As Long
Dim FilePos As Long
Dim FileLenght As Long
Dim Temp As Long
Dim Dictionary As Integer
Dim DictionaryPos As Integer
Dim OldDict As Integer
Dim OldPos As Integer
Dim TempDist As Integer
Dim DistCount As Integer
Dim x As Integer

    Temp = (CLng(1024) * DictionarySize) / 256 - 1
    For x = 0 To 16
        If 2 ^ x > Temp Then
            MaxDictBitPos = x
            Exit For
        End If
    Next
    Call Compression_MultiDictionary4_Init
    FileLenght = UBound(ByteArray)
    ReDim PosStream(FileLenght / 3)
    ReDim DistStream(FileLenght / 3)
    ReDim LengthStream(FileLenght / 3)
    ReDim ContStream(FileLenght / 15)
    FilePos = 0
    DictStr = vbNullString
    ExtraBits = 0
    TempDist = 0
    DistCount = 0
    Do Until FilePos > FileLenght
        ByteValue = ByteArray(FilePos)
        FilePos = FilePos + 1
        NewStr = DictStr & Chr$(ByteValue)
        Call Compression_MultiDictionary4_Search(NewStr, Dictionary, DictionaryPos)
        If Dictionary <> UsedDicts + 1 Then
            DictStr = NewStr
            OldDict = Dictionary
            OldPos = DictionaryPos
        Else
            Do While OldDict > (2 ^ NowBitLength) - 1
                Call Compression_Add_BitsToContStream(1, NowBitLength)
                Call Compression_Add_ASCtoArray(DistStream, DistPos, Chr$(255))
                NowBitLength = NowBitLength + 1
            Loop
            Call Compression_Add_BitsToContStream(CLng(OldDict), NowBitLength)
            If OldDict > 0 Then
                Call Compression_Add_ASCtoArray(DistStream, DistPos, Chr$(OldPos))
                Call Compression_Add_ASCtoArray(LengthStream, LengthPos, Chr$(Len(DictStr) - 2))
                OldDict = 0
            Else
                Call Compression_Add_ASCtoArray(PosStream, PosPos, Chr$(OldPos))
            End If
            Call Compression_Add_CharToDict4(DictStr)
            OldPos = ByteValue
            DictStr = Chr$(ByteValue)
        End If
    Loop
    Do While OldDict > (2 ^ NowBitLength) - 1
        Call Compression_Add_BitsToContStream(1, NowBitLength)
        Call Compression_Add_ASCtoArray(DistStream, DistPos, Chr$(255))
        NowBitLength = NowBitLength + 1
    Loop
    Call Compression_Add_BitsToContStream(CLng(OldDict), NowBitLength)
    If OldDict > 0 Then
        Call Compression_Add_ASCtoArray(DistStream, DistPos, Chr$(OldPos))
        Call Compression_Add_ASCtoArray(LengthStream, LengthPos, Chr$(Len(DictStr) - 2))
    Else
        Call Compression_Add_ASCtoArray(PosStream, PosPos, Chr$(OldPos))
    End If
    Call Compression_Add_BitsToContStream(1, NowBitLength)
    Call Compression_Add_ASCtoArray(DistStream, DistPos, Chr$(0))
    Do While CntBitCount > 0
        Call Compression_Add_BitsToContStream(0, 1)
    Loop
    ReDim Preserve PosStream(PosPos - 1)
    ReDim Preserve ContStream(CntPos - 1)
    ReDim Preserve LengthStream(LengthPos - 1)
    ReDim Preserve DistStream(DistPos - 1)
    ReDim ByteArray(UBound(ContStream) + UBound(LengthStream) + UBound(DistStream) + UBound(PosStream) + 4 + 9)
    ByteArray(0) = MaxDictBitPos
    ByteArray(1) = Int(UBound(ContStream) / &H10000) And &HFF
    ByteArray(2) = Int(UBound(ContStream) / &H100) And &HFF
    ByteArray(3) = UBound(ContStream) And &HFF
    ByteArray(4) = Int(UBound(LengthStream) / &H10000) And &HFF
    ByteArray(5) = Int(UBound(LengthStream) / &H100) And &HFF
    ByteArray(6) = UBound(LengthStream) And &HFF
    ByteArray(7) = Int(UBound(DistStream) / &H10000) And &HFF
    ByteArray(8) = Int(UBound(DistStream) / &H100) And &HFF
    ByteArray(9) = UBound(DistStream) And &HFF
    Call CopyMem(ByteArray(10), ContStream(0), UBound(ContStream) + 1)
    Call CopyMem(ByteArray(10 + UBound(ContStream) + 1), LengthStream(0), UBound(LengthStream) + 1)
    Call CopyMem(ByteArray(10 + UBound(ContStream) + UBound(LengthStream) + 2), DistStream(0), UBound(DistStream) + 1)
    Call CopyMem(ByteArray(10 + UBound(ContStream) + UBound(LengthStream) + UBound(DistStream) + 3), PosStream(0), UBound(PosStream) + 1)

End Sub

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Public Sub Compression_Compress_RLE(ByteArray() As Byte, IsCompressed As Boolean)

Dim x As Long
Dim Y As Long
Dim ByteCount As Long
Dim LastAsc As Integer
Dim TelSame As Long
Dim Times255 As Integer
Dim Same255 As Integer
Dim IsRun As Boolean
Dim ZeroCount As Integer
Dim LengthPos As Long
Dim NoLength As Boolean

    ReDim ContStream(200)
    ReDim LengthStream(200)
    ReDim OutStream(500)
    IsCompressed = False
    ByteCount = 0
    LastAsc = 0
    CntPos = 1
    OutPos = 0
    LengthPos = 0
    TelSame = 0
    ZeroCount = 0
    For x = 0 To UBound(ByteArray)
        If LastAsc = ByteArray(x) And x <> 0 Then IsRun = True Else IsRun = False
        If IsRun = False Then
            If TelSame = 1 Then
                TelSame = 0
                Call Compression_Add_CharToArray(OutStream, OutPos, CByte(LastAsc))
                ByteCount = ByteCount + 1
            ElseIf TelSame > 1 Then
                For Y = 1 To Int(ByteCount / 255)
                    Call Compression_Add_CharToArray(ContStream, CntPos, 255)
                Next
                ByteCount = ByteCount Mod 255
                If ByteCount = 0 Then ZeroCount = ZeroCount + 1
                Call Compression_Add_CharToArray(ContStream, CntPos, CByte(ByteCount))
                ByteCount = 0
                For Y = 1 To Int(TelSame / 255)
                    Call Compression_Add_CharToArray(LengthStream, LengthPos, 255)
                Next
                TelSame = TelSame Mod 255
                Call Compression_Add_CharToArray(LengthStream, LengthPos, CByte(TelSame))
                TelSame = 0
            End If
            Call Compression_Add_CharToArray(OutStream, OutPos, ByteArray(x))
            ByteCount = ByteCount + 1
        Else
            TelSame = TelSame + 1
        End If
        LastAsc = ByteArray(x)
    Next
    If IsRun = True Then
        If TelSame < 2 Then
            Call Compression_Add_CharToArray(OutStream, OutPos, CByte(LastAsc))
        Else
            For Y = 1 To Int(ByteCount / 255)
                Call Compression_Add_CharToArray(ContStream, CntPos, 255)
            Next
            ByteCount = ByteCount Mod 255
            Call Compression_Add_CharToArray(ContStream, CntPos, CByte(ByteCount))
            For Y = 1 To Int(TelSame / 255)
                Call Compression_Add_CharToArray(LengthStream, LengthPos, 255)
            Next
            TelSame = TelSame Mod 255
            Call Compression_Add_CharToArray(LengthStream, LengthPos, CByte(TelSame))
        End If
    End If
    ContStream(0) = CByte(ZeroCount)
    If CntPos > 1 Then IsCompressed = True
    Call Compression_Add_CharToArray(ContStream, CntPos, 0)  'No Run Till EOF
    ReDim Preserve ContStream(CntPos - 1)
    If LengthPos > 0 Then
        ReDim Preserve LengthStream(LengthPos - 1)
        NoLength = False
    Else
        NoLength = True
    End If
    ReDim Preserve OutStream(OutPos - 1)
    CntPos = UBound(ContStream) + 1
    LengthPos = 0
    If NoLength = False Then LengthPos = UBound(LengthStream) + 1
    OutPos = UBound(OutStream) + 1
    ReDim ByteArray(CntPos + LengthPos + OutPos - 1)
    Call CopyMem(ByteArray(0), ContStream(0), CntPos)
    If LengthPos > 0 Then
        Call CopyMem(ByteArray(CntPos), LengthStream(0), LengthPos)
    End If
    Call CopyMem(ByteArray(CntPos + LengthPos), OutStream(0), OutPos)

End Sub

Public Sub Compression_Compress_RLELoop(ByteArray() As Byte)

Dim NuSize As Long
Dim TimesRLE As Integer
Dim Filenr As Integer
Dim IsCompressed As Boolean

    Do
        NuSize = UBound(ByteArray)
        Call Compression_Compress_RLE(ByteArray, IsCompressed)
        TimesRLE = TimesRLE + 1
    Loop While IsCompressed = True
    ReDim Preserve ByteArray(UBound(ByteArray) + 1)
    ByteArray(UBound(ByteArray)) = TimesRLE

End Sub

Public Sub Compression_DeCompress(SrcFile As String, DestFile As String, Compression As CompressMethods)

Dim Dummy As Boolean

    Compression_File_Load SrcFile
    Select Case Compression
    Case RLE
        Compression_DeCompress_RLE CompressArray()
    Case RLE_Loop
        Compression_DeCompress_RLELoop CompressArray()
    Case LZW
        Compression_DeCompress_LZW CompressArray()
    End Select
    Compression_File_Save DestFile

End Sub

Public Sub Compression_DeCompress_LZW(ByteArray() As Byte)

Dim DictVal As Long
Dim TempByte As Long
Dim OldKarValue As Long
Dim DeComPByte() As Byte
Dim DeCompPos As Long
Dim FilePos As Long
Dim FileLenght As Long
Dim InpPos As Long
Dim Dictionary As Integer
Dim DictPos As Integer
Dim DictLen As Integer
Dim DistencePos As Long
Dim Temp As Long
Dim TempDist As Integer
Dim DistCount As Integer

    MaxDictBitPos = ByteArray(0)
    Call Compression_MultiDictionary4_Init
    CntPos = 10
    Temp = (CLng(ByteArray(1)) * 256) + ByteArray(2)
    Temp = CLng(Temp) * 256 + ByteArray(3)
    LengthPos = CntPos + Temp + 1
    Temp = (CLng(ByteArray(4)) * 256) + ByteArray(5)
    Temp = CLng(Temp) * 256 + ByteArray(6)
    DistencePos = LengthPos + Temp + 1
    Temp = (CLng(ByteArray(7)) * 256) + ByteArray(8)
    Temp = CLng(Temp) * 256 + ByteArray(9)
    PosPos = DistencePos + Temp + 1
    ReDim DistStream(500)
    DistCount = 0
    Do
        Dictionary = Compression_ReadFromArray_Bits(ByteArray, CntPos, NowBitLength)
        If Dictionary = 0 Then
            DictPos = Compression_ReadFromArray_ASC(ByteArray, PosPos)
            Call Compression_Add_ASCtoArray(DistStream, DistPos, Chr$(DictPos))
            Call Compression_Add_CharToDict4(Chr$(DictPos))
        Else
            DictPos = Compression_ReadFromArray_ASC(ByteArray, DistencePos)
            If DictPos = 0 Then Exit Do
            If DictPos = 255 And Dictionary = 1 Then
                NowBitLength = NowBitLength + 1
            Else
                DictLen = Compression_ReadFromArray_ASC(ByteArray, LengthPos) + 2
                Call Compression_Add_ASCtoArray(DistStream, DistPos, Mid$(Dict(Dictionary), DictPos, DictLen))
                Call Compression_Add_CharToDict4(Mid$(Dict(Dictionary), DictPos, DictLen))
            End If
        End If
    Loop
    DistPos = DistPos - 1
    ReDim ByteArray(DistPos)
    Call CopyMem(ByteArray(0), DistStream(0), DistPos + 1)

End Sub

Public Sub Compression_DeCompress_RLE(ByteArray() As Byte)

Dim x As Long
Dim CntCount As Long
Dim LastChar As Byte
Dim ByteCount As Long
Dim InpPos As Long
Dim ZeroCount As Integer
Dim LengthPos As Long

    ZeroCount = 0
    For x = 1 To UBound(ByteArray)
        If ByteArray(x) = 0 Then
            If ZeroCount = ByteArray(0) Then Exit For
            ZeroCount = ZeroCount + 1
        End If
        If ByteArray(x) <> 255 Then
            CntCount = CntCount + 1
        End If
    Next
    OutPos = 0
    CntPos = 1
    '    LengthPos = 0
    LengthPos = x + 1
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
            For x = 1 To UBound(ByteArray) - InpPos + 1
                LastChar = Compression_ReadFromArray_Char(ByteArray, InpPos)
                Call Compression_Add_CharToArray(OutStream, OutPos, LastChar)
            Next
        Else
            For x = 1 To ByteCount
                LastChar = Compression_ReadFromArray_Char(ByteArray, InpPos)
                Call Compression_Add_CharToArray(OutStream, OutPos, LastChar)
            Next
            If ByteCount = 255 Then
                Do
                    ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
                    For x = 1 To ByteCount
                        LastChar = Compression_ReadFromArray_Char(ByteArray, InpPos)
                        Call Compression_Add_CharToArray(OutStream, OutPos, LastChar)
                    Next
                Loop While ByteCount = 255
                ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
            Else
                ByteCount = Compression_ReadFromArray_Char(ByteArray, CntPos)
            End If
            For x = 1 To CntCount
                Call Compression_Add_CharToArray(OutStream, OutPos, LastChar)
            Next
            If CntCount = 255 Then
                Do
                    CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
                    For x = 1 To CntCount
                        Call Compression_Add_CharToArray(OutStream, OutPos, LastChar)
                    Next
                Loop While CntCount = 255
                CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
            Else
                CntCount = Compression_ReadFromArray_Char(ByteArray, LengthPos)
            End If
        End If
    Loop While InpPos <= UBound(ByteArray)
    ReDim ByteArray(OutPos - 1)
    Call CopyMem(ByteArray(0), OutStream(0), OutPos)

End Sub

Public Sub Compression_DeCompress_RLELoop(ByteArray() As Byte)

Dim x As Integer
Dim TimesRLE As Integer

    TimesRLE = ByteArray(UBound(ByteArray))
    ReDim Preserve ByteArray(UBound(ByteArray) - 1)
    For x = 1 To TimesRLE
        Call Compression_DeCompress_RLE(ByteArray)
    Next

End Sub

Private Sub Compression_File_Load(FilePath As String)

Dim FreeNum As Integer

    If LenB(FilePath) = 0 Then Exit Sub
    FreeNum = FreeFile
    Open FilePath For Binary As #FreeNum
    ReDim CompressArray(0 To LOF(FreeNum) - 1)
    Get #FreeNum, , CompressArray()
    Close #FreeNum

End Sub

Private Sub Compression_File_Save(FilePath As String)

Dim FreeNum As Integer

    If LenB(FilePath) = 0 Then Exit Sub
    FreeNum = FreeFile
    Open FilePath For Binary As #FreeNum
    Put #FreeNum, , CompressArray()
    Close #FreeNum

End Sub

Private Sub Compression_MultiDictionary4_Init()

Dim x As Integer
Dim Y As Integer

    MaxDict = (2 ^ MaxDictBitPos) - 1
    ReDim Dict(MaxDict)
    For x = 0 To 255
        Dict(0) = Dict(0) & Chr$(x)
    Next
    For x = 1 To MaxDict
        Dict(x) = vbNullString
    Next
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

Private Sub Compression_MultiDictionary4_Search(Char As String, DictNum As Integer, DictPos As Integer)

    If Len(Char) = 1 Then
        DictNum = 0
        DictPos = Asc(Char)
        Exit Sub
    Else
        DictNum = 1
        Do While DictNum <= UsedDicts
            DictPos = InStr(Dict(DictNum), Char)
            If DictPos <> 0 Then
                Exit Sub
            End If
            DictNum = DictNum + 1
        Loop
    End If

End Sub

Private Function Compression_ReadFromArray_ASC(WhichArray() As Byte, FromPos As Long) As Integer

    Compression_ReadFromArray_ASC = WhichArray(FromPos)
    FromPos = FromPos + 1

End Function

Private Function Compression_ReadFromArray_Bits(FromArray() As Byte, FromPos As Long, NumBits As Integer) As Long

Dim x As Integer
Dim Temp As Long

    For x = 1 To NumBits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - ReadBitPos)) > 0))
        ReadBitPos = ReadBitPos + 1
        If ReadBitPos = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While x < NumBits
                    Temp = Temp * 2
                    x = x + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            ReadBitPos = 0
        End If
    Next
    Compression_ReadFromArray_Bits = Temp

End Function

Private Function Compression_ReadFromArray_Char(FromArray() As Byte, FromPos As Long) As Byte

    Compression_ReadFromArray_Char = FromArray(FromPos)
    FromPos = FromPos + 1

End Function

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:49)  Decl: 36  Code: 579  Total: 615 Lines
':) CommentOnly: 3 (0.5%)  Commented: 1 (0.2%)  Empty: 76 (12.4%)  Max Logic Depth: 6
