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
Private Const DictionarySize As Long = 3

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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
    Case LZW
        Compression_Compress_LZW CompressArray()
    End Select
    Compression_File_Save DestFile

End Sub

Public Sub Compression_Compress_LZW(ByteArray() As Byte)

Dim ByteValue As Byte
Dim DictStr As String
Dim NewStr As String
Dim FilePos As Long
Dim FileLenght As Long
Dim Temp As Long
Dim Dictionary As Integer
Dim DictionaryPos As Integer
Dim OldDict As Integer
Dim OldPos As Integer
Dim X As Integer

    Temp = (CLng(1024) * DictionarySize) / 256 - 1
    
    For X = 0 To 16
        If 2 ^ X > Temp Then
            MaxDictBitPos = X
            Exit For
        End If
    Next X
    
    Compression_MultiDictionary4_Init
    FileLenght = UBound(ByteArray)
    ReDim PosStream(FileLenght / 3)
    ReDim DistStream(FileLenght / 3)
    ReDim LengthStream(FileLenght / 3)
    ReDim ContStream(FileLenght / 15)
    DictStr = vbNullString
    
    Do Until FilePos > FileLenght
        ByteValue = ByteArray(FilePos)
        FilePos = FilePos + 1
        NewStr = DictStr & Chr$(ByteValue)
        Compression_MultiDictionary4_Search NewStr, Dictionary, DictionaryPos
        If Dictionary <> UsedDicts + 1 Then
            DictStr = NewStr
            OldDict = Dictionary
            OldPos = DictionaryPos
        Else
            Do While OldDict > (2 ^ NowBitLength) - 1
                Compression_Add_BitsToContStream 1, NowBitLength
                Compression_Add_ASCtoArray DistStream, DistPos, Chr$(255)
                NowBitLength = NowBitLength + 1
            Loop
            Call Compression_Add_BitsToContStream(CLng(OldDict), NowBitLength)
            If OldDict > 0 Then
                Compression_Add_ASCtoArray DistStream, DistPos, Chr$(OldPos)
                Compression_Add_ASCtoArray LengthStream, LengthPos, Chr$(Len(DictStr) - 2)
                OldDict = 0
            Else
                Compression_Add_ASCtoArray PosStream, PosPos, Chr$(OldPos)
            End If
            Compression_Add_CharToDict4 DictStr
            OldPos = ByteValue
            DictStr = Chr$(ByteValue)
        End If
    Loop
    Do While OldDict > (2 ^ NowBitLength) - 1
        Compression_Add_BitsToContStream 1, NowBitLength
        Compression_Add_ASCtoArray DistStream, DistPos, Chr$(255)
        NowBitLength = NowBitLength + 1
    Loop
    Call Compression_Add_BitsToContStream(CLng(OldDict), NowBitLength)
    If OldDict > 0 Then
        Compression_Add_ASCtoArray DistStream, DistPos, Chr$(OldPos)
        Compression_Add_ASCtoArray LengthStream, LengthPos, Chr$(Len(DictStr) - 2)
    Else
        Compression_Add_ASCtoArray PosStream, PosPos, Chr$(OldPos)
    End If
    Compression_Add_BitsToContStream 1, NowBitLength
    Compression_Add_ASCtoArray DistStream, DistPos, vbNullChar
    Do While CntBitCount > 0
        Compression_Add_BitsToContStream 0, 1
    Loop
    ReDim Preserve PosStream(PosPos - 1)
    ReDim Preserve ContStream(CntPos - 1)
    ReDim Preserve LengthStream(LengthPos - 1)
    ReDim Preserve DistStream(DistPos - 1)
    ReDim ByteArray(UBound(ContStream) + UBound(LengthStream) + UBound(DistStream) + UBound(PosStream) + 13)
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
    CopyMem ByteArray(10), ContStream(0), UBound(ContStream) + 1
    CopyMem ByteArray(10 + UBound(ContStream) + 1), LengthStream(0), UBound(LengthStream) + 1
    CopyMem ByteArray(10 + UBound(ContStream) + UBound(LengthStream) + 2), DistStream(0), UBound(DistStream) + 1
    CopyMem ByteArray(10 + UBound(ContStream) + UBound(LengthStream) + UBound(DistStream) + 3), PosStream(0), UBound(PosStream) + 1

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
    Case LZW
        Compression_DeCompress_LZW CompressArray()
    End Select
    Compression_File_Save DestFile

End Sub

Public Sub Compression_DeCompress_LZW(ByteArray() As Byte)
Dim Dictionary As Integer
Dim DictPos As Integer
Dim DictLen As Integer
Dim DistencePos As Long
Dim Temp As Long

    MaxDictBitPos = ByteArray(0)
    Compression_MultiDictionary4_Init
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
    
    Do
        Dictionary = Compression_ReadFromArray_Bits(ByteArray, CntPos, NowBitLength)
        If Dictionary = 0 Then
            DictPos = Compression_ReadFromArray_ASC(ByteArray, PosPos)
            Compression_Add_ASCtoArray DistStream, DistPos, Chr$(DictPos)
            Compression_Add_CharToDict4 Chr$(DictPos)
        Else
            DictPos = Compression_ReadFromArray_ASC(ByteArray, DistencePos)
            If DictPos = 0 Then Exit Do
            If DictPos = 255 And Dictionary = 1 Then
                NowBitLength = NowBitLength + 1
            Else
                DictLen = Compression_ReadFromArray_ASC(ByteArray, LengthPos) + 2
                Compression_Add_ASCtoArray DistStream, DistPos, Mid$(Dict(Dictionary), DictPos, DictLen)
                Compression_Add_CharToDict4 Mid$(Dict(Dictionary), DictPos, DictLen)
            End If
        End If
    Loop
    DistPos = DistPos - 1
    ReDim ByteArray(DistPos) As Byte
    CopyMem ByteArray(0), DistStream(0), DistPos + 1

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
