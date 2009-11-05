VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This byte buffer was created for being able to automatically convert variables to byte arrays
'and to be able to access them by retrieving them in the order stored. The main intentions of this
'was for creating packets, so you could easily send binary packets (in a byte array) instead of
'having to send strings (which are larger then a byte array, along with slower).
'IE:
'The value "12340" in a string would be 5 bytes, while as an integer, only be 2 bytes. The really
'noticeable values come in at Longs, such as "1293049392" as a string (10 bytes)
'VS as a long (4 bytes).
'This buffer supports strings only of 256 characters long because the length of a string is stored
'as a byte. Though, this can easily be changed by changing Put_Byte / Get_Byte in
'Put_String / Get_String to Put_Integer / Get_Integer, or if you are sending huge strings, you can
'use longs. Try to keep it to the smallest variable you can, though, since it increases the length
'of every string you use.
'This is required to be in there since variable-length strings (in oppose to fixed length strings)
'can not be stored in the same manner as integers, bytes and longs. How would the program know
'if "ABC" was 2 strings "AB" and "C", or "A" and "BC". Instead, it appears "3ABC" so the program
'knows that the next 3 bytes are part of the same string.
'This works very similar to file I/O in binary. You have to retrieve in the same order as you store,
'and you can only retrieve what you have already stored. I did not add in error checking for
'recieving past the stored size just for the slightest speed increase (every ms counts!), along
'with that you shouldn't be even trying to retrieve past what you have stored.
Private PutBytePos As Long
Private GetBytePos As Long
Private ByteBuffer() As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Clear()

'Clear all the values so we can use the buffer from the start again

    PutBytePos = 0
    GetBytePos = 0
    ReDim ByteBuffer(0)

End Sub

Public Function Get_Buffer() As Byte()

'Return the byte array buffer

    Get_Buffer = ByteBuffer

End Function

Public Function Get_Byte() As Byte

'Retrieve a byte from the buffer (1 byte)

    On Error GoTo ErrOut
    If GetBytePos > UBound(ByteBuffer) Then GoTo ErrOut

    CopyMemory Get_Byte, ByteBuffer(GetBytePos), 1
    GetBytePos = GetBytePos + 1

Exit Function

ErrOut:

    If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: Could not acquire BYTE from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer)

End Function

Public Function Get_Integer() As Integer

'Retrieve an integer from the buffer (2 bytes)

    On Error GoTo ErrOut
    If GetBytePos + 1 > UBound(ByteBuffer) Then GoTo ErrOut

    CopyMemory Get_Integer, ByteBuffer(GetBytePos), 2
    GetBytePos = GetBytePos + 2

Exit Function

ErrOut:

    If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: Could not acquire INTEGER from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer)

End Function

Public Function Get_Long() As Long

'Retrieve a long from the buffer (4 bytes)

    On Error GoTo ErrOut
    If GetBytePos + 3 > UBound(ByteBuffer) Then GoTo ErrOut

    CopyMemory Get_Long, ByteBuffer(GetBytePos), 4
    GetBytePos = GetBytePos + 4

Exit Function

ErrOut:

    If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: Could not acquire INTEGER from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer)

End Function

Private Function Get_PutPos() As Long

'Return the put byte position

    Get_PutPos = PutBytePos

End Function

Public Function Get_ReadPos() As Long

'Return the read byte position

    Get_ReadPos = GetBytePos

End Function

Public Function Get_String() As String

Dim tempB() As Byte
Dim ArraySize As Byte

    On Error GoTo ErrOut

    'Retrieve a string from the buffer
    ArraySize = Get_Byte    'Get the size of the string

    'Check for a valid size before sizing the array
    If ArraySize = 0 Then GoTo ErrOut

    'Resize the temp byte array to fit the size of the string
    ReDim tempB(ArraySize - 1)

    'Copy the bytes for the string in the buffer to the temp byte array
    CopyMemory tempB(0), ByteBuffer(GetBytePos), ArraySize

    'Convert the byte array to Unicode
    Get_String = StrConv(tempB, vbUnicode)
    GetBytePos = GetBytePos + ArraySize

Exit Function

ErrOut:

    If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: Could not acquire STRING from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer) & " Size: " & ArraySize

End Function

Public Function Get_StringEX() As String

Dim tempB() As Byte
Dim ArraySize As Integer

    On Error GoTo ErrOut

    'Retrieve a very long string from the buffer
    ArraySize = Get_Integer 'Get the size of the string

    'Check for a valid size before sizing the array
    If ArraySize = 0 Then GoTo ErrOut

    'Resize the temp byte array to fit the size of the string
    ReDim tempB(ArraySize - 1)

    'Copy the bytes for the string in the buffer to the temp byte array
    CopyMemory tempB(0), ByteBuffer(GetBytePos), ArraySize

    'Convert the byte array to Unicode
    Get_StringEX = StrConv(tempB, vbUnicode)
    GetBytePos = GetBytePos + ArraySize

Exit Function

ErrOut:

    If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: Could not acquire STRING from buffer. GetBytePos: " & GetBytePos & " Ubound(ByteBuffer): " & UBound(ByteBuffer) & " Size: " & ArraySize

End Function

Public Sub Put_Byte(ByVal Value As Byte)

'Store a byte (1 byte)

    ReDim Preserve ByteBuffer(0 To PutBytePos + 1)
    CopyMemory ByteBuffer(PutBytePos), Value, 1
    PutBytePos = PutBytePos + 1

End Sub

Public Sub Put_Integer(ByVal Value As Integer)

'Store an integer (2 bytes)

    ReDim Preserve ByteBuffer(0 To PutBytePos + 2)
    CopyMemory ByteBuffer(PutBytePos), Value, 2
    PutBytePos = PutBytePos + 2

End Sub

Public Sub Put_Long(ByVal Value As Long)

'Store a long (4 bytes)

    ReDim Preserve ByteBuffer(0 To PutBytePos + 4)
    CopyMemory ByteBuffer(PutBytePos), Value, 4
    PutBytePos = PutBytePos + 4

End Sub

Public Sub Put_String(ByRef Value As String)
Dim tempB() As Byte

'Store a string

    'Check for invalid value
    If Value = vbNullString Then Exit Sub

    'Convert the string to a byte array
    tempB = StrConv(Value, vbFromUnicode)
    
    'Store a byte-long value that represents the size of the string
    If UBound(tempB) > 254 Then
        If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: The following string was too large to pack into the Put_String buffer! " & Value
        Exit Sub
    End If
    Put_Byte UBound(tempB) + 1

    'Resize the array to fit the string
    ReDim Preserve ByteBuffer(0 To PutBytePos + UBound(tempB) + 1)

    'Store the byte array of the string into the buffer byte array
    CopyMemory ByteBuffer(PutBytePos), tempB(0), UBound(tempB) + 1
    PutBytePos = PutBytePos + UBound(tempB) + 1

End Sub

Public Sub Put_StringEX(ByRef Value As String)

Dim tempB() As Byte
    'Store a very long string
    
    'Check for invalid value
    If Value = vbNullString Then Exit Sub
    
    'Convert the string to a byte array
    tempB = StrConv(Value, vbFromUnicode)

    'Store a byte-long value that represents the size of the string
    If UBound(tempB) > 32760 Then
        If DEBUG_PrintPacketReadErrors Then Debug.Print "Error: The following string was too large to pack into the Put_StringEX buffer! " & Value
        Exit Sub
    End If
    Put_Integer UBound(tempB) + 1
 
    'Resize the array to fit the string
    ReDim Preserve ByteBuffer(0 To PutBytePos + UBound(tempB) + 1)

    'Store the byte array of the string into the buffer byte array
    CopyMemory ByteBuffer(PutBytePos), tempB(0), UBound(tempB) + 1
    PutBytePos = PutBytePos + UBound(tempB) + 1

End Sub

Public Sub Set_Buffer(ByRef Value() As Byte)

    'Clear the values
    Clear

    'Set the byte buffer to the size of the array being sent in
    ReDim ByteBuffer(0 To UBound(Value))

    'Copy to the byte buffer
    CopyMemory ByteBuffer(0), Value(0), UBound(Value) + 1

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:48)  Decl: 27  Code: 236  Total: 263 Lines
':) CommentOnly: 54 (20.5%)  Commented: 2 (0.8%)  Empty: 96 (36.5%)  Max Logic Depth: 1

