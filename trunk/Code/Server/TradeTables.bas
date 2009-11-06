Attribute VB_Name = "TradeTables"
Option Explicit


Public Function TradeTable_NextOpen() As Byte

'*****************************************************************
'Finds the next open trade table
'*****************************************************************
Dim i As Long

    'Check for an unused table
    For i = 1 To NumTradeTables
    
        'Both users will be closed state, or neither, so check one
        If TradeTable(i).User1State = TRADESTATE_CLOSED Then
        
            'Table is free
            TradeTable_NextOpen = i
            Exit Function
            
        End If
        
    Next i
    
    'No free tables, make a new one if possible
    If NumTradeTables < 255 Then
        
        'Create the new trade table slot
        NumTradeTables = NumTradeTables + 1
        ReDim Preserve TradeTable(1 To NumTradeTables)
        TradeTable_NextOpen = NumTradeTables
        
    End If

End Function

Public Sub TradeTable_Create(ByVal UserIndex1 As Integer, ByVal UserIndex2 As Integer)

'*****************************************************************
'Creates a trade table for two users
'*****************************************************************
Dim TableIndex As Byte

    'Get the table index
    TableIndex = TradeTable_NextOpen
    If TableIndex = 0 Then Exit Sub
    
    'Clear the table just in case
    ZeroMemory TradeTable(TableIndex), Len(TradeTable(TableIndex))
    
    'Assign the values
    TradeTable(TableIndex).User1 = UserIndex1
    TradeTable(TableIndex).User2 = UserIndex2
    TradeTable(TableIndex).User1State = TRADESTATE_TRADING
    TradeTable(TableIndex).User2State = TRADESTATE_TRADING
    UserList(UserIndex1).flags.TradeTable = TableIndex
    UserList(UserIndex2).flags.TradeTable = TableIndex
    
    'Send the packet to the users to show the tables
    ConBuf.PreAllocate 2
    ConBuf.Put_Byte DataCode.User_Trade_Trade
    ConBuf.Put_String UserList(UserIndex2).Name
    ConBuf.Put_String UserList(UserIndex1).Name
    ConBuf.Put_Byte 2
    Data_Send ToIndex, UserIndex2, ConBuf.Get_Buffer
    
    ConBuf.PreAllocate 2
    ConBuf.Put_Byte DataCode.User_Trade_Trade
    ConBuf.Put_String UserList(UserIndex1).Name
    ConBuf.Put_String UserList(UserIndex2).Name
    ConBuf.Put_Byte 1
    Data_Send ToIndex, UserIndex1, ConBuf.Get_Buffer

End Sub



