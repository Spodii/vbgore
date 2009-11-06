Attribute VB_Name = "TradeTables"
Option Explicit

Public Sub TradeTable_RemoveItem(ByVal UserIndex As Integer, ByVal TableSlot As Byte)

'*****************************************************************
'User wants to remove an item from their trade table
'*****************************************************************
Dim TradeTableIndex As Byte
Dim UserTableIndex As Byte

    'Make sure the user has a trade table
    TradeTableIndex = UserList(UserIndex).flags.TradeTable
    If TradeTableIndex <= 0 Then Exit Sub 'Invalid table index
    
    'Get the user's index in the table
    UserTableIndex = TradeTable_GetUserTableIndex(TradeTableIndex, UserIndex)
    If UserTableIndex = 0 Then Exit Sub 'Error!
    
    'Remove the item
    If UserTableIndex = 1 Then
        If TradeTable(TradeTableIndex).Objs1(TableSlot).UserInvSlot > 0 Then
            TradeTable(TradeTableIndex).Objs1(TableSlot).Amount = 0
            TradeTable(TradeTableIndex).Objs1(TableSlot).UserInvSlot = 0
            TradeTable_SendSlotPacket TradeTableIndex, TableSlot, UserTableIndex
        End If
    Else
        If TradeTable(TradeTableIndex).Objs2(TableSlot).UserInvSlot > 0 Then
            TradeTable(TradeTableIndex).Objs2(TableSlot).Amount = 0
            TradeTable(TradeTableIndex).Objs2(TableSlot).UserInvSlot = 0
            TradeTable_SendSlotPacket TradeTableIndex, TableSlot, UserTableIndex
        End If
    End If

End Sub

Public Sub TradeTable_UpdateSlot(ByVal UserIndex As Integer, ByVal InvSlot As Byte, ByVal Amount As Long)

'*****************************************************************
'Update a trade table's slot (either adds, changes or remove an item or gold)
'*****************************************************************
Dim TradeTableIndex As Byte
Dim UserTableIndex As Byte
Dim PutTableSlot As Byte
Dim i As Long
    
    'Make sure the user has a trade table
    TradeTableIndex = UserList(UserIndex).flags.TradeTable
    If TradeTableIndex <= 0 Then Exit Sub 'Invalid table index
    
    'Get the user's index in the table
    UserTableIndex = TradeTable_GetUserTableIndex(TradeTableIndex, UserIndex)
    If UserTableIndex = 0 Then Exit Sub 'Error!
    
    'If the invslot = 0, then we are updating the gold
    If InvSlot = 0 Then
    
        'Make sure the user has enough gold
        If UserList(UserIndex).Stats.BaseStat(SID.Gold) < Amount Then Exit Sub
    
        'Update the table
        If UserTableIndex = 1 Then
            TradeTable(TradeTableIndex).Gold1 = Amount
            TradeTable_SendSlotPacket TradeTableIndex, 0, 1
        Else
            TradeTable(TradeTableIndex).Gold2 = Amount
            TradeTable_SendSlotPacket TradeTableIndex, 0, 2
        End If

    'If the invslot > 0 then we are updating the objects
    Else
    
        'Make sure the user has the object they entered, and enough of it
        If InvSlot > MAX_INVENTORY_SLOTS Then Exit Sub
        If UserList(UserIndex).Object(InvSlot).ObjIndex = 0 Then Exit Sub
        If UserList(UserIndex).Object(InvSlot).Amount < Amount Then Exit Sub
    
        If Amount <= 0 Then
        
            'Since the amount is equal to 0, we're removing the item
            For i = 1 To 9
                If UserTableIndex = 1 Then
                    If TradeTable(TradeTableIndex).Objs1(i).UserInvSlot = InvSlot Then
                        TradeTable(TradeTableIndex).Objs1(i).Amount = 0
                        TradeTable(TradeTableIndex).Objs1(i).UserInvSlot = 0
                        TradeTable_SendSlotPacket TradeTableIndex, i, 1
                        Exit Sub
                    End If
                Else
                    If TradeTable(TradeTableIndex).Objs2(i).UserInvSlot = InvSlot Then
                        TradeTable(TradeTableIndex).Objs1(i).Amount = 0
                        TradeTable(TradeTableIndex).Objs1(i).UserInvSlot = 0
                        TradeTable_SendSlotPacket TradeTableIndex, i, 2
                        Exit Sub
                    End If
                End If
            Next i
            Exit Sub
        
        Else
            
            'Make sure the user hasn't already put the item in the trade table
            For i = 1 To 9
                If UserTableIndex = 1 Then
                    If TradeTable(TradeTableIndex).Objs1(i).UserInvSlot = InvSlot Then Exit Sub
                Else
                    If TradeTable(TradeTableIndex).Objs2(i).UserInvSlot = InvSlot Then Exit Sub
                End If
            Next i
        
        End If
        
        'Find the next free slot
        PutTableSlot = 0
        If UserTableIndex = 1 Then
            Do
                PutTableSlot = PutTableSlot + 1
                If PutTableSlot > 9 Then Exit Sub   'No more room :(
            Loop While TradeTable(TradeTableIndex).Objs1(PutTableSlot).UserInvSlot > 0
        Else
            Do
                PutTableSlot = PutTableSlot + 1
                If PutTableSlot > 9 Then Exit Sub   'No more room :(
            Loop While TradeTable(TradeTableIndex).Objs2(PutTableSlot).UserInvSlot > 0
        End If
        
        'If we made it this far, we have an object and a slot to put it in, so place it there!
        If UserTableIndex = 1 Then
            TradeTable(TradeTableIndex).Objs1(PutTableSlot).UserInvSlot = InvSlot
            TradeTable(TradeTableIndex).Objs1(PutTableSlot).Amount = Amount
            TradeTable_SendSlotPacket TradeTableIndex, PutTableSlot, 1
        Else
            TradeTable(TradeTableIndex).Objs2(PutTableSlot).UserInvSlot = InvSlot
            TradeTable(TradeTableIndex).Objs2(PutTableSlot).Amount = Amount
            TradeTable_SendSlotPacket TradeTableIndex, PutTableSlot, 2
        End If
        
    End If

End Sub

Private Sub TradeTable_SendSlotPacket(ByVal TradeTableIndex As Byte, ByVal TableSlot As Byte, ByVal UserTableIndex As Byte)

'*****************************************************************
'Updates the clients of a trade table with the changes applied to a slot
'*****************************************************************
Dim Amount As Long
Dim ObjIndex As Integer
Dim GrhIndex As Long

    'If the tableslot > 0, then we need the object information
    If TableSlot > 0 Then
        If UserTableIndex = 1 Then
            ObjIndex = UserList(TradeTable(TradeTableIndex).User1).Object(TradeTable(TradeTableIndex).Objs1(TableSlot).UserInvSlot).ObjIndex
            If ObjIndex > 0 Then GrhIndex = ObjData.GrhIndex(ObjIndex) Else GrhIndex = 0
            Amount = TradeTable(TradeTableIndex).Objs1(TableSlot).Amount
        Else
            ObjIndex = UserList(TradeTable(TradeTableIndex).User2).Object(TradeTable(TradeTableIndex).Objs2(TableSlot).UserInvSlot).ObjIndex
            If ObjIndex > 0 Then GrhIndex = ObjData.GrhIndex(ObjIndex) Else GrhIndex = 0
            Amount = TradeTable(TradeTableIndex).Objs2(TableSlot).Amount
        End If
    
    'If the tableslot = 0, we're updating gold
    Else
        If UserTableIndex = 1 Then
            Amount = TradeTable(TradeTableIndex).Gold1
        Else
            Amount = TradeTable(TradeTableIndex).Gold2
        End If
    End If
    
    'If we're updating gold, we don't need to add the object index
    If TableSlot > 0 Then ConBuf.PreAllocate 11 Else ConBuf.PreAllocate 7
    
    ConBuf.Put_Byte DataCode.User_Trade_UpdateTrade
    ConBuf.Put_Byte UserTableIndex
    ConBuf.Put_Byte TableSlot
    ConBuf.Put_Long Amount
    
    'Put the object index only for an object
    If TableSlot > 0 Then ConBuf.Put_Long GrhIndex
    
    'Send the data to both the clients
    Data_Send ToIndex, TradeTable(TradeTableIndex).User1, ConBuf.Get_Buffer
    Data_Send ToIndex, TradeTable(TradeTableIndex).User2, ConBuf.Get_Buffer

End Sub

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

Private Function TradeTable_GetUserTableIndex(ByVal TradeTableIndex As Byte, ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Returns the user's index for the table, either 1 or 2 (or 0 for error)
'*****************************************************************

    'Find out what index the user is in the trade table, either 1 or 2
    If TradeTable(TradeTableIndex).User1 = UserIndex Then
        TradeTable_GetUserTableIndex = 1
    ElseIf TradeTable(TradeTableIndex).User2 = UserIndex Then
        TradeTable_GetUserTableIndex = 2
    Else
        'Oh crap! This user doesn't belong in this table!
        TradeTable_GetUserTableIndex = 0
    End If

End Function
