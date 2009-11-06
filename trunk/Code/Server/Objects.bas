Attribute VB_Name = "Objects"
Option Explicit


Public Sub Obj_CleanMapTile(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

'*****************************************************************
'Removes all the unused obj slots on a map tile
'Make sure you call this every time you remove an object from a tile!
'*****************************************************************
Dim NumObjs As Byte
Dim i As Long
Dim j As Long

    Log "Call Obj_CleanMapTile(" & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the map is in memory
    If MapInfo(Map).DataLoaded = 0 Then Exit Sub

    'Make sure we wern't given an empty map tile
    If MapInfo(Map).ObjTile(X, Y).NumObjs = 0 Then
        Log "Obj_CleanMapTile: NumObjs = 0 - aborting", CodeTracker '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check through all the object slots
    For i = 1 To MapInfo(Map).ObjTile(X, Y).NumObjs
        If MapInfo(Map).ObjTile(X, Y).ObjInfo(i).ObjIndex > 0 Then
            If MapInfo(Map).ObjTile(X, Y).ObjInfo(i).Amount > 0 Then
                
                'Object found, so raise the count
                NumObjs = NumObjs + 1
                
                'Move down in the array if possible
                If i > 1 Then   'We can't sort any lower then 1, so don't even try it
                    
                    'Loop through all the previous object slots
                    For j = 1 To (i - 1)    '(i - 1) is since we don't to check the slot it is already on!
                    
                        'If the object slot is unused then
                        If MapInfo(Map).ObjTile(X, Y).ObjInfo(j).ObjIndex = 0 Or MapInfo(Map).ObjTile(X, Y).ObjInfo(j).Amount = 0 Then
                
                            'Scoot the item's keester down to that slot (swap the used into the unused)
                            MapInfo(Map).ObjTile(X, Y).ObjInfo(j) = MapInfo(Map).ObjTile(X, Y).ObjInfo(i)
                            MapInfo(Map).ObjTile(X, Y).ObjLife(j) = MapInfo(Map).ObjTile(X, Y).ObjLife(i)
                            
                            'Clear the old object
                            ZeroMemory MapInfo(Map).ObjTile(X, Y).ObjInfo(i), Len(MapInfo(Map).ObjTile(X, Y).ObjInfo(i))
                            MapInfo(Map).ObjTile(X, Y).ObjLife(i) = 0
                        
                        End If
                        
                    Next j
                    
                End If
                
            End If
        End If
    Next i
    
    'Once all that code above has gone through, NumObjs should have the number of valid objects
    ' and the first object slots should be used (unused at the end), so if redim the array by
    ' the NumObjs value, all we will cut off is the unused slots! :)
    If NumObjs > 0 Then
        Log "Obj_CleanMapTile: Resizing ObjInfo() array (1 To " & NumObjs & ")", CodeTracker '//\\LOGLINE//\\
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjInfo(1 To NumObjs)
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjLife(1 To NumObjs)
    Else
        'We have no slots at all used, so kill the whole damn thing
        Log "Obj_CleanMapTile: Erasing ObjInfo() array", CodeTracker '//\\LOGLINE//\\
        Erase MapInfo(Map).ObjTile(X, Y).ObjInfo
        Erase MapInfo(Map).ObjTile(X, Y).ObjLife
    End If
    
    'Assign the value to the map array for later usage
    MapInfo(Map).ObjTile(X, Y).NumObjs = NumObjs

End Sub

Public Function Obj_ValidObjForClass(ByVal Class As Byte, ByVal ObjIndex As Integer) As Boolean

'*****************************************************************
'Checks if an object, by the object index, is useable by the passed class
'*****************************************************************
    
    'Check if theres a defined class requirement
    If ObjData.ClassReq(ObjIndex) > 0 Then
        
        'If Class AND ClassReq is true, then we meet the requirements
        Obj_ValidObjForClass = (Class And ObjData.ClassReq(ObjIndex))
        
    Else
        
        'No requirements
        Obj_ValidObjForClass = True
    
    End If

End Function

Public Sub Obj_Erase(ByVal Num As Integer, ByVal ObjSlot As Byte, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

'*****************************************************************
'Erase a object
'*****************************************************************

    Log "Call Obj_Erase(" & Num & "," & ObjSlot & "," & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\

    'Check for a valid index
    If ObjSlot > MapInfo(Map).ObjTile(X, Y).NumObjs Then
        Log "Obj_Erase: Invalid ObjSlot specified (" & ObjSlot & ")", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If
    
    'Check to erase every object
    If Num = -1 Then Num = MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount

    'Remove the amount
    Log "Obj_Erase: Removing " & Num & " objects from (" & Map & "," & X & "," & Y & ") - current amount = " & MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount, CodeTracker '//\\LOGLINE//\\
    MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount = MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount - Num
    
    'Check if they are all gone
    If MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).Amount <= 0 Then
        Log "Obj_Erase: Erasing object from client screens at (" & Map & "," & X & "," & Y & ")", CodeTracker '//\\LOGLINE//\\
        With MapInfo(Map).ObjTile(X, Y)
            .ObjInfo(ObjSlot).ObjIndex = 0
            .ObjInfo(ObjSlot).Amount = 0
            .ObjLife(ObjSlot) = 0
        End With
        ConBuf.PreAllocate 7
        ConBuf.Put_Byte DataCode.Server_EraseObject
        ConBuf.Put_Byte CByte(X)
        ConBuf.Put_Byte CByte(Y)
        ConBuf.Put_Long ObjData.GrhIndex(MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot).ObjIndex)
        Data_Send ToMap, 0, ConBuf.Get_Buffer, Map, PP_GroundObjects
    End If

End Sub

Public Sub Obj_ClosestFreeSpot(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByRef NewX As Byte, ByRef NewY As Byte, ByRef NewSlot As Byte)

'*****************************************************************
'Find the closest place to put an object
'*****************************************************************
Dim lX As Byte
Dim lY As Byte

    Log "Call Obj_ClosestFreeSpot(" & Map & "," & X & "," & Y & "," & NewX & "," & NewY & "," & NewSlot & ")", CodeTracker '//\\LOGLINE//\\
    
    'Check the defined location
    If Not (MapInfo(Map).Data(X, Y).Blocked And BlockedAll) Then
        If MapInfo(Map).ObjTile(X, Y).NumObjs < MaxObjsPerTile Then
            
            'Spot is useable
            NewX = X
            NewY = Y
            NewSlot = MapInfo(Map).ObjTile(X, Y).NumObjs + 1
            Exit Sub
            
        End If
    End If
    
    'Primary spot didn't work, so loop around it and check if those work
    If X > 0 Then
        If Y > 0 Then
            For lX = X - 1 To X + 1
                For lY = Y - 1 To Y + 1
                    If lX > 1 Then
                        If lX < MapInfo(Map).Width Then
                            If lY > 1 Then
                                If lY < MapInfo(Map).Height Then
                                    If MapInfo(Map).Data(lX, lY).Blocked = 0 Then
                                        If MapInfo(Map).ObjTile(lX, lY).NumObjs < MaxObjsPerTile Then
                                            
                                            'Spot is useable
                                            NewX = lX
                                            NewY = lY
                                            NewSlot = MapInfo(Map).ObjTile(lX, lY).NumObjs + 1
                                            Exit Sub
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next lY
            Next lX
        Else '//\\LOGLINE//\\
            Log "Obj_ClosestFreeSpot: X value is zero, can not subtract 1! Crash avoided!", CriticalError '//\\LOGLINE//\\
        End If
    Else    '//\\LOGLINE//\\
        Log "Obj_ClosestFreeSpot: X value is zero, can not subtract 1! Crash avoided!", CriticalError '//\\LOGLINE//\\
    End If
    
End Sub

Sub Obj_Make(Obj As Obj, ByVal ObjSlot As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal BypassUpdate As Byte = 0)

'*****************************************************************
'Create an object
'*****************************************************************

    Log "Call Obj_Make(N/A," & ObjSlot & "," & Map & "," & X & "," & Y & "," & BypassUpdate & ")", CodeTracker '//\\LOGLINE//\\

    'Make sure the ObjIndex isn't too high
    If ObjSlot > MaxObjsPerTile Then
        Log "Obj_Make: ObjSlot value too high (" & ObjSlot & ")", CriticalError '//\\LOGLINE//\\
        Exit Sub
    End If

    'Resize the object array to fit the slot
    If ObjSlot > MapInfo(Map).ObjTile(X, Y).NumObjs Then
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjInfo(1 To ObjSlot)
        ReDim Preserve MapInfo(Map).ObjTile(X, Y).ObjLife(1 To ObjSlot)
        MapInfo(Map).ObjTile(X, Y).NumObjs = ObjSlot
    End If
    
    'Add the object to the map slot
    MapInfo(Map).ObjTile(X, Y).ObjInfo(ObjSlot) = Obj
    MapInfo(Map).ObjTile(X, Y).ObjLife(ObjSlot) = timeGetTime
    
    'Clean the map tile just in case
    Obj_CleanMapTile Map, X, Y
    
    'Send the update to everyone on the map
    If BypassUpdate = 0 Then
        Log "Obj_Make: Updating object information with packet Server_MakeObject", CodeTracker '//\\LOGLINE//\\
        ConBuf.PreAllocate 7
        ConBuf.Put_Byte DataCode.Server_MakeObject
        ConBuf.Put_Long ObjData.GrhIndex(Obj.ObjIndex)
        ConBuf.Put_Byte X
        ConBuf.Put_Byte Y
        Data_Send ToMap, 0, ConBuf.Get_Buffer, Map, PP_GroundObjects
    End If
    
End Sub
