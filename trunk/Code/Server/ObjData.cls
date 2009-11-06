VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Object data array (is NOT ordered by the obj index)
Private cData() As udtObjData

'UBound of the cData array
Private cDataUBound As Integer

'Number of unused indexes in the cData array
Private cDataUnused As Integer

'Object timers (index is by obj index)
Private ObjTimer() As Long

'Maximum number of datas
Private MaxDatas As Integer

'Used to get the data index from the object index (ex: DataIndex = ObjIndexToDataIndex(ObjIndex) )
Private ObjIndexToDataIndex() As Integer

Public Sub SetDataUBound(ByVal NumDatas As Integer)
'*****************************************************************
'Updates the needed arrays with the highest object data index
'More info: http://www.vbgore.com/GameServer.ObjData.SetDataUBound
'*****************************************************************

    'Set the number of object datas
    ReDim ObjIndexToDataIndex(1 To NumDatas) As Integer
    ReDim ObjTimer(1 To NumDatas) As Long
    MaxDatas = NumDatas

End Sub

Public Sub CheckObjUnloading()
'*****************************************************************
'Loops through all the objects and checks if it is time to unload them
'More info: http://www.vbgore.com/GameServer.ObjData.CheckObjUnloading
'*****************************************************************
Dim i As Integer

    'Loop through all the objects in use
    For i = 1 To cDataUBound    'Only loop to the highest index
        
        'Check if the object is in use (pointer > 0)
        If cData(i).Pointer > 0 Then
        
            'Check if enough time has elapsed to unload the object
            If ObjTimer(cData(i).Pointer) < timeGetTime Then
            
                'Set the object as unused
                ObjTimer(cData(i).Pointer) = 0
                cData(i).Pointer = 0
                cDataUnused = cDataUnused + 1

            End If
        
        End If
    
    Next i

End Sub

Private Sub ReadyObj(ByVal ObjIndex As Integer)
'*****************************************************************
'Makes sure an object is ready (and updates its unload timer)
'More info: http://www.vbgore.com/GameServer.ObjData.ReadyObj
'*****************************************************************
Dim FileNum As Byte
Dim i As Integer

    'Make sure the object is loaded
    If ObjTimer(ObjIndex) = 0 Then
        
        'Find the next free cData index
        If cDataUnused > 0 Then
        
            'Theres an unused index, loop through them until we find it
            i = 0
            Do
                i = i + 1
            Loop While cData(i).Pointer > 0
            cDataUnused = cDataUnused - 1
        
        Else
        
            'There is no unused index, resize the array
            cDataUBound = cDataUBound + 1
            ReDim Preserve cData(1 To cDataUBound) As udtObjData
            i = cDataUBound
            
        End If

        'Load the object
        FileNum = FreeFile
        Open ServerTempPath & "o" & ObjIndex & ".temp" For Binary Access Read As #FileNum
            Get #FileNum, , cData(i)
        Close #FileNum
        cData(i).Pointer = ObjIndex
        ObjIndexToDataIndex(ObjIndex) = i

    End If
    
    'Set the timer
    ObjTimer(ObjIndex) = timeGetTime + ObjMemoryLife

End Sub

Public Property Get Name(ByVal Index As Integer) As String
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            Name = cData(ObjIndexToDataIndex(Index)).Name
        End If
    End If
End Property

Public Property Get ObjType(ByVal Index As Integer) As Byte
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ObjType = cData(ObjIndexToDataIndex(Index)).ObjType
        End If
    End If
End Property

Public Property Get GrhIndex(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            GrhIndex = cData(ObjIndexToDataIndex(Index)).GrhIndex
        End If
    End If
End Property

Public Property Get SpriteBody(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            SpriteBody = cData(ObjIndexToDataIndex(Index)).SpriteBody
        End If
    End If
End Property

Public Property Get SpriteWeapon(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            SpriteWeapon = cData(ObjIndexToDataIndex(Index)).SpriteWeapon
        End If
    End If
End Property

Public Property Get SpriteHair(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            SpriteHair = cData(ObjIndexToDataIndex(Index)).SpriteHair
        End If
    End If
End Property

Public Property Get SpriteHead(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            SpriteHead = cData(ObjIndexToDataIndex(Index)).SpriteHead
        End If
    End If
End Property

Public Property Get SpriteWings(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            SpriteWings = cData(ObjIndexToDataIndex(Index)).SpriteWings
        End If
    End If
End Property

Public Property Get WeaponType(ByVal Index As Integer) As Byte
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            WeaponType = cData(ObjIndexToDataIndex(Index)).WeaponType
        End If
    End If
End Property

Public Property Get WeaponRange(ByVal Index As Integer) As Byte
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            WeaponRange = cData(ObjIndexToDataIndex(Index)).WeaponRange
        End If
    End If
End Property

Public Property Get UseGrh(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            UseGrh = cData(ObjIndexToDataIndex(Index)).UseGrh
        End If
    End If
End Property

Public Property Get ProjectileRotateSpeed(ByVal Index As Integer) As Byte
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ProjectileRotateSpeed = cData(ObjIndexToDataIndex(Index)).ProjectileRotateSpeed
        End If
    End If
End Property

Public Property Get Value(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            Value = cData(ObjIndexToDataIndex(Index)).Value
        End If
    End If
End Property

Public Property Get RepHP(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            RepHP = cData(ObjIndexToDataIndex(Index)).RepHP
        End If
    End If
End Property

Public Property Get RepMP(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            RepMP = cData(ObjIndexToDataIndex(Index)).RepMP
        End If
    End If
End Property

Public Property Get RepSP(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            RepSP = cData(ObjIndexToDataIndex(Index)).RepSP
        End If
    End If
End Property

Public Property Get RepHPP(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            RepHPP = cData(ObjIndexToDataIndex(Index)).RepHPP
        End If
    End If
End Property

Public Property Get RepMPP(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            RepMPP = cData(ObjIndexToDataIndex(Index)).RepMPP
        End If
    End If
End Property

Public Property Get RepSPP(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            RepSPP = cData(ObjIndexToDataIndex(Index)).RepSPP
        End If
    End If
End Property

Public Property Get ReqStr(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ReqStr = cData(ObjIndexToDataIndex(Index)).ReqStr
        End If
    End If
End Property

Public Property Get ReqAgi(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ReqAgi = cData(ObjIndexToDataIndex(Index)).ReqAgi
        End If
    End If
End Property

Public Property Get ReqMag(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ReqMag = cData(ObjIndexToDataIndex(Index)).ReqMag
        End If
    End If
End Property

Public Property Get ReqLvl(ByVal Index As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ReqLvl = cData(ObjIndexToDataIndex(Index)).ReqLvl
        End If
    End If
End Property

Public Property Get Stacking(ByVal Index As Integer) As Integer
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            Stacking = cData(ObjIndexToDataIndex(Index)).Stacking
        End If
    End If
End Property

Public Property Get UseSfx(ByVal Index As Integer) As Byte
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            UseSfx = cData(ObjIndexToDataIndex(Index)).UseSfx
        End If
    End If
End Property

Public Property Get ClassReq(ByVal Index As Integer) As Byte
    If Index > 0 Then
        If Index <= MaxDatas Then
            ReadyObj Index
            ClassReq = cData(ObjIndexToDataIndex(Index)).ClassReq
        End If
    End If
End Property

Public Property Get AddStat(ByVal Index As Integer, ByVal StatIndex As Integer) As Long
    If Index > 0 Then
        If Index <= MaxDatas Then
            If StatIndex > 0 Then
                If StatIndex <= NumStats Then
                    ReadyObj Index
                    AddStat = cData(ObjIndexToDataIndex(Index)).AddStat(StatIndex)
                End If
            End If
        End If
    End If
End Property
