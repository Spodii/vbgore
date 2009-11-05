Attribute VB_Name = "General"
'**************************************************************************************
'**************************************************************************************
'***                                   HOW TO USE                                   ***
'**************************************************************************************
'**************************************************************************************
'*** To use this conversion tool, you must place the old variable formats in the    ***
'*** OldData module so it can load correctly. Place the variable format you wish to ***
'*** change to in the NewData module. Note that changes may need to take place in-  ***
'*** between the loading and saving process, so use this if needed.                 ***
'***                                                                                ***
'*** Once you finish a conversion, make sure you copy the variables from NewData    ***
'*** to OldData, since they are now your most up-to-date ones. :)                   ***
'**************************************************************************************
'**************************************************************************************
Option Explicit
Private NumObjDatas As Long
Private OldObj() As OldObjData.ObjData
Private NewObj() As NewObjData.ObjData
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Sub OldToNew()
Dim i As Long
Dim j As Long

    'Convert the old to the new - because the types arn't the same (if they were, we wouldn't
    ' even be needing to use this code!), we have to copy every piece of information over
    ' variable by variable. Use this section to also do algorithms to set the variables, too,
    ' such as if you wanted to add in a Value for objects, and wanted to make the value
    ' a combination of a bunch of the object variables.
    For i = 1 To NumObjDatas
        For j = 1 To OldObjData.NumStats
            NewObj(i).AddStat(j) = OldObj(i).AddStat(j)
        Next j
        NewObj(i).GrhIndex = OldObj(i).GrhIndex
        NewObj(i).Name = OldObj(i).Name
        NewObj(i).ObjType = OldObj(i).ObjType
        NewObj(i).Price = OldObj(i).Price
        NewObj(i).RepHP = OldObj(i).RepHP
        NewObj(i).RepHPP = OldObj(i).RepHPP
        NewObj(i).RepMP = OldObj(i).RepMP
        NewObj(i).RepMPP = OldObj(i).RepMPP
        NewObj(i).RepSP = OldObj(i).RepSP
        NewObj(i).RepSPP = OldObj(i).RepSPP
        NewObj(i).SpriteBody = OldObj(i).SpriteBody
        NewObj(i).SpriteHair = OldObj(i).SpriteHair
        NewObj(i).SpriteHead = OldObj(i).SpriteHead
        NewObj(i).SpriteHelm = OldObj(i).SpriteHelm
        NewObj(i).SpriteWeapon = OldObj(i).SpriteWeapon
        NewObj(i).WeaponType = OldObj(i).WeaponType
    Next i

End Sub

Sub Main()
Dim FileNum As Byte

    'Load the file paths
    InitFilePaths
    
    'Get the number of objects
    FileNum = FreeFile
    Open OBJsPath & "Count.obj" For Binary As FileNum
        Get FileNum, , NumObjDatas
    Close FileNum
    
    'Resize our arrays
    ReDim OldObj(0 To NumObjDatas)
    ReDim NewObj(0 To NumObjDatas)

    'Load the objects
    Load_OBJs
    
    'Save the backups
    Save_OBJs_Backup
    
    'Convert the old variables to the new
    OldToNew
    
    'Save the objects
    Save_OBJs
    
    'Done
    MsgBox "Objects conversion successful!" & vbCrLf & _
           "Old type size: " & Len(OldObj(0)) & vbCrLf & _
           "New type size: " & Len(NewObj(0)) & vbCrLf & vbCrLf & _
           "Backups were made and placed in the following folder: " & vbCrLf & _
           OBJsPath & "Backups\ folder!" & vbCrLf & vbCrLf & _
           "Be sure to copy your NewObjData to OldObjData so the next time it will load properly!", vbOKOnly

End Sub

Sub Load_OBJs()
Dim Object As Long
Dim FileNum As Byte

    'Fill Object List
    FileNum = FreeFile
    For Object = 1 To NumObjDatas
        Open OBJsPath & Object & ".obj" For Binary As FileNum
            Get FileNum, , OldObj(Object)
        Close FileNum
    Next Object

End Sub

Sub Save_OBJs()
Dim Object As Long
Dim FileNum As Byte
    
    'Get the number of objects
    FileNum = FreeFile

    'Fill Object List
    For Object = 1 To NumObjDatas
        Open OBJsPath & Object & ".obj" For Binary As FileNum
            Put FileNum, , NewObj(Object)
        Close FileNum
    Next Object

End Sub

Sub Save_OBJs_Backup()
Dim Object As Long
Dim FileNum As Byte
    
    'Get the number of objects
    FileNum = FreeFile
    
    'Ensure the path exists
    MakeSureDirectoryPathExists OBJsPath & "Backup\"

    'Fill Object List
    For Object = 1 To NumObjDatas
        Open OBJsPath & "Backup\" & Object & ".obj" For Binary As FileNum
            Put FileNum, , OldObj(Object)
        Close FileNum
    Next Object

End Sub
