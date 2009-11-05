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
Private OldUser() As OldUserData.User
Private NewUser() As NewUserData.User
Private CharPaths() As String
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Sub OldToNew()
Dim j As Long
Dim User As Long

    'Convert the old to the new - because the types arn't the same (if they were, we wouldn't
    ' even be needing to use this code!), we have to copy every piece of information over
    ' variable by variable. Use this section to also do algorithms to set the variables, too,
    ' such as if you wanted to add in a Value for objects, and wanted to make the value
    ' a combination of a bunch of the object variables.
    
    '*** Keep in mind you only have to convert those which need to be saved! ***
    
    For User = 1 To UBound(CharPaths)
        NewUser(User).ArmorEqpSlot = OldUser(User).ArmorEqpSlot
        NewUser(User).Char.Body = OldUser(User).Char.Body
        NewUser(User).Char.CharIndex = OldUser(User).Char.CharIndex
        NewUser(User).Char.Desc = OldUser(User).Char.Desc
        NewUser(User).Char.Hair = OldUser(User).Char.Hair
        NewUser(User).Char.Head = OldUser(User).Char.Head
        NewUser(User).Char.Heading = OldUser(User).Char.Heading
        NewUser(User).Char.HeadHeading = OldUser(User).Char.HeadHeading
        NewUser(User).Char.Weapon = OldUser(User).Char.Weapon
        NewUser(User).CompletedQuests = OldUser(User).CompletedQuests
        NewUser(User).Desc = OldUser(User).Desc
        For j = 1 To UBound(OldUser(User).Object)
            NewUser(User).Object(j).Amount = OldUser(User).Object(j).Amount
            NewUser(User).Object(j).Equipped = OldUser(User).Object(j).Equipped
            NewUser(User).Object(j).ObjIndex = OldUser(User).Object(j).ObjIndex
        Next j
        NewUser(User).Pos.Map = OldUser(User).Pos.Map
        NewUser(User).Pos.X = OldUser(User).Pos.X
        NewUser(User).Pos.Y = OldUser(User).Pos.Y
        For j = 1 To UBound(OldUser(User).Quest)
            NewUser(User).Quest(j) = OldUser(User).Quest(j)
        Next j
        NewUser(User).Skills.Bless = OldUser(User).Skills.Bless
        NewUser(User).Skills.IronSkin = OldUser(User).Skills.IronSkin
        NewUser(User).Skills.Protect = OldUser(User).Skills.Protect
        NewUser(User).Skills.Strengthen = OldUser(User).Skills.Strengthen
        NewUser(User).Skills.WarCurse = OldUser(User).Skills.WarCurse
        NewUser(User).WeaponEqpSlot = OldUser(User).WeaponEqpSlot
        NewUser(User).WeaponType = OldUser(User).WeaponType
        For j = 1 To UBound(OldUser(User).MailID)
            NewUser(User).MailID(j) = OldUser(User).MailID(j)
        Next j
        For j = 1 To UBound(OldUser(User).KnownSkills)
            NewUser(User).KnownSkills(j) = OldUser(User).KnownSkills(j)
        Next j
        
        'NOTE - THIS REPLACES THE LOADING ROUTINE FOR USERSTATS CLASS MODULE!
        For j = 1 To OldUserData.NumStats
            NewUser(User).BaseStats(j) = OldUser(User).BaseStats(j)
        Next j
        
    Next User

End Sub

Sub Main()

    If MsgBox("This will NOT create a backup of your character files - you must do this yourself!" & vbCrLf & _
        "Press NO now to quit, or press YES to continue with the conversion.", vbYesNo) = vbNo Then End

    'Load the file paths
    InitFilePaths
    
    'Get the character paths
    CharPaths = AllFilesInFolders(CharPath, False)
    
    'Resize our arrays
    ReDim OldUser(1 To UBound(CharPaths))
    ReDim NewUser(1 To UBound(CharPaths))

    'Load the users
    Load_Users
    
    'Convert the old variables to the new
    OldToNew
    
    'Save the users
    Save_Users
    
    'Done
    MsgBox "User conversion successful!" & vbCrLf & _
           "Old type size: " & Len(OldUser(0)) & vbCrLf & _
           "New type size: " & Len(NewUser(0)) & vbCrLf & _
           "Be sure to copy your NewUserData to OldUserData so it will load correctly next time!", vbOKOnly
           
End Sub

Sub Load_Users()
Dim User As Long
Dim FileNum As Byte
Dim i As Integer
Dim j As Long

    'Fill the character list
    FileNum = FreeFile
    For User = 1 To UBound(CharPaths)
        Open CharPaths(User) For Binary As FileNum
            Get FileNum, , OldUser(User).ArmorEqpSlot
            Get FileNum, , OldUser(User).Char
            Get FileNum, , i
            OldUser(User).CompletedQuests = Space$(i)
            Get FileNum, , OldUser(User).CompletedQuests
            Get FileNum, , OldUser(User).Desc
            Get FileNum, , OldUser(User).Object
            Get FileNum, , OldUser(User).Pos
            Get FileNum, , OldUser(User).Quest()
            Get FileNum, , OldUser(User).Skills
            Get FileNum, , OldUser(User).WeaponEqpSlot
            Get FileNum, , OldUser(User).WeaponType
            Get FileNum, , OldUser(User).MailID
            Get FileNum, , OldUser(User).KnownSkills
            
            'NOTE - THIS REPLACES THE LOADING ROUTINE FOR USERSTATS CLASS MODULE!
            For j = 1 To OldUserData.NumStats
                Get FileNum, , OldUser(User).BaseStats(j)
            Next j
            
        Close FileNum
    Next User

End Sub

Sub Save_Users()
Dim User As Long
Dim FileNum As Byte
Dim j As Long
Dim i As Integer

    'Fill the character list
    FileNum = FreeFile
    For User = 1 To UBound(CharPaths)
        Open CharPaths(User) For Binary As FileNum
            Put FileNum, , NewUser(User).ArmorEqpSlot
            Put FileNum, , NewUser(User).Char
            i = CInt(Len(NewUser(User).CompletedQuests))
            Put FileNum, , i
            Put FileNum, , NewUser(User).CompletedQuests
            Put FileNum, , NewUser(User).Desc
            Put FileNum, , NewUser(User).Object
            Put FileNum, , NewUser(User).Pos
            Put FileNum, , NewUser(User).Quest()
            Put FileNum, , NewUser(User).Skills
            Put FileNum, , NewUser(User).WeaponEqpSlot
            Put FileNum, , NewUser(User).WeaponType
            Put FileNum, , NewUser(User).MailID
            Put FileNum, , NewUser(User).KnownSkills
            
            'NOTE - THIS REPLACES THE LOADING ROUTINE FOR USERSTATS CLASS MODULE!
            For j = 1 To NewUserData.NumStats
                Put FileNum, , NewUser(User).BaseStats(j)
            Next j
            
        Close FileNum
    Next User

End Sub
