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
Private NumNPCs As Long
Private OldNPC() As OldNPCData.NPC
Private NewNPC() As NewNPCData.NPC
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Sub OldToNew()
Dim i As Long
Dim j As Long

    'Convert the old to the new - because the types arn't the same (if they were, we wouldn't
    ' even be needing to use this code!), we have to copy every piece of information over
    ' variable by variable. Use this section to also do algorithms to set the variables, too,
    ' such as if you wanted to add in a Value for objects, and wanted to make the value
    ' a combination of a bunch of the object variables.
    For i = 1 To NumNPCs
        NewNPC(i).Attackable = OldNPC(i).Attackable
        For j = 1 To UBound(NewNPC(i).BaseStat())
            NewNPC(i).BaseStat(j) = OldNPC(i).BaseStat(j)
        Next j
        NewNPC(i).Char.Body = OldNPC(i).Char.Body
        NewNPC(i).Char.CharIndex = OldNPC(i).Char.CharIndex
        NewNPC(i).Char.Desc = OldNPC(i).Char.Desc
        NewNPC(i).Char.Hair = OldNPC(i).Char.Hair
        NewNPC(i).Char.Head = OldNPC(i).Char.Head
        NewNPC(i).Char.HeadHeading = OldNPC(i).Char.HeadHeading
        NewNPC(i).Char.Heading = OldNPC(i).Char.Heading
        NewNPC(i).Char.Weapon = OldNPC(i).Char.Weapon
        NewNPC(i).Desc = OldNPC(i).Desc
        NewNPC(i).GiveEXP = OldNPC(i).GiveEXP
        NewNPC(i).GiveGLD = OldNPC(i).GiveGLD
        NewNPC(i).Hostile = OldNPC(i).Hostile
        NewNPC(i).Movement = OldNPC(i).Movement
        NewNPC(i).Name = OldNPC(i).Name
        NewNPC(i).NumVendItems = OldNPC(i).NumVendItems
        NewNPC(i).Quest = OldNPC(i).Quest
        NewNPC(i).StartPos.Map = OldNPC(i).StartPos.Map
        NewNPC(i).StartPos.X = OldNPC(i).StartPos.X
        NewNPC(i).StartPos.Y = OldNPC(i).StartPos.Y
        NewNPC(i).VendItems = OldNPC(i).VendItems
    Next i

End Sub

Sub Main()
Dim FileNum As Byte

    'Load the file paths
    InitFilePaths
    
    'Get the number of objects
    FileNum = FreeFile
    Open NPCsPath & "Count.npc" For Binary As FileNum
        Get FileNum, , NumNPCs
    Close FileNum
    
    'Resize our arrays
    ReDim OldNPC(0 To NumNPCs)
    ReDim NewNPC(0 To NumNPCs)

    'Load the objects
    Load_NPCs
    
    'Save the backups
    Save_NPCs_Backup
    
    'Convert the old variables to the new
    OldToNew
    
    'Save the objects
    Save_NPCs
    
    'Done
    MsgBox "NPCs conversion successful!" & vbCrLf & _
           "Old type size: " & Len(OldNPC(0)) & vbCrLf & _
           "New type size: " & Len(NewNPC(0)) & vbCrLf & vbCrLf & _
           "Backups were made and placed in the following folder: " & vbCrLf & _
           NPCsPath & "Backups\ folder!" & vbCrLf & vbCrLf & _
           "Be sure to copy your NewNPCData to OldNPCData so the next time it will load properly!", vbOKOnly

End Sub

Sub Load_NPCs()
Dim NPC As Long
Dim FileNum As Byte

    'Fill Object List
    FileNum = FreeFile
    For NPC = 1 To NumNPCs
        Open NPCsPath & NPC & ".npc" For Binary As FileNum
            Get FileNum, , NewNPC(NPC)
        Close FileNum
    Next NPC

End Sub

Sub Save_NPCs()
Dim NPC As Long
Dim FileNum As Byte

    'Get the number of objects
    FileNum = FreeFile
    
    'Erase old files
    For NPC = 1 To NumNPCs
        Kill NPCsPath & NPC & ".npc"
    Next NPC

    'Fill Object List
    FileNum = FreeFile
    For NPC = 1 To NumNPCs
        Open NPCsPath & NPC & ".npc" For Binary As FileNum
            Put FileNum, , NewNPC(NPC)
        Close FileNum
    Next NPC

End Sub

Sub Save_NPCs_Backup()
Dim NPC As Long
Dim FileNum As Byte
    
    'Get the number of objects
    FileNum = FreeFile
    
    'Ensure the path exists
    MakeSureDirectoryPathExists NPCsPath & "Backup\"

    'Fill Object List
    FileNum = FreeFile
    For NPC = 1 To NumNPCs
        Open NPCsPath & "Backup\" & NPC & ".npc" For Binary As FileNum
            Put FileNum, , NewNPC(NPC)
        Close FileNum
    Next NPC

End Sub


