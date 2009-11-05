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
Private NumQuests As Long
Private OldQuests() As OldQuestData.Quest
Private NewQuests() As NewQuestData.Quest
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Sub OldToNew()
Dim i As Long

    'Convert the old to the new - because the types arn't the same (if they were, we wouldn't
    ' even be needing to use this code!), we have to copy every piece of information over
    ' variable by variable. Use this section to also do algorithms to set the variables, too,
    ' such as if you wanted to add in a Value for objects, and wanted to make the value
    ' a combination of a bunch of the object variables.
    For i = 1 To NumQuests
        NewQuests(i).AcceptLearnSkill = OldQuests(i).AcceptLearnSkill
        NewQuests(i).AcceptReqLvl = OldQuests(i).AcceptReqLvl
        NewQuests(i).AcceptReqObj = OldQuests(i).AcceptReqObj
        NewQuests(i).AcceptReqObjAmount = OldQuests(i).AcceptReqObjAmount
        NewQuests(i).AcceptRewExp = OldQuests(i).AcceptRewExp
        NewQuests(i).AcceptRewGold = OldQuests(i).AcceptRewGold
        NewQuests(i).AcceptRewObj = OldQuests(i).AcceptRewObj
        NewQuests(i).AcceptRewObjAmount = OldQuests(i).AcceptRewObjAmount
        NewQuests(i).AcceptTxt = OldQuests(i).AcceptTxt
        NewQuests(i).FinishLearnSkill = OldQuests(i).FinishLearnSkill
        NewQuests(i).FinishReqNPC = OldQuests(i).FinishReqNPC
        NewQuests(i).FinishReqNPCAmount = OldQuests(i).FinishReqNPCAmount
        NewQuests(i).FinishReqObj = OldQuests(i).FinishReqObj
        NewQuests(i).FinishReqObjAmount = OldQuests(i).FinishReqObjAmount
        NewQuests(i).FinishRewExp = OldQuests(i).FinishRewExp
        NewQuests(i).FinishRewGold = OldQuests(i).FinishRewGold
        NewQuests(i).FinishRewObj = OldQuests(i).FinishRewObj
        NewQuests(i).FinishRewObjAmount = OldQuests(i).FinishRewObjAmount
        NewQuests(i).FinishTxt = OldQuests(i).FinishTxt
        NewQuests(i).IncompleteTxt = OldQuests(i).IncompleteTxt
        NewQuests(i).Name = OldQuests(i).Name
        NewQuests(i).Redoable = OldQuests(i).Redoable
        NewQuests(i).StartTxt = OldQuests(i).StartTxt
    Next i

End Sub

Sub Main()
Dim FileNum As Byte

    'Load the file paths
    InitFilePaths
    
    'Get the number of Quests
    FileNum = FreeFile
    Open QuestsPath & "Count.quest" For Binary As FileNum
        Get FileNum, , NumQuests
    Close FileNum
    
    'Resize our arrays
    ReDim OldQuests(0 To NumQuests)
    ReDim NewQuests(0 To NumQuests)

    'Load the Quests
    Load_Quests
    
    'Save the backups
    Save_Quests_Backup
    
    'Convert the old variables to the new
    OldToNew
    
    'Save the Quests
    Save_Quests
    
    'Done
    MsgBox "Quests conversion successful!" & vbCrLf & _
           "Old type size: " & Len(OldQuests(0)) & vbCrLf & _
           "New type size: " & Len(NewQuests(0)) & vbCrLf & vbCrLf & _
           "Backups were made and placed in the following folder: " & vbCrLf & _
           QuestsPath & "Backups\ folder!" & vbCrLf & vbCrLf & _
           "Be sure to copy your NewQuestData to OldQuestData so the next time it will load properly!", vbOKOnly

End Sub

Sub Load_Quests()
Dim Quests As Long
Dim FileNum As Byte

    'Fill Quests List
    FileNum = FreeFile
    For Quests = 1 To NumQuests
        Open QuestsPath & Quests & ".quest" For Binary As FileNum
            Get FileNum, , OldQuests(Quests)
        Close FileNum
    Next Quests

End Sub

Sub Save_Quests()
Dim Quests As Long
Dim FileNum As Byte
    
    'Get the number of quests
    FileNum = FreeFile

    'Fill Quests List
    For Quests = 1 To NumQuests
        Open QuestsPath & Quests & ".quest" For Binary As FileNum
            Put FileNum, , NewQuests(Quests)
        Close FileNum
    Next Quests

End Sub

Sub Save_Quests_Backup()
Dim Quests As Long
Dim FileNum As Byte
    
    'Get the number of quests
    FileNum = FreeFile
    
    'Ensure the path exists
    MakeSureDirectoryPathExists QuestsPath & "Backup\"

    'Fill Quests List
    For Quests = 1 To NumQuests
        Open QuestsPath & "Backup\" & Quests & ".quest" For Binary As FileNum
            Put FileNum, , OldQuests(Quests)
        Close FileNum
    Next Quests

End Sub


