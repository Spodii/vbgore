Attribute VB_Name = "General"
'**       ____        _________   ______   ______  ______   _______           **
'**       \   \      /   /     \ /  ____\ /      \|      \ |   ____|          **
'**        \   \    /   /|      |  /     |        |       ||  |____           **
'***        \   \  /   / |     /| |  ___ |        |      / |   ____|         ***
'****        \   \/   /  |     \| |  \  \|        |   _  \ |  |____         ****
'******       \      /   |      |  \__|  |        |  | \  \|       |      ******
'********      \____/    |_____/ \______/ \______/|__|  \__\_______|    ********
'*******************************************************************************
'*******************************************************************************
'************ vbGORE - Visual Basic 6.0 Graphical Online RPG Engine ************
'************            Official Release: Version 0.1.1            ************
'************                 http://www.vbgore.com                 ************
'*******************************************************************************
'*******************************************************************************
'***** Source Distribution Information: ****************************************
'*******************************************************************************
'** If you wish to distribute this source code, you must distribute as-is     **
'** from the vbGORE website unless permission is given to do otherwise. This  **
'** comment block must remain in-tact in the distribution. If you wish to     **
'** distribute modified versions of vbGORE, please contact Spodi (info below) **
'** before distributing the source code. You may never label the source code  **
'** as the "Official Release" or similar unless the code and content remains  **
'** unmodified from the version downloaded from the official website.         **
'** You may also never sale the source code without permission first. If you  **
'** want to sell the code, please contact Spodi (below). This is to prevent   **
'** people from ripping off other people by selling an insignificantly        **
'** modified version of open-source code just to make a few quick bucks.      **
'*******************************************************************************
'***** Creating Engines With vbGORE: *******************************************
'*******************************************************************************
'** If you plan to create an engine with vbGORE that, please contact Spodi    **
'** before doing so. You may not sell the engine unless told elsewise (the    **
'** engine must has substantial modifications), and you may not claim it as   **
'** all your own work - credit must be given to vbGORE, along with a link to  **
'** the vbGORE homepage. Failure to gain approval from Spodi directly to      **
'** make a new engine with vbGORE will result in first a friendly reminder,   **
'** followed by much more drastic measures.                                   **
'*******************************************************************************
'***** Helping Out vbGORE: *****************************************************
'*******************************************************************************
'** If you want to help out with vbGORE's progress, theres a few things you   **
'** can do:                                                                   **
'**  *Donate - Great way to keep a free project going. :) Info and benifits   **
'**        for donating can be found at:                                      **
'**        http://www.vbgore.com/modules.php?name=Content&pa=showpage&pid=11  **
'**  *Contribute - Check out our forums, contribute ideas, report bugs, or    **
'**        create tutorials for the Knowledge Base. :)                        **
'**  *Ads - Advertisements have been placed on the site for those who can     **
'**        not or do not want to donate. Not donating is understandable - not **
'**        everyone has access to credit cards / paypal or spair money laying **
'**        around. These ads allow for a free way for you to help out the     **
'**        site. Those who do donate have the option to hide/remove the ads.  **
'*******************************************************************************
'***** Conact Information: *****************************************************
'*******************************************************************************
'** Please contact the creator of vbGORE (Spodi) directly with any questions: **
'** AIM: Spodii                          Yahoo: Spodii                        **
'** MSN: Spodii@hotmail.com              Email: spodi@vbgore.com              **
'** 2nd Email: spodii@hotmail.com        Website: http://www.vbgore.com       **
'*******************************************************************************
'***** Credits: ****************************************************************
'*******************************************************************************
'** Below are credits to those who have helped with the project or who have   **
'** distributed source code which has help this project's creation. The below **
'** is listed in no particular order of significance:                         **
'**                                                                           **
'** ORE (Aaron Perkins): Used as base engine and for learning experience      **
'**   http://www.baronsoft.com/                                               **
'** SOX (Trevor Herselman): Used for all the networking                       **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=35239&lngWId=1      **
'** Compression Methods (Marco v/d Berg): Provided compression algorithms     **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1      **
'** All Files In Folder (Jorge Colaccini): Algorithm implimented into engine  **
'**   http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=51435&lngWId=1      **
'** Game Programming Wiki (All community): Help on many different subjects    **
'**   http://wwww.gpwiki.org/                                                 **
'** ORE Maraxus's Edition (Maraxus): Used the map editor from this project    **
'**                                                                           **
'** Also, all the members of the vbGORE community who have submitted          **
'** tutorials, bugs, suggestions, criticism and have just stuck around!!      **
'** Big thanks goes to Van, Nex666 and ChAsE01!                               **
'**                                                                           **
'** If you feel you belong in these credits, please contact Spodi (above).    **
'*******************************************************************************
'*******************************************************************************

Option Explicit

'Quest
Public Type Quest
    Name As String                  'The quest's name
    StartTxt As String              'What the NPC says to the player to explain the crisis
    AcceptTxt As String             'What the NPC says when the player accepts the quest
    IncompleteTxt As String         'What the NPC says to the player when they return without completing the quest
    FinishTxt As String             'What the NPC says when the player finishes the quest
    AcceptReqLvl As Long            'What level the user is required to be to accept
    AcceptReqObj As Integer         'Index of the object the user is required to have to accept
    AcceptReqObjAmount As Integer   'How much of the object the user must have before accepting
    AcceptRewExp As Long            'Amount of Exp the user gets for accepting the quest
    AcceptRewGold As Long           'Amount of gold the user gets for accepting the quest
    AcceptRewObj As Integer         'Object the user gets for accepting the quest
    AcceptRewObjAmount As Integer   'Amount of the object the user gets for accepting the quest
    AcceptLearnSkill As Byte        'Skill the user learns for accepting the quest (by SkID value)
    FinishReqObj As Integer         'Object the user must have to finish the quest
    FinishReqObjAmount As Integer   'Amount of the object the user must have to finish the quest
    FinishReqNPC As Integer         'Index of the NPC the user must kill to finish the quest
    FinishReqNPCAmount As Integer   'How many of the NPCs the user must kill to finish the quest
    FinishRewExp As Long            'Exp the user gets for finishing the quest
    FinishRewGold As Long           'How much gold the user gets for finishing the quest
    FinishRewObj As Integer         'The index of the object the user gets for finishing the quest
    FinishRewObjAmount As Integer   'The amount of the object the user gets for finishing the quest
    FinishLearnSkill As Byte        'Skill the user learns for finishing the quest (by SkID value)
    Redoable As Byte                'If the quest can be done infinite times
End Type

Public QuestPath As String

'Our current quest
Public QuestNum As Integer
Public OpenQuest As Quest

Sub Main()
Dim FilePath As String

    QuestPath = App.Path & "\Quests\"

    'Show the main form
    frmMain.Show

    'Check for the first quest
    If Command$ = "" Then
        If Engine_FileExist(QuestPath & "1.quest", vbNormal) Then Editor_LoadQuest 1
    Else
        FilePath = Mid$(Command$, 2, Len(Command$) - 2) 'Retrieve the filepath from Command$ and crop off the "'s
        Editor_LoadQuest Val(Right$(FilePath, Len(FilePath) - Len(QuestPath)))
    End If

End Sub

Sub Editor_LoadQuest(ByVal QuestID As Integer)

'*****************************************************************
'Load the selected quest
'*****************************************************************
Dim FileNum As Byte

    'Check that the file exists
    If Engine_FileExist(QuestPath & QuestID & ".quest", vbNormal) = False Then
        MsgBox "The selected quest file (" & QuestPath & QuestID & ".quest) does not exist!", vbOKOnly
        Exit Sub
    End If

    QuestNum = QuestID

    'Open the file
    FileNum = FreeFile
    Open QuestPath & QuestID & ".quest" For Binary As #FileNum
        Get #FileNum, , OpenQuest
    Close #FileNum
    
    'Fill in the information
    With frmMain
        .Caption = "Quest Editor - Quest: " & QuestNum
        .AAAmountTxt.Text = OpenQuest.AcceptReqObjAmount
        .AALvlTxt.Text = OpenQuest.AcceptReqLvl
        .AAObjTxt.Text = OpenQuest.AcceptReqObj
        .AcceptTxt.Text = OpenQuest.AcceptTxt
        .ARExpTxt.Text = OpenQuest.AcceptRewExp
        .ARGoldTxt.Text = OpenQuest.AcceptRewGold
        .ARObjAmountTxt.Text = OpenQuest.AcceptRewObjAmount
        .ARObjTxt.Text = OpenQuest.AcceptRewObj
        .ARSkillTxt.Text = OpenQuest.AcceptLearnSkill
        .FANPCAmountTxt.Text = OpenQuest.FinishReqNPCAmount
        .FANPCTxt.Text = OpenQuest.FinishReqNPC
        .FAObjAmountTxt.Text = OpenQuest.FinishReqObjAmount
        .FAObjtxt.Text = OpenQuest.FinishReqObj
        .FRExpTxt.Text = OpenQuest.FinishRewExp
        .FRGoldTxt.Text = OpenQuest.FinishRewGold
        .FRObjAmountTxt.Text = OpenQuest.FinishRewObjAmount
        .FRObjTxt.Text = OpenQuest.FinishRewObj
        .FRSkillTxt.Text = OpenQuest.FinishLearnSkill
        .StartTxt.Text = OpenQuest.StartTxt
        .NameTxt.Text = OpenQuest.Name
        .FinishTxt.Text = OpenQuest.FinishTxt
        .IncompleteTxt.Text = OpenQuest.IncompleteTxt
        .RedoChk.Value = OpenQuest.Redoable
    End With

End Sub

Sub Editor_SaveQuest(ByVal QuestID As Integer)

'*****************************************************************
'Save the selected quest
'*****************************************************************
Dim FileNum As Byte
Dim Num As Integer

    QuestNum = QuestID
    
    'Fill in the information
    With frmMain
        .Caption = "Quest Editor - Quest: " & QuestNum
        OpenQuest.AcceptReqObjAmount = .AAAmountTxt.Text
        OpenQuest.AcceptReqLvl = .AALvlTxt.Text
        OpenQuest.AcceptReqObj = .AAObjTxt.Text
        OpenQuest.AcceptTxt = .AcceptTxt.Text
        OpenQuest.AcceptRewExp = .ARExpTxt.Text
        OpenQuest.AcceptRewGold = .ARGoldTxt.Text
        OpenQuest.AcceptRewObjAmount = .ARObjAmountTxt.Text
        OpenQuest.AcceptRewObj = .ARObjTxt.Text
        OpenQuest.AcceptLearnSkill = .ARSkillTxt.Text
        OpenQuest.FinishReqNPCAmount = .FANPCAmountTxt.Text
        OpenQuest.FinishReqNPC = .FANPCTxt.Text
        OpenQuest.FinishReqObjAmount = .FAObjAmountTxt.Text
        OpenQuest.FinishReqObj = .FAObjtxt.Text
        OpenQuest.FinishRewExp = .FRExpTxt.Text
        OpenQuest.FinishRewGold = .FRGoldTxt.Text
        OpenQuest.FinishRewObjAmount = .FRObjAmountTxt.Text
        OpenQuest.FinishRewObj = .FRObjTxt.Text
        OpenQuest.FinishLearnSkill = .FRSkillTxt.Text
        OpenQuest.StartTxt = .StartTxt.Text
        OpenQuest.Name = .NameTxt.Text
        OpenQuest.FinishTxt = .FinishTxt.Text
        OpenQuest.IncompleteTxt = .IncompleteTxt.Text
        OpenQuest.Redoable = .RedoChk.Value
    End With
    
    'Check to update the number of quests
    FileNum = FreeFile
    Open QuestPath & "Count.quest" For Binary As #FileNum
        Get #FileNum, , Num
    Close #FileNum
    If Num < QuestID Then
        Open QuestPath & "Count.quest" For Binary As #FileNum
            Put #FileNum, , QuestID
        Close #FileNum
    End If
        
    'Open the file
    FileNum = FreeFile
    Open QuestPath & QuestID & ".quest" For Binary As #FileNum
        Put #FileNum, , OpenQuest
    Close #FileNum
    
    'Saved
    MsgBox "Quest " & QuestID & " saved successfully!", vbOKOnly

End Sub

Function Engine_FileExist(file As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(file, FileType) <> "")

End Function
