Attribute VB_Name = "DataIDs"
Option Explicit

'********** Emoticons ************
Public Const NumEmotes As Byte = 10
Public Type EmoID
    Dots As Byte
    Exclimation As Byte
    Question As Byte
    Surprised As Byte
    Heart As Byte
    Hearts As Byte
    HeartBroken As Byte
    Utensils As Byte
    Meat As Byte
    ExcliQuestion As Byte
End Type
Public EmoID As EmoID

'********** Classes ************
'Classes work by using bitwise operations, so each class ID must be a power of 2 (1, 2, 4, 8, 16, 32, 64, or 128)
'If you want more clases, change the classes to "Integer"
'To set class requirements, OR the values together
'EX:
'ClassReq = Warrior OR Rogue
'This means the class must be a Warrior or Rogue
'To check the values, use AND
'EX:
'If ClassReq AND UserClass Then
'   User meets requirements
'Else
'   User doesn't meet requirements
'End If
Public Type ClassID
    Warrior As Byte
    Mage As Byte
    Rogue As Byte
    NoReq As Byte
End Type
Public ClassID As ClassID

'********** Packets ************
'Data String Codenames (Reduces all data transfers to 1 byte tags)
Public Type DataCode
    Comm_Talk As Byte
    Comm_Shout As Byte
    Comm_Emote As Byte
    Comm_Whisper As Byte
    Comm_FontType_Talk As Byte
    Comm_FontType_Fight As Byte
    Comm_FontType_Info As Byte
    Comm_FontType_Quest As Byte
    Comm_UseBubble As Byte  'Do not use this alone - OR it onto Comm_Talk!
    Server_MailMessage As Byte
    Server_MailBox As Byte
    Server_MailItemTake As Byte
    Server_MailItemRemove As Byte
    Server_MailDelete As Byte
    Server_MailCompose As Byte
    Server_UserCharIndex As Byte
    Server_SetUserPosition As Byte
    Server_MakeChar As Byte
    Server_EraseChar As Byte
    Server_MoveChar As Byte
    Server_ChangeChar As Byte
    Server_MakeObject As Byte
    Server_EraseObject As Byte
    Server_PlaySound As Byte
    Server_PlaySound3D As Byte
    Server_Who As Byte
    Server_CharHP As Byte
    Server_CharMP As Byte
    Server_IconCursed As Byte
    Server_IconWarCursed As Byte
    Server_IconBlessed As Byte
    Server_IconStrengthened As Byte
    Server_IconProtected As Byte
    Server_IconIronSkin As Byte
    Server_IconSpellExhaustion As Byte
    Server_SetCharDamage As Byte
    Server_Help As Byte
    Server_Disconnect As Byte
    Server_Connect As Byte
    Server_Message As Byte
    Server_SetCharSpeed As Byte
    Server_MakeProjectile As Byte
    Server_MakeSlash As Byte
    Server_MailObjUpdate As Byte
    Server_MakeEffect As Byte
    Server_PTD As Byte
    Server_SendQuestInfo As Byte
    Map_LoadMap As Byte
    Map_DoneLoadingMap As Byte
    Map_SendName As Byte
    Map_RequestPositionUpdate As Byte
    User_Target As Byte
    User_KnownSkills As Byte
    User_Attack As Byte
    User_SetInventorySlot As Byte
    User_Desc As Byte
    User_Login  As Byte
    User_NewLogin As Byte
    User_Get As Byte
    User_Drop As Byte
    User_Use As Byte
    User_Move As Byte
    User_Rotate As Byte
    User_LeftClick As Byte
    User_RightClick As Byte
    User_LookLeft As Byte
    User_LookRight As Byte
    User_AggressiveFace As Byte
    User_Blink As Byte
    User_Trade_StartNPCTrade As Byte
    User_Trade_BuyFromNPC As Byte
    User_Trade_SellToNPC As Byte
    User_Bank_Open As Byte
    User_Bank_PutItem As Byte
    User_Bank_TakeItem As Byte
    User_Bank_UpdateSlot As Byte
    User_BaseStat As Byte
    User_ModStat As Byte
    User_CastSkill As Byte
    User_ChangeInvSlot As Byte
    User_Emote As Byte
    User_StartQuest As Byte
    User_SetWeaponRange As Byte
    User_RequestMakeChar As Byte
    User_RequestUserCharIndex As Byte
    User_ChangeServer As Byte
    User_AddFriend As Byte
    User_ConfirmPosition As Byte
    User_RemoveFriend As Byte
    GM_Approach As Byte
    GM_Summon As Byte
    GM_Kick As Byte
    GM_Raise As Byte
    GM_SetGMLevel As Byte
    GM_Thrall As Byte
    GM_DeThrall As Byte
    GM_BanIP As Byte
    GM_UnBanIP As Byte
End Type
Public DataCode As DataCode

'********** Character Stats/Skills ************
Public Const NumStats As Byte = 18
Public Const NumSkills As Byte = 7
Public Const FirstModStat As Byte = 9   'The lowest number of the first stat that can be modded
Public Type StatOrder
    'These can NOT be modded (theres no ModStat())
    MinMAN As Byte
    MinHP As Byte
    MinSTA As Byte
    Gold As Byte
    Points As Byte
    EXP As Byte
    ELU As Byte
    ELV As Byte
    'These CAN be modded (ModStat() is used)
    MaxHIT As Byte
    MinHIT As Byte
    DEF As Byte
    MaxHP As Byte
    MaxSTA As Byte
    MaxMAN As Byte
    Str As Byte
    Agi As Byte
    Mag As Byte
    Speed As Byte   'Speed works as + (Speed / 2) on the client since just + Speed would be too drastic (8 would double the normal speed)
End Type
Public SID As StatOrder 'Stat ID
Public Type SkillID
    Bless As Byte
    Protection As Byte
    Strengthen As Byte
    Warcry As Byte
    Heal As Byte
    IronSkin As Byte
    SpikeField As Byte
End Type
Public SkID As SkillID  'Skill IDs

Public Sub InitDataCommands()

    'Load the values for the data commands
    
    With EmoID
        .Dots = 1
        .Exclimation = 2
        .Question = 3
        .Surprised = 4
        .Heart = 5
        .Hearts = 6
        .HeartBroken = 7
        .Utensils = 8
        .Meat = 9
        .ExcliQuestion = 10
    End With

    With SkID
        .Bless = 1
        .Heal = 2
        .IronSkin = 3
        .Protection = 4
        .Strengthen = 5
        .Warcry = 6
        .SpikeField = 7
    End With
    
    With ClassID
        'These values must be based off of powers of 2!
        .Warrior = 1    '2 ^ 0
        .Mage = 2       '2 ^ 1
        .Rogue = 4      '2 ^ 2 ... etc
        
        'This sets every bit to 1, which means that it will work with every class
        .NoReq = 255
    End With

    With SID
        'These can NOT be modded
        .MinHP = 1
        .MinMAN = 2
        .MinSTA = 3
        .Gold = 4
        .Points = 5
        .EXP = 6
        .ELU = 7
        .ELV = 8
        'These CAN be modded
        .MaxHIT = 9
        .MaxHP = 10
        .MaxMAN = 11
        .MaxSTA = 12
        .MinHIT = 13
        .DEF = 14
        .Agi = 15
        .Mag = 16
        .Str = 17
        .Speed = 18
    End With

    With DataCode
        .User_RequestMakeChar = 1
        .GM_Thrall = 2
        .Server_IconSpellExhaustion = 3
        .Comm_Shout = 4
        .Server_UserCharIndex = 5
        .Comm_Emote = 6
        .Server_SetUserPosition = 7
        .Map_LoadMap = 8
        .Map_DoneLoadingMap = 9
        .GM_Raise = 10
        .GM_Kick = 11
        .Server_CharHP = 12
        .GM_Summon = 13
        .User_ChangeServer = 14
        .Map_SendName = 15
        .User_Attack = 16
        .Server_MakeChar = 17
        .Server_EraseChar = 18
        .Server_MoveChar = 19
        .Server_ChangeChar = 20
        .Server_MakeObject = 21
        .Server_EraseObject = 22
        .User_KnownSkills = 23
        .User_SetInventorySlot = 24
        .User_StartQuest = 25
        .Server_Connect = 26
        .Server_PlaySound = 27
        .User_Login = 28
        .User_NewLogin = 29
        .Comm_Whisper = 30
        .Server_Who = 31
        .User_Move = 32
        .User_Rotate = 33
        .User_LeftClick = 34
        .User_RightClick = 35
        .Map_RequestPositionUpdate = 36
        .User_Get = 37
        .User_Drop = 38
        .User_Use = 39
        .GM_Approach = 40
        .Comm_Talk = 41
        .Server_SetCharDamage = 42
        .User_ChangeInvSlot = 43
        .User_Emote = 44
        .Server_CharMP = 45
        .Server_Disconnect = 46
        .User_LookLeft = 47
        .User_LookRight = 48
        .User_Blink = 49
        .User_AggressiveFace = 50
        .User_Trade_BuyFromNPC = 51
        .User_BaseStat = 52
        .User_ModStat = 53
        .Comm_FontType_Fight = 54
        .Comm_FontType_Info = 55
        .Comm_FontType_Quest = 56
        .Comm_FontType_Talk = 57
        .Server_Help = 58
        .User_Desc = 59
        .Server_PTD = 60
        .User_Target = 61
        .User_Trade_StartNPCTrade = 62
        .User_Trade_SellToNPC = 63
        .User_CastSkill = 64
        .Server_IconCursed = 65
        .Server_IconWarCursed = 66
        .Server_IconBlessed = 67
        .Server_IconStrengthened = 68
        .Server_IconProtected = 69
        .Server_IconIronSkin = 70
        .Server_MailBox = 71
        .Server_MailMessage = 72
        .User_RequestUserCharIndex = 73
        .Server_MailItemTake = 74
        .Server_MailObjUpdate = 75
        .Server_MailDelete = 76
        .Server_MailCompose = 77
        .GM_SetGMLevel = 78
        .Server_Message = 79
        .GM_DeThrall = 80
        .Server_PlaySound3D = 81
        .Server_SetCharSpeed = 82
        .User_SetWeaponRange = 83
        .Server_MakeProjectile = 84
        .Server_MakeSlash = 85
        .Server_MakeEffect = 86
        .User_Bank_Open = 87
        .User_Bank_PutItem = 88
        .User_Bank_TakeItem = 89
        .User_Bank_UpdateSlot = 90
        .GM_BanIP = 91
        .GM_UnBanIP = 92
        .Server_SendQuestInfo = 93
        .User_ConfirmPosition = 94

        'Value 128 can be used over again since this does not count as an ID in itself - just ignore this variable! ;)
        .Comm_UseBubble = 128
        
    End With

End Sub

