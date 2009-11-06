Attribute VB_Name = "Skills"
Option Explicit

'General
Private Const MaxSummons As Byte = 3        'Maximum number of characters on player can summon

'Bless
Private Const Bless_Cost As Single = 0.5    'Magic * Bless_Cost
Private Const Bless_Length As Long = 300 'How long the skill lasts
Private Const Bless_Exhaust As Long = 3500  'Exhaustion time
Private Const Bless_Sfx As Byte = 8

'Protection
Private Const Pro_Cost As Single = 0.5      'Magic * Pro_Cost
Private Const Pro_Length As Long = 300000
Private Const Pro_Exhaust As Long = 2000
Private Const Pro_Sfx As Byte = 8

'Strengthen
Private Const Str_Cost As Single = 0.5      'Magic * Str_Cost
Private Const Str_Length As Long = 300000
Private Const Str_Exhaust As Long = 2000
Private Const Str_Sfx As Byte = 8

'Warcry
Private Const Warcry_Cost As Single = 0.5   'Strength * Warcry_Cost
Private Const Warcry_Length As Long = 15000
Private Const Warcry_Exhaust As Long = 1500

'Heal
Private Const Heal_Cost As Single = 0.5     'Magic * Heal_Cost
Private Const Heal_Value As Single = 1.5    'Magic * Heal_Value = MinHP Raised
Public Const Heal_Exhaust As Long = 1000
Public Const Heal_ClassReq As Integer = -1   'Class requirements - unfortunately, since its a constant, you can't use ClassID, but have to use the value directly

'Summon bandit
Private Const SumBandit_Cost As Single = 1       'Magic * SummonBandit_Cost
Private Const SumBandit_Exhaust As Long = 3000
Private Const SumBandit_Length As Long = 300000  'Summoned NPC automatically dispells (dies) after this time goes by

Public Sub Skill_SummonBandit_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Summon a bandit to fight with you
'*****************************************************************
Dim CharIndex As Integer
Dim tIndex As Integer

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.SummonBandit) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * SumBandit_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If

    'Make sure the user doesn't have too many summons already
    If UserList(CasterIndex).NumSlaves >= MaxSummons Then
        Data_Send ToIndex, CasterIndex, cMessage(127).Data
        Exit Sub
    End If

    'Summon the NPC
    tIndex = Load_NPC(2, 1, SumBandit_Length)
    
    'Check for an invalid index (load failed)
    If tIndex < 1 Then Exit Sub
    
    'Find a legal position
    Server_ClosestLegalPos UserList(CasterIndex).Pos, NPCList(tIndex).Pos
    
    'Check if the position is legal
    If Not Server_LegalPos(NPCList(tIndex).Pos.Map, NPCList(tIndex).Pos.X, NPCList(tIndex).Pos.Y, 0) Then
        NPC_Close tIndex
        Exit Sub
    End If
    
    'Set up the NPC's information
    NPCList(tIndex).ChatID = 0
    NPCList(tIndex).Attackable = 1
    NPCList(tIndex).Hostile = 1
    NPCList(tIndex).AI = 7
    NPCList(tIndex).Name = "Summoned " & NPCList(tIndex).Name
    NPCList(tIndex).BaseStat(SID.Agi) = NPCList(tIndex).BaseStat(SID.Agi) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 10)
    NPCList(tIndex).BaseStat(SID.DEF) = NPCList(tIndex).BaseStat(SID.DEF) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 10)
    NPCList(tIndex).BaseStat(SID.MinHIT) = NPCList(tIndex).BaseStat(SID.MinHIT) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 10)
    NPCList(tIndex).BaseStat(SID.MaxHIT) = NPCList(tIndex).BaseStat(SID.MaxHIT) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 10)
    NPCList(tIndex).BaseStat(SID.Speed) = NPCList(tIndex).BaseStat(SID.Speed) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 20)
    NPCList(tIndex).BaseStat(SID.MaxHP) = NPCList(tIndex).BaseStat(SID.MaxHP) + UserList(CasterIndex).Stats.ModStat(SID.Mag)
    NPCList(tIndex).BaseStat(SID.MinHP) = NPCList(tIndex).BaseStat(SID.MaxHP)
    NPCList(tIndex).ModStat(SID.MaxHP) = NPCList(tIndex).BaseStat(SID.MaxHP)
    NPC_UpdateModStats tIndex
    
    'Set up the NPC on the map / char array
    MapInfo(NPCList(tIndex).Pos.Map).Data(NPCList(tIndex).Pos.X, NPCList(tIndex).Pos.Y).NPCIndex = tIndex
    CharIndex = Server_NextOpenCharIndex
    NPCList(tIndex).Char.CharIndex = CharIndex
    CharList(CharIndex).Index = tIndex
    CharList(CharIndex).CharType = CharType_NPC
    
    'Bind the NPC to the user
    UserList(CasterIndex).NumSlaves = UserList(CasterIndex).NumSlaves + 1
    ReDim Preserve UserList(CasterIndex).SlaveNPCIndex(1 To UserList(CasterIndex).NumSlaves)
    UserList(CasterIndex).SlaveNPCIndex(UserList(CasterIndex).NumSlaves) = tIndex
    NPCList(tIndex).OwnerIndex = CasterIndex
    
    'Display the NPC
    NPC_Spawn tIndex
    NPC_MakeChar ToMap, CasterIndex, tIndex, NPCList(tIndex).Pos.Map, NPCList(tIndex).Pos.X, NPCList(tIndex).Pos.Y
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + SumBandit_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Display the effect on the map - this must be done AFTER the NPC is made
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.SummonBandit
    ConBuf.Put_Integer NPCList(tIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell

    'Reduce the user's mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * SumBandit_Cost)
    
End Sub

Public Function Skill_ValidSkillForClass(ByVal Class As Integer, ByVal SkillID As Byte) As Boolean

'*****************************************************************
'Check if the SkillID can be used by the class
'For skills with no defined requirements, theres no requirements
'Heal only has a requirement as an example
'*****************************************************************
Dim ClassReq As Integer

    'Sort by skill id
    Select Case SkillID
        Case SkID.Heal: ClassReq = Heal_ClassReq
    End Select
    
    'Treat 0 as "all classes can use"
    If ClassReq <> 0 Then
    
        'Check the ClassReq VS the passed class
        Skill_ValidSkillForClass = (Class And ClassReq)

    Else

        'No requirements
        Skill_ValidSkillForClass = True

    End If

End Function

Public Sub Skill_Bless_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises all the character's stats
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Bless) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.BlessCounter > 0 Then
        If NPCList(TargetIndex).Skills.Bless > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.PreAllocate 9 + Len(UserList(CasterIndex).Name)  '4 + "bless" = 9
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "bless"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
            
        End If
    End If
    
    'Display the bless icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Bless = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.BlessCounter = timeGetTime + Bless_Length
    NPCList(TargetIndex).Skills.Bless = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Bless_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Send the message to the caster
    ConBuf.PreAllocate 3 + Len(NPCList(TargetIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 40
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

    'Display the effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Bless
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Bless_Sfx
    ConBuf.Put_Byte UserList(CasterIndex).Pos.X
    ConBuf.Put_Byte UserList(CasterIndex).Pos.Y
    Data_Send ToPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
End Sub

Public Sub Skill_Protection_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises all the character's stats
'*****************************************************************
    
    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > timeGetTime Then Exit Sub
    If NPCList(CasterIndex).Counters.ActionDelay > timeGetTime Then Exit Sub
    
    'Check for enough mana to cast
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.Mag) * Pro_Cost) Then Exit Sub

    'Reduce the mana
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Pro_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.ProtectCounter > 0 Then
        If NPCList(TargetIndex).Skills.Protect > NPCList(CasterIndex).ModStat(SID.Mag) Then
            'Power of what we are casting is weaker then what is already applied
            Exit Sub
        End If
    End If
    
    'Display the protection icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Protect = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.ProtectCounter = timeGetTime + Pro_Length
    NPCList(TargetIndex).Skills.Protect = NPCList(CasterIndex).BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    NPCList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Pro_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons

    'Display the effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Protection
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Pro_Sfx
    ConBuf.Put_Byte NPCList(CasterIndex).Pos.X
    ConBuf.Put_Byte NPCList(CasterIndex).Pos.Y
    Data_Send ToNPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
End Sub

Public Sub Skill_Strengthen_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises all the character's stats
'*****************************************************************
    
    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > timeGetTime Then Exit Sub
    If NPCList(CasterIndex).Counters.ActionDelay > timeGetTime Then Exit Sub
    
    'Check for enough mana to cast
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.Mag) * Str_Cost) Then Exit Sub

    'Reduce the mana
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Str_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.StrengthenCounter > 0 Then
        If NPCList(TargetIndex).Skills.Strengthen > NPCList(CasterIndex).ModStat(SID.Mag) Then
            'Power of what we are casting is weaker then what is already applied
            Exit Sub
        End If
    End If
    
    'Display the strengthen icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Strengthen = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.StrengthenCounter = timeGetTime + Str_Length
    NPCList(TargetIndex).Skills.Strengthen = NPCList(CasterIndex).BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    NPCList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Str_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons

    'Display the effect on the map
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Strengthen
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Str_Sfx
    ConBuf.Put_Byte NPCList(CasterIndex).Pos.X
    ConBuf.Put_Byte NPCList(CasterIndex).Pos.Y
    Data_Send ToNPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
End Sub

Public Sub Skill_Bless_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises all the character's stats
'*****************************************************************
    
    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > timeGetTime Then Exit Sub
    If NPCList(CasterIndex).Counters.ActionDelay > timeGetTime Then Exit Sub
    
    'Check for enough mana to cast
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.Mag) * Bless_Cost) Then Exit Sub

    'Reduce the mana
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Bless_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.BlessCounter > 0 Then
        If NPCList(TargetIndex).Skills.Bless > NPCList(CasterIndex).ModStat(SID.Mag) Then
            'Power of what we are casting is weaker then what is already applied
            Exit Sub
        End If
    End If
    
    'Display the bless icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Bless = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.BlessCounter = timeGetTime + Bless_Length
    NPCList(TargetIndex).Skills.Bless = NPCList(CasterIndex).BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    NPCList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Bless_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons

    'Display the effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Bless
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Bless_Sfx
    ConBuf.Put_Byte NPCList(CasterIndex).Pos.X
    ConBuf.Put_Byte NPCList(CasterIndex).Pos.Y
    Data_Send ToNPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Upate the NPC's stats that was casted on
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
End Sub

Public Sub Skill_Strengthen_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises the character's damage
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Strengthen) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.StrengthenCounter > 0 Then
        If NPCList(TargetIndex).Skills.Strengthen > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.PreAllocate 10 + Len(UserList(CasterIndex).Name) '4 + "strengthen" = 10
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "strengthen"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
            
        End If
    End If
    
    'Display the strengthen icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Strengthen = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.StrengthenCounter = timeGetTime + Str_Length
    NPCList(TargetIndex).Skills.Strengthen = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Str_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Send the message to the caster
    ConBuf.PreAllocate 3 + Len(NPCList(TargetIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 46
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

    'Display the effect on the map
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Strengthen
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Str_Sfx
    ConBuf.Put_Byte UserList(CasterIndex).Pos.X
    ConBuf.Put_Byte UserList(CasterIndex).Pos.Y
    Data_Send ToPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
End Sub

Public Sub Skill_Protection_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises the character's defence
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Protection) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.ProtectCounter > 0 Then
        If NPCList(TargetIndex).Skills.Protect > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.PreAllocate 14 + Len(UserList(CasterIndex).Name) '4 + "protection" = 14
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "protection"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
            
        End If
    End If
    
    'Display the protection icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Protect = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.ProtectCounter = timeGetTime + Pro_Length
    NPCList(TargetIndex).Skills.Protect = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Pro_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Send the message to the caster
    ConBuf.PreAllocate 3 + Len(NPCList(TargetIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 44
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Protection
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Pro_Sfx
    ConBuf.Put_Byte UserList(CasterIndex).Pos.X
    ConBuf.Put_Byte UserList(CasterIndex).Pos.Y
    Data_Send ToPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
End Sub

Public Sub Skill_Bless_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises all the character's stats
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    
    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Bless) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost)
    
    'Cast on the target
    If UserList(TargetIndex).Counters.BlessCounter > 0 Then
        If UserList(TargetIndex).Skills.Bless > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.PreAllocate 9 + Len(UserList(CasterIndex).Name)  '4 + "bless" = 9
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "bless"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
            
        End If
    End If
    
    'Display the bless icon (only if it isn't already displayed)
    If UserList(TargetIndex).Skills.Bless = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    UserList(TargetIndex).Counters.BlessCounter = timeGetTime + Bless_Length
    UserList(TargetIndex).Skills.Bless = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    UserList(TargetIndex).Stats.Update = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Bless_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Send the message to the caster
    If TargetIndex <> CasterIndex Then
        ConBuf.PreAllocate 3 + Len(UserList(TargetIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 40
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
        
    End If
    
    'Send the message to the target
    ConBuf.PreAllocate 5 + Len(UserList(CasterIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 41
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Skills.Bless
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Display the effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Bless
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Bless_Sfx
    ConBuf.Put_Byte UserList(CasterIndex).Pos.X
    ConBuf.Put_Byte UserList(CasterIndex).Pos.Y
    Data_Send ToPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
End Sub

Public Sub Skill_Strengthen_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises the character's damage
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    
    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Strengthen) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost)
    
    'Cast on the target
    If UserList(TargetIndex).Counters.StrengthenCounter > 0 Then
        If UserList(TargetIndex).Skills.Strengthen > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.PreAllocate 14 + Len(UserList(CasterIndex).Name) '4 + "strengthen" = 14
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "strengthen"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
            
        End If
    End If
    
    'Display the strengthen icon (only if it isn't already displayed)
    If UserList(TargetIndex).Skills.Strengthen = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    UserList(TargetIndex).Counters.StrengthenCounter = timeGetTime + Str_Length
    UserList(TargetIndex).Skills.Strengthen = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    UserList(TargetIndex).Stats.Update = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Str_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Send the message to the caster
    If TargetIndex <> CasterIndex Then
        ConBuf.PreAllocate 3 + Len(UserList(TargetIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 46
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
        
    End If
    
    'Send the message to the target
    ConBuf.PreAllocate 5 + Len(UserList(CasterIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 47
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Skills.Strengthen
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Strengthen
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Str_Sfx
    ConBuf.Put_Byte UserList(CasterIndex).Pos.X
    ConBuf.Put_Byte UserList(CasterIndex).Pos.Y
    Data_Send ToPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
End Sub

Public Sub Skill_Protection_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Raises the character's defence
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    
    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Protection) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost)
    
    'Cast on the target
    If UserList(TargetIndex).Counters.ProtectCounter > 0 Then
        If UserList(TargetIndex).Skills.Protect > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.PreAllocate 14 + Len(UserList(CasterIndex).Name) '4 + "protection" = 14
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "protection"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
            
        End If
    End If
    
    'Display the protection icon (only if it isn't already displayed)
    If UserList(TargetIndex).Skills.Protect = 0 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    End If
    
    'Apply the spell's effects
    UserList(TargetIndex).Counters.ProtectCounter = timeGetTime + Pro_Length
    UserList(TargetIndex).Skills.Protect = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    UserList(TargetIndex).Stats.Update = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Pro_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons
    
    'Send the message to the caster
    If TargetIndex <> CasterIndex Then
        ConBuf.PreAllocate 3 + Len(UserList(TargetIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 44
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
        
    End If
    
    'Send the message to the target
    ConBuf.PreAllocate 5 + Len(UserList(CasterIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 45
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Skills.Protect
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Protection
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    'Play sound effect
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_PlaySound3D
    ConBuf.Put_Byte Pro_Sfx
    ConBuf.Put_Byte UserList(CasterIndex).Pos.X
    ConBuf.Put_Byte UserList(CasterIndex).Pos.Y
    Data_Send ToPCArea, CasterIndex, ConBuf.Get_Buffer, , PP_Sound
    
End Sub

Public Sub Skill_Heal_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Sub

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Heal) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < UserList(CasterIndex).Stats.BaseStat(SID.Mag) * Heal_Cost Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Cost)

    'Check for a valid range
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Heal_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons

    'Create casting effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell

    'Cast on the target
    UserList(TargetIndex).Stats.BaseStat(SID.MinHP) = UserList(TargetIndex).Stats.BaseStat(SID.MinHP) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Value)

    'Message to the caster
    If CasterIndex <> TargetIndex Then
        ConBuf.PreAllocate 3 + Len(UserList(TargetIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 42
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    End If
    
    'Message to the target
    ConBuf.PreAllocate 5 + Len(UserList(CasterIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 43
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

End Sub

Public Function Skill_Heal_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Boolean

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > timeGetTime Then Exit Function
    If NPCList(CasterIndex).Counters.ActionDelay > timeGetTime Then Exit Function

    'Check for enough mana
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < NPCList(CasterIndex).BaseStat(SID.Mag) * Heal_Cost Then Exit Function
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Heal_Cost)

    'Apply spell exhaustion
    NPCList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Heal_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_StatusIcons

    'Create casting effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map, PP_DisplaySpell

    'Cast on the target
    NPC_Heal TargetIndex, (NPCList(CasterIndex).ModStat(SID.Mag) * Heal_Value)
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    Skill_Heal_NPCtoNPC = True
    
End Function

Public Sub Skill_Heal_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Sub

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Heal) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < UserList(CasterIndex).Stats.BaseStat(SID.Mag) * Heal_Cost Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Cost)

    'Check for a valid range
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Sub

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Heal_Exhaust
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons

    'Create casting effect
    ConBuf.PreAllocate 6
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell

    'Cast on the target
    NPC_Heal TargetIndex, (UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Value)
    
    'Message to the caster
    ConBuf.PreAllocate 3 + Len(NPCList(TargetIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 42
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

End Sub

Public Sub Skill_IronSkin_PC(ByVal UserIndex As Integer)

'*****************************************************************
'Decreases user attack by 50% to increase defence by 200%
'*****************************************************************

    'Check for invalid values
    If UserIndex = 0 Then Exit Sub
  
    'Check for the skill in the user posession
    If UserList(UserIndex).KnownSkills(SkID.IronSkin) = 0 Then
        Data_Send ToIndex, UserIndex, cMessage(37).Data
        Exit Sub
    End If

    'Check if still exhausted
    If UserList(UserIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    UserList(UserIndex).Counters.SpellExhaustion = timeGetTime + 2000
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_StatusIcons

    'Remove the Iron Skin
    If UserList(UserIndex).Skills.IronSkin = 1 Then
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_StatusIcons

    Else 'Enable the Iron Skin
        UserList(UserIndex).Skills.IronSkin = 1
        
        ConBuf.PreAllocate 4
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map, PP_StatusIcons
    End If
    
    UserList(UserIndex).Stats.Update = 1

End Sub

Public Sub Skill_SpikeField(ByVal CasterIndex As Integer)

'*****************************************************************
'Forms a field of spikes around the user
'      |3|
'    |3|4|3|
'  |3|2|2|2|3|
'  |3|2|1|2|3|
'  |4|3|U|3|4|
'    |3|4|3|
'*****************************************************************

'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'--- WARNING! THIS SKILL IS CODED POORLY AND I DO NOT RECOMMEND YOU USE IT IN YOUR GAME! ---
'-------------------------------------------------------------------------------------------
'--- Theres problems that will arise if you cast it on the side of the map, along with   ---
'--- it just isn't a good structure as it is                                             ---
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

Dim aMap As Integer
Dim aX As Integer
Dim aY As Integer
Dim Damage As Long

    'Check for spell exhaustion
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub

    'Check if the user knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.SpikeField) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 2) Then
        Data_Send ToIndex, CasterIndex, cMessage(38).Data
        Exit Sub
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) \ 2)

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + 3000
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_StatusIcons

    'Set the values to shorter variables
    Damage = UserList(CasterIndex).Stats.BaseStat(SID.Mag) + 5
    aMap = UserList(CasterIndex).Pos.Map
    aX = UserList(CasterIndex).Pos.X
    aY = UserList(CasterIndex).Pos.Y

    'Loop through all the tiles, damaging any NPC on them
    'NORTH
    On Error Resume Next
    If UserList(CasterIndex).Char.HeadHeading = NORTH Or UserList(CasterIndex).Char.HeadHeading = NORTHEAST Then
        If MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 1).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 2, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY).NPCIndex, CasterIndex, Damage * 0.25

        If MapInfo(aMap).Data(aX - 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 1).NPCIndex, CasterIndex, Damage
        If MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 2).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 1, aY - 3).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY - 3).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 1, aY - 3).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 3).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX, aY - 4).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 4).NPCIndex, CasterIndex, Damage * 0.25

        'EAST
    ElseIf UserList(CasterIndex).Char.HeadHeading = EAST Or UserList(CasterIndex).Char.HeadHeading = SOUTHEAST Then
        If MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 2).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 2).NPCIndex, CasterIndex, Damage * 0.25

        If MapInfo(aMap).Data(aX + 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY).NPCIndex, CasterIndex, Damage
        If MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX + 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX + 3, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 3, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 3, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 3, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 3, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 3, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX + 4, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 4, aY).NPCIndex, CasterIndex, Damage * 0.25

        'SOUTH
    ElseIf UserList(CasterIndex).Char.HeadHeading = SOUTH Or UserList(CasterIndex).Char.HeadHeading = SOUTHWEST Then
        If MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 1).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 2, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY).NPCIndex, CasterIndex, Damage * 0.25

        If MapInfo(aMap).Data(aX - 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 1).NPCIndex, CasterIndex, Damage
        If MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 2).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX + 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 1, aY + 3).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 3).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 1, aY + 3).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 3).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX, aY + 4).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 4).NPCIndex, CasterIndex, Damage * 0.25

        'WEST
    ElseIf UserList(CasterIndex).Char.HeadHeading = WEST Or UserList(CasterIndex).Char.HeadHeading = NORTHWEST Then
        If MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 2).NPCIndex, CasterIndex, Damage * 0.25
        If MapInfo(aMap).Data(aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX, aY + 2).NPCIndex, CasterIndex, Damage * 0.25

        If MapInfo(aMap).Data(aX - 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY).NPCIndex, CasterIndex, Damage
        If MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX - 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 2, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY).NPCIndex, CasterIndex, Damage * 0.5
        If MapInfo(aMap).Data(aX - 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 3, aY - 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 3, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 3, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 3, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapInfo(aMap).Data(aX - 3, aY + 1).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 3, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapInfo(aMap).Data(aX - 4, aY).NPCIndex > 0 Then NPC_Damage MapInfo(aMap).Data(aX - 4, aY).NPCIndex, CasterIndex, Damage * 0.25

    End If

    'Display the user casting it on other people's screens
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.SpikeField
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map, PP_DisplaySpell
    
    On Error GoTo 0
    
End Sub

Public Sub Skill_Warcry_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Lower the stats of all attackable hostiles in range
'*****************************************************************
Dim LoopC As Integer
Dim WarCursePower As Integer

    'Check if still exhausted
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    
    'Check if the user knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Warcry) = 0 Then
        Data_Send ToIndex, CasterIndex, cMessage(37).Data
        Exit Sub
    End If

    'Check for enough endurance
    If UserList(CasterIndex).Stats.BaseStat(SID.MinSTA) < Int(UserList(CasterIndex).Stats.ModStat(SID.Str) * Warcry_Cost) Then
        Data_Send ToIndex, CasterIndex, cMessage(48).Data
        Exit Sub
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinSTA) = UserList(CasterIndex).Stats.BaseStat(SID.MinSTA) - Int(UserList(CasterIndex).Stats.ModStat(SID.Str) * Warcry_Cost)

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = timeGetTime + Warcry_Exhaust

    'Cast on all attackable hostile NPCs in the PC area
    Data_Send ToIndex, CasterIndex, cMessage(49).Data

    'Loop through all the alive and active NPCs
    WarCursePower = UserList(CasterIndex).Stats.ModStat(SID.Str)
    For LoopC = 1 To LastNPC
        If NPCList(LoopC).Flags.NPCActive Then
            If NPCList(LoopC).Flags.NPCAlive Then
                If NPCList(LoopC).Pos.Map = UserList(CasterIndex).Pos.Map Then
                    If NPCList(LoopC).Attackable Then
                        If NPCList(LoopC).Hostile Then
                            If NPCList(LoopC).OwnerIndex = 0 Then
                                If NPCList(LoopC).Skills.WarCurse <= WarCursePower Then
                                    If Server_RectDistance(UserList(CasterIndex).Pos.X, UserList(CasterIndex).Pos.Y, NPCList(LoopC).Pos.X, NPCList(LoopC).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then
    
                                        'Tell the users in the screen that the NPC is weaker
                                        ConBuf.PreAllocate 3 + Len(NPCList(LoopC).Name)
                                        ConBuf.Put_Byte DataCode.Server_Message
                                        ConBuf.Put_Byte 50
                                        ConBuf.Put_String NPCList(LoopC).Name
                                        Data_Send ToNPCArea, LoopC, ConBuf.Get_Buffer
                                        
                                        'Warcurse icon
                                        If NPCList(LoopC).Skills.WarCurse = 0 Then
                                            ConBuf.PreAllocate 4
                                            ConBuf.Put_Byte DataCode.Server_IconWarCursed
                                            ConBuf.Put_Byte 1
                                            ConBuf.Put_Integer NPCList(LoopC).Char.CharIndex
                                            Data_Send ToMap, 0, ConBuf.Get_Buffer, NPCList(LoopC).Pos.Map, PP_StatusIcons
                                        End If
                                        NPCList(LoopC).Skills.WarCurse = WarCursePower
                                        NPCList(LoopC).Counters.WarCurseCounter = timeGetTime + Warcry_Length
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
       End If
    Next LoopC
    
End Sub
