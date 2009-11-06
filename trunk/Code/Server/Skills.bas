Attribute VB_Name = "Skills"
Option Explicit

'Bless
Private Const Bless_Cost As Single = 0.5    'Magic * Bless_Cost
Private Const Bless_Length As Long = 300000 'How long the skill lasts
Private Const Bless_Exhaust As Long = 3500  'Exhaustion time

'Protection
Private Const Pro_Cost As Single = 0.5      'Magic * Pro_Cost
Private Const Pro_Length As Long = 300000
Private Const Pro_Exhaust As Long = 2000

'Strengthen
Private Const Str_Cost As Single = 0.5      'Magic * Str_Cost
Private Const Str_Length As Long = 300000
Private Const Str_Exhaust As Long = 2000

'Warcry
Private Const Warcry_Cost As Single = 0.5   'Strength * Warcry_Cost
Private Const Warcry_Length As Long = 15000
Private Const Warcry_Exhaust As Long = 1500

'Heal
Private Const Heal_Cost As Single = 0.5     'Magic * Heal_Cost
Private Const Heal_Value As Single = 1.5    'Magic * Heal_Value = MinHP Raised
Public Const Heal_Exhaust As Long = 1000

Public Function Skill_Bless_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises all the character's stats
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Bless) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.BlessCounter > 0 Then
        If NPCList(TargetIndex).Skills.Bless > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "bless"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Function
            
        End If
    End If
    
    'Display the bless icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Bless = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.BlessCounter = CurrentTime + Bless_Length
    NPCList(TargetIndex).Skills.Bless = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Bless_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Send the message to the caster
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 40
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Bless
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    Skill_Bless_PCtoNPC = 1
    
End Function

Public Function Skill_Protection_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises all the character's stats
'*****************************************************************
    
    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > CurrentTime Then Exit Function
    If NPCList(CasterIndex).Counters.ActionDelay > CurrentTime Then Exit Function
    
    'Check for enough mana to cast
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.Mag) * Pro_Cost) Then Exit Function

    'Reduce the mana
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Pro_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.ProtectCounter > 0 Then
        If NPCList(TargetIndex).Skills.Protect > NPCList(CasterIndex).ModStat(SID.Mag) Then
            'Power of what we are casting is weaker then what is already applied
            Exit Function
        End If
    End If
    
    'Display the protection icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Protect = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.ProtectCounter = CurrentTime + Pro_Length
    NPCList(TargetIndex).Skills.Protect = NPCList(CasterIndex).BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    NPCList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Pro_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Protection
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    Skill_Protection_NPCtoNPC = 1
    
End Function

Public Function Skill_Strengthen_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises all the character's stats
'*****************************************************************
    
    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > CurrentTime Then Exit Function
    If NPCList(CasterIndex).Counters.ActionDelay > CurrentTime Then Exit Function
    
    'Check for enough mana to cast
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.Mag) * Str_Cost) Then Exit Function

    'Reduce the mana
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Str_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.StrengthenCounter > 0 Then
        If NPCList(TargetIndex).Skills.Strengthen > NPCList(CasterIndex).ModStat(SID.Mag) Then
            'Power of what we are casting is weaker then what is already applied
            Exit Function
        End If
    End If
    
    'Display the strengthen icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Strengthen = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.StrengthenCounter = CurrentTime + Str_Length
    NPCList(TargetIndex).Skills.Strengthen = NPCList(CasterIndex).BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    NPCList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Str_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Strengthen
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    Skill_Strengthen_NPCtoNPC = 1
    
End Function

Public Function Skill_Bless_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises all the character's stats
'*****************************************************************
    
    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > CurrentTime Then Exit Function
    If NPCList(CasterIndex).Counters.ActionDelay > CurrentTime Then Exit Function
    
    'Check for enough mana to cast
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.Mag) * Bless_Cost) Then Exit Function

    'Reduce the mana
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Bless_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.BlessCounter > 0 Then
        If NPCList(TargetIndex).Skills.Bless > NPCList(CasterIndex).ModStat(SID.Mag) Then
            'Power of what we are casting is weaker then what is already applied
            Exit Function
        End If
    End If
    
    'Display the bless icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Bless = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.BlessCounter = CurrentTime + Bless_Length
    NPCList(TargetIndex).Skills.Bless = NPCList(CasterIndex).BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    NPCList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Bless_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Bless
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    Skill_Bless_NPCtoNPC = 1
    
End Function

Public Function Skill_Strengthen_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises the character's damage
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Strengthen) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.StrengthenCounter > 0 Then
        If NPCList(TargetIndex).Skills.Strengthen > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "strengthen"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Function
            
        End If
    End If
    
    'Display the strengthen icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Strengthen = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.StrengthenCounter = CurrentTime + Str_Length
    NPCList(TargetIndex).Skills.Strengthen = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Str_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Send the message to the caster
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 46
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Strengthen
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    Skill_Strengthen_PCtoNPC = 1
    
End Function

Public Function Skill_Protection_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises the character's defence
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Protection) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost)
    
    'Cast on the target
    If NPCList(TargetIndex).Counters.ProtectCounter > 0 Then
        If NPCList(TargetIndex).Skills.Protect > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "protection"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Function
            
        End If
    End If
    
    'Display the protection icon (only if it isn't already displayed)
    If NPCList(TargetIndex).Skills.Protect = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    NPCList(TargetIndex).Counters.ProtectCounter = CurrentTime + Pro_Length
    NPCList(TargetIndex).Skills.Protect = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    NPCList(TargetIndex).Flags.UpdateStats = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Pro_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Send the message to the caster
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 44
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Protection
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    NPCList(TargetIndex).Flags.UpdateStats = 1
    Skill_Protection_PCtoNPC = 1
    
End Function

Public Function Skill_Bless_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises all the character's stats
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function
    
    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Bless) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Bless_Cost)
    
    'Cast on the target
    If UserList(TargetIndex).Counters.BlessCounter > 0 Then
        If UserList(TargetIndex).Skills.Bless > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "bless"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Function
            
        End If
    End If
    
    'Display the bless icon (only if it isn't already displayed)
    If UserList(TargetIndex).Skills.Bless = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    UserList(TargetIndex).Counters.BlessCounter = CurrentTime + Bless_Length
    UserList(TargetIndex).Skills.Bless = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    UserList(TargetIndex).Stats.Update = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Bless_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Send the message to the caster
    If TargetIndex <> CasterIndex Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 40
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
        
    End If
    
    'Send the message to the target
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 41
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Skills.Bless
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Bless
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    Skill_Bless_PCtoPC = 1
    
End Function

Public Function Skill_Strengthen_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises the character's damage
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function
    
    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Strengthen) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Str_Cost)
    
    'Cast on the target
    If UserList(TargetIndex).Counters.StrengthenCounter > 0 Then
        If UserList(TargetIndex).Skills.Strengthen > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "strengthen"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Function
            
        End If
    End If
    
    'Display the strengthen icon (only if it isn't already displayed)
    If UserList(TargetIndex).Skills.Strengthen = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    UserList(TargetIndex).Counters.StrengthenCounter = CurrentTime + Str_Length
    UserList(TargetIndex).Skills.Strengthen = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    UserList(TargetIndex).Stats.Update = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Str_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Send the message to the caster
    If TargetIndex <> CasterIndex Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 46
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
        
    End If
    
    'Send the message to the target
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 47
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Skills.Strengthen
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Strengthen
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    Skill_Strengthen_PCtoPC = 1
    
End Function

Public Function Skill_Protection_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Raises the character's defence
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function
    
    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Protection) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for enough mana to cast
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    
    'Check for a valid target distance
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function
    
    'Reduce the mana
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Pro_Cost)
    
    'Cast on the target
    If UserList(TargetIndex).Counters.ProtectCounter > 0 Then
        If UserList(TargetIndex).Skills.Protect > UserList(CasterIndex).Stats.ModStat(SID.Mag) Then
            
            'Power of what we are casting is weaker then what is already applied
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Server_Message
            ConBuf.Put_Byte 39
            ConBuf.Put_String "protection"
            ConBuf.Put_String UserList(CasterIndex).Name
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Function
            
        End If
    End If
    
    'Display the protection icon (only if it isn't already displayed)
    If UserList(TargetIndex).Skills.Protect = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    End If
    
    'Apply the spell's effects
    UserList(TargetIndex).Counters.ProtectCounter = CurrentTime + Pro_Length
    UserList(TargetIndex).Skills.Protect = UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    UserList(TargetIndex).Stats.Update = 1
    
    'Add the spell exhaustion and display it
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Pro_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Send the message to the caster
    If TargetIndex <> CasterIndex Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 44
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
        
    End If
    
    'Send the message to the target
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 45
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Skills.Protect
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Display the effect on the map
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Protection
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    Skill_Protection_PCtoPC = 1
    
End Function

Public Function Skill_Heal_PCtoPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function
    If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(TargetIndex).Flags.UserLogged = 0 Then Exit Function

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Heal) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < UserList(CasterIndex).Stats.BaseStat(SID.Mag) * Heal_Cost Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Cost)

    'Check for a valid range
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Heal_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Create casting effect
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Cast on the target
    UserList(TargetIndex).Stats.BaseStat(SID.MinHP) = UserList(TargetIndex).Stats.BaseStat(SID.MinHP) + (UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Value)

    'Message to the caster
    If CasterIndex <> TargetIndex Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 42
        ConBuf.Put_String UserList(TargetIndex).Name
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        
        'Face the caster to the target
        UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, UserList(TargetIndex).Pos)
        UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    End If
    
    'Message to the target
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 43
    ConBuf.Put_String UserList(CasterIndex).Name
    ConBuf.Put_Integer UserList(CasterIndex).Stats.BaseStat(SID.Mag)
    Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
    
    'Successfully casted
    Skill_Heal_PCtoPC = 1

End Function

Public Function Skill_Heal_NPCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

    'Check for invalid values
    If NPCList(CasterIndex).Counters.SpellExhaustion > CurrentTime Then Exit Function
    If NPCList(CasterIndex).Counters.ActionDelay > CurrentTime Then Exit Function

    'Check for enough mana
    If NPCList(CasterIndex).BaseStat(SID.MinMAN) < NPCList(CasterIndex).BaseStat(SID.Mag) * Heal_Cost Then Exit Function
    NPCList(CasterIndex).BaseStat(SID.MinMAN) = NPCList(CasterIndex).BaseStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.Mag) * Heal_Cost)

    'Apply spell exhaustion
    NPCList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Heal_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    'Create casting effect
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    'Cast on the target
    NPC_Heal TargetIndex, (NPCList(CasterIndex).ModStat(SID.Mag) * Heal_Value)
    
    'Face the caster to the target
    If CasterIndex <> TargetIndex Then
        NPCList(CasterIndex).Char.Heading = Server_FindDirection(NPCList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
        NPCList(CasterIndex).Char.HeadHeading = NPCList(CasterIndex).Char.Heading
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_Rotate
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        ConBuf.Put_Byte NPCList(CasterIndex).Char.Heading
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If
    
    'Successfully casted
    Skill_Heal_NPCtoNPC = 1

End Function

Public Function Skill_Heal_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Byte

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

    'Check for invalid values
    If UserList(CasterIndex).Flags.SwitchingMaps Then Exit Function
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function
    If UserList(CasterIndex).Flags.UserLogged = 0 Then Exit Function

    'Check if the caster knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Heal) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < UserList(CasterIndex).Stats.BaseStat(SID.Mag) * Heal_Cost Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Cost)

    'Check for a valid range
    If Server_CheckTargetedDistance(CasterIndex) = 0 Then Exit Function

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Heal_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Create casting effect
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Cast on the target
    NPC_Heal TargetIndex, (UserList(CasterIndex).Stats.ModStat(SID.Mag) * Heal_Value)
    
    'Message to the caster
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 42
    ConBuf.Put_String NPCList(TargetIndex).Name
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
    
    'Face the caster to the target
    UserList(CasterIndex).Char.Heading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    UserList(CasterIndex).Char.HeadHeading = UserList(CasterIndex).Char.Heading
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_Rotate
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte UserList(CasterIndex).Char.Heading
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    Skill_Heal_PCtoNPC = 1

End Function

Public Function Skill_IronSkin_PC(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Decreases user attack by 50% to increase defence by 200%
'*****************************************************************

    'Check for invalid values
    If UserIndex = 0 Then Exit Function
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Function

    'Check for the skill in the user posession
    If UserList(UserIndex).KnownSkills(SkID.IronSkin) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Function
    End If

    'Check if still exhausted
    If UserList(UserIndex).Counters.SpellExhaustion > 0 Then Exit Function
    UserList(UserIndex).Counters.SpellExhaustion = CurrentTime + 2000
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map

    'Remove the Iron Skin
    If UserList(UserIndex).Skills.IronSkin = 1 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 0
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map

    Else 'Enable the Iron Skin
        UserList(UserIndex).Skills.IronSkin = 1
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.IronSkin
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconIronSkin
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        Data_Send ToMap, UserIndex, ConBuf.Get_Buffer, UserList(UserIndex).Pos.Map
    End If
    
    UserList(UserIndex).Stats.Update = 1
    
    'Successfully casted
    Skill_IronSkin_PC = 1

End Function

Public Function Skill_SpikeField(ByVal CasterIndex As Integer) As Byte

'*****************************************************************
'Forms a field of spikes around the user
'      |3|
'    |3|4|3|
'  |3|2|2|2|3|
'  |3|2|1|2|3|
'  |4|3|U|3|4|
'    |3|4|3|
'*****************************************************************

Dim aMap As Integer
Dim aX As Integer
Dim aY As Integer
Dim Damage As Integer

    'Check for spell exhaustion
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function

    'Check if the user knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.SpikeField) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * 0.5) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 38
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) = UserList(CasterIndex).Stats.BaseStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.Mag) * 0.5)

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + 3000
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Set the values to shorter variables
    Damage = UserList(CasterIndex).Stats.BaseStat(SID.Mag) + 5
    aMap = UserList(CasterIndex).Pos.Map
    aX = UserList(CasterIndex).Pos.X
    aY = UserList(CasterIndex).Pos.Y

    'Loop through all the tiles, damaging any NPC on them
    'NORTH
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
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.SpikeField
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    
    'Successfully casted
    Skill_SpikeField = 1

End Function

Public Function Skill_Warcry_PC(ByVal CasterIndex As Integer) As Byte

'*****************************************************************
'Lower the stats of all attackable hostiles in range
'*****************************************************************
Dim LoopC As Integer
Dim WarCursePower As Integer

    'Check if still exhausted
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Function
    
    'Check if the user knows the skill
    If UserList(CasterIndex).KnownSkills(SkID.Warcry) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 37
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If

    'Check for enough endurance
    If UserList(CasterIndex).Stats.BaseStat(SID.MinSTA) < Int(UserList(CasterIndex).Stats.ModStat(SID.Str) * Warcry_Cost) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 48
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Function
    End If
    UserList(CasterIndex).Stats.BaseStat(SID.MinSTA) = UserList(CasterIndex).Stats.BaseStat(SID.MinSTA) - Int(UserList(CasterIndex).Stats.ModStat(SID.Str) * Warcry_Cost)

    'Apply spell exhaustion
    UserList(CasterIndex).Counters.SpellExhaustion = CurrentTime + Warcry_Exhaust
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Warcry
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Cast on all attackable hostile NPCs in the PC area
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 49
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

    'Loop through all the alive and active NPCs
    WarCursePower = UserList(CasterIndex).Stats.ModStat(SID.Str)
    For LoopC = 1 To NumNPCs
        If NPCList(LoopC).Flags.NPCActive Then
            If NPCList(LoopC).Flags.NPCAlive Then
                If NPCList(LoopC).Pos.Map = UserList(CasterIndex).Pos.Map Then
                    If NPCList(LoopC).Attackable Then
                        If NPCList(LoopC).Hostile Then
                            If NPCList(LoopC).Skills.WarCurse <= WarCursePower Then
                                If Server_RectDistance(UserList(CasterIndex).Pos.X, UserList(CasterIndex).Pos.Y, NPCList(LoopC).Pos.X, NPCList(LoopC).Pos.Y, MaxServerDistanceX, MaxServerDistanceY) Then

                                    'Tell the users in the screen that the NPC is weaker
                                    ConBuf.Clear
                                    ConBuf.Put_Byte DataCode.Server_Message
                                    ConBuf.Put_Byte 50
                                    ConBuf.Put_String NPCList(LoopC).Name
                                    Data_Send ToNPCArea, LoopC, ConBuf.Get_Buffer
                                    
                                    'Warcurse icon
                                    If NPCList(LoopC).Skills.WarCurse = 0 Then
                                        ConBuf.Clear
                                        ConBuf.Put_Byte DataCode.Server_IconWarCursed
                                        ConBuf.Put_Byte 1
                                        ConBuf.Put_Integer NPCList(LoopC).Char.CharIndex
                                        Data_Send ToMap, 0, ConBuf.Get_Buffer, NPCList(LoopC).Pos.Map
                                    End If
                                    NPCList(LoopC).Skills.WarCurse = WarCursePower
                                    NPCList(LoopC).Counters.WarCurseCounter = CurrentTime + Warcry_Length
                                    
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next LoopC
    
    'Successfully casted
    Skill_Warcry_PC = 1

End Function

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:47)  Decl: 1  Code: 1038  Total: 1039 Lines
':) CommentOnly: 84 (8.1%)  Commented: 3 (0.3%)  Empty: 135 (13%)  Max Logic Depth: 8
