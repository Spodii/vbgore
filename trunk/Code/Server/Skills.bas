Attribute VB_Name = "Skills"
Option Explicit

Public Sub Skill_Bless(ByVal TargetIndex As Integer, ByVal CasterIndex As Integer, ByVal TargetType As Byte, ByVal CasterType As Byte)

'*****************************************************************
'Increases all of the user's stats by modbless / 3
'*****************************************************************

'Check for invalid values

    If CasterType < 1 Then Exit Sub
    If CasterType > 2 Then Exit Sub
    If TargetType < 1 Then Exit Sub
    If TargetType > 2 Then Exit Sub
    If CasterType = CharType_PC Then
        If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Sub
        If UserList(TargetIndex).Flags.DownloadingMap Then Exit Sub
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'Check if the user knows the skill
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).KnownSkills(SkID.Bless) = 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You do not know that spell!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    End If

    'Check for enough mana
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Stats.ModStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.MaxMAN) * 0.15) Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "Not enough mana."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).ModStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.MaxMAN) * 0.15) Then Exit Sub
    End If

    'Check if still exhausted
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'If skill is already on the target, we have to make sure the spell power is either equal or greater
    If CasterType = CharType_PC Then
        'PC -> PC
        If TargetType = CharType_PC Then
            If UserList(TargetIndex).Counters.BlessCounter > 0 Then
                If UserList(TargetIndex).Skills.Bless > UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "Magical interference trying to cast bless on " & UserList(CasterIndex).Name
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
                End If
            End If
            'PC -> NPC
        ElseIf TargetType = CharType_NPC Then
            If NPCList(TargetIndex).Counters.BlessCounter > 0 Then
                If NPCList(TargetIndex).Skills.Bless > UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "Magical interference trying to cast bless on " & NPCList(TargetIndex).Name
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
                End If
            End If
        End If
    ElseIf CasterType = CharType_NPC Then
        'NPC -> PC
        If TargetType = CharType_PC Then
            If UserList(TargetIndex).Counters.BlessCounter > 0 Then
                If UserList(TargetIndex).Skills.Bless > NPCList(CasterIndex).ModStat(SID.DefensiveMag) Then Exit Sub
            End If
            'NPC -> NPC
        ElseIf TargetType = CharType_NPC Then
            If NPCList(TargetIndex).Counters.BlessCounter > 0 Then
                If NPCList(TargetIndex).Skills.Bless > NPCList(CasterIndex).ModStat(SID.DefensiveMag) Then Exit Sub
            End If
        End If
    End If

    'Add spell exhaustion
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Counters.SpellExhaustion = 3500
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).Counters.SpellExhaustion = 3500
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If

    'Apply the spell's effects
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Stats.ModStat(SID.MinMAN) = UserList(CasterIndex).Stats.ModStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.MaxMAN) * 0.15)

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        If TargetType = CharType_PC Then
            ConBuf.Put_String "You blessed " & UserList(TargetIndex).Name & "."
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_String "You blessed " & NPCList(TargetIndex).Name & "."
        End If
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String UserList(CasterIndex).Name & " blessed you with a power of " & UserList(CasterIndex).Skills.Bless & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

            UserList(TargetIndex).Counters.BlessCounter = 300000
            UserList(TargetIndex).Skills.Bless = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag)
        ElseIf TargetType = CharType_NPC Then
            NPCList(TargetIndex).Counters.BlessCounter = 300000
            NPCList(TargetIndex).Skills.Bless = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag)
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Bless
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).ModStat(SID.MinMAN) = NPCList(CasterIndex).ModStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.MaxMAN) * 0.15)

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String NPCList(CasterIndex).Name & " blessed you with a power of " & NPCList(CasterIndex).Skills.Bless & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

            UserList(TargetIndex).Counters.BlessCounter = 300000
            UserList(TargetIndex).Skills.Bless = NPCList(CasterIndex).BaseStat(SID.DefensiveMag)
        ElseIf TargetType = CharType_NPC Then
            NPCList(TargetIndex).Counters.BlessCounter = 300000
            NPCList(TargetIndex).Skills.Bless = NPCList(CasterIndex).BaseStat(SID.DefensiveMag)
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Bless
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconBlessed
        ConBuf.Put_Byte 1
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    End If

End Sub

Public Sub Skill_Heal(ByVal TargetIndex As Integer, ByVal CasterIndex As Integer, ByVal TargetType As Byte, ByVal CasterType As Byte)

'*****************************************************************
'Heal the target at the cost of mana
'*****************************************************************

'Check for invalid values

    If CasterType < 1 Then Exit Sub
    If CasterType > 2 Then Exit Sub
    If TargetType < 1 Then Exit Sub
    If TargetType > 2 Then Exit Sub
    If CasterType = CharType_PC Then
        If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Sub
        If UserList(TargetIndex).Flags.DownloadingMap Then Exit Sub
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'Check if the caster knows the skill (NPCs that dont know heal shouldn't even be calling this)
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).KnownSkills(SkID.Heal) = 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You do not know that skill!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    End If

    'Check for enough mana
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Stats.ModStat(SID.MinMAN) < UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag) * 0.5 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "Not enough mana."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).ModStat(SID.MinMAN) < NPCList(CasterIndex).BaseStat(SID.DefensiveMag) * 0.5 Then Exit Sub
    End If

    'Apply spell exhaustion
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Counters.SpellExhaustion = 1000
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).Counters.SpellExhaustion = 1000
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
        ConBuf.Put_Byte 1
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If

    'Create casting effect
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Heal
    If CasterType = CharType_PC Then
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    ElseIf CasterType = CharType_NPC Then
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If

    'Reduce the caster's mana
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Stats.ModStat(SID.MinMAN) = UserList(CasterIndex).Stats.ModStat(SID.MinMAN) - (UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) * 0.5)
    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).ModStat(SID.MinMAN) = NPCList(CasterIndex).ModStat(SID.MinMAN) - (NPCList(CasterIndex).ModStat(SID.DefensiveMag) * 0.5)
    End If

    'Cast on the target
    If TargetType = CharType_PC Then
        UserList(TargetIndex).Stats.ModStat(SID.MinHP) = UserList(TargetIndex).Stats.ModStat(SID.MinHP) + UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag)
    ElseIf TargetType = CharType_PC Then
        NPCList(TargetIndex).ModStat(SID.MinHP) = NPCList(TargetIndex).ModStat(SID.MinHP) + NPCList(CasterIndex).ModStat(SID.DefensiveMag)
    End If

    'Say the information
    If TargetIndex = CasterIndex Then
        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You healed yourself."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        End If
    Else
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        If TargetType = CharType_PC Then
            ConBuf.Put_String "You healed " & UserList(TargetIndex).Name & "."
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_String "You healed " & NPCList(TargetIndex).Name & "."
        End If
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String UserList(CasterIndex).Name & " healed you " & UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag) & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer
        End If

    End If

End Sub

Public Sub Skill_IronSkin(ByVal UserIndex As Integer)

'*****************************************************************
'Decreases user attack by 50% to increase defence by 200%
'*****************************************************************

'Check for invalid values

    If UserIndex = 0 Then Exit Sub
    If UserList(UserIndex).Flags.SwitchingMaps Then Exit Sub
    If UserList(UserIndex).Flags.DownloadingMap Then Exit Sub

    'Check for the skill in the user posession
    If UserList(UserIndex).KnownSkills(SkID.IronSkin) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You do not know that skill!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check if still exhausted
    If UserList(UserIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    UserList(UserIndex).Counters.SpellExhaustion = 2000
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

End Sub

Public Sub Skill_Protection(ByVal TargetIndex As Integer, ByVal CasterIndex As Integer, ByVal TargetType As Byte, ByVal CasterType As Byte)

'*****************************************************************
'Increase the user's armor value by modprotect / 5
'*****************************************************************

'Check for invalid values

    If CasterType < 1 Then Exit Sub
    If CasterType > 2 Then Exit Sub
    If TargetType < 1 Then Exit Sub
    If TargetType > 2 Then Exit Sub
    If CasterType = CharType_PC Then
        If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Sub
        If UserList(TargetIndex).Flags.DownloadingMap Then Exit Sub
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'Check if the user knows the skill
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).KnownSkills(SkID.Protection) = 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You do not know that spell!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    End If

    'Check for enough mana
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Stats.ModStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.MaxMAN) * 0.15) Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "Not enough mana."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).ModStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.MaxMAN) * 0.15) Then Exit Sub
    End If

    'Check if still exhausted
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'If skill is already on the target, we have to make sure the spell power is either equal or greater
    If CasterType = CharType_PC Then
        'PC -> PC
        If TargetType = CharType_PC Then
            If UserList(TargetIndex).Counters.ProtectCounter > 0 Then
                If UserList(TargetIndex).Skills.Protect > UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "Magical interference trying to cast protection on " & UserList(CasterIndex).Name
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
                End If
            End If
            'PC -> NPC
        ElseIf TargetType = CharType_NPC Then
            If NPCList(TargetIndex).Counters.ProtectCounter > 0 Then
                If NPCList(TargetIndex).Skills.Protect > UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "Magical interference trying to cast protection on " & NPCList(TargetIndex).Name
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
                End If
            End If
        End If
    ElseIf CasterType = CharType_NPC Then
        'NPC -> PC
        If TargetType = CharType_PC Then
            If UserList(TargetIndex).Counters.ProtectCounter > 0 Then
                If UserList(TargetIndex).Skills.Protect > NPCList(CasterIndex).ModStat(SID.DefensiveMag) Then Exit Sub
            End If
            'NPC -> NPC
        ElseIf TargetType = CharType_NPC Then
            If NPCList(TargetIndex).Counters.ProtectCounter > 0 Then
                If NPCList(TargetIndex).Skills.Protect > NPCList(CasterIndex).ModStat(SID.DefensiveMag) Then Exit Sub
            End If
        End If
    End If

    'Add spell exhaustion
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Counters.SpellExhaustion = 3500
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).Counters.SpellExhaustion = 3500
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If

    'Apply the spell's effects
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Stats.ModStat(SID.MinMAN) = UserList(CasterIndex).Stats.ModStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.MaxMAN) * 0.15)

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        If TargetType = CharType_PC Then
            ConBuf.Put_String "You protected " & UserList(TargetIndex).Name & "."
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_String "You protected " & NPCList(TargetIndex).Name & "."
        End If
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String UserList(CasterIndex).Name & " protected you with a power of " & UserList(CasterIndex).Skills.Protect & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

            UserList(TargetIndex).Counters.ProtectCounter = 300000
            UserList(TargetIndex).Skills.Protect = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag)
        ElseIf TargetType = CharType_NPC Then
            NPCList(TargetIndex).Counters.ProtectCounter = 300000
            NPCList(TargetIndex).Skills.Protect = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag)
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Protection
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).ModStat(SID.MinMAN) = NPCList(CasterIndex).ModStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.MaxMAN) * 0.15)

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String NPCList(CasterIndex).Name & " protected you with a power of " & NPCList(CasterIndex).Skills.Protect & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

            UserList(TargetIndex).Counters.ProtectCounter = 300000
            UserList(TargetIndex).Skills.Protect = NPCList(CasterIndex).BaseStat(SID.DefensiveMag)
        ElseIf TargetType = CharType_NPC Then
            NPCList(TargetIndex).Counters.ProtectCounter = 300000
            NPCList(TargetIndex).Skills.Protect = NPCList(CasterIndex).BaseStat(SID.DefensiveMag)
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Protection
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconProtected
        ConBuf.Put_Byte 1
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    End If

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

Dim aMap As Integer
Dim aX As Integer
Dim aY As Integer
Dim Damage As Integer

'Check if the user knows the skill

    If UserList(CasterIndex).KnownSkills(SkID.SpikeField) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You do not know that spell!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check for enough mana
    If UserList(CasterIndex).Stats.ModStat(SID.MinMAN) < 1 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "Not enough mana."
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check if still exhausted
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    UserList(CasterIndex).Counters.SpellExhaustion = 3000
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Set the values to shorter variables
    Damage = UserList(CasterIndex).Stats.BaseStat(SID.OffensiveMag) + 5
    aMap = UserList(CasterIndex).Pos.Map
    aX = UserList(CasterIndex).Pos.x
    aY = UserList(CasterIndex).Pos.Y

    'Loop through all the tiles, damaging any NPC on them
    'NORTH
    If UserList(CasterIndex).Char.HeadHeading = NORTH Then
        If MapData(aMap, aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 1).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 2, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY).NPCIndex, CasterIndex, Damage * 0.25

        If MapData(aMap, aX - 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 1).NPCIndex, CasterIndex, Damage
        If MapData(aMap, aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 2).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 1, aY - 3).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY - 3).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 1, aY - 3).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 3).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX, aY - 4).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 4).NPCIndex, CasterIndex, Damage * 0.25

        'EAST
    ElseIf UserList(CasterIndex).Char.HeadHeading = EAST Then
        If MapData(aMap, aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 2).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 2).NPCIndex, CasterIndex, Damage * 0.25

        If MapData(aMap, aX + 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY).NPCIndex, CasterIndex, Damage
        If MapData(aMap, aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX + 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX + 3, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 3, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 3, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 3, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 3, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 3, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX + 4, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 4, aY).NPCIndex, CasterIndex, Damage * 0.25

        'SOUTH
    ElseIf UserList(CasterIndex).Char.HeadHeading = SOUTH Then
        If MapData(aMap, aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 1).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 2, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY).NPCIndex, CasterIndex, Damage * 0.25

        If MapData(aMap, aX - 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 1).NPCIndex, CasterIndex, Damage
        If MapData(aMap, aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 2).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX + 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 1, aY + 3).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 3).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 3).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 1, aY + 3).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 3).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX, aY + 4).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 4).NPCIndex, CasterIndex, Damage * 0.25

        'WEST
    ElseIf UserList(CasterIndex).Char.HeadHeading = WEST Then
        If MapData(aMap, aX + 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX + 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX + 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX + 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 2).NPCIndex, CasterIndex, Damage * 0.25
        If MapData(aMap, aX, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX, aY + 2).NPCIndex, CasterIndex, Damage * 0.25

        If MapData(aMap, aX - 1, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 1, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY - 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX - 1, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY).NPCIndex, CasterIndex, Damage
        If MapData(aMap, aX - 1, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 1).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX - 1, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 1, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 2, aY - 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY - 2).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 2, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 2, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY).NPCIndex, CasterIndex, Damage * 0.5
        If MapData(aMap, aX - 2, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY + 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 2, aY + 2).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 2, aY + 2).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 3, aY - 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 3, aY - 1).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 3, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 3, aY).NPCIndex, CasterIndex, Damage * 0.333
        If MapData(aMap, aX - 3, aY + 1).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 3, aY + 1).NPCIndex, CasterIndex, Damage * 0.333

        If MapData(aMap, aX - 4, aY).NPCIndex > 0 Then NPC_Damage MapData(aMap, aX - 4, aY).NPCIndex, CasterIndex, Damage * 0.25

    End If

    'Display the user casting it on other people's screens
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.SpikeField
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

End Sub

Public Sub Skill_Strengthen(ByVal TargetIndex As Integer, ByVal CasterIndex As Integer, ByVal TargetType As Byte, ByVal CasterType As Byte)

'*****************************************************************
'Increase the user's armor value by modstrengthen / 5
'*****************************************************************

'Check for invalid values

    If CasterType < 1 Then Exit Sub
    If CasterType > 2 Then Exit Sub
    If TargetType < 1 Then Exit Sub
    If TargetType > 2 Then Exit Sub
    If CasterType = CharType_PC Then
        If UserList(TargetIndex).Flags.SwitchingMaps Then Exit Sub
        If UserList(TargetIndex).Flags.DownloadingMap Then Exit Sub
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'Check if the user knows the skill
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).KnownSkills(SkID.Strengthen) = 0 Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "You do not know that spell!"
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    End If

    'Check for enough mana
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Stats.ModStat(SID.MinMAN) < Int(UserList(CasterIndex).Stats.ModStat(SID.MaxMAN) * 0.15) Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String "Not enough mana."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
            Exit Sub
        End If
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).ModStat(SID.MinMAN) < Int(NPCList(CasterIndex).ModStat(SID.MaxMAN) * 0.15) Then Exit Sub
    End If

    'Check if still exhausted
    If CasterType = CharType_PC Then
        If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    ElseIf CasterType = CharType_NPC Then
        If NPCList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    End If

    'If skill is already on the target, we have to make sure the spell power is either equal or greater
    If CasterType = CharType_PC Then
        'PC -> PC
        If TargetType = CharType_PC Then
            If UserList(TargetIndex).Counters.StrengthenCounter > 0 Then
                If UserList(TargetIndex).Skills.Strengthen > UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "Magical interference trying to cast strengthen on " & UserList(CasterIndex).Name
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
                End If
            End If
            'PC -> NPC
        ElseIf TargetType = CharType_NPC Then
            If NPCList(TargetIndex).Counters.StrengthenCounter > 0 Then
                If NPCList(TargetIndex).Skills.Strengthen > UserList(CasterIndex).Stats.ModStat(SID.DefensiveMag) Then
                    ConBuf.Clear
                    ConBuf.Put_Byte DataCode.Comm_Talk
                    ConBuf.Put_String "Magical interference trying to cast strengthen on " & NPCList(TargetIndex).Name
                    ConBuf.Put_Byte DataCode.Comm_FontType_Info
                    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
                End If
            End If
        End If
    ElseIf CasterType = CharType_NPC Then
        'NPC -> PC
        If TargetType = CharType_PC Then
            If UserList(TargetIndex).Counters.StrengthenCounter > 0 Then
                If UserList(TargetIndex).Skills.Strengthen > NPCList(CasterIndex).ModStat(SID.DefensiveMag) Then Exit Sub
            End If
            'NPC -> NPC
        ElseIf TargetType = CharType_NPC Then
            If NPCList(TargetIndex).Counters.StrengthenCounter > 0 Then
                If NPCList(TargetIndex).Skills.Strengthen > NPCList(CasterIndex).ModStat(SID.DefensiveMag) Then Exit Sub
            End If
        End If
    End If

    'Add spell exhaustion
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Server_IconSpellExhaustion
    ConBuf.Put_Byte 1
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Counters.SpellExhaustion = 3500
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map
    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).Counters.SpellExhaustion = 3500
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map
    End If

    'Apply the spell's effects
    If CasterType = CharType_PC Then
        UserList(CasterIndex).Stats.ModStat(SID.MinMAN) = UserList(CasterIndex).Stats.ModStat(SID.MinMAN) - Int(UserList(CasterIndex).Stats.ModStat(SID.MaxMAN) * 0.15)

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        If TargetType = CharType_PC Then
            ConBuf.Put_String "You strengthened " & UserList(TargetIndex).Name & "."
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_String "You strengthened " & NPCList(TargetIndex).Name & "."
        End If
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String UserList(CasterIndex).Name & " strengthened you with a power of " & UserList(CasterIndex).Skills.Strengthen & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

            UserList(TargetIndex).Counters.StrengthenCounter = 300000
            UserList(TargetIndex).Skills.Strengthen = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag)
        ElseIf TargetType = CharType_NPC Then
            NPCList(TargetIndex).Counters.StrengthenCounter = 300000
            NPCList(TargetIndex).Skills.Strengthen = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag)
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Strengthen
        ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    ElseIf CasterType = CharType_NPC Then
        NPCList(CasterIndex).ModStat(SID.MinMAN) = NPCList(CasterIndex).ModStat(SID.MinMAN) - Int(NPCList(CasterIndex).ModStat(SID.MaxMAN) * 0.15)

        If TargetType = CharType_PC Then
            ConBuf.Clear
            ConBuf.Put_Byte DataCode.Comm_Talk
            ConBuf.Put_String NPCList(CasterIndex).Name & " strengthened you with a power of " & NPCList(CasterIndex).Skills.Strengthen & "."
            ConBuf.Put_Byte DataCode.Comm_FontType_Info
            Data_Send ToIndex, TargetIndex, ConBuf.Get_Buffer

            UserList(TargetIndex).Counters.StrengthenCounter = 300000
            UserList(TargetIndex).Skills.Strengthen = NPCList(CasterIndex).BaseStat(SID.DefensiveMag)
        ElseIf TargetType = CharType_NPC Then
            NPCList(TargetIndex).Counters.StrengthenCounter = 300000
            NPCList(TargetIndex).Skills.Strengthen = NPCList(CasterIndex).BaseStat(SID.DefensiveMag)
        End If

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Strengthen
        ConBuf.Put_Integer NPCList(CasterIndex).Char.CharIndex
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Server_IconStrengthened
        ConBuf.Put_Byte 1
        If TargetType = CharType_PC Then
            ConBuf.Put_Integer UserList(TargetIndex).Char.CharIndex
        ElseIf TargetType = CharType_NPC Then
            ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
        End If
        Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, NPCList(CasterIndex).Pos.Map

    End If

End Sub

Public Sub Skill_Warcry(ByVal CasterIndex As Integer)

'Cry out and curse all enemies in the screen that are hostile and attackable

Dim LoopC As Integer
Dim WarCursePower As Integer

'Check if the user knows the skill

    If UserList(CasterIndex).KnownSkills(SkID.Warcry) = 0 Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "You do not know that skill!"
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check for enough endurance
    If UserList(CasterIndex).Stats.ModStat(SID.MinSTA) < 1 Then '(3 * UserList(CasterIndex).Stats.ModWarcry) Then
        ConBuf.Clear
        ConBuf.Put_Byte DataCode.Comm_Talk
        ConBuf.Put_String "Not enough stamina."
        ConBuf.Put_Byte DataCode.Comm_FontType_Info
        Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
        Exit Sub
    End If

    'Check if still exhausted
    If UserList(CasterIndex).Counters.SpellExhaustion > 0 Then Exit Sub
    UserList(CasterIndex).Counters.SpellExhaustion = 1000
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.User_CastSkill
    ConBuf.Put_Byte SkID.Warcry
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    Data_Send ToMap, CasterIndex, ConBuf.Get_Buffer, UserList(CasterIndex).Pos.Map

    'Cast on all NPCs in the PC area
    ConBuf.Clear
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "You warcry!"
    ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer

    For LoopC = 1 To NumNPCs
        If NPCList(LoopC).Flags.NPCAlive Then
            If NPCList(LoopC).Flags.NPCAlive Then
                If NPCList(LoopC).Pos.Map = UserList(CasterIndex).Pos.Map Then
                    If NPCList(LoopC).Attackable Then
                        WarCursePower = UserList(CasterIndex).Stats.BaseStat(SID.DefensiveMag) - (NPCList(LoopC).ModStat(SID.Immunity) * 0.5)
                        If NPCList(LoopC).Skills.WarCurse <= WarCursePower Then
                            If Server_Distance(UserList(CasterIndex).Pos.x, UserList(CasterIndex).Pos.Y, NPCList(LoopC).Pos.x, NPCList(LoopC).Pos.Y) <= Max_Server_Distance Then
                                NPCList(LoopC).Skills.WarCurse = WarCursePower
                                NPCList(LoopC).Counters.WarCurseCounter = 30000 '30 seconds
                                ConBuf.Clear
                                ConBuf.Put_Byte DataCode.Comm_Talk
                                ConBuf.Put_String NPCList(LoopC).Name & " appears weaker."
                                ConBuf.Put_Byte DataCode.Comm_FontType_Info
                                Data_Send ToNPCArea, LoopC, ConBuf.Get_Buffer
                                ConBuf.Clear
                                ConBuf.Put_Byte DataCode.Server_IconWarCursed
                                ConBuf.Put_Byte 1
                                ConBuf.Put_Integer NPCList(LoopC).Char.CharIndex
                                Data_Send ToMap, 0, ConBuf.Get_Buffer, NPCList(LoopC).Pos.Map
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next LoopC

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Sep-05 23:47)  Decl: 1  Code: 1038  Total: 1039 Lines
':) CommentOnly: 84 (8.1%)  Commented: 3 (0.3%)  Empty: 135 (13%)  Max Logic Depth: 8
