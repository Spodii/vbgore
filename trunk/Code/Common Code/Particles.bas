Attribute VB_Name = "Particles"
Option Explicit
Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    ShiftX As Single            'How much to shift every particle (is reset after move)
    ShiftY As Single
    Gfx As Byte                 'Particle texture used
    Used As Boolean             'If the effect is in use
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Direction As Integer        'Misc variable (depends on the effect)
    Particles() As Particle     'Information on each particle
    Progression As Single       'Misc variable (depends on the effect)
    PartVertex() As TLVERTEX    'Used to point render particles
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindToChar As Integer       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
    BoundToMap As Byte          'If the effect is bound to the map - these kinds of effects should always loop
End Type
Public NumEffects As Byte   'Maximum number of effects at once
Public Effect() As Effect   'List of all the active effects

'Constants With The Order Number For Each Effect
Public Const EffectNum_Fire As Byte = 1             'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Heal As Byte = 3             'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless As Byte = 4            'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection As Byte = 5       ' (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen As Byte = 6       ' which float up and fade out
Public Const EffectNum_Rain As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8 'Template for creating particle effects through equations - a page with some equations can be found here: http://www.vbgore.com/modules.php?name=Forums&file=viewtopic&t=221
Public Const EffectNum_Waterfall As Byte = 9        'Waterfall effect

Function Effect_EquationTemplate_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Byte = 1) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_EquationTemplate_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_EquationTemplate_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
Dim X As Single
Dim Y As Single
Dim R As Single
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    R = (Index / 20) * EXP(Index / Effect(EffectIndex).Progression Mod 3)
    X = R * Cos(Index)
    Y = R * Sin(Index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long
Dim TargetX As Integer  'Bound character's position
Dim TargetY As Integer
Dim TargetI As Integer  'Bound character's index
Dim TargetA As Single   'Angle which the effect will be heading towards the bound character

'Calculate The Time Difference

    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_EquationTemplate_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the shifting values for if the screen moves
                Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + (LastOffsetX - ParticleOffsetX)
                Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + (LastOffsetY - ParticleOffsetY)

                'Set The Particle Information On The Particle Vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Bless_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Bless_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Bless_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

Dim a As Single
Dim X As Single
Dim Y As Single

'Get the positions

    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long
Dim TargetX As Integer  'Bound character's position
Dim TargetY As Integer
Dim TargetI As Integer  'Bound character's index
Dim TargetA As Single   'Angle which the effect will be heading towards the bound character

'Calculate The Time Difference

    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Update position through character binding
    If Effect(EffectIndex).BindToChar Then
        TargetI = Effect(EffectIndex).BindToChar
        TargetX = CharList(TargetI).RealPos.X
        TargetY = CharList(TargetI).RealPos.Y
        TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, TargetX, TargetY) + 180
        Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
        Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed

        'Unbind when character is reached
        If Abs(Effect(EffectIndex).X - TargetX) < 8 Then
            If Abs(Effect(EffectIndex).Y - TargetY) < 8 Then
                Effect(EffectIndex).BindToChar = 0
                Effect(EffectIndex).Progression = 0
            End If
        End If

    End If

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Random clear
            If Int(Rnd * 10000 * ElapsedTime) = 0 Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Bless_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the shifting values for if the screen moves
                Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + (LastOffsetX - ParticleOffsetX)
                Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + (LastOffsetY - ParticleOffsetY)

                'Set The Particle Information On The Particle Vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Fire_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Byte = 180, Optional ByVal Progression As Byte = 1) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Fire_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Fire_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

'Reset the particle

    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

'Calculate The Time Difference

    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Fire_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the shifting values for if the screen moves
                Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + (LastOffsetX - ParticleOffsetX)
                Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + (LastOffsetY - ParticleOffsetY)

                'Set The Particle Information On The Particle Vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_FToDW(F As Single) As Long
Dim Buf As D3DXBuffer
Dim TempVal As Long

    'Converts a single into a long (Float to DWORD)
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, F
    D3DX.BufferGetData Buf, 0, 4, 1, TempVal
    Effect_FToDW = TempVal

End Function

Function Effect_Heal_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Byte = 1) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Heal_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Heal_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.05 + (Rnd * 0.05)

End Sub

Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long
Dim TargetX As Integer  'Bound character's position
Dim TargetY As Integer
Dim TargetI As Integer  'Bound character's index
Dim TargetA As Single   'Angle which the effect will be heading towards the bound character

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Update position through character binding
    If Effect(EffectIndex).BindToChar Then
        TargetI = Effect(EffectIndex).BindToChar
        
        'Check if the character the effect was bound to was removed - if so, end the effect
        If TargetI <= 0 Then Effect(EffectIndex).Progression = 0
        If TargetI > LastChar Then Effect(EffectIndex).Progression = 0
        If CharList(TargetI).Active = 0 Then Effect(EffectIndex).Progression = 0
        
        'If we still have the effect going, move it towards the bound character
        If Effect(EffectIndex).Progression <> 0 Then
            TargetX = CharList(TargetI).RealPos.X
            TargetY = CharList(TargetI).RealPos.Y
            TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, TargetX, TargetY) + 180
            Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
            Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
    
            'Unbind when character is reached
            If Abs(Effect(EffectIndex).X - TargetX) < 8 Then
                If Abs(Effect(EffectIndex).Y - TargetY) < 8 Then
                    Effect(EffectIndex).BindToChar = 0
                    Effect(EffectIndex).Progression = 0
                End If
            End If
            
        End If

    End If

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Random clear
            If Int(Rnd * 100 * ElapsedTime) = 0 Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Heal_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the shifting values for if the screen moves
                Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + (LastOffsetX - ParticleOffsetX)
                Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + (LastOffsetY - ParticleOffsetY)

                'Set The Particle Information On The Particle Vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)

Dim LoopC As Long

'Check If To Kill All Effects

    If KillAll = True Then

        'Loop Through Every Effect
        For LoopC = 1 To NumEffects

            'Stop The Effect
            Effect(LoopC).Used = False

        Next
        
    Else

        'Stop The Selected Effect
        Effect(EffectIndex).Used = False
        
    End If

End Sub

Private Function Effect_NextOpenSlot() As Integer

Dim EffectIndex As Integer

'Find The Next Open Effect Slot

    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While Effect(EffectIndex).Used = True    'Check Next If Effect Is In Use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

End Function

Function Effect_Protection_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Protection_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Protection_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

Dim a As Single
Dim X As Single
Dim Y As Single

'Get the positions

    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long
Dim TargetX As Integer  'Bound character's position
Dim TargetY As Integer
Dim TargetI As Integer  'Bound character's index
Dim TargetA As Single   'Angle which the effect will be heading towards the bound character

'Calculate The Time Difference

    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Update position through character binding
    If Effect(EffectIndex).BindToChar Then
        TargetI = Effect(EffectIndex).BindToChar
        TargetX = CharList(TargetI).RealPos.X
        TargetY = CharList(TargetI).RealPos.Y
        TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, TargetX, TargetY) + 180
        Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
        Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed

        'Unbind when character is reached
        If Abs(Effect(EffectIndex).X - TargetX) < 8 Then
            If Abs(Effect(EffectIndex).Y - TargetY) < 8 Then
                Effect(EffectIndex).BindToChar = 0
                Effect(EffectIndex).Progression = 0
            End If
        End If

    End If

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Random clear
            If Int(Rnd * 10000 * ElapsedTime) = 0 Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Protection_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the shifting values for if the screen moves
                Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + (LastOffsetX - ParticleOffsetX)
                Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + (LastOffsetY - ParticleOffsetY)

                'Set The Particle Information On The Particle Vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Public Sub Effect_Render(ByVal EffectIndex As Integer)

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
    
    'Set the render state to point blitting
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Set the texture
    D3DDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)
    LastTexture = -1

    'Draw all the particles at once
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Snow_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Snow_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (ScreenWidth + 400)), Rnd * (ScreenHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (ScreenWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > ScreenWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > ScreenHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (ScreenWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.8, 0

End Sub

Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (ScreenWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (ScreenHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Apply shift values
            Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + Effect(EffectIndex).ShiftX
            Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + Effect(EffectIndex).ShiftY

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Reset the particle
                Effect_Snow_Reset EffectIndex, LoopC

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

    'Remove shift values
    Effect(EffectIndex).ShiftX = 0
    Effect(EffectIndex).ShiftY = 0

End Sub

Function Effect_Strengthen_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Strengthen_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Strengthen_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(a) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(a) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long
Dim TargetX As Integer  'Bound character's position
Dim TargetY As Integer
Dim TargetI As Integer  'Bound character's index
Dim TargetA As Single   'Angle which the effect will be heading towards the bound character

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Update position through character binding
    If Effect(EffectIndex).BindToChar Then
        TargetI = Effect(EffectIndex).BindToChar
        TargetX = CharList(TargetI).RealPos.X
        TargetY = CharList(TargetI).RealPos.Y
        TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, TargetX, TargetY) + 180
        Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
        Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed

        'Unbind when character is reached
        If Abs(Effect(EffectIndex).X - TargetX) < 8 Then
            If Abs(Effect(EffectIndex).Y - TargetY) < 8 Then
                Effect(EffectIndex).BindToChar = 0
                Effect(EffectIndex).Progression = 0
            End If
        End If

    End If

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Random clear
            If Int(Rnd * 10000 * ElapsedTime) = 0 Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Strengthen_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the shifting values for if the screen moves
                Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + (LastOffsetX - ParticleOffsetX)
                Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + (LastOffsetY - ParticleOffsetY)

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_UpdateAll()
Dim LoopC As Long

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used Then

            'Find out which effect is selected, then update it
            If Effect(LoopC).EffectNum = EffectNum_Fire Then Effect_Fire_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Snow Then Effect_Snow_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Heal Then Effect_Heal_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Bless Then Effect_Bless_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Protection Then Effect_Protection_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Rain Then Effect_Rain_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update LoopC
            
            'Render the effect
            Effect_Render LoopC

        End If

    Next

End Sub

Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Rain_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (ScreenWidth + 400)), Rnd * (ScreenHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > ScreenWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > ScreenHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (ScreenWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.4, 0

End Sub

Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if the particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            
            'Apply shift values
            Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX + Effect(EffectIndex).ShiftX
            Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + Effect(EffectIndex).ShiftY

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (ScreenWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (ScreenHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Reset the particle
                Effect_Rain_Reset EffectIndex, LoopC

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

    'Remove shift values
    Effect(EffectIndex).ShiftX = 0
    Effect(EffectIndex).ShiftY = 0

End Sub

Public Sub Effect_Begin(ByVal EffectIndex As Integer, ByVal X As Single, ByVal Y As Single, ByVal GfxIndex As Byte, ByVal Particles As Byte, Optional ByVal Direction As Single = 180)

'*****************************************************************
'A very simplistic form of initialization for particle effects, should only be used for starting map-based effects
'*****************************************************************
Dim RetNum As Byte

    Select Case EffectIndex
        Case EffectNum_Fire
            RetNum = Effect_Fire_Begin(X, Y, GfxIndex, Particles, Direction, 1)
            Effect(RetNum).BoundToMap = 1
        Case EffectNum_Waterfall
            RetNum = Effect_Waterfall_Begin(X, Y, GfxIndex, Particles)
            Effect(RetNum).BoundToMap = 1
    End Select
    
End Sub

Function Effect_Waterfall_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Waterfall_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Waterfall_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
Dim i As Byte

    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 130), 0, 8 + (Rnd * 6), 0, 0
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 0
    End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0
    
End Sub

Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long
Dim TargetX As Integer  'Bound character's position
Dim TargetY As Integer
Dim TargetI As Integer  'Bound character's index
Dim TargetA As Single   'Angle which the effect will be heading towards the bound character

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Set the shifting values for if the screen moves
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime
    
                'Random clear
                If Int(Rnd * 10000 * ElapsedTime) = 0 Then .sngA = 0
    
                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
    
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, LoopC
    
                Else
                
                    'Set the shifting values for if the screen moves
                    .sngX = .sngX + (LastOffsetX - ParticleOffsetX)
                    .sngY = .sngY + (LastOffsetY - ParticleOffsetY)

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).X = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub
