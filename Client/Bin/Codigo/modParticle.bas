Attribute VB_Name = "Mod_DX8_Particle"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' / Módulo extraído de vbGore y modificado por Dunkansdk (Emanuel Matías D'Urso)

' - BOSKORCHA AO

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Type Effect

    X       As Single                       'Location of effect
    Y       As Single
    GoToX   As Single                       'Location to move to
    GoToY   As Single
    KillWhenAtTarget    As Boolean          'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost  As Boolean          'Kill the effect if the target is lost (sets progression = 0)
    Gfx         As Byte                     'Particle texture used
    Used        As Boolean                  'If the effect is in use
    EffectNum   As Byte                     'What number of effect that is used
    Modifier    As Integer                  'Misc variable (depends on the effect)
    FloatSize   As Long                     'The size of the particles
    Direction   As Integer                  'Misc variable (depends on the effect)
    Particles() As Particle                 'Information on each particle
    Progression As Single                   'Progression state, best to design where 0 = effect ends
    looping As Boolean
    PartVertex()    As TLVERTEX             'Used to point render particles
    PreviousFrame   As Long                 'Tick time of the last frame
    PreviousMove As Long
    NowMove As Long
    ParticleCount   As Integer              'Number of particles total
    ParticlesLeft   As Integer              'Number of particles left - only for non-repetitive effects
    BindToChar      As Integer              'Setting this value will bind the effect to move towards the character
    BindSpeed       As Single               'How fast the effect moves towards the character
    BoundToMap      As Byte                 'If the effect is bound to the map or not (used only by the map editor)
    TAngle          As Single
    
End Type

Public NumEffects   As Byte     'Maximum number of effects at once
Public Effect()     As Effect   'List of all the active effects

'Constants With The Order Number For Each Effect
Public Const EffectNum_Fire                 As Byte = 1 'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow                 As Byte = 2 'Snow that covers the screen - weather effect
Public Const EffectNum_Rain                 As Byte = 3 'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_Waterfall            As Byte = 4 'Waterfall effect
Public Const EffectNum_Smoke                As Byte = 5 'Smoke effect
Public Const EffectNum_Water                As Byte = 6 'Water effect
Public Const EffectNum_Teleport             As Byte = 7 'Teleport
Public Const EffectNum_Curse_Apocalipsis    As Byte = 8 'Curve directs the user
Public Const EffectNum_End_Apocalipsis      As Byte = 9 'Explsión after spell
Public Const EffectNum_Curse_Inmovilizar    As Byte = 10 'Curve directs the user and inmo.
Public Const EffectNum_Lissajous            As Byte = 11 'Curva Lissajous, end inmo.
Public Const EffectNum_Meditation           As Byte = 12


Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)

Function Effect_End_Apocalipsis_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_End_Apocalipsis_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_End_Apocalipsis_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_End_Apocalipsis  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(18)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_End_Apocalipsis_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_End_Apocalipsis_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_End_Apocalipsis_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
Dim rG As Single
    
    rG = (Rnd * 0.5)
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    R = (Index / 250) * ((Effect(EffectIndex).Progression / 10) ^ 2)
    X = R * Round(Cos(Index), 0) + (Index * Rnd * 0.07) * Sgn(Cos(Index))
    Y = R * Round(Sin(Index), 0) + (Index * Rnd * 0.07) * Sgn(Sin(Index))
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    If RandomNumber(1, 2) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2 + rG, 0.2, 0.9, 0.2 + (Rnd * 0.2)
    Else
        Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2 + rG, 1, 0.9, 0.2 + (Rnd * 0.2)
    End If

End Sub

Private Sub Effect_End_Apocalipsis_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_End_Apocalipsis_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.2 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 50 Then

                    'Reset the particle
                    Effect_End_Apocalipsis_Reset EffectIndex, LoopC

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Fire_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
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
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Fire_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Private Function Effect_FToDW(F As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
'*****************************************************************
Dim buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, F
    DirectD3D8.BufferGetData buf, 0, 4, 1, Effect_FToDW

End Function

Function Effect_Curse_Apocalipsis_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Curse_Apocalipsis_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Curse_Apocalipsis     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles      'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
    Effect(EffectIndex).TAngle = 0
    
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        'Effect_Curse_Apocalipsis_Reset EffectIndex, LoopC
    Next LoopC
    
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Curse_Apocalipsis_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    'Static TargetA As Single
    Dim Co As Single
    Dim sI As Single
    'Calculate the angle
    
    If Effect(EffectIndex).TAngle = 0 And Effect(EffectIndex).GoToX <> -30000 Then Effect(EffectIndex).TAngle = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
    sI = Sin(Effect(EffectIndex).TAngle * DegreeToRadian)
    Co = Cos(Effect(EffectIndex).TAngle * DegreeToRadian)
    
    'Reset the particle
    If RandomNumber(1, 2) = 2 Then
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * 20, Si * Sin(Effect(EffectIndex).Progression * 3) * 20, 0, 0
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * 25, Effect(EffectIndex).Y + sI * Sin(Effect(EffectIndex).Progression) * 25, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2 + (Rnd * 0.5), 1, 0.5 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    Else
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * -20, Si * Sin(Effect(EffectIndex).Progression * 3) * -20, 0, 0
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * -25, Effect(EffectIndex).Y + sI * Sin(Effect(EffectIndex).Progression) * -25, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2 + (Rnd * 0.5), 0.2, 0.7 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    End If
    

End Sub

Private Sub Effect_Curse_Apocalipsis_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 3) = 3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Curse_Apocalipsis_Reset EffectIndex, LoopC
                    
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
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'*****************************************************************
'Kills (stops) a single effect or all effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Kill
'*****************************************************************
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
'*****************************************************************
'Finds the next open effects index
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot
'*****************************************************************
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

    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000

End Function

Public Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
'***************************************************
'Update an effect's position if the screen has moved
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateOffset
'***************************************************
    If UserCharIndex <> 0 Then
        If EffectIndex <> charlist(UserCharIndex).Aura Then
            Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX) '- Sgn(CharList(UserCharIndex).ScrollDirectionX) * 0.5 * (LastOffsetX - ParticleOffsetX)
            Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY) '- Sgn(CharList(UserCharIndex).ScrollDirectionY) * 0.5 * (LastOffsetY - ParticleOffsetY)
        End If
    End If
    
End Sub

Public Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)
 
'***************************************************
'Updates the binding of a particle effect to a target, if
'the effect is bound to a character
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding
'***************************************************
Dim TargetI As Integer
Dim TargetA As Single
Dim LoopC As Integer
Dim RetNum As Integer

    Effect(EffectIndex).NowMove = timeGetTime
    
    If Effect(EffectIndex).PreviousMove + 10 < Effect(EffectIndex).NowMove Then
        Effect(EffectIndex).PreviousMove = timeGetTime

        'Update position through character binding
        If Effect(EffectIndex).BindToChar > 0 Then
     
            'Store the character index
            TargetI = Effect(EffectIndex).BindToChar
     
            'Check for a valid binding index
            If TargetI > LastChar Then
                Effect(EffectIndex).BindToChar = 0
                If Effect(EffectIndex).KillWhenTargetLost Then
                    Effect(EffectIndex).Progression = 0
                    Exit Sub
                End If
            ElseIf charlist(TargetI).Active = 0 Then
                Effect(EffectIndex).BindToChar = 0
                If Effect(EffectIndex).KillWhenTargetLost Then
                    Effect(EffectIndex).Progression = 0
                    Exit Sub
                End If
            Else
     
                'Calculate the X and Y positions
                Effect(EffectIndex).GoToX = Engine_TPtoSPX(charlist(Effect(EffectIndex).BindToChar).Pos.X) + 8
                Effect(EffectIndex).GoToY = Engine_TPtoSPY(charlist(Effect(EffectIndex).BindToChar).Pos.Y) + 8
     
            End If
     
        End If
     
        'Move to the new position if needed
        If Effect(EffectIndex).GoToX > -30000 Or Effect(EffectIndex).GoToY > -30000 Then
            If Effect(EffectIndex).GoToX <> Effect(EffectIndex).X Or Effect(EffectIndex).GoToY <> Effect(EffectIndex).Y Then
     
                'Calculate the angle
                TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
                'Update the position of the effect
                Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
                Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
                
                'Check if the effect is close enough to the target to just stick it at the target
                If Effect(EffectIndex).GoToX > -30000 Then
                    If Abs(Effect(EffectIndex).X - Effect(EffectIndex).GoToX) < 10 Then Effect(EffectIndex).X = Effect(EffectIndex).GoToX
                End If
                If Effect(EffectIndex).GoToY > -30000 Then
                    If Abs(Effect(EffectIndex).Y - Effect(EffectIndex).GoToY) < 10 Then Effect(EffectIndex).Y = Effect(EffectIndex).GoToY
                End If
     
                'Check if the position of the effect is equal to that of the target
                If Effect(EffectIndex).X = Effect(EffectIndex).GoToX Then
                    If Effect(EffectIndex).Y = Effect(EffectIndex).GoToY Then
     
                        'For some effects, if the position is reached, we want to end the effect
                        If Effect(EffectIndex).KillWhenAtTarget Then
                            
                            'Explode on impact
                            If Effect(EffectIndex).Progression <> 0 Then
                            
                                If Effect(EffectIndex).EffectNum = EffectNum_Curse_Apocalipsis Then
                                    RetNum = Effect_End_Apocalipsis_Begin(Effect(EffectIndex).X, Effect(EffectIndex).Y, 2, 300, 40)  'Apocalipsis
                                ElseIf Effect(EffectIndex).EffectNum = EffectNum_Curse_Inmovilizar Then
                                    RetNum = Effect_Lissajous_Begin(Effect(EffectIndex).X, Effect(EffectIndex).Y, 1, 250, 1)  'Inmovilizar
                            '    ElseIf Effect(EffectIndex).EffectNum = EffectNum_Necro Then
                            '        Effect(EffectIndex).TAngle = 0
                            '        RetNum = Effect_Green_Begin(Effect(EffectIndex).x, Effect(EffectIndex).y, 2, 300, 40)  'Apocalipsis
                            '    ElseIf Effect(EffectIndex).EffectNum = EffectNum_Curse Then
                            '        Effect(EffectIndex).TAngle = 0
                            '        RetNum = Effect_Lissajous_Begin(Effect(EffectIndex).x, Effect(EffectIndex).y, 1, 250, 1)  'Inmovilizar
                                End If
                            
                                If RetNum > 0 Then
                                    Effect(RetNum).BindToChar = Effect(EffectIndex).BindToChar
                                    Effect(RetNum).BindSpeed = 10
                                End If
                            
                            End If
                            
                            Effect(EffectIndex).BindToChar = 0
                            Effect(EffectIndex).Progression = 0
                            Effect(EffectIndex).GoToX = Effect(EffectIndex).X
                            Effect(EffectIndex).GoToY = Effect(EffectIndex).Y
                        End If
                        
                        Exit Sub    'The effect is at the right position, don't update
     
                    End If
                End If
     
            End If
        End If
    End If
End Sub

Public Sub Effect_Render(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Render
'*****************************************************************

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    DirectDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
        
    'Set the render state to point blitting
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the last texture to a random number to force the engine to reload the texture
    LastTexture = -65489

    'Set the texture
    DirectDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)

    'Draw all the particles at once
    DirectDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))
                    
    'Reset the render state back to normal
    If SetRenderStates Then DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin
'*****************************************************************
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
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Snow_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset
'*****************************************************************

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
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.5, 0

End Sub

Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update
'*****************************************************************
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

End Sub

Sub Effect_UpdateAll()
'*****************************************************************
'Updates all of the effects and renders them
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
'*****************************************************************
Dim LoopC As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used Then
        
            'Update the effect position if the screen has moved
            Effect_UpdateOffset LoopC
        
            'Update the effect position if it is binded
            Effect_UpdateBinding LoopC

            'Find out which effect is selected, then update it
            If Effect(LoopC).EffectNum = EffectNum_Fire Then Effect_Fire_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Snow Then Effect_Snow_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Rain Then Effect_Rain_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Smoke Then Effect_Smoke_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Water Then Effect_Water_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Teleport Then Effect_Teleport_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Curse_Apocalipsis Then Effect_Curse_Apocalipsis_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_End_Apocalipsis Then Effect_End_Apocalipsis_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Curse_Inmovilizar Then Effect_Curse_Inmovilizar_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Lissajous Then Effect_Lissajous_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Meditation Then Effect_Meditation_Update LoopC
            
            'Render the effect
            Effect_Render LoopC, False

        End If

    Next
    
    'Set the render state back for normal rendering
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin
'*****************************************************************
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
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Rain_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset
'*****************************************************************

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
    Effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.8, 1, 0.4, Rnd * 0.068

End Sub

Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update
'*****************************************************************
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

End Sub

Public Sub Effect_Begin(ByVal EffectIndex As Integer, ByVal X As Single, ByVal Y As Single, ByVal GfxIndex As Byte, ByVal Particles As Integer, Optional ByVal Direction As Single = 180, Optional ByVal BindToMap As Boolean = False)
'*****************************************************************
'A very simplistic form of initialization for particle effects
'Should only be used for starting map-based effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Begin
'*****************************************************************
Dim RetNum As Byte

    X = Engine_TPtoSPX(X)
    Y = Engine_TPtoSPY(Y)

    Select Case EffectIndex
        Case EffectNum_Fire ' 1
            RetNum = Effect_Fire_Begin(X, Y, GfxIndex, Particles, Direction, 1)
        Case EffectNum_Waterfall ' 9
            'RetNum = Effect_Waterfall_Begin(x, y, 2, 100)
        Case EffectNum_Smoke ' 11
            RetNum = Effect_Smoke_Begin(X, Y, 1, 100)
        Case EffectNum_Water ' 12
            RetNum = Effect_Water_Begin(X, Y, 1, 100)
        Case EffectNum_Teleport ' 13
            RetNum = Effect_Teleport_Begin(X - 1, Y - 1, 1, 300, 48)
    End Select
    
    'Bind the effect to the map if needed
    If BindToMap Then Effect(RetNum).BoundToMap = 1
    
End Sub

Function Effect_Waterfall_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
'*****************************************************************
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
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Waterfall_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
'*****************************************************************

  
    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 90), Effect(EffectIndex).Y + (Rnd * 120), 0, 2 + (Rnd * 6), 0, 2
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 90), Effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 2
    End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.5 + (Rnd * 0.4), 0
    
End Sub

Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
    
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, LoopC
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).X = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub

Function Effect_Smoke_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Smoke_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Smoke_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Smoke      'Set the effect number
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
    Effect(EffectIndex).FloatSize = Effect_FToDW(30)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Smoke_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Smoke_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Smoke_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 40, Effect(EffectIndex).Y - 40 + Rnd * 40, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 5, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 11, 1, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2, 0.2, 0.01 + (Rnd * 0.2), 0.01 + (Rnd * 0.01)

End Sub

Private Sub Effect_Smoke_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Smoke_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

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
                    Effect_Smoke_Reset EffectIndex, LoopC

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Water_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Water_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Water_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Water      'Set the effect number
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
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Water_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Water_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Water_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 3 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 11, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 14, 0, 2
    Effect(EffectIndex).Particles(Index).ResetColor 0, 0.1, 0.3, 0.1 + (Rnd * 1), 0.027 + (Rnd * 0.06)

End Sub

Private Sub Effect_Water_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Water_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

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
                    Effect_Water_Reset EffectIndex, LoopC

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Teleport_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Teleport_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Teleport     'Set the effect number
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
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Teleport_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Teleport_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
Dim A As Single
Dim X As Single
Dim Y As Single


    If Rnd * 10 < 5 Then
        'Get the positions
        A = Rnd * 360 * DegreeToRadian
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier) / 2.2 '* (0.8 + Rnd * 0.2)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)
    Else
        A = Rnd * 360 * DegreeToRadian
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier / 2) / 2.2 '* (0.8 + Rnd * 0.2)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier / 2)
    End If
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 1, Rnd * 0.5, Rnd * 0.5, 0.6 + (Rnd * 0.4), 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_Teleport_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    'If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

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
                    Effect_Teleport_Reset EffectIndex, LoopC

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(0, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB * 2, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Curse_Inmovilizar_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Curse_Inmovilizar_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Curse_Inmovilizar     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles                   'Set the number of particles
    Effect(EffectIndex).Used = True                                 'Enabled the effect
    Effect(EffectIndex).X = X                                       'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                                       'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                                   'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(12)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
    
    Effect(EffectIndex).TAngle = 0
    
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
    Next LoopC
    
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Curse_Inmovilizar_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    Dim Co As Single
    Dim sI As Single
    Dim rG As Single
       
    rG = (Rnd * 0.4)
    'Calculate the angle
    
    If Effect(EffectIndex).TAngle = 0 And Effect(EffectIndex).GoToX <> -30000 Then Effect(EffectIndex).TAngle = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
    sI = Sin(Effect(EffectIndex).TAngle * DegreeToRadian)
    Co = Cos(Effect(EffectIndex).TAngle * DegreeToRadian)
    
    'Reset the particle
    If RandomNumber(1, 2) = 2 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * 15, Effect(EffectIndex).Y + sI * Sin(Effect(EffectIndex).Progression) * 15, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor rG, 0.4, rG, 0.5 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * -15, Effect(EffectIndex).Y + sI * Sin(Effect(EffectIndex).Progression) * -15, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor rG, rG, 0.4, 0.2 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    End If

End Sub

Private Sub Effect_Curse_Inmovilizar_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 3) = 3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Curse_Inmovilizar_Reset EffectIndex, LoopC
                    
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
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Lissajous_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Lissajous_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Lissajous       'Set the effect number
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
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Lissajous_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Lissajous_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

Dim X As Single
Dim Y As Single

Dim A As Single
Dim B As Integer
Dim aL As Single
Dim rG As Single

    aL = 3.1415 / 2
    A = 3
    B = 4
    
    rG = (Rnd * 0.4)
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.14
    
    X = A * Sin(Effect(EffectIndex).Progression * aL / 12) * 7 + (Rnd * 10) - 5
    Y = B * Sin(Effect(EffectIndex).Progression / 12) * 7 - 10 + (Rnd * 10) - 5
    
    'Reset the particle
    
    If RandomNumber(1, 2) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor rG, 0.4, rG, 0.9, 0.2 + (Rnd * 0.2)
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - X, Effect(EffectIndex).Y - Y - 20, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor rG, rG, 0.4, 0.5 + (Rnd * 0.2), 0.2
    End If
    

End Sub

Private Sub Effect_Lissajous_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    'Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.001
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 4) = 1 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 200 Then

                    'Reset the particle
                    Effect_Lissajous_Reset EffectIndex, LoopC

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Meditation_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Meditation_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Meditation       'Set the effect number
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
        Effect(EffectIndex).PartVertex(LoopC).RHW = 1
        Effect_Meditation_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Meditation_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

Dim X As Single
Dim Y As Single

Dim A As Single
Dim B As Integer
Dim aL As Single
Dim rG As Single

    aL = 3.1415 / 2
    A = 3
    B = 6
    
    rG = (Rnd * 0.4)
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.14
    
    X = A * Sin(Effect(EffectIndex).Progression * aL / 15) * 20 + (Rnd * 10) - 5
    Y = B * Sin(Effect(EffectIndex).Progression / 15) * 10 - 10 + (Rnd * 10) - 25
    
    'Reset the particle
    
    If RandomNumber(1, 2) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2 + (Rnd * 0.5), 0.2, 0.3, 0.2 + (Rnd * 0.2)
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - X, Effect(EffectIndex).Y - Y - 20, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2 + (Rnd * 0.5), 0.2, 0.6 + (Rnd * 0.4), 0.2
    End If
    

End Sub

Private Sub Effect_Meditation_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    'Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.001
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 4) = 1 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > -1 Then

                    'Reset the particle
                    Effect_Meditation_Reset EffectIndex, LoopC

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

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

