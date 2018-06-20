Attribute VB_Name = "AI"
'Argentum Online 0.12.2
'Copyright (C) 2002 Mï¿½rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 nï¿½mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Cï¿½digo Postal 1900
'Pablo Ignacio Mï¿½rquez

Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    
    'Pretorianos
    SacerdotePretorianoAi = 20
    GuerreroPretorianoAi = 21
    MagoPretorianoAi = 22
    CazadorPretorianoAi = 23
    ReyPretoriano = 24
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visiï¿½n que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'                        Modulo AI_NPC
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'AI de los NPC
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½

Private Sub GuardiasAI(ByVal NPCIndex As Integer, ByVal DelCaos As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            'ï¿½ES CRIMINAL?
                            If Not DelCaos Then
                                If criminal(UI) Then
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            Else
                                If Not criminal(UI) Then
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If  'not inmovil
        Next headingloop
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NPCIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/01/2010 (ZaMa)
'28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
'14/09/200*: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'**************************************************************
On Error GoTo XD
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    Dim UserProtected As Boolean
    
    atacoPJ = False
    
    With Npclist(NPCIndex)
1        For headingloop = eHeading.NORTH To eHeading.WEST
2            nPos = .Pos
3            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
4                Call HeadtoPos(headingloop, nPos)
45                If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
56                    UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
57                    NPCI = MapData(nPos.map, nPos.X, nPos.Y).NPCIndex
                    If UI > 0 And Not atacoPJ Then
58                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
59                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                        
61                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            
                            atacoPJ = True
62                            If .Movement = NpcObjeto Then
                                ' Los npc objeto no atacan siempre al mismo usuario
63                                If RandomNumber(1, 3) = 3 Then atacoPJ = False
                            End If
                            
64                            If atacoPJ Then
65                                If .flags.LanzaSpells Then
                                    If .flags.AtacaDoble Then
66                                        If (RandomNumber(0, 1)) Then
67                                            If NpcAtacaUser(NPCIndex, UI) Then
68                                                Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                    
77                                    Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
78                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                End If
                            End If
79                            If NpcAtacaUser(NPCIndex, UI) Then
71                                Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                            End If
                            Exit Sub

                        End If
                    ElseIf NPCI > 0 Then
72                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
73                            Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
74                            Call SistemaCombate.NpcAtacaNpc(NPCIndex, NPCI, False)
                            Exit Sub
                        End If
                   End If
                End If
            End If  'inmo
        Next headingloop
    End With
    
75    Call RestoreOldMovement(NPCIndex)
XD:

End Sub

Private Sub HostilBuenoAI(ByVal NPCIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
    Dim nPos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).name = .flags.AttackedBy Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                End If
                                
                                If NpcAtacaUser(NPCIndex, UI) Then
                                    Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NPCIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
    Dim tHeading As Byte
    Dim UserIndex As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                        
                        If UserList(UserIndex).flags.Muerto = 0 Then
                            If Not UserProtected Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                Exit Sub
                            End If
                        End If
                        
                    End If
                End If
            Next i
            
        ' No esta inmobilizado
        Else
            
            ' Tiene prioridad de seguir al usuario al que le pertenece si esta en el rango de vision
            Dim OwnerIndex As Integer
            
            OwnerIndex = .Owner
            If OwnerIndex > 0 Then
            
                'Is it in it's range of vision??
                If Abs(UserList(OwnerIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(OwnerIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        ' va hacia el si o esta invi ni oculto
                        If UserList(OwnerIndex).flags.invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.EnConsulta And Not UserList(OwnerIndex).flags.Ignorado Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, OwnerIndex)
                                
                            tHeading = FindDirection(.Pos, UserList(OwnerIndex).Pos)
                            Call MoveNPCChar(NPCIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
            
            ' No le pertenece a nadie o el dueño no esta en el rango de vision, sigue a cualquiera
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        With UserList(UserIndex)
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or .flags.Ignorado Or .flags.EnConsulta
                            
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And _
                                .flags.AdminPerseguible And Not UserProtected Then
                                
                                If Npclist(NPCIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                
                                tHeading = FindDirection(Npclist(NPCIndex).Pos, .Pos)
                                Call MoveNPCChar(NPCIndex, tHeading)
                                Exit Sub
                            End If
                            
                        End With
                        
                    End If
                End If
            Next i
            
            'Si llega aca es que no habï¿½a ningï¿½n usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
            
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NPCIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: Marco Vanotti (MarKoxX)
'Last Modify Date: 08/16/2008
'08/16/2008: MarKoxX - Now pets that do melï¿½ attacks have to be near the enemy to attack.
'**************************************************************
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim i As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer

    With Npclist(NPCIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select

            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)

                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        If UserList(UI).name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(NPCIndex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(NPCIndex).Pos) <= 1 Then
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NPCIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NPCIndex, UI)
                                        End If
                                    End If
                                 End If
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(UI).name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro del ejército real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(NPCIndex)
                                    Exit Sub
                                End If
                            End If
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NPCIndex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(NPCIndex).Pos) <= 1 Then
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NPCIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NPCIndex, UI)
                                        End If
                                    End If
                                 End If
                                 
                                 tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                 Call MoveNPCChar(NPCIndex, tHeading)
                                 
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NPCIndex As Integer)
    With Npclist(NPCIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        End If
    End With
End Sub

Private Sub PersigueCiudadano(ByVal NPCIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
'***************************************************
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)
        For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                    If Not criminal(UserIndex) Then
                    
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                        
                        If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And _
                            UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                            
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                            End If
                            tHeading = FindDirection(.Pos, UserList(UserIndex).Pos)
                            Call MoveNPCChar(NPCIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                    
               End If
            End If
            
        Next i
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub PersigueCriminal(ByVal NPCIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
'***************************************************
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If criminal(UserIndex) Then
                            With UserList(UserIndex)
                                 
                                 UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                                 UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                                 
                                 If .flags.Muerto = 0 And .flags.invisible = 0 And _
                                    .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                     
                                     If Npclist(NPCIndex).flags.LanzaSpells > 0 Then
                                           Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                     End If
                                     Exit Sub
                                End If
                            End With
                        End If
                        
                   End If
                End If
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If criminal(UserIndex) Then
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado
                            
                            If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And _
                               UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                End If
                                If .flags.Inmovilizado = 1 Then Exit Sub
                                tHeading = FindDirection(.Pos, UserList(UserIndex).Pos)
                                Call MoveNPCChar(NPCIndex, tHeading)
                                Exit Sub
                           End If
                        End If
                        
                   End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub SeguirAmo(ByVal NPCIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tHeading As Byte
    Dim UI As Integer
    
    With Npclist(NPCIndex)
        If .Target = 0 And .targetNPC = 0 Then
            UI = .MaestroUser
            
            If UI > 0 Then
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        If UserList(UI).flags.Muerto = 0 _
                                And UserList(UI).flags.invisible = 0 _
                                And UserList(UI).flags.Oculto = 0 _
                                And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NPCIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NPCIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tHeading As Byte
    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With Npclist(NPCIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.map, X, Y).NPCIndex
                        If NI > 0 Then
                            If .targetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NPCIndex, NI)
                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NPCIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
                For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                       NI = MapData(.Pos.map, X, Y).NPCIndex
                       If NI > 0 Then
                            If .targetNPC = NI Then
                                 bNoEsta = True
                                 If .Numero = ELEMENTALFUEGO Then
                                     Call NpcLanzaUnSpellSobreNpc(NPCIndex, NI)
                                     If Npclist(NI).NPCtype = DRAGON Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NPCIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NI)
                                    End If
                                 End If
                                 If .flags.Inmovilizado = 1 Then Exit Sub
                                 If .targetNPC = 0 Then Exit Sub
                                 tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.map, X, Y).NPCIndex).Pos)
                                 Call MoveNPCChar(NPCIndex, tHeading)
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NPCIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Public Sub AiNpcObjeto(ByVal NPCIndex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 14/09/2009 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'***************************************************
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)
        For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
            
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                    With UserList(UserIndex)
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                        
                        If .flags.Muerto = 0 And .flags.invisible = 0 And _
                            .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            
                            ' No quiero que ataque siempre al primero
                            If RandomNumber(1, 3) < 3 Then
                                If Npclist(NPCIndex).flags.LanzaSpells > 0 Then
                                     Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                End If
                            
                                Exit Sub
                            End If
                        End If
                    End With
               End If
            End If
            
        Next i
    End With

End Sub

Sub NPCAI(ByVal NPCIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: ZaMa
'Last Modify Date: 15/11/2009
'08/16/2008: MarKoxX - Now pets that do melï¿½ attacks have to be near the enemy to attack.
'15/11/2009: ZaMa - Implementacion de npc objetos ai.
'**************************************************************

On Error GoTo ErrorHandler
    With Npclist(NPCIndex)
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then
            'Busca a alguien para atacar
            'ï¿½Es un guardia?
            If .NPCtype = eNPCType.GuardiaReal Then
2                Call GuardiasAI(NPCIndex, False)
3            ElseIf .NPCtype = eNPCType.Guardiascaos Then
4                Call GuardiasAI(NPCIndex, True)
5            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
6                Call HostilMalvadoAI(NPCIndex)
7            ElseIf .Hostile And .Stats.Alineacion = 0 Then
8                Call HostilBuenoAI(NPCIndex)
9            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
11        Select Case .Movement
            Case TipoAI.MueveAlAzar
12                If .flags.Inmovilizado = 1 Then Exit Sub
13                If .NPCtype = eNPCType.GuardiaReal Then
14                    If RandomNumber(1, 12) = 3 Then
15                        Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
16                    End If
                    
17                    Call PersigueCriminal(NPCIndex)
                    
18                ElseIf .NPCtype = eNPCType.Guardiascaos Then
19                    If RandomNumber(1, 12) = 3 Then
21                        Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    
23                    Call PersigueCiudadano(NPCIndex)
                    
                Else
24                    If RandomNumber(1, 12) = 3 Then
26                        Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
25                    End If
                End If
            
            'Va hacia el usuario cercano
27            Case TipoAI.NpcMaloAtacaUsersBuenos
28                Call IrUsuarioCercano(NPCIndex)
            
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
29                Call SeguirAgresor(NPCIndex)
            
            'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
33                Call PersigueCriminal(NPCIndex)
            
            Case TipoAI.SigueAmo
34                If .flags.Inmovilizado = 1 Then Exit Sub
35                Call SeguirAmo(NPCIndex)
36                If RandomNumber(1, 12) = 3 Then
37                    Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
38                End If
            
39            Case TipoAI.NpcAtacaNpc
41                Call AiNpcAtacaNpc(NPCIndex)
                
42            Case TipoAI.NpcObjeto
43                Call AiNpcObjeto(NPCIndex)
                
46            Case TipoAI.NpcPathfinding
44                If .flags.Inmovilizado = 1 Then Exit Sub
49                If ReCalculatePath(NPCIndex) Then
51                    Call PathFindingAI(NPCIndex)
                    'Existe el camino?
52                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
56                        Call MoveNPCChar(NPCIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
58                    End If
                Else
59                    If Not PathEnd(NPCIndex) Then
61                        Call FollowPath(NPCIndex)
                    Else
62                        .PFINFO.PathLenght = 0
                    End If
                End If
        End Select
    End With
Exit Sub

ErrorHandler:
    Call LogError("lINEA " & Erl() & " .NPCAI " & Npclist(NPCIndex).name & " " & Npclist(NPCIndex).MaestroUser & " " & Npclist(NPCIndex).MaestroNpc & " mapa:" & Npclist(NPCIndex).Pos.map & " x:" & Npclist(NPCIndex).Pos.X & " y:" & Npclist(NPCIndex).Pos.Y & " Mov:" & Npclist(NPCIndex).Movement & " TargU:" & Npclist(NPCIndex).Target & " TargN:" & Npclist(NPCIndex).targetNPC)
    Dim MiNPC As Npc
    MiNPC = Npclist(NPCIndex)
    'Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function UserNear(ByVal NPCIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'Returns True if there is an user adjacent to the npc position.
'***************************************************

    With Npclist(NPCIndex)
        UserNear = Not Int(Distance(.Pos.X, .Pos.Y, UserList(.PFINFO.targetUser).Pos.X, _
                    UserList(.PFINFO.targetUser).Pos.Y)) > 1
    End With
End Function

Function ReCalculatePath(ByVal NPCIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'Returns true if we have to seek a new path
'***************************************************

    If Npclist(NPCIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NPCIndex) And Npclist(NPCIndex).PFINFO.PathLenght = Npclist(NPCIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True
    End If
End Function

Function PathEnd(ByVal NPCIndex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'Returns if the npc has arrived to the end of its path
'***************************************************
    PathEnd = Npclist(NPCIndex).PFINFO.CurPos = Npclist(NPCIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NPCIndex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'Moves the npc.
'***************************************************
    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    With Npclist(NPCIndex)
        tmpPos.map = .Pos.map
        tmpPos.X = .PFINFO.Path(.PFINFO.CurPos).Y ' invertï¿½ las coordenadas
        tmpPos.Y = .PFINFO.Path(.PFINFO.CurPos).X
        
        'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
        
        tHeading = FindDirection(.Pos, tmpPos)
        
        MoveNPCChar NPCIndex, tHeading
        
        .PFINFO.CurPos = .PFINFO.CurPos + 1
    End With
End Function

Function PathFindingAI(ByVal NPCIndex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'This function seeks the shortest path from the Npc
'to the user's location.
'***************************************************
    Dim Y As Long
    Dim X As Long
    
    With Npclist(NPCIndex)
        For Y = .Pos.Y - 10 To .Pos.Y + 10    'Makes a loop that looks at
             For X = .Pos.X - 10 To .Pos.X + 10   '5 tiles in every direction
                
                 'Make sure tile is legal
                 If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                    
                     'look for a user
                     If MapData(.Pos.map, X, Y).UserIndex > 0 Then
                         'Move towards user
                          Dim tmpUserIndex As Integer
                          tmpUserIndex = MapData(.Pos.map, X, Y).UserIndex
                          With UserList(tmpUserIndex)
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
                                Npclist(NPCIndex).PFINFO.Target.X = .Pos.Y
                                Npclist(NPCIndex).PFINFO.Target.Y = .Pos.X 'ops!
                                Npclist(NPCIndex).PFINFO.targetUser = tmpUserIndex
                                Call SeekPath(NPCIndex)
                                Exit Function
                            End If
                        End With
                    End If
                End If
            Next X
        Next Y
    End With
End Function

Sub NpcLanzaUnSpell(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: -
'Last Modify Date: -
'**************************************************************
    With UserList(UserIndex)
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
    End With
    
    Dim k As Integer
    k = RandomNumber(1, Npclist(NPCIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NPCIndex, UserIndex, Npclist(NPCIndex).Spells(k))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NPCIndex As Integer, ByVal targetNPC As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim k As Integer
    k = RandomNumber(1, Npclist(NPCIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NPCIndex, targetNPC, Npclist(NPCIndex).Spells(k))
End Sub
