Attribute VB_Name = "SistemaCombate"

'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Public Function PoderEvasionEscudo(ByVal userIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PoderEvasionEscudo = (UserList(userIndex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(userIndex).Clase).Escudo) / 2
End Function

Public Function PoderEvasion(ByVal userIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Dim lTemp As Long
    With UserList(userIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function PoderAtaqueArma(ByVal userIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(userIndex)
        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal userIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(userIndex)
        If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal userIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(userIndex)
        If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.Clase).AtaqueWrestling
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueWrestling
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueWrestling
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueWrestling
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal userIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    Arma = UserList(userIndex).Invent.WeaponEqpObjIndex
    
    If Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(userIndex)
            Skill = eSkill.Proyectiles
            'Envia arrows. - maTih.-
            Dim MunicionObjIndex    As Integer
            MunicionObjIndex = UserList(userIndex).Invent.MunicionEqpObjIndex
            'Tiene munición?
            If MunicionObjIndex <> 0 Then
                Call mod_DunkanGeneral.Enviar_FlechaANpc(UserList(userIndex).Char.CharIndex, NpcIndex, ObjData(MunicionObjIndex).GrhIndex)
            End If
        Else
            PoderAtaque = PoderAtaqueArma(userIndex)
            Skill = eSkill.Armas
        End If
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(userIndex)
        Skill = eSkill.Wrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal userIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    UserEvasion = PoderEvasion(userIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(userIndex)
    
    SkillTacticas = UserList(userIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(userIndex).Stats.UserSkills(eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y))
                    Call WriteMultiMessage(userIndex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
                End If
            End If
        End If
    End If
End Function

Public Function CalcularDaño(ByVal userIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/04/2010 (ZaMa)
'01/04/2010: ZaMa - Modifico el daño de wrestling.
'01/04/2010: ZaMa - Agrego bonificadores de wrestling para los guantes.
'***************************************************
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim DañoMaxArma As Long
    Dim DañoMinArma As Long
    Dim objIndex As Integer
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(userIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.Minhit, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.Minhit, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                            DañoArma = RandomNumber(Arma.Minhit, Arma.MaxHIT)
                            DañoMaxArma = Arma.MaxHIT
                            matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon daño es 1
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                        DañoArma = RandomNumber(Arma.Minhit, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If
                End If
            Else ' Ataca usuario
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.Minhit, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.Minhit, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.Clase).DañoArmas
                        DañoArma = 1 ' Si usa la espada mataDragones daño es 1
                        DañoMaxArma = 1
                    Else
                        DañoArma = RandomNumber(Arma.Minhit, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If
                End If
            End If
        Else
            ModifClase = ModClase(.Clase).DañoWrestling
            
            ' Daño sin guantes
            DañoMinArma = 4
            DañoMaxArma = 9
            
            ' Plus de guantes (en slot de anillo)
            objIndex = .Invent.AnilloEqpObjIndex
            If objIndex > 0 Then
                If ObjData(objIndex).Guante = 1 Then
                    DañoMinArma = DañoMinArma + ObjData(objIndex).Minhit
                    DañoMaxArma = DañoMaxArma + ObjData(objIndex).MaxHIT
                End If
            End If
            
            DañoArma = RandomNumber(DañoMinArma, DañoMaxArma)
            
        End If
        
        DañoUsuario = RandomNumber(.Stats.Minhit, .Stats.MaxHIT)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDaño = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DañoUsuario) * ModifClase
        End If
    End With
End Function

Public Sub RestarCriminalidad(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim EraCriminal As Boolean
    EraCriminal = criminal(userIndex)
    
    With UserList(userIndex).Reputacion
        If .BandidoRep > 0 Then
             .BandidoRep = .BandidoRep - vlASALTO
             If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
             .LadronesRep = .LadronesRep - (vlCAZADOR * 10)
             If .LadronesRep < 0 Then .LadronesRep = 0
        End If
    End With
    
    If EraCriminal And Not criminal(userIndex) Then
        Call RefreshCharStatus(userIndex)
    End If
End Sub




Public Sub UsuarioAtaca(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/02/2012
'                   Agrego checkeos de la IA (maTih.-)
'***************************************************
 
    Dim Index As Integer
    Dim AttackPos As WorldPos
   
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(userIndex, False) Then Exit Sub
   
    'Check Spell-Magic interval
    If Not IntervaloPermiteMagiaGolpe(userIndex) Then
        'Check Attack interval
        If Not IntervaloPermiteAtacar(userIndex) Then
            Exit Sub
        End If
    End If
   
    With UserList(userIndex)
       
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
       
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub
        End If
       
    #If ConBots Then
        'Está el bot?
        If MapData(AttackPos.map, AttackPos.X, AttackPos.Y).botIndex <> 0 Then
           'Checkeo que esté invocado.
           If ia_Bot(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).botIndex).Invocado Then
              ia_DamageHit MapData(AttackPos.map, AttackPos.X, AttackPos.Y).botIndex, userIndex
           End If
        End If
    #End If
       
        Index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).userIndex
       
        'Look for user
        If Index > 0 Then
            Call UsuarioAtacaUsuario(userIndex, Index)
            Call WriteUpdateHP(userIndex)
            Call WriteUpdateHP(Index)
            Exit Sub
        End If
       
        Index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex
       
       
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        Call WriteUpdateHP(userIndex)
        Call WriteUpdateMana(userIndex)
       
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
           
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    Dim ProbEvadir As Long
    Dim Skill As eSkill
    
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
    
    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
       UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
       UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    'Esta usando un arma ???
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
            Skill = eSkill.Proyectiles
            'Envia arrow to user. - maTih.-
            Call mod_DunkanGeneral.Enviar_FlechaANpc(UserList(AtacanteIndex).Char.CharIndex, VictimaIndex, ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).GrhIndex)
        Else
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)
            Skill = eSkill.Armas
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
        Skill = eSkill.Wrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    ' Se reduce la evasion un 25%
    If UserList(VictimaIndex).flags.Meditando = True Then
        ProbEvadir = (100 - ProbExito) * 0.75
        ProbExito = MinimoInt(90, 100 - ProbEvadir)
    End If
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        'Fallo ???
        If Not UsuarioImpacto Then
            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
                Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
            End If
        End If
    End If

    
    Call FlushBuffer(VictimaIndex)
    
    Exit Function
    
Errhandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets
'                    inválidos, y evitar un doble chequeo innecesario
'***************************************************

On Error GoTo Errhandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    
    With UserList(AtacanteIndex)
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Function
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acción
            If .Clase = eClass.Bandit Then
                
                
            'y ahora, el ladrón puede llegar a paralizar con el golpe.
            ElseIf .Clase = eClass.Thief Then
                
            End If
            
            Call UserDañoUser(AtacanteIndex, VictimaIndex)
        Else
            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            End If
            
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
        End If
        

        End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
Errhandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.Description)
End Function

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar
'11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal
'***************************************************
    
On Error GoTo Errhandler

    Dim Daño As Long
    Dim Lugar As Byte
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim Obj As ObjData
    Dim Resist As Byte
    
    Daño = CalcularDaño(AtacanteIndex)
    
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(.Invent.BarcoObjIndex)
             Daño = Daño + RandomNumber(Obj.Minhit, Obj.MaxHIT)
        End If
        
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    Daño = Daño - absorbido
                    If Daño < 0 Then Daño = 1
                End If
            
            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData
                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                    absorbido = absorbido + defbarco - Resist
                    Daño = Daño - absorbido
                    If Daño < 0 Then Daño = 1
                End If
        End Select
        
        Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, Daño)
        Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, Daño)
              
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - Daño

        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    
                    ' Si es arma arrojadiza..
                    If ObjData(.Invent.WeaponEqpObjIndex).Municion = 0 Then
                        ' Si acuchilla
                        If ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1 Then
                            Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, Daño)
                        End If
                    End If
                Else
                End If
            Else

            End If
                    
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, Daño)
            End If
            'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
            Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, Daño)
        End If
        
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
            
            ' No cuenta la muerte si estaba en estado atacable
            If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                
                Call ContarMuerte(VictimaIndex, AtacanteIndex)
            End If
            
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex)
        Else
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)
        End If
    End With
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Sub
    
Errhandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDañoUser. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Sub

Public Function PuedeApuñalar(ByVal userIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(userIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                        Or UserList(userIndex).Clase = eClass.Assasin
        End If
    End If
End Function


Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal victimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 05/05/2010
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
'***************************************************

    If TriggerZonaPelea(AttackerIndex, victimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    Dim VictimaEsAtacable As Boolean
    
    If Not criminal(AttackerIndex) Then
        If Not criminal(victimIndex) Then
            ' Si la victima no es atacable por el agresor, entonces se hace pk
            VictimaEsAtacable = UserList(victimIndex).flags.AtacablePor = AttackerIndex
            If Not VictimaEsAtacable Then Call VolverCriminal(AttackerIndex)
        End If
    End If
    
    With UserList(victimIndex)
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(victimIndex)
            Call WriteConsoleMsg(victimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, victimIndex, mod_DunkanProtocol.Send_CreateMeditation(.Char.CharIndex, 0, 0))
        End If
    End With
    
    EraCriminal = criminal(AttackerIndex)
    
    ' Si ataco a un atacable, no suma puntos de bandido
    If Not VictimaEsAtacable Then
        With UserList(AttackerIndex).Reputacion
            If Not criminal(victimIndex) Then
                .BandidoRep = .BandidoRep + vlASALTO
                If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                
                .NobleRep = .NobleRep * 0.5
                If .NobleRep < 0 Then .NobleRep = 0
            Else
                .NobleRep = .NobleRep + vlNoble
                If .NobleRep > MAXREP Then .NobleRep = MAXREP
            End If
        End With
    End If
    
    If criminal(AttackerIndex) Then

        
        If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
    ElseIf EraCriminal Then
        Call RefreshCharStatus(AttackerIndex)
    End If

    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(victimIndex)
    Call FlushBuffer(victimIndex)
End Sub


Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal victimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 02/04/2010
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
'***************************************************
On Error GoTo Errhandler

    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(victimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un espíritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Servidor en deathMatch?
    PuedeAtacar = Server_Info.DeathMathc
    
    'No es death?
    If Not PuedeAtacar Then
       'Checkea, los 2 criminalES? no puede.
       If criminal(victimIndex) And criminal(AttackerIndex) Then
          Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a los de tu equipo.!", FontTypeNames.FONTTYPE_FIGHT)
          Exit Function
       ElseIf Not criminal(victimIndex) And (Not criminal(AttackerIndex)) Then
       'Checkea, los 2 ciudadanos.
          Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a los de tu equipo.!", FontTypeNames.FONTTYPE_FIGHT)
          Exit Function
       Else
          'Uno de cada equipo, puede.
          PuedeAtacar = True
       End If
    End If
    PuedeAtacar = True
Exit Function

Errhandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.Description)
End Function


Private Function SameClan(ByVal userIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Returns True if both players belong to the same clan.
'Last Modification: 16/11/2009
'***************************************************
    SameClan = (UserList(userIndex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And _
                UserList(userIndex).GuildIndex <> 0
End Function

Private Function SameParty(ByVal userIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Returns True if both players belong to the same party.
'Last Modification: 16/11/2009
'***************************************************
    SameParty = UserList(userIndex).PartyIndex = UserList(OtherUserIndex).PartyIndex And _
                UserList(userIndex).PartyIndex <> 0
End Function

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "¡¡" & UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "¡¡Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Sub
