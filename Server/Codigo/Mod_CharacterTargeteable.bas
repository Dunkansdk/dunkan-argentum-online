Attribute VB_Name = "Mod_CharacterTargeteable"
Option Explicit

' @author : maTih.-
' @date   : 20110107
' @note   : Module of manager to the characters targets.

'declare constants to be system.
Private Const TARGETS_SLOTS_OBJ     As Byte = 2
Private Const TARGETS_SLOTS_USER    As Byte = 7
Private Const TARGETS_SLOTS_NPC     As Byte = 6

Public Sub CharTarget_SendTargets(ByVal UserIndex As Integer, ByVal targetUser As Integer, ByVal targetNPC As Integer, ByVal TargetOBJ As Integer)

' @author : maTih.-
' @date   : 20110107
' @note   : prepare and send targets to be user.

Dim endString       As String
Dim nPosOBJ         As WorldPos

    With UserList(UserIndex)
    
    'Analize the posible targets allways its npc.
    If targetNPC > 0 Then
        
        'Set Distance.
        
        If Distancia(Npclist(targetNPC).Pos, .Pos) <= TARGETS_SLOTS_NPC Then
        
        'Where NPC-Hostil?
                
        If Npclist(targetNPC).Hostile = 1 Then
            
            endString = CharTarget_GetTargetNPCHostile(targetNPC, UserIndex)
            
        Else
            
            endString = CharTarget_GetTargetNPC(targetNPC)
        
        End If
        
        End If

    End If
    
    'Analize the posible targets allways its npc.
    
    If targetUser > 0 Then
    
        'Where User in Distance < 7?
        
        If Distancia(.Pos, UserList(targetUser).Pos) <= TARGETS_SLOTS_USER Then
        
            endString = CharTarget_GetTargetUseR(UserIndex, targetUser)
            
        End If
    
    End If
    
    If TargetOBJ > 0 Then
        
        'Generate world pos by targetObjPos
        
        nPosOBJ.map = .flags.TargetObjMap
        nPosOBJ.X = .flags.TargetObjX
        nPosOBJ.Y = .flags.TargetObjY
        
        'Analize the distance.
        
        If Distancia(.Pos, nPosOBJ) <= TARGETS_SLOTS_OBJ Then
        
            endString = CharTarget_GetTargetObj(TargetOBJ)
        
        End If
        
        Exit Sub
        
    End If
    
    If HayAgua(.Pos.map, .flags.TargetX, .flags.TargetY) Then
    
            endString = CharTarget_GetWhater
            
    End If
    
    End With
    
    WriteDaoUpdateTargets UserIndex, endString
    
End Sub

Public Function CharTarget_GetWhater() As String

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its targets of obj.

CharTarget_GetWhater = "Mar|Pescar"

End Function

Public Function CharTarget_GetTargetObj(ByVal ObjIndex As Integer) As String
    
' @author : maTih.-
' @date   : 20110107
' @note   : Returns its targets of obj.
    
Dim tmpTarget   As String

    If ObjIndex <= 0 Or ObjIndex > UBound(ObjData()) Then
        CharTarget_GetTargetObj = "Invalid obj."
        Exit Function
    End If
    
    With ObjData(ObjIndex)
        
        'We are is ObjExplotable.
            If CharTarget_ObjExplotable(ObjIndex, tmpTarget) Then
                CharTarget_GetTargetObj = .name & "|" & tmpTarget
            End If
        
    End With

End Function

Private Function CharTarget_ObjExplotable(ByVal ObjIndex As Integer, ByRef finaltyString As String) As Boolean

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its Obj explotable.

    With ObjData(ObjIndex)
    
    Dim EsYunque        As Boolean
    Dim EsYacimiento    As Boolean
    Dim EsArbolElfico   As Boolean
    Dim EsFragua        As Boolean
    
    EsYunque = (.OBJType = eOBJType.otyunque)
    EsYacimiento = (.OBJType = eOBJType.otYacimiento)
    EsArbolElfico = (.OBJType = eOBJType.otArbolElfico)
    EsFragua = (.OBJType = eOBJType.otFragua)
    
    CharTarget_ObjExplotable = (EsYunque Or EsYacimiento Or EsArbolElfico Or EsFragua)
    
    If EsYunque Then
        finaltyString = "Crear item"
    ElseIf EsYacimiento Then
        finaltyString = "Explotar"
    ElseIf EsArbolElfico Then
        finaltyString = "Talar"
    ElseIf EsFragua Then
        finaltyString = "UtilizaR"
    End If
    
    End With
    
End Function

Public Function CharTarget_GetTargetNPCHostile(ByVal NPCTarget As Integer, ByVal UserIndex As Integer) As String

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its targets of NPC Hostiles.

Dim tmpString       As String
Dim refSkills       As String

    With Npclist(NPCTarget)
    
        tmpString = .name & "|" & (.Stats.MinHp & "/" & .Stats.MaxHp) & "|" & CharTarget_GetSpellsToNpc(NPCTarget)
        
        If .flags.Domable Then
            
                CharTarget_GetUserDomableNPC UserIndex, NPCTarget, refSkills
                tmpString = tmpString & "|Domable[Puedes domarlo:" & refSkills & "]"
            
        End If
        
    End With
    
    CharTarget_GetTargetNPCHostile = tmpString
    
End Function

Public Function CharTarget_GetUserDomableNPC(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByRef skillsPoints As String) As Boolean

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its NPC is domable by UserIndex

Dim RequieredPoints     As Byte
Dim MySkillPoints       As Byte

With UserList(UserIndex)

            MySkillPoints = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))
            
            'BONIFICACIONES DE DRUIDAS.
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                RequieredPoints = Npclist(NpcIndex).flags.Domable * 0.8
            
            ElseIf .Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
                RequieredPoints = Npclist(NpcIndex).flags.Domable * 0.89
                
            Else
                RequieredPoints = Npclist(NpcIndex).flags.Domable
            End If
            
            If RequieredPoints <= MySkillPoints Then
                CharTarget_GetUserDomableNPC = True
                skillsPoints = "Si"
            Else
                CharTarget_GetUserDomableNPC = False
                skillsPoints = "No"
            End If
End With

End Function

Public Function CharTarget_GetSpellsToNpc(ByVal NpcIndex As Integer)

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its List of spells NPCIndex.

    Dim loopX       As Long
    Dim inSPells    As String:
    Dim spellIndex  As Integer
    
    With Npclist(NpcIndex)
    
        If .flags.LanzaSpells Then
                
                spellIndex = .Spells(.flags.LanzaSpells)
                
                inSPells = "Lanza: " & Hechizos(spellIndex).Nombre
                
        End If
        
    End With
     
    CharTarget_GetSpellsToNpc = inSPells
    
End Function

Public Function CharTarget_GetTargetNPC(ByVal NPCTarget As Integer) As String

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its targets of NPC not hostile.

CharTarget_GetTargetNPC = Npclist(NPCTarget).name

    With Npclist(NPCTarget)
        
        If .Comercia = 1 Then
        
            CharTarget_GetTargetNPC = CharTarget_GetTargetNPC & "|Comerciar"
        
        ElseIf .NPCtype = eNPCType.Banquero Then
            CharTarget_GetTargetNPC = CharTarget_GetTargetNPC & "|Iniciar deposito"
        
        End If
    
    End With

End Function

Public Function CharTarget_GetTargetUseR(ByVal UserTargeteador As Integer, ByVal UserTarget As Integer) As String

' @author : maTih.-
' @date   : 20110107
' @note   : Returns its targets of targetUserIndex.

    Dim MyGuild         As Integer
    Dim MyParty         As Integer
    
    Dim otherGuild      As Integer
    Dim otherParty      As Integer
    
    Dim nString         As String
    
    MyGuild = UserList(UserTargeteador).GuildIndex
    MyParty = UserList(UserTargeteador).PartyIndex
    
    otherGuild = UserList(UserTarget).GuildIndex
    otherParty = UserList(UserTarget).PartyIndex
    
    nString = UserList(UserTarget).name & "|Comerciar|"
    
    'Analiza CLANES.
    
   If MyGuild <> 0 Then
        'Somos lideres?
        Debug.Print modGuilds.GuildLeader(MyGuild)
        If UCase$(modGuilds.GuildLeader(MyGuild)) = UCase$(UserList(UserTarget).name) Then
            'Si tiene clan, pero al que clickeo NO, entonces lo invita.
            If otherGuild = 0 Then
                nString = nString & "Invitar al clan"
            ElseIf otherGuild = MyGuild Then
            'Si tiene clan, y son del mismo clan, entonces envia "Expulsar"
                nString = nString & "Expulsar del clan"
            End If
        End If
    End If
    
    'Analiza PARTYs.
    
    'Si está en una party...
    If MyParty <> 0 Then
        'Si es el lider de la misma.
        If Parties(MyParty).EsPartyLeader(UserTargeteador) Then
            'Si el otro NO tiene party le envia "Invitar"
            If otherParty = 0 Then
                nString = nString & "|Invitar a la party"
            ElseIf otherParty = MyParty Then
            'Si el otro tiene party AND es la misma que quien clikea, y somos lideres, Expulsar.
                nString = nString & "|Expulsar de la party"
            End If
        End If
    End If
    
    CharTarget_GetTargetUseR = nString
    
End Function
