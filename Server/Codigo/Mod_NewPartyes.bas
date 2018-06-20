Attribute VB_Name = "Mod_NewPartyes"
Option Explicit

'************************************************************
'..................NUEVO SISTEMA DE PARTYES..................
'************************************************************
'************************************************************
'..................Escrito por maTih(28/01)..................
'************************************************************

'CANTIDAD DE MIEMBROS MAXIMOS.
Private Const MAX_MIEMBROS          As Integer = 4

'CANTIDAD DE MAXIMAS PARTYS.
Private Const MAX_PARTYS            As Integer = 500

Type newPartys
     UserIndex(1 To MAX_MIEMBROS)   As Integer  'Punteros de los usuarios.
     Nivel                          As Byte     'Nivel de la party.
     ExpAcumulada                   As Double   'Experiencia acumulada en la party.
     PartyLider                     As Integer  'Lider de la party.
     CantidadMiembros               As Byte     'Para repartir exp.
End Type

Private UltimoPartyIndex            As Byte     'Con esto indexeamos el array.

Public newParty(1 To MAX_PARTYS) As newPartys   'Declaracion de uso.

Sub NewParty_ResetearSlot(ByVal SlotIndex As Integer, ByVal ResetAll As Boolean)

' \ Author  : maTih.-
' \ Note    : Resetea uno o todos los slots de el array de partys..

Dim loopX       As Long

With newParty(SlotIndex)

    If Not ResetAll Then
    
             .CantidadMiembros = 0
             .ExpAcumulada = 0
             .Nivel = 0
             .PartyLider = 0
             
             For loopX = 1 To MAX_MIEMBROS
             .UserIndex(loopX) = 0
             Next loopX
    
        Exit Sub
    
    End If
    
End With
    
    'Si estamos acá hay que resetear todos los slots.
    
    For loopX = 1 To MAX_PARTYS
        With newParty(loopX)
             .CantidadMiembros = 0
             .ExpAcumulada = 0
             .Nivel = 0
             .PartyLider = 0
             .UserIndex(1) = 0
             .UserIndex(2) = 0
             .UserIndex(3) = 0
        End With
    Next loopX
    
End Sub

Sub NewParty_SolicitarIngreso(ByVal targetUser As Integer, ByVal SolicitanteIndex As Integer)

' \ Author  : maTih.-
' \ Note    : SolicitanteIndex solicita ingresar a la party de TargetUser

Dim otherPI     As Integer
Dim myPI        As Integer
Dim MyLVL       As Byte

otherPI = UserList(targetUser).IndexParty
myPI = UserList(SolicitanteIndex).IndexParty

MyLVL = UserList(SolicitanteIndex).Stats.ELV

'Si a quien clickea no tiene ninguna party.
If otherPI <= 0 Then
    WriteConsoleMsg SolicitanteIndex, UserList(targetUser).name & " No es lider de ninguna party!", FontTypeNames.FONTTYPE_CITIZEN
    Exit Sub
End If

'Si quien clickea tiene una party.
If myPI > 0 Then
    'Si es el mismo partyIndex.
    If myPI = otherPI Then
        WriteConsoleMsg SolicitanteIndex, "Ya integras la party de " & UserList(targetUser).name & " !", FontTypeNames.FONTTYPE_CITIZEN
        Exit Sub
    End If
    
    'Si no...
    WriteConsoleMsg SolicitanteIndex, "Ya integras una party!", FontTypeNames.FONTTYPE_CITIZEN
    Exit Sub
End If

'Checkeo si la party está llena.

If NewParty_PartyLlena(newParty(otherPI).CantidadMiembros, newParty(otherPI).Nivel) Then
    WriteConsoleMsg SolicitanteIndex, "La party está llena!!", FontTypeNames.FONTTYPE_CITIZEN
    Exit Sub
End If

'Checkeo rango del nivel con el lider.

With UserList(newParty(otherPI).PartyLider)
    
    'Si su nivel - nivel del lider es menor o mayor a 4 entonces no puede.
    If (.Stats.ELV - MyLVL) < 4 Or (.Stats.ELV - MyLVL) > 4 Then
        WriteConsoleMsg SolicitanteIndex, "No puedes ingresar a esa party por que el creador tiene un rango mayor/menor a 4 que el tuyo.", FontTypeNames.FONTTYPE_CITIZEN
        Exit Sub
    End If
    
End With

    'Tenemos todo para enviar la solicitud.

    WriteConsoleMsg newParty(otherPI).PartyLider, UserList(SolicitanteIndex).name & " Está solicitando ingresar a vuestra party, para aceptarlo tipea /ACEPTARPARTY " & UserList(SolicitanteIndex).name, FontTypeNames.FONTTYPE_CITIZEN
        
    'Para aceptarlo, guardamos el puntero con el que indexamos el array aca.
    UserList(SolicitanteIndex).SolicitandoParty = otherPI
    
End Sub

Sub NewParty_LiderAcepta(ByVal AceptadoIndex As Integer)

' \ Author  : maTih.-
' \ Note    : Lider acepta a AceptadoIndex.

Dim aspirantIndex       As Integer
Dim LeaderIndex         As Integer

'Aspirante a que Party.

aspirantIndex = UserList(AceptadoIndex).SolicitandoParty

If aspirantIndex <= 0 Then Exit Sub

'Checkeo si la party está abierta.

'Si el primer usuario es nulo entonces no hay tal party.
If newParty(aspirantIndex).UserIndex(1) = 0 Then Exit Sub

LeaderIndex = newParty(aspirantIndex).PartyLider

'Bien , checkeo que no se halla llenado.

If NewParty_PartyLlena(newParty(aspirantIndex).CantidadMiembros, newParty(aspirantIndex).Nivel) Then
    WriteConsoleMsg LeaderIndex, "La party se ha llenado!!", FontTypeNames.FONTTYPE_CITIZEN
    Exit Sub
End If

'Bueno , ahora si, aceptar al usuario..

    With newParty(aspirantIndex)
    
        .CantidadMiembros = .CantidadMiembros + 1
        .UserIndex(.CantidadMiembros) = AceptadoIndex
        NewParty_MsgParty aspirantIndex, UserList(LeaderIndex).name & " Ha aceptado a " & UserList(AceptadoIndex).name & " en la party!"
                
    End With

End Sub

Sub NewParty_SalirParty(ByVal CloseUserIndex As Integer)

' \ Author  : maTih.-
' \ Note    : CloseUserIndex sale de su party.

    Dim myPI        As Integer
    Dim MyNumero    As Byte
    Dim expForMy    As Long
    
    'Mi partyIndex aca.
    
    myPI = UserList(CloseUserIndex).IndexParty
    
    If myPI <= 0 Then Exit Sub
    
    'Experiencia que me corresponde de la total.
    
    'No bonificamos el % extra si se retira.
    expForMy = (newParty(myPI).ExpAcumulada / newParty(myPI).CantidadMiembros)
    
    'Sumamos.
    UserList(CloseUserIndex).Stats.Exp = UserList(CloseUserIndex).Stats.Exp + expForMy
    
    'Updateamos cliente.
    WriteUpdateExp CloseUserIndex
    
    'Checkeamos si subió de nivel.
    CheckUserLevel CloseUserIndex
    
    'Restamos al puntero del array la exp.
    newParty(myPI).ExpAcumulada = newParty(myPI).ExpAcumulada - expForMy
    
    'Informamos a los otros usuarios
    NewParty_MsgParty myPI, UserList(CloseUserIndex).name & " Se fué de la party!"
    
    'Encontramos al usuario en la party y lo seteamos como 0
    
    MyNumero = NewParty_FoundNumber(CloseUserIndex, myPI)
    
    newParty(myPI).UserIndex(MyNumero) = 0
    
    UserList(CloseUserIndex).IndexParty = 0
    
End Sub

Function NewParty_FoundNumber(ByVal Index As Integer, ByVal PartyIndex As Integer) As Byte

' \ Author  : maTih.-
' \ Note    : Encuentra el numero de usuarios para el array de userIndex (de la party)

Dim loopX       As Long

For loopX = 1 To MAX_MIEMBROS
    With newParty(PartyIndex)
    
        If .UserIndex(loopX) > 0 Then
            If .UserIndex(loopX) = Index Then
                NewParty_FoundNumber = Index
                Exit Function
            End If
        End If
    
    End With
Next loopX

End Function

Sub NewParty_MsgParty(ByVal IndexParty As Integer, ByVal Message As String)

' \ Author  : maTih.-
' \ Note    : Envia un mensaje a un index de party.

    Dim loopX       As Long
    
    With newParty(IndexParty)
    
        For loopX = 1 To .CantidadMiembros
            
            If UserList(.UserIndex(loopX)).ConnID <> -1 Then
                WriteConsoleMsg .UserIndex(loopX), Message, FontTypeNames.FONTTYPE_CITIZEN
            End If
            
        Next loopX
    
    End With

End Sub

Sub NewParty_EnviarStatus(ByVal PartyIndex As Integer)

' \ Author  : maTih.-
' \ Note    : Actualiza todos los clientes de PartyIndex

Dim loopX       As Long

With newParty(PartyIndex)

    For loopX = 1 To MAX_MIEMBROS   '.CantidadMiembros
        
        If .UserIndex(loopX) > 0 Then
            If UserList(.UserIndex(loopX)).ConnID <> -1 Then
                'WriteUpdatePartyStatus .userindex(loopx), NewParty_PrepareStrings(partyIndex)
            End If
        End If
        
    Next loopX
End With

End Sub

Sub NewParty_CrearParty(ByVal CreadorIndex As Integer)

' \ Author  : maTih.-
' \ Note    : CreadorIndex crea una party.
    
    UltimoPartyIndex = UltimoPartyIndex + 1
    
    With newParty(UltimoPartyIndex)
        
        'Guardamos el Creador como primer usuario y como Lider.
        .UserIndex(1) = CreadorIndex
        .PartyLider = CreadorIndex
        
        'Cantidad de miembros es 1.
        .CantidadMiembros = 1
        
        'Reseteamos exp acumulada.
        .ExpAcumulada = 0
        
        'Obtenemos el nivel de la party segun su skill.
        .Nivel = NewParty_PosibleNivel(NewParty_GetSkill(CreadorIndex))
        
        'Informamos al usuario.
        
        WriteConsoleMsg CreadorIndex, "Has creado una party!", FontTypeNames.FONTTYPE_CITIZEN
        WriteConsoleMsg CreadorIndex, "El nivel de la misma es : " & .Nivel & " el bonus de experiencia es de : " & NewParty_PorcentajeByLVL(.Nivel) & " %", FontTypeNames.FONTTYPE_CITIZEN
        
    End With
    
End Sub

Sub NewParty_RepartirExp(ByVal IndexParty As Integer)

' \ Author  : maTih.-
' \ Note    : Reparte la experiencia de todos los usuarios de party.

    With newParty(IndexParty)
        
        Dim loopX       As Long
        Dim expToUser   As Long
        
        expToUser = (.ExpAcumulada / .CantidadMiembros) + Porcentaje(.ExpAcumulada, NewParty_PorcentajeByLVL(.Nivel))
        
        For loopX = 1 To MAX_MIEMBROS
            'Recorremos todos los usuarios si es que hay.
            If .UserIndex(loopX) > 0 Then
                If UserList(.UserIndex(loopX)).ConnID <> -1 Then
                    'Les damos la experiencia
                    UserList(.UserIndex(loopX)).Stats.Exp = UserList(.UserIndex(loopX)).Stats.Exp + expToUser
                    
                    'Updateo de clientes.
                    
                    WriteUpdateExp .UserIndex(loopX)
                    
                    'Checkeo del nivel.
                    
                    CheckUserLevel .UserIndex(loopX)
                    
                    'Mensaje.
                    
                    WriteConsoleMsg .UserIndex(loopX), "Se repartió la experiencia de la party! Obtubiste : " & expToUser & " Puntos de experiencia.!", FontTypeNames.FONTTYPE_CITIZEN
                    
                End If
            End If
        Next loopX
            
    End With

End Sub

Function NewParty_PrepareStringHead(ByVal PartyIndex As Integer) As String

' \ Author  : maTih.-
' \ Note    : Devuelve el string con cabezas de los usuarios en la party.

Dim loopX       As Long
Dim prepareS    As String

With newParty(PartyIndex)

    For loopX = 1 To MAX_MIEMBROS
        
        If .UserIndex(loopX) > 0 Then
            If UserList(.UserIndex(loopX)).ConnID <> -1 Then
                prepareS = prepareS & UserList(.UserIndex(loopX)).Char.Head & "|"
            End If
        End If
    Next loopX

End With

NewParty_PrepareStringHead = prepareS

End Function

Function NewParty_PrepareStringName(ByVal PartyIndex As Integer) As String

' \ Author  : maTih.-
' \ Note    : Devuelve el string con los nombres de los usuarios en la party.

Dim loopX       As Long
Dim prepareS    As String

With newParty(PartyIndex)

    For loopX = 1 To MAX_MIEMBROS
        
        If .UserIndex(loopX) > 0 Then
            If UserList(.UserIndex(loopX)).ConnID <> -1 Then
                prepareS = prepareS & UserList(.UserIndex(loopX)).name & "|"
            End If
        End If
    Next loopX

End With

NewParty_PrepareStringName = prepareS

End Function

Function NewParty_PorcentajeByLVL(ByVal PartyNivel As Byte) As Byte

' \ Author  : maTih.-
' \ Note    : Devuelve el porcentaje extra segun el nivel de la party.

Dim endPorc         As Byte

    Select Case PartyNivel
        
            Case 1
            
            '% de party nivel 1.
            
            endPorc = 3
            
            Case 2
            
            '% de party nivel 2.
            
            endPorc = 6
            
            Case 3
            
            '% de party nivel 3.
            
            endPorc = 8
        
    End Select
    
    NewParty_PorcentajeByLVL = endPorc

End Function

Function NewParty_PosibleNivel(ByVal SkillMontaraz As Byte) As Byte

' \ Author  : maTih.-
' \ Note    : Devuelve el nivel para la party segun el skill del usuario.

Dim endLevel    As Byte

    Select Case SkillMontaraz
    
           'Manejo de el nivel que puede crear el usuario segun su skill.
           Case 20 To 49
        
           endLevel = 1
           
           Case 50 To 74
           
           endLevel = 2
           
           Case 75 To 100
           
           endLevel = 3
    
    End Select
    
NewParty_PosibleNivel = endLevel

End Function

Function NewParty_PuedeCrear(ByVal CreadorIndex As Integer, ByRef errorMsg As String) As Boolean

' \ Author  : maTih.-
' \ Note    : Devuelve Si creadorIndex puede crear una party.
    
    NewParty_PuedeCrear = False
    
    With UserList(CreadorIndex)
    
    'Si está muerto no puede crear una party.
    
    If .flags.Muerto = 1 Then
        errorMsg = "Estás muerto!!"
        Exit Function
    End If
    
    'Si el skill Es < 20 no puede.
    
    If NewParty_GetSkill(CreadorIndex) < 20 Then
        errorMsg = "Necesitas 20 skills en montaraz para poder iniciar una party!"
        Exit Function
    End If
    
    'Si ya está en una party.
    
    If .IndexParty > 0 Then
        errorMsg = "Ya integras una party!"
        Exit Function
    End If
    
    'Si los slots de partys están llenos.
    
    If UltimoPartyIndex >= MAX_PARTYS Then
        errorMsg = "No hay más slots para crear partys, por favor comuniquelo a un administrador."
        Exit Function
    End If
    
    'Estamos acá, entonces si puede.
    
    NewParty_PuedeCrear = True
    End With

End Function

Function NewParty_GetSkill(ByVal CreadorIndex As Integer) As Byte

' \ Author  : maTih.-
' \ Note    : Devuelve La cantidad de skills que tiene en Montaraz.
    
    'TODO : Cambiar cuando se agrege el skill.
    NewParty_GetSkill = UserList(CreadorIndex).Stats.UserSkills(eSkill.Supervivencia)

End Function

Function NewParty_PartyLlena(ByVal CantidadMiembros As Byte, ByVal NivelParty As Byte) As Boolean

' \ Author  : maTih.-
' \ Note    : Devuelve Si la party está llena segun su nivel y sus miembros.

Dim endBool         As Boolean

    Select Case NivelParty
    
           Case 1
            endBool = (CantidadMiembros >= 2)
            
           Case 2       'Nivel 2.
           
            endBool = (CantidadMiembros >= 3)
            
           Case 3       'Nivel 3.
            
            endBool = (CantidadMiembros >= 4)
            
    End Select
    
NewParty_PartyLlena = endBool

End Function
