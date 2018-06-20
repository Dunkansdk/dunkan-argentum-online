Attribute VB_Name = "modHechizos"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700



Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim hIndex As Integer
Dim j As Integer

With UserList(UserIndex)
    hIndex = ObjData(.Invent.Object(Slot).objIndex).HechizoIndex
    
    If Not TieneHechizo(hIndex, UserIndex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If .Stats.UserHechizos(j) = 0 Then Exit For
        Next j
            
        If .Stats.UserHechizos(j) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes espacio para m�s hechizos.", FontTypeNames.FONTTYPE_INFO)
        Else
            .Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If
End With

End Sub
            
Sub DecirPalabrasMagicas(ByVal SpellWords As String, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 17/11/2009
'25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
'17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
'***************************************************
On Error Resume Next
With UserList(UserIndex)
    If .flags.AdminInvisible <> 1 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(SpellWords, .Char.CharIndex, vbCyan, True))
        
        ' Si estaba oculto, se vuelve visible
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.invisible = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                Call SetInvisible(UserIndex, .Char.CharIndex, False)
            End If
        End If
    End If
End With
    Exit Sub
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010
'Last Modification By: ZaMa
'06/11/09 - Corregida la bonificaci�n de man� del mimetismo en el druida con flauta m�gica equipada.
'19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
'12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
'***************************************************
Dim DruidManaBonus As Single

    With UserList(UserIndex)
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
            
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If .Clase = eClass.Mage Then
                If .Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call WriteConsoleMsg(UserIndex, "No posees un b�culo lo suficientemente poderoso para poder lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un b�culo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
        
        DruidManaBonus = 1
        If .Clase = eClass.Druid Then
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                ' 50% menos de mana requerido para mimetismo
                If Hechizos(HechizoIndex).Mimetiza = 1 Then
                    DruidManaBonus = 0.5
                    
                ' 30% menos de mana requerido para invocaciones
                ElseIf Hechizos(HechizoIndex).tipo = uInvocacion Then
                    DruidManaBonus = 0.7
                
                ' 10% menos de mana requerido para las demas magias, excepto apoca
                ElseIf HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
                    DruidManaBonus = 0.9
                End If
            End If
            
            ' Necesita tener la barra de mana completa para invocar una mascota
            If Hechizos(HechizoIndex).Warp = 1 Then
                If .Stats.MinMAN <> .Stats.MaxMAN Then
                    Call WriteConsoleMsg(UserIndex, "Debes poseer toda tu man� para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                ' Si no tiene mascotas, no tiene sentido que lo use
                ElseIf .NroMascotas = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Debes poseer alguna mascota para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
        
        If .Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido * DruidManaBonus Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente man�.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
    End With
    
    PuedeLanzar = True
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer

    With UserList(UserIndex)
        PosCasteadaX = .flags.TargetX
        PosCasteadaY = .flags.TargetY
        PosCasteadaM = .flags.TargetMap
        
        H = .flags.Hechizo
        
        If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
            b = True
            For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
                For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                            'hay un user
                            If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
                            End If
                        End If
                    End If
                Next TempY
            Next TempX
        
            Call InfoHechizo(UserIndex)
        End If
    End With
End Sub


Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal spellIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    
    Dim HechizoCasteado As Boolean
    Dim ManaRequerida As Integer
    
    Select Case Hechizos(spellIndex).tipo
           
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)
    End Select

    If HechizoCasteado Then
        With UserList(UserIndex)
            
            ManaRequerida = Hechizos(spellIndex).ManaRequerido
            
            If Hechizos(spellIndex).Warp = 1 Then ' Invoc� una mascota
            ' Consume toda la mana
                ManaRequerida = .Stats.MinMAN
            Else
                ' Bonificaciones en hechizos
                If .Clase = eClass.Druid Then
                    ' Solo con flauta equipada
                    If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                        ' 30% menos de mana para invocaciones
                        ManaRequerida = ManaRequerida * 0.7
                    End If
                End If
            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Update user stats
        Call WriteUpdateMana(UserIndex)
        End With
    End If
    
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal spellIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010
'18/11/2009: ZaMa - Optimizacion de codigo.
'12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
'***************************************************
    
    Dim HechizoCasteado As Boolean
    Dim ManaRequerida As Integer
    
    Select Case Hechizos(spellIndex).tipo
        Case TipoHechizo.uEstado
            ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(UserIndex, HechizoCasteado)
        
        Case TipoHechizo.uPropiedades
            ' Afectan HP,MANA,STAMINA,ETC
            HechizoCasteado = HechizoPropUsuario(UserIndex)
    End Select

    If HechizoCasteado Then
        With UserList(UserIndex)
            
            ManaRequerida = Hechizos(spellIndex).ManaRequerido
            
            ' Bonificaciones para druida
            If .Clase = eClass.Druid Then
                ' Solo con flauta magica
                If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                    If Hechizos(spellIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        
                    ElseIf spellIndex <> APOCALIPSIS_SPELL_INDEX Then
                        ' 10% menos de mana para todo menos apoca y descarga
                        ManaRequerida = ManaRequerida * 0.9
                    End If
                End If
            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            Call WriteUpdateMana(UserIndex)
            .flags.targetUser = 0
        End With
    End If

End Sub

Sub LanzarHechizo(ByVal spellIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 02/16/2010
'24/01/2007 ZaMa - Optimizacion de codigo.
'02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
'***************************************************
On Error GoTo Errhandler

With UserList(UserIndex)
    
    If .flags.EnConsulta Then
        Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si est�s en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If PuedeLanzar(UserIndex, spellIndex) Then
        Select Case Hechizos(spellIndex).Target
            Case TargetType.uUsuarios
                If .flags.targetUser > 0 Then
                    If Abs(UserList(.flags.targetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, spellIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Este hechizo act�a s�lo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
                End If
            

            
            Case TargetType.uUsuariosYnpc
                If .flags.targetUser > 0 Then
                    If Abs(UserList(.flags.targetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, spellIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Target inv�lido.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case TargetType.uTerreno
                Call HandleHechizoTerreno(UserIndex, spellIndex)
        End Select
        
    End If
    
    If .Counters.Trabajando Then _
        .Counters.Trabajando = .Counters.Trabajando - 1
    
    If .Counters.Ocultando Then _
        .Counters.Ocultando = .Counters.Ocultando - 1

End With

Exit Sub

Errhandler:
    Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.Description & _
        " Hechizo: " & Hechizos(spellIndex).Nombre & "(" & spellIndex & _
        "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & ").")
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 28/04/2010
'Handles the Spells that afect the Stats of an User
'24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
'26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
'26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
'02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
'06/28/2008 NicoNZ - Agregu� que se le de valor al flag Inmovilizado.
'17/11/2008: NicoNZ - Agregado para quitar la penalizaci�n de vida en el ring y cambio de ecuacion.
'13/02/2009: ZaMa - Arreglada ecuacion para quitar vida tras resucitar en rings.
'23/11/2009: ZaMa - Optimizacion de codigo.
'28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
'***************************************************


Dim HechizoIndex As Integer
Dim targetIndex As Integer

With UserList(UserIndex)
    HechizoIndex = .flags.Hechizo
    targetIndex = .flags.targetUser
    
    ' <-------- Agrega Invisibilidad ---------->
    If Hechizos(HechizoIndex).Invisibilidad = 1 Then
        If UserList(targetIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "�El usuario est� muerto!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
        If UserList(targetIndex).Counters.Saliendo Then
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "�El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "�No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
                HechizoCasteado = False
                Exit Sub
            End If
        End If
        
        'NO VALE invi en el servidor.
        If Server_Info.Invisibilidad = False Then
           Call WriteConsoleMsg(UserIndex, "No est� permitido el hechizo en el servidor.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Sub
        End If
        
        'No usar invi mapas InviSinEfecto
        If MapInfo(UserList(targetIndex).Pos.map).InviSinEfecto > 0 Then
            Call WriteConsoleMsg(UserIndex, "�La invisibilidad no funciona aqu�!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        HechizoCasteado = CanSupportUser(UserIndex, targetIndex, True)
        If Not HechizoCasteado Then Exit Sub
        
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(targetIndex).flags.Privilegios And PlayerType.User Then
                HechizoCasteado = False
                Exit Sub
            End If
        End If
       
        UserList(targetIndex).flags.invisible = 1
        Call SetInvisible(targetIndex, UserList(targetIndex).Char.CharIndex, True)
    
        Call InfoHechizo(UserIndex)
        HechizoCasteado = True
    End If
    
    ' <-------- Agrega Mimetismo ---------->
    If Hechizos(HechizoIndex).Mimetiza = 1 Then
        If UserList(targetIndex).flags.Muerto = 1 Then
            Exit Sub
        End If
        
        If UserList(targetIndex).flags.Navegando = 1 Then
            Exit Sub
        End If
        If .flags.Navegando = 1 Then
            Exit Sub
        End If
        
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(targetIndex).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
        
        If .flags.Mimetizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.AdminInvisible = 1 Then Exit Sub
        
        'copio el char original al mimetizado
        
        .CharMimetizado.body = .Char.body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.body = UserList(targetIndex).Char.body
        .Char.Head = UserList(targetIndex).Char.Head
        .Char.CascoAnim = UserList(targetIndex).Char.CascoAnim
        .Char.ShieldAnim = UserList(targetIndex).Char.ShieldAnim
        .Char.WeaponAnim = GetWeaponAnim(UserIndex, UserList(targetIndex).Invent.WeaponEqpObjIndex)
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
       
       Call InfoHechizo(UserIndex)
       HechizoCasteado = True
    End If
    
    ' <-------- Agrega Envenenamiento ---------->
    If Hechizos(HechizoIndex).Envenena = 1 Then
        If UserIndex = targetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        UserList(targetIndex).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        HechizoCasteado = True
    End If
    
    ' <-------- Cura Envenenamiento ---------->
    If Hechizos(HechizoIndex).CuraVeneno = 1 Then
    
        'Verificamos que el usuario no este muerto
        If UserList(targetIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "�El usuario est� muerto!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        HechizoCasteado = CanSupportUser(UserIndex, targetIndex)
        If Not HechizoCasteado Then Exit Sub
            
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(targetIndex).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
            
        UserList(targetIndex).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        HechizoCasteado = True
    End If
    
    ' <-------- Agrega Maldicion ---------->
    If Hechizos(HechizoIndex).Maldicion = 1 Then
        If UserIndex = targetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        UserList(targetIndex).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        HechizoCasteado = True
    End If
    
    ' <-------- Remueve Maldicion ---------->
    If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
            UserList(targetIndex).flags.Maldicion = 0
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
    End If
    
    ' <-------- Agrega Bendicion ---------->
    If Hechizos(HechizoIndex).Bendicion = 1 Then
            UserList(targetIndex).flags.Bendicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
    End If
    
    ' <-------- Agrega Paralisis/Inmobilidad ---------->
    If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
        If UserIndex = targetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
         If UserList(targetIndex).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
            
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
            End If
            
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
            If UserList(targetIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(targetIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserIndex, " �El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(targetIndex)
                Exit Sub
            End If
            
            If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(targetIndex).flags.Inmovilizado = 1
            UserList(targetIndex).flags.Paralizado = 1
            UserList(targetIndex).Counters.Paralisis = IntervaloParalizado
            
            Call WriteParalizeOK(targetIndex)
            Call FlushBuffer(targetIndex)
        End If
    End If
    
    ' <-------- Remueve Paralisis/Inmobilidad ---------->
    If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
        ' Remueve si esta en ese estado
        If UserList(targetIndex).flags.Paralizado = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, targetIndex, True)
            If Not HechizoCasteado Then Exit Sub
            
            UserList(targetIndex).flags.Inmovilizado = 0
            UserList(targetIndex).flags.Paralizado = 0
            
            'no need to crypt this
            Call WriteParalizeOK(targetIndex)
            Call InfoHechizo(UserIndex)
        
        End If
    End If
    
    ' <-------- Remueve Estupidez (Aturdimiento) ---------->
    If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
    
        ' Remueve si esta en ese estado
        If UserList(targetIndex).flags.Estupidez = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, targetIndex)
            If Not HechizoCasteado Then Exit Sub
        
            UserList(targetIndex).flags.Estupidez = 0
            
            'no need to crypt this
            Call WriteDumbNoMore(targetIndex)
            Call FlushBuffer(targetIndex)
            Call InfoHechizo(UserIndex)
        
        End If
    End If
    
    ' <-------- Revive ---------->
    If Hechizos(HechizoIndex).Revivir = 1 Then
        If UserList(targetIndex).flags.Muerto = 1 Then
            
            'No vale resu en el servidor.
            If Server_Info.Resucitar = False Then
               Call WriteConsoleMsg(UserIndex, "No est� permitido revivir en el servidor.", FontTypeNames.FONTTYPE_FIGHT)
               Exit Sub
            End If
            
            'revisamos si necesita vara
            If .Clase = eClass.Mage Then
                If .Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call WriteConsoleMsg(UserIndex, "Necesitas un b�culo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    End If
                End If
            ElseIf .Clase = eClass.Bard Then
                If .Invent.AnilloEqpObjIndex <> LAUDELFICO And .Invent.AnilloEqpObjIndex <> LAUDMAGICO Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento m�gico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                End If
            ElseIf .Clase = eClass.Druid Then
                If .Invent.AnilloEqpObjIndex <> FLAUTAELFICA And .Invent.AnilloEqpObjIndex <> FLAUTAMAGICA Then
                    Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento m�gico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                End If
            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, targetIndex, True)
            If Not HechizoCasteado Then Exit Sub
    
            Dim EraCriminal As Boolean
            EraCriminal = criminal(UserIndex)
            
            If Not criminal(targetIndex) Then
                If targetIndex <> UserIndex Then
                    .Reputacion.NobleRep = .Reputacion.NobleRep + 500
                    If .Reputacion.NobleRep > MAXREP Then _
                        .Reputacion.NobleRep = MAXREP
                    Call WriteConsoleMsg(UserIndex, "�Los Dioses te sonr�en, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            If EraCriminal And Not criminal(UserIndex) Then
                Call RefreshCharStatus(UserIndex)
            End If
            
            With UserList(targetIndex)
                'Pablo Toxic Waste (GD: 29/04/07)
                .Stats.MinAGU = 0
                .flags.Sed = 1
                .Stats.MinHam = 0
                .flags.Hambre = 1
                Call InfoHechizo(UserIndex)
                .Stats.MinMAN = 0
                .Stats.MinSta = 0
            End With
            
            'Agregado para quitar la penalizaci�n de vida en el ring y cambio de ecuacion. (NicoNZ)
            If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                If .flags.Privilegios And PlayerType.User Then
                    .Stats.MinHp = .Stats.MinHp * (1 - UserList(targetIndex).Stats.ELV * 0.015)
                End If
            End If
            
            If (.Stats.MinHp <= 0) Then
                Call UserDie(UserIndex)
                Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
            Else
                Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = True
            End If
            
            If UserList(targetIndex).flags.Traveling = 1 Then
                UserList(targetIndex).Counters.goHome = 0
                UserList(targetIndex).flags.Traveling = 0
                'Call WriteConsoleMsg(TargetIndex, "Tu viaje ha sido cancelado.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteMultiMessage(targetIndex, eMessages.CancelHome)
            End If
            
            Call RevivirUsuario(targetIndex)
        Else
            HechizoCasteado = False
        End If
    
    End If
    
    ' <-------- Agrega Ceguera ---------->
    If Hechizos(HechizoIndex).Ceguera = 1 Then
        If UserIndex = targetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
            End If
            UserList(targetIndex).flags.Ceguera = 1
            UserList(targetIndex).Counters.Ceguera = IntervaloParalizado / 3

            Call FlushBuffer(targetIndex)
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
    End If
    
    ' <-------- Agrega Estupidez (Aturdimiento) ---------->
    If Hechizos(HechizoIndex).Estupidez = 1 Then
        If UserIndex = targetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
            End If
            If UserList(targetIndex).flags.Estupidez = 0 Then
                UserList(targetIndex).flags.Estupidez = 1
                UserList(targetIndex).Counters.Ceguera = IntervaloParalizado
            End If
            Call WriteDumb(targetIndex)
            Call FlushBuffer(targetIndex)
    
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
    End If
End With

End Sub



Sub InfoHechizo(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/07/2009
'25/07/2009: ZaMa - Code improvements.
'25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
'***************************************************
    Dim spellIndex As Integer
    Dim tUser As Integer
    Dim tNpc As Integer
    
    With UserList(UserIndex)
        spellIndex = .flags.Hechizo
        tUser = .flags.targetUser
        tNpc = .flags.targetNPC
        
        Call DecirPalabrasMagicas(Hechizos(spellIndex).PalabrasMagicas, UserIndex)
        
        If tUser > 0 Then
            ' Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible = 1 And UserIndex = tUser Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops))
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y))
            Else
                    'Envia hechizo a usuario - maTih.-
                    Call mod_DunkanGeneral.Enviar_HechizoAUser(UserIndex, tUser, Hechizos(spellIndex).EffectIndex, Hechizos(spellIndex).loops)
                    Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)
            End If
        ElseIf tNpc > 0 Then
                    'Envia hechizo a NPC - maTih.-
                Call mod_DunkanGeneral.Enviar_HechizoANpc(UserIndex, tNpc, Hechizos(spellIndex).EffectIndex, Hechizos(spellIndex).loops)
                Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y))
        End If
        
        If tUser > 0 Then
            If UserIndex <> tUser Then
                If .showName Then
                    Call WriteConsoleMsg(UserIndex, Hechizos(spellIndex).HechizeroMsg & " " & UserList(tUser).Name, FontTypeNames.FONTTYPE_FIGHT)
                Else
                    Call WriteConsoleMsg(UserIndex, Hechizos(spellIndex).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
                End If
                Call WriteConsoleMsg(tUser, .Name & " " & Hechizos(spellIndex).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, Hechizos(spellIndex).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
            End If
        ElseIf tNpc > 0 Then
            Call WriteConsoleMsg(UserIndex, Hechizos(spellIndex).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With

End Sub

Public Function HechizoPropUsuario(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 28/04/2010
'02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
'28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
'***************************************************

Dim spellIndex As Integer
Dim Da�o As Long
Dim targetIndex As Integer

spellIndex = UserList(UserIndex).flags.Hechizo
targetIndex = UserList(UserIndex).flags.targetUser
      
With UserList(targetIndex)
    If .flags.Muerto Then
        Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
          
    ' <-------- Aumenta Hambre ---------->
    If Hechizos(spellIndex).SubeHam = 1 Then
        
        Call InfoHechizo(UserIndex)
        
        Da�o = RandomNumber(Hechizos(spellIndex).MinHam, Hechizos(spellIndex).MaxHam)
        
        .Stats.MinHam = .Stats.MinHam + Da�o
        If .Stats.MinHam > .Stats.MaxHam Then _
            .Stats.MinHam = .Stats.MaxHam
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Da�o & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & Da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
    ' <-------- Quita Hambre ---------->
    ElseIf Hechizos(spellIndex).SubeHam = 2 Then
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        Else
            Exit Function
        End If
        
        Call InfoHechizo(UserIndex)
        
        Da�o = RandomNumber(Hechizos(spellIndex).MinHam, Hechizos(spellIndex).MaxHam)
        
        .Stats.MinHam = .Stats.MinHam - Da�o
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Da�o & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & Da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        If .Stats.MinHam < 1 Then
            .Stats.MinHam = 0
            .flags.Hambre = 1
        End If
        
    End If
    
    ' <-------- Aumenta Sed ---------->
    If Hechizos(spellIndex).SubeSed = 1 Then
        
        Call InfoHechizo(UserIndex)
        
        Da�o = RandomNumber(Hechizos(spellIndex).MinSed, Hechizos(spellIndex).MaxSed)
        
        .Stats.MinAGU = .Stats.MinAGU + Da�o
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        
 
             
        If UserIndex <> targetIndex Then
          Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Da�o & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
          Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & Da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        Else
          Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    
    ' <-------- Quita Sed ---------->
    ElseIf Hechizos(spellIndex).SubeSed = 2 Then
        
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        
        Call InfoHechizo(UserIndex)
        
        Da�o = RandomNumber(Hechizos(spellIndex).MinSed, Hechizos(spellIndex).MaxSed)
        
        .Stats.MinAGU = .Stats.MinAGU - Da�o
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Da�o & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & Da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        If .Stats.MinAGU < 1 Then
            .Stats.MinAGU = 0
            .flags.Sed = 1
        End If
        
 
        
    End If
    
    ' <-------- Aumenta Agilidad ---------->
    If Hechizos(spellIndex).SubeAgilidad = 1 Then
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, targetIndex) Then Exit Function
        
        Call InfoHechizo(UserIndex)
        Da�o = RandomNumber(Hechizos(spellIndex).MinAgilidad, Hechizos(spellIndex).MaxAgilidad)
        
        .flags.DuracionEfecto = 1200
        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + Da�o
        If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then _
            .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
        
        .flags.TomoPocion = True
    
    ' <-------- Quita Agilidad ---------->
    ElseIf Hechizos(spellIndex).SubeAgilidad = 2 Then
        
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        
        Call InfoHechizo(UserIndex)
        
        .flags.TomoPocion = True
        Da�o = RandomNumber(Hechizos(spellIndex).MinAgilidad, Hechizos(spellIndex).MaxAgilidad)
        .flags.DuracionEfecto = 700
        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - Da�o
        If .Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS

    End If
    
    ' <-------- Aumenta Fuerza ---------->
    If Hechizos(spellIndex).SubeFuerza = 1 Then
    
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, targetIndex) Then Exit Function
        
        Call InfoHechizo(UserIndex)
        Da�o = RandomNumber(Hechizos(spellIndex).MinFuerza, Hechizos(spellIndex).MaxFuerza)
        
        .flags.DuracionEfecto = 1200
    
        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + Da�o
        If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then _
            .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
        
        .flags.TomoPocion = True
    
    ' <-------- Quita Fuerza ---------->
    ElseIf Hechizos(spellIndex).SubeFuerza = 2 Then
    
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        
        Call InfoHechizo(UserIndex)
        
        .flags.TomoPocion = True
        
        Da�o = RandomNumber(Hechizos(spellIndex).MinFuerza, Hechizos(spellIndex).MaxFuerza)
        .flags.DuracionEfecto = 700
        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - Da�o
        If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS

    End If
    
    ' <-------- Cura salud ---------->
    If Hechizos(spellIndex).SubeHP = 1 Then
        
        'Verifica que el usuario no este muerto
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "�El usuario est� muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(UserIndex, targetIndex) Then Exit Function
           
        Da�o = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
        Da�o = Da�o + Porcentaje(Da�o, 3 * UserList(UserIndex).Stats.ELV)
        
        Call InfoHechizo(UserIndex)
    
        .Stats.MinHp = .Stats.MinHp + Da�o
        If .Stats.MinHp > .Stats.MaxHp Then _
            .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(targetIndex)
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Da�o & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & Da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    ' <-------- Quita salud (Da�a) ---------->
    ElseIf Hechizos(spellIndex).SubeHP = 2 Then
        
        If UserIndex = targetIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        
        Da�o = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
        
        Da�o = Da�o + Porcentaje(Da�o, 3 * UserList(UserIndex).Stats.ELV)
        
        If Hechizos(spellIndex).StaffAffected Then
            If UserList(UserIndex).Clase = eClass.Mage Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Da�o = (Da�o * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                Else
                    Da�o = Da�o * 0.7 'Baja da�o a 70% del original
                End If
            End If
        End If
        
        If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDELFICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
            Da�o = Da�o * 1.04  'laud magico de los bardos
        End If
        
        'cascos antimagia
        If (.Invent.CascoEqpObjIndex > 0) Then
            Da�o = Da�o - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        'anillos
        If (.Invent.AnilloEqpObjIndex > 0) Then
            Da�o = Da�o - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        If Da�o < 0 Then Da�o = 0
        
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        
        Call InfoHechizo(UserIndex)
        
        .Stats.MinHp = .Stats.MinHp - Da�o
        
        Call WriteUpdateHP(targetIndex)
        
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & Da�o & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & Da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        
        'Muere
        If .Stats.MinHp < 1 Then
        
            If .flags.AtacablePor <> UserIndex Then

                Call ContarMuerte(targetIndex, UserIndex)
            End If
            
            .Stats.MinHp = 0
            Call ActStats(targetIndex, UserIndex)
            Call UserDie(targetIndex)
        End If
        
    End If
    
    ' <-------- Aumenta Mana ---------->
    If Hechizos(spellIndex).SubeMana = 1 Then
        
        Call InfoHechizo(UserIndex)
        .Stats.MinMAN = .Stats.MinMAN + Da�o
        If .Stats.MinMAN > .Stats.MaxMAN Then _
            .Stats.MinMAN = .Stats.MaxMAN
        
        Call WriteUpdateMana(targetIndex)
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Da�o & " puntos de man� a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & Da�o & " puntos de man�.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Da�o & " puntos de man�.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    
    ' <-------- Quita Mana ---------->
    ElseIf Hechizos(spellIndex).SubeMana = 2 Then
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        
        Call InfoHechizo(UserIndex)
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Da�o & " puntos de man� a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & Da�o & " puntos de man�.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Da�o & " puntos de man�.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        .Stats.MinMAN = .Stats.MinMAN - Da�o
        If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
        
        Call WriteUpdateMana(targetIndex)
        
    End If
    
    ' <-------- Aumenta Stamina ---------->
    If Hechizos(spellIndex).SubeSta = 1 Then
        Call InfoHechizo(UserIndex)
        .Stats.MinSta = .Stats.MinSta + Da�o
        If .Stats.MinSta > .Stats.MaxSta Then _
            .Stats.MinSta = .Stats.MaxSta

        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Da�o & " puntos de energ�a a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & Da�o & " puntos de energ�a.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Da�o & " puntos de energ�a.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    ' <-------- Quita Stamina ---------->
    ElseIf Hechizos(spellIndex).SubeSta = 2 Then
        If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
        If UserIndex <> targetIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        End If
        
        Call InfoHechizo(UserIndex)
        
        If UserIndex <> targetIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & Da�o & " puntos de energ�a a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & Da�o & " puntos de energ�a.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & Da�o & " puntos de energ�a.", FontTypeNames.FONTTYPE_FIGHT)
        End If


        
    End If
End With

HechizoPropUsuario = True

Call FlushBuffer(targetIndex)

End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, ByVal targetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 28/04/2010
'Checks if caster can cast support magic on target user.
'***************************************************
     
 On Error GoTo Errhandler
 
    With UserList(CasterIndex)
        
        ' Te podes curar a vos mismo
        If CasterIndex = targetIndex Then
            CanSupportUser = True
            Exit Function
        End If
        
         ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, targetIndex) = TRIGGER6_PERMITE Then
            CanSupportUser = True
            Exit Function
        End If
     
        ' Victima criminal?
        If criminal(targetIndex) Then
        
            ' Casteador Ciuda?
            If Not criminal(CasterIndex) Then
            
                ' Armadas no pueden ayudar
                If esArmada(CasterIndex) Then
                    Call WriteConsoleMsg(CasterIndex, "Los miembros del ej�rcito real no pueden ayudar a los criminales.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                
                ' Si el ciuda tiene el seguro puesto no puede ayudar
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else
                    ' Penalizacion
                    If DoCriminal Then
                        Call VolverCriminal(CasterIndex)
                    Else
                        Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)
                    End If
                End If
            End If
            
        ' Victima ciuda o army
        Else
            ' Casteador es caos? => No Pueden ayudar ciudas
            If esCaos(CasterIndex) Then
                Call WriteConsoleMsg(CasterIndex, "Los miembros de la legi�n oscura no pueden ayudar a los ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
                
            ' Casteador ciuda/army?
            ElseIf Not criminal(CasterIndex) Then
                
                ' Esta en estado atacable?
                If UserList(targetIndex).flags.AtacablePor > 0 Then
                    
                    ' No esta atacable por el casteador?
                    If UserList(targetIndex).flags.AtacablePor <> CasterIndex Then
                    
                        ' Si es armada no puede ayudar
                        If esArmada(CasterIndex) Then
                            Call WriteConsoleMsg(CasterIndex, "Los miembros del ej�rcito real no pueden ayudar a ciudadanos en estado atacable.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function
                        End If
    
                        ' Seguro puesto?
                        If .flags.Seguro Then
                            Call WriteConsoleMsg(CasterIndex, "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function
                        Else
                            Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)
                        End If
                    End If
                End If
    
            End If
        End If
    End With
    
    CanSupportUser = True

    Exit Function
    
Errhandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.Number & " - " & Err.Description & _
                  " CasterIndex: " & CasterIndex & ", TargetIndex: " & targetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim loopC As Byte

With UserList(UserIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
        'Actualiza el inventario
        If .Stats.UserHechizos(Slot) > 0 Then
            Call ChangeUserHechizo(UserIndex, Slot, .Stats.UserHechizos(Slot))
        Else
            Call ChangeUserHechizo(UserIndex, Slot, 0)
        End If
    Else
        'Actualiza todos los slots
        For loopC = 1 To MAXUSERHECHIZOS
            'Actualiza el inventario
            If .Stats.UserHechizos(loopC) > 0 Then
                Call ChangeUserHechizo(UserIndex, loopC, .Stats.UserHechizos(loopC))
            Else
                Call ChangeUserHechizo(UserIndex, loopC, 0)
            End If
        Next loopC
    End If
End With

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    
    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
    
    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call WriteChangeSpellSlot(UserIndex, Slot)
    Else
        Call WriteChangeSpellSlot(UserIndex, Slot)
    End If

End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal HechizoDesplazado As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

With UserList(UserIndex)
    If Dire = 1 Then 'Mover arriba
        If HechizoDesplazado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direcci�n.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
            .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado - 1)
            .Stats.UserHechizos(HechizoDesplazado - 1) = TempHechizo
        End If
    Else 'mover abajo
        If HechizoDesplazado = MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direcci�n.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
            .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado + 1)
            .Stats.UserHechizos(HechizoDesplazado + 1) = TempHechizo
        End If
    End If
End With

End Sub

Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    With UserList(UserIndex)
        'Si estamos en la arena no hacemos nada
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            'pierdo nobleza...
            .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts
            If .Reputacion.NobleRep < 0 Then
                .Reputacion.NobleRep = 0
            End If
            
            'gano bandido...
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts
            If .Reputacion.BandidoRep > MAXREP Then _
                .Reputacion.BandidoRep = MAXREP
            Call WriteMultiMessage(UserIndex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)

        End If
        
        If Not EraCriminal And criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub
