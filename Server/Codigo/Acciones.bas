Attribute VB_Name = "Acciones"
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tempIndex As Integer
    
On Error Resume Next
    '�Rango Visi�n? (ToxicWaste)
    If (Abs(UserList(userIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(userIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    'ai barco?
    'If MapData(map, X, Y).EmbarcacionIndex <> 0 Then
       'Anota.
       'Call mod_Embarcaciones.Pasajero(userIndex, MapData(map, X, Y).EmbarcacionIndex)
    'End If
    
    '�Posicion valida?
    If InMapBounds(Map, X, Y) Then
        With UserList(userIndex)
        
            If MapData(Map, X, Y).userIndex <> 0 Then
                tempIndex = MapData(Map, X, Y).userIndex

                
            End If
                
            If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(Map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.targetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    '�Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(userIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(userIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(userIndex)
                
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(userIndex, "El sacerdote no puede curarte debido a que est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(userIndex)) Then
                        Call RevivirUsuario(userIndex)
                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(userIndex) Then
                        'curamos totalmente
                        .Stats.MinHp = .Stats.MaxHp
                        Call WriteUpdateUserStats(userIndex)
                    End If
                End If
                
            '�Es un obj?
            ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetOBJ = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, userIndex)
                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(Map, X, Y, userIndex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(Map, X, Y, userIndex)
                    Case eOBJType.otLe�a    'Le�a
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(Map, X, Y, userIndex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetOBJ = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y, userIndex)
                    
                End Select
            
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetOBJ = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y + 1, userIndex)
                End Select
            
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetOBJ = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y + 1, userIndex)
                End Select
            End If
        End With
    End If
End Sub

Public Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Agrego foros faccionarios
'***************************************************

On Error Resume Next

    Dim Pos As WorldPos
    
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    
    If Distancia(Pos, UserList(userIndex).Pos) > 2 Then
        Call WriteConsoleMsg(userIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If SendPosts(userIndex, ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ForoID) Then
        Call WriteShowForumForm(userIndex)
    End If
    
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

If Not (Distance(UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, Map, X, Y, 0)
                    Call Bloquear(True, Map, X - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(userIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, Map, X - 1, Y, 1)
                Call Bloquear(True, Map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(userIndex).flags.TargetOBJ = MapData(Map, X, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(userIndex, "La puerta est� cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(userIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(userIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

With UserList(userIndex)
    If Distancia(Pos, .Pos) > 2 Then
        Call WriteConsoleMsg(userIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
        Call WriteConsoleMsg(userIndex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
        Suerte = 1
    End If
    
    exito = RandomNumber(1, Suerte)
    
    If exito = 1 Then
        If MapInfo(.Pos.Map).Zona <> Ciudad Then
            Obj.ObjIndex = FOGATA
            Obj.amount = 1
            
            Call WriteConsoleMsg(userIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
            Call MakeObj(Obj, Map, X, Y)
            
            'Las fogatas prendidas se deben eliminar
            Dim Fogatita As New cGarbage
            Fogatita.Map = Map
            Fogatita.X = X
            Fogatita.Y = Y
            Call TrashCollector.Add(Fogatita)
            
            Call SubirSkill(userIndex, eSkill.Supervivencia, True)
        Else
            Call WriteConsoleMsg(userIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(userIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(userIndex, eSkill.Supervivencia, False)
    End If

End With

End Sub
