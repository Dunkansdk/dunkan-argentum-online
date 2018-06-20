Attribute VB_Name = "InvUsuario"
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

Option Explicit

Public Function TieneObjetosRobables(ByVal userIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim objIndex As Integer

For i = 1 To UserList(userIndex).CurrentInventorySlots
    objIndex = UserList(userIndex).Invent.Object(i).objIndex
    If objIndex > 0 Then
            If (ObjData(objIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(objIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i
End Function

Function ClasePuedeUsarItem(ByVal userIndex As Integer, ByVal objIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo manejador

    Dim flag As Boolean
    
    'Admins can use ANYTHING!
    If UserList(userIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(objIndex).ClaseProhibida(1) <> 0 Then
            Dim i As Integer
            For i = 1 To NUMCLASES
                If ObjData(objIndex).ClaseProhibida(i) = UserList(userIndex).Clase Then
                    ClasePuedeUsarItem = False
                    sMotivo = "Tu clase no puede usar este objeto."
                    Exit Function
                End If
            Next i
        End If
    End If
    
    ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Integer

With UserList(userIndex)
    For j = 1 To UserList(userIndex).CurrentInventorySlots
        If .Invent.Object(j).objIndex > 0 Then
             
             If ObjData(.Invent.Object(j).objIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(userIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, userIndex, j)
        
        End If
    Next j
    
    '[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
    'es transportado a su hogar de origen ;)
    If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
        
        Dim DeDonde As WorldPos
        
        Select Case .Hogar
            Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                DeDonde = Lindos
            Case eCiudad.cUllathorpe
                DeDonde = Ullathorpe
            Case eCiudad.cBanderbill
                DeDonde = Banderbill
            Case Else
                DeDonde = Nix
        End Select
        
        Call WarpUserChar(userIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)
    
    End If
    '[/Barrin]
End With

End Sub

Sub LimpiarInventario(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Integer

With UserList(userIndex)
    For j = 1 To .CurrentInventorySlots
        .Invent.Object(j).objIndex = 0
        .Invent.Object(j).Amount = 0
        .Invent.Object(j).Equipped = 0
    Next j

    .Invent.NroItems = 0
    
    .Invent.ArmourEqpObjIndex = 0
    .Invent.ArmourEqpSlot = 0
    
    .Invent.WeaponEqpObjIndex = 0
    .Invent.WeaponEqpSlot = 0
    
    .Invent.CascoEqpObjIndex = 0
    .Invent.CascoEqpSlot = 0
    
    .Invent.EscudoEqpObjIndex = 0
    .Invent.EscudoEqpSlot = 0
    
    .Invent.AnilloEqpObjIndex = 0
    .Invent.AnilloEqpSlot = 0
    
    .Invent.MunicionEqpObjIndex = 0
    .Invent.MunicionEqpSlot = 0
    
    .Invent.BarcoObjIndex = 0
    .Invent.BarcoSlot = 0
    
    .Invent.MochilaEqpObjIndex = 0
    .Invent.MochilaEqpSlot = 0

End With

End Sub

Sub QuitarUserInvItem(ByVal userIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    If Slot < 1 Or Slot > UserList(userIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(userIndex).Invent.Object(Slot)
        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(userIndex, Slot)
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(userIndex).Invent.NroItems = UserList(userIndex).Invent.NroItems - 1
            .objIndex = 0
            .Amount = 0
        End If
    End With

Exit Sub

Errhandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

Dim NullObj As UserOBJ
Dim loopC As Long

With UserList(userIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
    
        'Actualiza el inventario
        If .Invent.Object(Slot).objIndex > 0 Then
            Call ChangeUserInv(userIndex, Slot, .Invent.Object(Slot))
        Else
            Call ChangeUserInv(userIndex, Slot, NullObj)
        End If
    
    Else
    
    'Actualiza todos los slots
        For loopC = 1 To .CurrentInventorySlots
            'Actualiza el inventario
            If .Invent.Object(loopC).objIndex > 0 Then
                Call ChangeUserInv(userIndex, loopC, .Invent.Object(loopC))
            Else
                Call ChangeUserInv(userIndex, loopC, NullObj)
            End If
        Next loopC
    End If
    
    Exit Sub
End With

Errhandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.Description)

End Sub

Sub DropObj(ByVal userIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Obj As Obj

With UserList(userIndex)
    If num > 0 Then
    
        If num > .Invent.Object(Slot).Amount Then num = .Invent.Object(Slot).Amount
        
        Obj.objIndex = .Invent.Object(Slot).objIndex
        Obj.Amount = num
        
        If (ItemNewbie(Obj.objIndex) And (.flags.Privilegios And PlayerType.User)) Then
            Call WriteConsoleMsg(userIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Check objeto en el suelo
        If MapData(.Pos.map, X, Y).ObjInfo.objIndex = 0 Or MapData(.Pos.map, X, Y).ObjInfo.objIndex = Obj.objIndex Then
            If num + MapData(.Pos.map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
                num = MAX_INVENTORY_OBJS - MapData(.Pos.map, X, Y).ObjInfo.Amount
            End If
            
            
            Call MakeObj(Obj, map, X, Y)
            Call QuitarUserInvItem(userIndex, Slot, num)
            Call UpdateUserInv(False, userIndex, Slot)
            
            If ObjData(Obj.objIndex).OBJType = eOBJType.otBarcos Then
                Call WriteConsoleMsg(userIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
            End If
            
            'Agrega a la lista de objetos - maTih.-
            Dim CleanPos    As WorldPos
            
            CleanPos.map = .Pos.map
            CleanPos.X = X
            CleanPos.Y = Y
                        
            Call mod_DunkanLimpieza.Agregar(CleanPos)
            
            If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiró cantidad:" & num & " Objeto:" & ObjData(Obj.objIndex).Name)
            
            'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(Obj.objIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.objIndex).Name & " Mapa: " & map & " X: " & X & " Y: " & Y)
            ElseIf Obj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(Obj.objIndex).NoLog <> 1 Then
                    Call LogDesarrollo(.Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.objIndex).Name & " Mapa: " & map & " X: " & X & " Y: " & Y)
                End If
            End If
        Else
            Call WriteConsoleMsg(userIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End With

End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

With MapData(map, X, Y)
    .ObjInfo.Amount = .ObjInfo.Amount - num
    
    If .ObjInfo.Amount <= 0 Then
        .ObjInfo.objIndex = 0
        .ObjInfo.Amount = 0
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
    End If
End With

End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    
    If Obj.objIndex > 0 And Obj.objIndex <= UBound(ObjData) Then
    
        With MapData(map, X, Y)
            If .ObjInfo.objIndex = Obj.objIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
            Else
                .ObjInfo = Obj
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.objIndex).GrhIndex, X, Y))
            End If
        End With
    End If

End Sub

Function MeterItemEnInventario(ByVal userIndex As Integer, ByRef MiObj As Obj) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim X As Integer
    Dim Y As Integer
    Dim Slot As Byte
    
    With UserList(userIndex)
        '¿el user ya tiene un objeto del mismo tipo?
        Slot = 1
        Do Until .Invent.Object(Slot).objIndex = MiObj.objIndex And _
                 .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
           Slot = Slot + 1
           If Slot > .CurrentInventorySlots Then
                 Exit Do
           End If
        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
           Slot = 1
           Do Until .Invent.Object(Slot).objIndex = 0
               Slot = Slot + 1
               If Slot > .CurrentInventorySlots Then
                   Call WriteConsoleMsg(userIndex, "No puedes cargar más objetos.", FontTypeNames.FONTTYPE_FIGHT)
                   MeterItemEnInventario = False
                   Exit Function
               End If
           Loop
           .Invent.NroItems = .Invent.NroItems + 1
        End If
    
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot < MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.objIndex) Then
                Call WriteConsoleMsg(userIndex, "No puedes contener objetos especiales en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                MeterItemEnInventario = False
                Exit Function
            End If
        End If
        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
           'Menor que MAX_INV_OBJS
           .Invent.Object(Slot).objIndex = MiObj.objIndex
           .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
           .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
        End If
    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, userIndex, Slot)
    
    
    Exit Function
Errhandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.Description)
End Function

Sub GetObj(ByVal userIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 18/12/2009
'18/12/2009: ZaMa - Oro directo a la billetera.
'***************************************************

    Dim Obj As ObjData
    Dim MiObj As Obj
    Dim ObjPos As String
    
    With UserList(userIndex)
        '¿Hay algun obj?
        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objIndex > 0 Then
            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objIndex).Agarrable <> 1 Then
                Dim X As Integer
                Dim Y As Integer
                Dim Slot As Byte
                
                X = .Pos.X
                Y = .Pos.Y
                
                Obj = ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objIndex)
                MiObj.Amount = MapData(.Pos.map, X, Y).ObjInfo.Amount
                MiObj.objIndex = MapData(.Pos.map, X, Y).ObjInfo.objIndex

                ' Oro directo a la billetera!
                If Obj.OBJType = otGuita Then
                    .Stats.GLD = .Stats.GLD + MiObj.Amount
                    'Lukea oro, actualizo ranking ***
                    Dim targetRankPos As Byte
        
                    targetRankPos = mod_DunkanRankings.IngresaOro(userIndex)
        
                    If targetRankPos <> 0 Then Call mod_DunkanRankings.ActualizarOros(userIndex, targetRankPos)
                    'Quitamos el objeto
                    Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)

                Else
                    If MeterItemEnInventario(userIndex, MiObj) Then
                    
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)
                        If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.objIndex).Name)
        
                        'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                        'Es un Objeto que tenemos que loguear?
                        If ObjData(MiObj.objIndex).Log = 1 Then
                            ObjPos = " Mapa: " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.objIndex).Name & ObjPos)
                        ElseIf MiObj.Amount > MAX_INVENTORY_OBJS - 1000 Then 'Es mucha cantidad?
                            'Si no es de los prohibidos de loguear, lo logueamos.
                            If ObjData(MiObj.objIndex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.objIndex).Name & ObjPos)
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(userIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With

End Sub

Sub Desequipar(ByVal userIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(userIndex)
        With .Invent
            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).objIndex = 0 Then
                Exit Sub
            End If
            
            Obj = ObjData(.Object(Slot).objIndex)
        End With
        
        Select Case Obj.OBJType
            Case eOBJType.otWeapon
                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(userIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otFlechas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
            
            Case eOBJType.otAnillo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0
                End With
            
            Case eOBJType.otArmadura
                With .Invent
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0
                End With
                
                Call DarCuerpoDesnudo(userIndex, .flags.Mimetizado = 1)
                With .Char
                    Call ChangeUserChar(userIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                End With
                 
            Case eOBJType.otCASCO
                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(userIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otESCUDO
                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(userIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otMochilas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0
                End With
                
                Call InvUsuario.TirarTodosLosItemsEnMochila(userIndex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End Select
    End With
    
    Call WriteUpdateHP(userIndex)
    Call WriteUpdateMana(userIndex)
    Call UpdateUserInv(False, userIndex, Slot)
    
    Exit Sub

Errhandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.Description)

End Sub

Function SexoPuedeUsarItem(ByVal userIndex As Integer, ByVal objIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo Errhandler
    
    If ObjData(objIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(userIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(objIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(userIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."
    
    Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal userIndex As Integer, ByVal objIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

    If ObjData(objIndex).Real = 1 Then
        If Not criminal(userIndex) Then
            FaccionPuedeUsarItem = esArmada(userIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(objIndex).Caos = 1 Then
        If criminal(userIndex) Then
            FaccionPuedeUsarItem = esCaos(userIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
    
    If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineación no puede usar este objeto."

End Function

Sub EquiparInvItem(ByVal userIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 14/01/2010 (ZaMa)
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
'14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
'*************************************************

On Error GoTo Errhandler

    'Equipa un item del inventario
    Dim Obj As ObjData
    Dim objIndex As Integer
    Dim sMotivo As String
    
    With UserList(userIndex)
        objIndex = .Invent.Object(Slot).objIndex
        Obj = ObjData(objIndex)
        
        If Obj.Newbie = 1 And Not EsNewbie(userIndex) Then
             Call WriteConsoleMsg(userIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
                
        Select Case Obj.OBJType
            Case eOBJType.otWeapon
               If ClasePuedeUsarItem(userIndex, objIndex, sMotivo) And _
                  FaccionPuedeUsarItem(userIndex, objIndex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(userIndex, Slot)
                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(userIndex, .Invent.WeaponEqpSlot)
                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = objIndex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then _
                        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(userIndex, objIndex)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(userIndex, objIndex)
                        Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
               Else
                    Call WriteConsoleMsg(userIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otAnillo
               If ClasePuedeUsarItem(userIndex, objIndex, sMotivo) And _
                  FaccionPuedeUsarItem(userIndex, objIndex, sMotivo) Then
                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(userIndex, Slot)
                            
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.AnilloEqpObjIndex > 0 Then
                            Call Desequipar(userIndex, .Invent.AnilloEqpSlot)
                        End If
                
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.AnilloEqpObjIndex = objIndex
                        .Invent.AnilloEqpSlot = Slot
               Else
                    Call WriteConsoleMsg(userIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otFlechas
               If ClasePuedeUsarItem(userIndex, objIndex, sMotivo) And _
                  FaccionPuedeUsarItem(userIndex, objIndex, sMotivo) Then
                        
                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(userIndex, Slot)
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.MunicionEqpObjIndex > 0 Then
                            Call Desequipar(userIndex, .Invent.MunicionEqpSlot)
                        End If
                
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.MunicionEqpObjIndex = objIndex
                        .Invent.MunicionEqpSlot = Slot
                        
               Else
                    Call WriteConsoleMsg(userIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otArmadura
                If .flags.Navegando = 1 Then Exit Sub
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(userIndex, objIndex, sMotivo) And _
                   SexoPuedeUsarItem(userIndex, objIndex, sMotivo) And _
                   CheckRazaUsaRopa(userIndex, objIndex, sMotivo) And _
                   FaccionPuedeUsarItem(userIndex, objIndex, sMotivo) Then
                   
                   'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(userIndex, Slot)
                        Call DarCuerpoDesnudo(userIndex, .flags.Mimetizado = 1)
                        If Not .flags.Mimetizado = 1 Then
                            Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(userIndex, .Invent.ArmourEqpSlot)
                    End If
            
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = objIndex
                    .Invent.ArmourEqpSlot = Slot
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.body = Obj.Ropaje
                    Else
                        .Char.body = Obj.Ropaje
                        Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    .flags.Desnudo = 0
                Else
                    Call WriteConsoleMsg(userIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otCASCO
                If .flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(userIndex, objIndex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(userIndex, Slot)
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = objIndex
                    .Invent.CascoEqpSlot = Slot
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = Obj.CascoAnim
                    Else
                        .Char.CascoAnim = Obj.CascoAnim
                        Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otESCUDO
                If .flags.Navegando = 1 Then Exit Sub
                
                 If ClasePuedeUsarItem(userIndex, objIndex, sMotivo) And _
                     FaccionPuedeUsarItem(userIndex, objIndex, sMotivo) Then
        
                     'Si esta equipado lo quita
                     If .Invent.Object(Slot).Equipped Then
                         Call Desequipar(userIndex, Slot)
                         If .flags.Mimetizado = 1 Then
                             .CharMimetizado.ShieldAnim = NingunEscudo
                         Else
                             .Char.ShieldAnim = NingunEscudo
                             Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                         End If
                         Exit Sub
                     End If
             
                     'Quita el anterior
                     If .Invent.EscudoEqpObjIndex > 0 Then
                         Call Desequipar(userIndex, .Invent.EscudoEqpSlot)
                     End If
             
                     'Lo equipa
                     
                     .Invent.Object(Slot).Equipped = 1
                     .Invent.EscudoEqpObjIndex = objIndex
                     .Invent.EscudoEqpSlot = Slot
                     If .flags.Mimetizado = 1 Then
                         .CharMimetizado.ShieldAnim = Obj.ShieldAnim
                     Else
                         .Char.ShieldAnim = Obj.ShieldAnim
                         
                         Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                     End If
                 Else
                     Call WriteConsoleMsg(userIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                 End If
                 
            Case eOBJType.otMochilas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(userIndex, Slot)
                    Exit Sub
                End If
                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(userIndex, .Invent.MochilaEqpSlot)
                End If
                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = objIndex
                .Invent.MochilaEqpSlot = Slot
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                Call WriteAddSlots(userIndex, Obj.MochilaType)
        End Select
    End With
    
    'Actualiza
    Call UpdateUserInv(False, userIndex, Slot)
    
    Exit Sub
    
Errhandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description)
End Sub

Public Function CheckRazaUsaRopa(ByVal userIndex As Integer, ItemIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo Errhandler

    With UserList(userIndex)
        'Verifica si la raza puede usar la ropa
        If .Raza = eRaza.Humano Or _
           .Raza = eRaza.Elfo Or _
           .Raza = eRaza.Drow Then
                CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
                CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.Raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False
        End If
    End With
    
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
    Exit Function
    
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal userIndex As Integer, ByVal Slot As Byte, Optional esU As Boolean = False)
'*************************************************
'Author: Unknown
'Last modified: 10/12/2009
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
'17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
'27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
'08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
'10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
'*************************************************

    Dim Obj As ObjData
    Dim objIndex As Integer
    Dim TargObj As ObjData
    Dim MiObj As Obj
    
    With UserList(userIndex)

        If .Invent.Object(Slot).Amount = 0 Then Exit Sub

        Obj = ObjData(.Invent.Object(Slot).objIndex)

        If Obj.Newbie = 1 And Not EsNewbie(userIndex) Then
            Call WriteConsoleMsg(userIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.OBJType = eOBJType.otWeapon Then
            If Obj.proyectil = 1 Then
                
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(userIndex, False) Then Exit Sub
            Else
                'dagas
                If Not IntervaloPermiteUsar(userIndex) Then Exit Sub
            End If
        Else
            If Not IntervaloPermiteUsar(userIndex) Then Exit Sub
        End If
        
        objIndex = .Invent.Object(Slot).objIndex
        .flags.TargetObjInvIndex = objIndex
        .flags.TargetObjInvSlot = Slot
        
        Select Case Obj.OBJType
            Case eOBJType.otUseOnce
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + Obj.MinHam
                If .Stats.MinHam > .Stats.MaxHam Then _
                    .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0

                'Sonido
                
                If objIndex = e_ObjetosCriticos.Manzana Or objIndex = e_ObjetosCriticos.Manzana2 Or objIndex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.SOUND_COMIDA)
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                
                Call UpdateUserInv(False, userIndex, Slot)
        
            Case eOBJType.otGuita
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).objIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, userIndex, Slot)
                
            Case eOBJType.otWeapon
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(objIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(userIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteMultiMessage(userIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)

                End If
            
            Case eOBJType.otPociones
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not IntervaloPermiteGolpeUsar(userIndex, False) Then
                    Call WriteConsoleMsg(userIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .flags.TomoPocion = True
                .flags.TipoPocion = Obj.TipoPocion
                        
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                            .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(userIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                            .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(userIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case 3 'Pocion roja, restaura HP
                        'Usa el item
                        .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        If .Stats.MinHp > .Stats.MaxHp Then _
                            .Stats.MinHp = .Stats.MaxHp
                        
                        QuitarUserInvItem userIndex, Slot, 1
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                    
                    WriteUpdateHP userIndex
                    
                    Case 4 'Pocion azul, restaura MANA
                        'Usa el item
                        'nuevo calculo para recargar mana
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                        If .Stats.MinMAN > .Stats.MaxMAN Then _
                            .Stats.MinMAN = .Stats.MaxMAN
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                        QuitarUserInvItem userIndex, Slot, 1
                        
                        WriteUpdateMana userIndex
                        
                    Case 5 ' Pocion violeta
                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(userIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(userIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case 6  ' Pocion Negra
                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(userIndex, Slot, 1)
                            Call UserDie(userIndex)
                            Call WriteConsoleMsg(userIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                        End If
               End Select
               
               Call UpdateUserInv(False, userIndex, Slot)
        
             Case eOBJType.otBebidas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
 
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
                Call UpdateUserInv(False, userIndex, Slot)
            
            Case eOBJType.otLlaves
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.TargetOBJ = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetOBJ)
                '¿El objeto clickeado es una puerta?
                If TargObj.OBJType = eOBJType.otPuertas Then
                    '¿Esta cerrada?
                    If TargObj.Cerrada = 1 Then
                          '¿Cerrada con llave?
                          If TargObj.Llave > 0 Then
                             If TargObj.clave = Obj.clave Then
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex).IndexCerrada
                                .flags.TargetOBJ = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex
                                Call WriteConsoleMsg(userIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(userIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          Else
                             If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(userIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetOBJ = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.objIndex
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(userIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          End If
                    Else
                          Call WriteConsoleMsg(userIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)
                          Exit Sub
                    End If
                End If
            
            Case eOBJType.otBotellaLlena
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
 
                MiObj.Amount = 1
                MiObj.objIndex = ObjData(.Invent.Object(Slot).objIndex).IndexCerrada
                Call QuitarUserInvItem(userIndex, Slot, 1)
                
                Call UpdateUserInv(False, userIndex, Slot)
            
            Case eOBJType.otPergaminos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Stats.MaxMAN > 0 Then
                    If .flags.Hambre = 0 And _
                        .flags.Sed = 0 Then
                        Call AgregarHechizo(userIndex, Slot)
                        Call UpdateUserInv(False, userIndex, Slot)
                    Else
                        Call WriteConsoleMsg(userIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(userIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
                End If
            Case eOBJType.otMinerales
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                Call WriteMultiMessage(userIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
            Case eOBJType.otInstrumentos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(userIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Obj.Real Then '¿Es el Cuerno Real?
                    If FaccionPuedeUsarItem(userIndex, objIndex) Then
                        If MapInfo(.Pos.map).Pk = False Then
                            Call WriteConsoleMsg(userIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(userIndex)
                            Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(userIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
                    If FaccionPuedeUsarItem(userIndex, objIndex) Then
                        If MapInfo(.Pos.map).Pk = False Then
                            Call WriteConsoleMsg(userIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(userIndex)
                            Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(userIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                'Si llega aca es porque es o Laud o Tambor o Flauta
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(userIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
               
            Case eOBJType.otBarcos
                'Verifica si esta aproximado al agua antes de permitirle navegar
                If .Stats.ELV < 25 Then
                    If .Clase <> eClass.Worker And .Clase <> eClass.Pirat Then
                        Call WriteConsoleMsg(userIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else
                        If .Stats.ELV < 20 Then
                            Call WriteConsoleMsg(userIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                End If
                
                If ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False)) _
                        And .flags.Navegando = 0) _
                        Or .flags.Navegando = 1 Then

                Else
                    Call WriteConsoleMsg(userIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
                End If
                
        End Select
    
    End With

End Sub


Sub TirarTodo(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    With UserList(userIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        Call TirarTodosLosItems(userIndex)
        
        Dim Cantidad As Long
        Cantidad = .Stats.GLD - CLng(.Stats.ELV) * 10000
        
    End With

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With ObjData(Index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                    (.Caos <> 1 Or .NoSeCae = 0) And _
                    .OBJType <> eOBJType.otLlaves And _
                    .OBJType <> eOBJType.otBarcos And _
                    .NoSeCae = 0
    End With

End Function

Sub TirarTodosLosItems(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
'***************************************************

    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    Dim DropAgua As Boolean
    
    With UserList(userIndex)
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).objIndex
            If ItemIndex > 0 Then
                 If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.objIndex = ItemIndex

                    DropAgua = True
                    ' Es pirata?
                    If .Clase = eClass.Pirat Then
                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 476 Then
                            ' Limitación por nivel, después dropea normalmente
                            If .Stats.ELV >= 20 And .Stats.ELV <= 25 Then
                                ' No dropea en agua
                                DropAgua = False
                            End If
                        End If
                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(userIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                 End If
            End If
        Next i
    End With
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 23/11/2009
'07/11/09: Pato - Fix bug #2819911
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(userIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = 1 To UserList(userIndex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).objIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.objIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(userIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Sub TirarTodosLosItemsEnMochila(ByVal userIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/09 (Budi)
'***************************************************
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(userIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).objIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.objIndex = ItemIndex
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(userIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Public Function getObjType(ByVal objIndex As Integer) As eOBJType
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If objIndex > 0 Then
        getObjType = ObjData(objIndex).OBJType
    End If
    
End Function
