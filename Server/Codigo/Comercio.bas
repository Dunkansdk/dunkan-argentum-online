Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio
    Compra = 1
    Venta = 2
End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, ByVal userIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
'  - 06/13/08 (NicoNZ)
'*************************************************
    Dim precio As Long
    Dim Objeto As Obj
    
    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Sub
        ElseIf Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(userIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call Ban(UserList(userIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & Cantidad)
            UserList(userIndex).flags.Ban = 1
            Call WriteErrorMsg(userIndex, "Has sido baneado por el Sistema AntiCheat.")
            Call FlushBuffer(userIndex)
            Call CloseSocket(userIndex)
            Exit Sub
        ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).amount > 0 Then
            Exit Sub
        End If
        
        If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).amount Then Cantidad = Npclist(UserList(userIndex).flags.targetNPC).Invent.Object(Slot).amount
        
        Objeto.amount = Cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        
        precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor / Descuento(userIndex) * Cantidad) + 0.5)

        If UserList(userIndex).Stats.GLD < precio Then
            Call WriteConsoleMsg(userIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If MeterItemEnInventario(userIndex, Objeto) = False Then
            'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
            Call WriteTradeOK(userIndex)
            Exit Sub
        End If
        
        UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD - precio
        
        Call QuitarNpcInvItem(UserList(userIndex).flags.targetNPC, CByte(Slot), Cantidad)
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(userIndex).Name & " compró del NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(userIndex).Name & " compró del NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(userIndex).Name & " compró " & ObjData(Objeto.ObjIndex).Name)
        End If
        
    ElseIf Modo = eModoComercio.Venta Then
        
        If Cantidad > UserList(userIndex).Invent.Object(Slot).amount Then Cantidad = UserList(userIndex).Invent.Object(Slot).amount
        
        Objeto.amount = Cantidad
        Objeto.ObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
        
        If Objeto.ObjIndex = 0 Then
            Exit Sub
        ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
            Call WriteConsoleMsg(userIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
            Call WriteTradeOK(userIndex)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then
            If Npclist(NpcIndex).Name <> "SR" Then
                Call WriteConsoleMsg(userIndex, "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
                Call WriteTradeOK(userIndex)
                Exit Sub
            End If
        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then
            If Npclist(NpcIndex).Name <> "SC" Then
                Call WriteConsoleMsg(userIndex, "Las armaduras de la legión oscura sólo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
                Call WriteTradeOK(userIndex)
                Exit Sub
            End If
        ElseIf UserList(userIndex).Invent.Object(Slot).amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(userIndex).Invent.Object()) Or Slot > UBound(UserList(userIndex).Invent.Object()) Then
            Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
            Exit Sub
        ElseIf UserList(userIndex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(userIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
            Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
            Call WriteTradeOK(userIndex)
            Exit Sub
        End If
        
        Call QuitarUserInvItem(userIndex, Slot, Cantidad)
        
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + precio
        
        'Vende objeto, actualizo ranking ***
        Dim targetRankPos As Byte
        
        targetRankPos = mod_DunkanRankings.IngresaOro(userIndex)
        
        If targetRankPos <> 0 Then Call mod_DunkanRankings.ActualizarOros(userIndex, targetRankPos)
        
        If UserList(userIndex).Stats.GLD > MAXORO Then _
            UserList(userIndex).Stats.GLD = MAXORO
        
        Dim NpcSlot As Integer
        NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.amount)
        
        If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            Npclist(NpcIndex).Invent.Object(NpcSlot).amount = Npclist(NpcIndex).Invent.Object(NpcSlot).amount + Objeto.amount
            If Npclist(NpcIndex).Invent.Object(NpcSlot).amount > MAX_INVENTORY_OBJS Then
                Npclist(NpcIndex).Invent.Object(NpcSlot).amount = MAX_INVENTORY_OBJS
            End If
        End If
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(userIndex).Name & " vendió al NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(userIndex).Name & " vendió al NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
        
    End If
    
    Call UpdateUserInv(True, userIndex, 0)
    Call WriteUpdateUserStats(userIndex)
    Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
    Call WriteTradeOK(userIndex)
        
    Call SubirSkill(userIndex, eSkill.Comerciar, True)
End Sub


Public Sub IniciarComercioNPC(ByVal userIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Call EnviarNpcInv(userIndex, UserList(userIndex).flags.targetNPC)
    UserList(userIndex).flags.Comerciando = True

End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    SlotEnNPCInv = 1
    Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto _
      And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).amount + Cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1
        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
            SlotEnNPCInv = SlotEnNPCInv + 1
            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    
    End If
    
End Function

Private Function Descuento(ByVal userIndex As Integer) As Single
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Descuento = 1 + UserList(userIndex).Stats.UserSkills(eSkill.Comerciar) / 100
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal userIndex As Integer, ByVal NpcIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last Modified: 06/14/08
'Last Modified By: Nicolás Ezequiel Bouhid (NicoNZ)
'*************************************************
    Dim Slot As Byte
    Dim val As Single
    Dim Usable As Boolean
    
    For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then
            Dim thisObj As Obj
            
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
            thisObj.amount = Npclist(NpcIndex).Invent.Object(Slot).amount
            
            val = (ObjData(thisObj.ObjIndex).Valor) / Descuento(userIndex)
            Usable = ClasePuedeUsarItem(userIndex, thisObj.ObjIndex)
            
            Call WriteChangeNPCInventorySlot(userIndex, Slot, thisObj, val, Usable)
        Else
            Dim DummyObj As Obj
            Call WriteChangeNPCInventorySlot(userIndex, Slot, DummyObj, 0, True)
        End If
    Next Slot
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePrice(ByVal ObjIndex As Integer) As Single
'*************************************************
'Author: Nicolás (NicoNZ)
'
'*************************************************
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    
    SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA
End Function
