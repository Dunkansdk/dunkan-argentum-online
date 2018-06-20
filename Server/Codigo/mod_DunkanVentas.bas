Attribute VB_Name = "mod_DunkanVentas"
' programado por maTih.-

Option Explicit

Type Venta
     invObj(1 To 20)     As Integer      'Inventario.
     invAmount(1 To 20)  As Integer      'Inventario.
     invPrecio(1 To 20)  As Long
     VendiendoAhora      As Boolean      'Si tieneventa.
End Type

Sub Cerrar(ByVal UserIndex As Integer)

' @ Cierra la venta.

With UserList(UserIndex).Venta

     Dim i  As Long
     
     .VendiendoAhora = False
         
     For i = 1 To 20
         Call QuitarSlot(UserIndex, i, 10000)
     Next i

     Call Protocol.WriteConsoleMsg(UserIndex, "Tu venta ha cerrado.", FontTypeNames.FONTTYPE_GM)

End With

End Sub

Sub Nueva(ByVal UserIndex As Integer, ByRef objPrecio() As Long, ByRef ObjAmount() As Integer, ByRef ObjIndex() As Integer)

' @ Usuario empieza la venta : D

Dim i   As Long

With UserList(UserIndex).Venta
    
     .VendiendoAhora = True
     
     'Llena el inventario.
     For i = 1 To UBound(ObjAmount())
         .invObj(i) = ObjIndex(i)
         .invAmount(i) = ObjAmount(i)
         .invPrecio(i) = objPrecio(i)
     Next i
    
     'Envia mensaje
     If MapInfo(UserList(UserIndex).Pos.map).NumUsers <> 1 Then
        Call modSendData.SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead( _
        "Estoy vendiendo objetos!", UserList(UserIndex).Char.CharIndex, _
        vbGreen))
     End If
    
End With
    
End Sub

Sub Compar(ByVal buyIndex As Integer, ByVal saleIndex As Integer, ByVal Slot As Byte, ByVal amount As Integer)

' @ Compra de la venta de saleIndex un slot.

Dim tempObj     As Obj

With UserList(saleIndex).Venta

     If Not .invAmount(Slot) <> 0 Then Exit Sub

     'Hay objeto?
     If .invObj(Slot) <> 0 Then
        'Compra más?
        If amount > .invAmount(Slot) Then amount = .invAmount(Slot)
           'Setea el objeto.
           tempObj.amount = amount
           tempObj.ObjIndex = .invObj(Slot)
           'Tiene el oro?
           If (.invPrecio(Slot) * amount) < UserList(buyIndex).Stats.GLD Then
              'Se lo saca.
              UserList(buyIndex).Stats.GLD = UserList(buyIndex).Stats.GLD - (.invPrecio(Slot) * amount)
              'Actualiza
              Call Protocol.WriteUpdateGold(buyIndex)
              'Suma el oro.
              UserList(saleIndex).Stats.GLD = UserList(saleIndex).Stats.GLD + (.invPrecio(Slot) * amount)
              Call Protocol.WriteUpdateGold(saleIndex)
              'Envia
              Call Protocol.WriteConsoleMsg(saleIndex, UserList(buyIndex).name & _
              " Compró " & ObjData(.invObj(Slot)).name & " (Cantidad : " & amount & ")", FontTypeNames.FONTTYPE_GM)
              'Ganas el oro
              Call Protocol.WriteConsoleMsg(saleIndex, "Has ganado " & Format$((.invPrecio(Slot) * amount), "#,###") & " monedas de oro.", FontTypeNames.FONTTYPE_GM)
              'Quita obj.
              QuitarSlot saleIndex, Slot, amount
              EnviarSlot buyIndex, Slot, .invAmount(Slot), .invPrecio(Slot), .invObj(Slot)
           Else
              Call Protocol.WriteConsoleMsg(buyIndex, "No tienes suficiente oro.", FontTypeNames.FONTTYPE_DIOS)
           End If
        End If
End With

End Sub

Sub QuitarSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal amount As Integer)

' @ Quitar una cantidad de items a un slot.

With UserList(UserIndex).Venta
     
     .invAmount(Slot) = .invAmount(Slot) - amount
     
     'Quita todoS?
     If .invAmount(Slot) <= 0 Then
        'Borra.
        .invAmount(Slot) = 0
        .invObj(Slot) = 0
        .invPrecio(Slot) = (0)
     End If
     
End With

End Sub

Sub Enviar(ByVal UserIndex As Integer, ByVal vendedorIndex As Integer)

' @ Envia la lista de vendedorIndex.

Dim i   As Long

For i = 1 To 20
    
    With UserList(vendedorIndex).Venta
         If .invAmount(i) <> 0 And .invObj(i) <> 0 Then
            Call EnviarSlot(UserIndex, i, .invAmount(i), .invPrecio(i), .invObj(i))
         End If
    End With
    
Next i

End Sub

Sub EnviarSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal amount As Integer, ByVal precio As Long, ByVal ObjIndex As Integer)

' @ Envia un slot.

Dim tmpGrhIndex As Integer

With UserList(UserIndex)

     Call .outgoingData.WriteByte(ServerPacketID.EnviarVenta)
     
     Call .outgoingData.WriteByte(Slot)
     
     Call .outgoingData.WriteInteger(ObjIndex)
     
     If ObjIndex <> 0 Then
        tmpGrhIndex = ObjData(ObjIndex).GrhIndex
     Else
        tmpGrhIndex = 0
     End If
     
     Call .outgoingData.WriteInteger(tmpGrhIndex)

     Call .outgoingData.WriteInteger(amount)

     Call .outgoingData.WriteLong(precio)
     
End With

End Sub
