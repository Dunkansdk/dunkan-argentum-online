Attribute VB_Name = "Mod_Ventas"
'************************************************************
'..................NUEVO SISTEMA DE PARTYES..................
'************************************************************
'************************************************************
'..................Escrito por maTih(28/01)..................
'************************************************************

Option Explicit

Public Const MAX_VENTA_SLOT As Byte = 20       'Maximos slots.

Public Type VentaObj
       ObjIndex                 As Integer      'ObjIndex del slot.
       Amount                   As Integer      'Cantidad de slot.
       Precio                   As Long         'Precio del objeto.
End Type

Public Type VentaInventario
        Objs(1 To MAX_VENTA_SLOT) As VentaObj   'Objetos.
        CantidadItems             As Byte       'Cuantos items.
        Vendiendo                 As Boolean    'Está vendiendo?
End Type



Sub Ventas_SendInvent(ByVal SendIndex As Integer, ByVal TargetIndex As Integer)

' \ Author : maTih.-
' \ Note   : Envia el inventario de venta de TargetIndex

With UserList(TargetIndex).VentaInv

Dim ObjIndex    As Integer
Dim loopX       As Long

For loopX = 1 To .CantidadItems
    ObjIndex = .Objs(loopX).ObjIndex
    
        If ObjIndex <> 0 Then
            WriteDaoUpdateVentaSlot SendIndex, TargetIndex, loopX
        End If
        
Next loopX

End With

    WriteDaoShowVentaForm SendIndex

End Sub

Sub Ventas_UpdateSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

' \ Author : maTih.-
' \ Note   : Actualiza un slot de la venta.

    

End Sub

Sub Ventas_AddSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer, ByVal Precie As Long)

' \ Author : maTih.-
' \ Note   : Agrega un objeto al inventario de la venta

With UserList(UserIndex).VentaInv

    .CantidadItems = .CantidadItems + 1
    
    .Objs(.CantidadItems).Amount = Amount
    .Objs(.CantidadItems).ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    .Objs(.CantidadItems).Precio = Precie

    Ventas_UpdateSlot UserIndex, .CantidadItems
End With

End Sub

Sub Ventas_QuitSlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer)

' \ Author : maTih.-
' \ Note   : Quita un objeto al inventario de la venta

Dim sacoTodo    As Boolean

    With UserList(UserIndex).VentaInv
            
            'Si saca mas de los que hay saca todos.
            If Amount > .Objs(Slot).Amount Then
                Amount = .Objs(Slot).Amount
                sacoTodo = True
            End If
            
            .Objs(Slot).Amount = .Objs(Slot).Amount - Amount
            
            'Si saco todos borro el objeto.
            
            If sacoTodo Then
                .Objs(Slot).Amount = 0
                .Objs(Slot).ObjIndex = 0
                .Objs(Slot).Precio = 0
            End If
            
            Ventas_UpdateSlot UserIndex, Slot
    
    End With

End Sub

Sub Ventas_ComprarSlot(ByVal UserIndex As Integer, ByVal VendedorIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer)

' \ Author : maTih.-
' \ Note   : Compra un objeto del inventario de vendedorIndex.

With UserList(VendedorIndex)

Dim dObj        As Obj

'Si el slot no es válido..
If Slot <= 0 Or Slot > .VentaInv.CantidadItems Then Exit Sub

    'Si quiere comprar más de los que hay,compra todos.
    If Amount > .VentaInv.Objs(Slot).Amount Then
        Amount = .VentaInv.Objs(Slot).Amount
    End If
    
    'No tiene el oro suficiente.
    If UserList(UserIndex).Stats.GLD < (.VentaInv.Objs(Slot).Precio * Amount) Then
        WriteConsoleMsg UserIndex, "No tienes tanto oro!", FontTypeNames.FONTTYPE_CITIZEN
        Exit Sub
    End If
    
    'Compró el item.
    
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - (.VentaInv.Objs(Slot).Precio * Amount)
    
    'Meto el objeto en el inventario.
    
    dObj.ObjIndex = .VentaInv.Objs(Slot).ObjIndex
    dObj.Amount = Amount
    
    If Not MeterItemEnInventario(UserIndex, dObj) Then
        TirarItemAlPiso UserList(UserIndex).Pos, dObj
    End If
    
    'Updateo el cliente.
    WriteUpdateGold UserIndex
    
    'Informo al vendedor.
    
        WriteConsoleMsg VendedorIndex, UserList(UserIndex).name & " Compro algunos objetos de tu venta!", FontTypeNames.FONTTYPE_CITIZEN
        
    'Le doy el oro.
    
        .Stats.GLD = .Stats.GLD + (.VentaInv.Objs(Slot).Precio * Amount)
    
    'Si compró todos, limpio el slot.
        
    If Amount >= .VentaInv.Objs(Slot).Amount Then
        .VentaInv.Objs(Slot).Amount = 0
        .VentaInv.Objs(Slot).ObjIndex = 0
        .VentaInv.Objs(Slot).Precio = 0
    
    Else
    
    'Si no, quito la cantidad que compró.
    
        .VentaInv.Objs(Slot).Amount = .VentaInv.Objs(Slot).Amount - Amount
        
    End If
    
End With

End Sub
