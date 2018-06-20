Attribute VB_Name = "mod_DunkanItem"
' Programado por maTih.-
' Version original : 25/04/2012

Option Explicit

Type ObjItems
     Duracion       As Integer          'Máximos usos antes de que se rompa.
     Precio         As Long             'Valor del arreglarlo.
     ObjNoUsable    As Integer          'Item "roto"
End Type

Sub CargarObj(ByVal objNum As Integer, ByRef Leedor As clsIniReader)

' @ Carga la durabilidad del objeto.

Dim duracionData     As Byte
Dim precioArreglo    As Long
Dim objNoUsableIndex As Integer

duracionData = val(Leedor.GetValue("OBJ" & objNum, "Duracion"))

'Duracion ?
If duracionData <> 0 Then
   
   'Setea la duracion.
   ObjData(objNum).OBJItem.Duracion = duracionData
   
   objNoUsableIndex = val(Leedor.GetValue("OBJ" & objNum, "ItemNoUsable"))
   
   'Si tiene otro objjeto.
   If objNoUsableIndex <> 0 Then
      ObjData(objNum).OBJItem.ObjNoUsable = objNoUsableIndex
   Else
      ObjData(objNum).OBJItem.ObjNoUsable = objNum
   End If
   
   'Carga el precio.
   precioArreglo = val(Leedor.GetValue("OBJ" & objNum, "ValorArreglo"))
   
   'Si tiene setea.
   If precioArreglo <> 0 Then
      ObjData(objNum).OBJItem.Precio = precioArreglo
   End If

End If


End Sub

Sub CheckearDuracion(ByVal Slot As Byte, ByVal userIndex As Integer)

' @ Comprueba si al slot le queda duración.

Dim ObjToUser   As Obj

With UserList(userIndex).Invent.Object(Slot)
     
     If .Duracion <> 0 Then
        .Duracion = .Duracion - 1
        
        'Nuevo obj.
        ObjToUser.Amount = 1
        ObjToUser.ObjIndex = ObjData(.ObjIndex).OBJItem.ObjNoUsable
        
        'Fin de la duración : P
        QuitarUserInvItem userIndex, Slot, 1
        
        'Le damos el item "roto"
        MeterItemEnInventario userIndex, ObjToUser
        
        'Le avisamos : D
        WriteConsoleMsg userIndex, "Tu : " & ObjData(ObjToUser.ObjIndex).name & " se ha roto, por lo tanto a quedado inutilizable!", FontTypeNames.FONTTYPE_DIOS
     End If
        
End With

End Sub

Sub CargarInventario(ByVal userIndex As Integer, ByRef Leedor As clsIniReader)

' @ Carga la duración de los objs del inv.

Dim loopX       As Long
Dim NumItems    As Byte

NumItems = val(Leedor.GetValue("Inventory", "CantidadItems"))

'No tiene objetos.
If Not NumItems <> 0 Then Exit Sub

For loopX = 1 To MAX_INVENTORY_SLOTS
    With UserList(userIndex).Invent.Object(loopX)
         .Duracion = GetDuracionString(Leedor.GetValue("Inventory", "Obj" & loopX))
    End With
Next loopX


End Sub

Sub GuardarInventario(ByVal userIndex As Integer, ByRef UserPath As String)

' @ Guarda las duraciones de los obj en el inventario.

Dim loopX   As Long

For loopX = 1 To MAX_INVENTORY_SLOTS
    'Hay objeto en el slot.
    If UserList(userIndex).Invent.Object(loopX).ObjIndex <> 0 Then
       'guarda
        WriteVar UserPath, "Inventory", "Obj" & loopX, CStr(UserList(userIndex).Invent.Object(loopX).ObjIndex) & "-" & CStr(UserList(userIndex).Invent.Object(loopX).Amount) & "-" & CStr(UserList(userIndex).Invent.Object(loopX).Equipped) & "-" & CStr(UserList(userIndex).Invent.Object(loopX).Duracion)
    End If
Next loopX

End Sub

Function GetDuracionString(ByRef AllString As String) As Integer

' @ Devuelve la duración de un item.

Dim AllStrings() As String

AllStrings() = Split(AllString, "-")

GetDuracionString = 0

'Tiene duración.
If UBound(AllStrings()) <> 2 Then
   GetDuracionString = val(AllStrings(UBound(AllStrings())))
End If

End Function
