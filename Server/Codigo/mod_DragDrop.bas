Attribute VB_Name = "mod_DragDrop"

' *****************************************************
' ********************* DUNKAN AO *********************
' *****************************************************
Option Explicit

Sub DragToUser(ByVal userIndex As Integer, ByVal tIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer)

' @ Author : maTih.-
'            Drag un slot a un usuario.

Dim tObj    As Obj
Dim tString As String
Dim Espacio As Boolean

'Preparo el objeto.
tObj.Amount = Amount
tObj.objIndex = UserList(userIndex).Invent.Object(Slot).objIndex

Espacio = MeterItemEnInventario(tIndex, tObj)

'No tiene espacio.
If Not Espacio Then
   WriteConsoleMsg userIndex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_CITIZEN
   Exit Sub
End If

'Quito el objeto.
QuitarUserInvItem userIndex, Slot, Amount

'Hago un update de su inventario.
UpdateUserInv False, userIndex, Slot

'Preparo el mensaje para userINdex (quien dragea)

tString = "Le has arrojado"

If tObj.Amount <> 1 Then
   tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
Else
   tString = tString & " Tu " & ObjData(tObj.objIndex).Name
End If

tString = tString & " ah " & UserList(tIndex).Name

'Envio el mensaje
WriteConsoleMsg userIndex, tString, FontTypeNames.FONTTYPE_CITIZEN

'Preparo el mensaje para el otro usuario (quien recibe)
tString = UserList(userIndex).Name & " Te ha arrojado"

If tObj.Amount <> 1 Then
   tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
Else
   tString = tString & " su " & ObjData(tObj.objIndex).Name
End If

'Envio el mensaje al otro usuario
WriteConsoleMsg userIndex, tString, FontTypeNames.FONTTYPE_CITIZEN

End Sub

Sub DragToNPC(ByVal userIndex As Integer, ByVal tNpc As Integer, ByVal Slot As Byte, ByVal Amount As Integer)

' @ Author : maTih.-
'            Drag un slot a un npc.
On Error GoTo errHandler

Dim teniaOro    As Long
Dim teniaObj    As Integer
Dim tmpIndex    As Integer

tmpIndex = UserList(userIndex).Invent.Object(Slot).objIndex
teniaOro = UserList(userIndex).Stats.GLD
teniaObj = UserList(userIndex).Invent.Object(Slot).Amount

'Es un banquero?
If Npclist(tNpc).NPCtype = eNPCType.Banquero Then
   Call UserDejaObj(userIndex, Slot, Amount)
   'No tiene más el mismo amount que antes? entonces depositó.
   If teniaObj <> UserList(userIndex).Invent.Object(Slot).Amount Then
      WriteConsoleMsg userIndex, "Has depositado " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_CITIZEN
      UpdateUserInv False, userIndex, Slot
   End If
'Es un npc comerciante?
ElseIf Npclist(tNpc).Comercia = 1 Then
   'El npc compra cualquier tipo de items?
   If Not Npclist(tNpc).TipoItems <> eOBJType.otCualquiera Or Npclist(tNpc).TipoItems = ObjData(UserList(userIndex).Invent.Object(Slot).objIndex).OBJType Then
      Call Comercio(eModoComercio.Venta, userIndex, tNpc, Slot, Amount)
      'Ganó oro? si es así es porque lo vendió.
      If teniaOro <> UserList(userIndex).Stats.GLD Then
         WriteConsoleMsg userIndex, "Le has vendido al " & Npclist(tNpc).Name & " " & Amount & " - " & ObjData(tmpIndex).Name, FontTypeNames.FONTTYPE_CITIZEN
      End If
   Else
      WriteConsoleMsg userIndex, "El npc no está interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_CITIZEN
   End If
End If

Exit Sub

errHandler:

End Sub

Sub DragToPos(ByVal userIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Slot As Byte, ByVal Amount As Integer)

' @ Author : maTih.-
'            Drag un slot a una posición.

Dim errorFound  As String
Dim tObj        As Obj
Dim tString     As String

'No puede dragear en esa pos?
If Not CanDragToPos(UserList(userIndex).Pos.Map, X, Y, errorFound) Then
   WriteConsoleMsg userIndex, errorFound, FontTypeNames.FONTTYPE_CITIZEN
   Exit Sub
End If

'Creo el objeto.
tObj.objIndex = UserList(userIndex).Invent.Object(Slot).objIndex
tObj.Amount = Amount

'Agrego el objeto a la posición.
MakeObj tObj, UserList(userIndex).Pos.Map, X, Y

'Quito el objeto.
QuitarUserInvItem userIndex, Slot, Amount

'Actualizo el inventario
UpdateUserInv False, userIndex, Slot

'Preparo el mensaje.
tString = "Has arrojado "

If tObj.Amount <> 1 Then
   tString = tString & tObj.Amount & " - " & ObjData(tObj.objIndex).Name
Else
   tString = "tu " & ObjData(tObj.objIndex).Name
End If

'ENvio.
WriteConsoleMsg userIndex, tString, FontTypeNames.FONTTYPE_CITIZEN

End Sub

Function CanDragToPos(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByRef Error As String) As Boolean

' @ Author : maTih.-
'            Devuelve si se puede dragear un item a x posición.

CanDragToPos = False

'Zona segura?
If Not MapInfo(Map).Pk Then
   Error = "No está permitido arrojar objetos al suelo en zonas seguras."
   Exit Function
End If

'Ya hay objeto?
If Not MapData(Map, X, Y).ObjInfo.objIndex = 0 Then
   Error = "Hay un objeto en esa posición!"
   Exit Function
End If

'Tile bloqueado?
If Not MapData(Map, X, Y).Blocked = 0 Then
   Error = "No puedes arrojar objetos en esa posición"
   Exit Function
End If

CanDragToPos = True

End Function

Function CanDragObj(ByVal objIndex As Integer, ByVal Navegando As Boolean, ByRef Error As String) As Boolean

' @ Author : maTih.-
'            Devuelve si un objeto es drageable.
CanDragObj = False

If objIndex < 1 Or objIndex > UBound(ObjData()) Then Exit Function

'Objeto newbie?
If ObjData(objIndex).Newbie <> 0 Then
   Error = "No puedes arrojar objetos newbies!"
   Exit Function
End If

'Está navgeando?
If Navegando Then
   Error = "No puedes arrojar un barco si estás navegando!"
   Exit Function
End If

CanDragObj = True

End Function
