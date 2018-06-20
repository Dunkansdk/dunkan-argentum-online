Attribute VB_Name = "mod_DunkanCs"
Option Explicit

Type BuyList
     Message    As String
     Objeto     As Obj
End Type

Type buyData
     Buys()     As BuyList
End Type

Public Buy_File               As String
Public Buy(1 To NUMCLASES)    As buyData


Sub Initialize(ByVal eClase As Byte)

'
' @ Objetos a comprar.

Buy_File = App.Path & "\Clases\" & ListaClases(eClase) & ".txt"

If Not FileExist(Buy_File) Then Exit Sub

Dim cantLines   As Byte
Dim loopX       As Long
Dim loader_Line As New clsIniReader
Dim tempStr     As String

loader_Line.Initialize Buy_File

cantLines = Val(loader_Line.GetValue("INIT", "Lineas"))

'No hay lineas
If Not cantLines <> 0 Then Exit Sub

'Resize array
ReDim Buy(eClase).Buys(1 To cantLines) As BuyList

For loopX = 1 To cantLines
    With Buy(eClase).Buys(loopX)
    'Llena la lista
     tempStr = loader_Line.GetValue("LIST" & CStr(loopX), "Objeto")
     .Message = loader_Line.GetValue("LIST" & CStr(loopX), "Mensaje")
     .Objeto.objIndex = Val(ReadField(1, tempStr, Asc("-")))
     .Objeto.Amount = Val(ReadField(2, tempStr, Asc("-")))
    End With
Next loopX

End Sub

Sub Comprar(ByVal slotList As Byte, ByVal userIndex As Integer)

'
' @ Compra un item

With Buy(UserList(userIndex).Clase).Buys(slotList)
     
     'Mete el item
     MeterItemEnInventario userIndex, .Objeto
     
     Call WriteConsoleMsg(userIndex, "Compraste " & .Objeto.Amount & " (" & ObjData(.Objeto.objIndex).Name & ".", FontTypeNames.FONTTYPE_CITIZEN)
     
End With

End Sub

Public Function GetList(ByVal eClase As eClass) As String

'
' @ Devuelve la lista en un solo string

Dim i   As Long

For i = 1 To UBound(Buy(eClase).Buys())
    GetList = GetList & Buy(eClase).Buys(i).Message & ","
Next i

End Function
