Attribute VB_Name = "Mod_DX8_Invent"
' maTih.-

Option Explicit

Type ObjVenta
     amount     As Integer
     objIndex   As Integer
     GrhIndex   As Integer
     Precio     As Long
End Type

Public Venta_Inventory(1 To 20) As ObjVenta
Public Venta_SelectSlot         As Byte

Public Sub AgregarSlot(ByVal slot As Byte, ByVal GrhIndex As Integer, ByVal objIndex As Integer, ByVal amount As Integer, ByVal Precio As Long)

' @ Agrega un slot.

With Venta_Inventory(slot)

     .objIndex = objIndex
     .GrhIndex = GrhIndex
     .Precio = Precio
     .amount = amount
     
End With

End Sub

Public Sub DibujarInventario()

' @ Dibuja los items de un inventario

 Dim RECTtemp   As RECT
 Dim LoopX      As Long
 Dim PosY       As Integer
 Dim PosX       As Integer
 
 DoEvents
 
 With RECTtemp
    .bottom = frmVilla.picInv.ScaleHeight
    .Right = frmVilla.picInv.ScaleWidth
 End With

 DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
 DirectDevice.BeginScene

 For LoopX = 1 To 20
     With Venta_Inventory(LoopX)
        'Posición del TileY
          PosY = PosicionY(LoopX)
          PosX = PosicionX(LoopX)
          
          'Hay obj?
          If .objIndex <> 0 Then
             'Nos aseguramos de tener un grh.
             If .GrhIndex <> 0 And .amount <> 0 Then
                DDrawTransGrhIndextoSurface .GrhIndex, PosX, PosY, 1
                
                'Dibuja amount.
                Engine_RenderText PosX, PosY, CStr(.amount), -1
                
                'Item seleccionado?
                If LoopX = Venta_SelectSlot Then
                   'Dibuja
                   Engine_Draw_Box PosX, PosY, 32, 33, D3DColorARGB(40, 255, 255, 9)
                End If
                
             End If
          End If
     End With
 Next LoopX
  
 DirectDevice.EndScene
 DirectDevice.Present RECTtemp, ByVal 0, frmVilla.picInv.hWnd, ByVal 0

End Sub

Function PosicionY(ByVal nowState As Byte) As Integer

' @ Posición Y para un item.

Select Case nowState
       Case Is <= 5
            PosicionY = 10
       Case Is > 5
            PosicionY = 50
       Case Is > 10
            PosicionY = 90
       Case Is > 15
            PosicionY = 130
End Select

End Function

Function PosicionX(ByVal nowState As Byte) As Integer

' @ Posición X para un slot.

Select Case nowState
       Case 1 To 5
            PosicionX = (28 * nowState) - 18
       Case 6 To 10
            PosicionX = (28 * (nowState - 5)) - 15
       Case 11 To 15
            PosicionX = (28 * (nowState - 10)) - 15
       Case 16 To 20
            PosicionX = (28 * (nowState - 15)) - 15
End Select

End Function

Function CliCkearItem(ByVal X As Integer, ByVal Y As Integer) As Byte

' @ Devuelve el item clickeado.

If (X >= 20 And X <= 35) And (Y >= 10 And Y <= 55) Then
    CliCkearItem = 1
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 10 And Y <= 55) Then
    CliCkearItem = 2
    Exit Function
End If

If (X >= 75 And X <= 100) And (Y >= 10 And Y <= 55) Then
    CliCkearItem = 3
    Exit Function
End If

If (X >= 105 And X <= 130) And (Y >= 10 And Y <= 55) Then
    CliCkearItem = 4
    Exit Function
End If

If (X >= 135 And X <= 170) And (Y >= 10 And Y <= 55) Then
    CliCkearItem = 5
    Exit Function
End If

If (X >= 20 And X <= 35) And (Y >= 57 And Y <= 90) Then
    CliCkearItem = 6
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 57 And Y <= 90) Then
    CliCkearItem = 7
    Exit Function
End If

If (X >= 75 And X <= 100) And (Y >= 57 And Y <= 90) Then
    CliCkearItem = 8
    Exit Function
End If

If (X >= 105 And X <= 130) And (Y >= 57 And Y <= 90) Then
    CliCkearItem = 9
    Exit Function
End If

If (X >= 135 And X <= 170) And (Y >= 57 And Y <= 90) Then
    CliCkearItem = 10
    Exit Function
End If

If (X >= 20 And X <= 35) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 11
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 12
    Exit Function
End If

If (X >= 20 And X <= 35) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 13
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 14
    Exit Function
End If

If (X >= 20 And X <= 35) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 15
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 16
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 17
    Exit Function
End If

If (X >= 20 And X <= 35) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 18
    Exit Function
End If

If (X >= 45 And X <= 70) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 19
    Exit Function
End If

If (X >= 20 And X <= 35) And (Y >= 40 And Y <= 55) Then
    CliCkearItem = 20
    Exit Function
End If

End Function

Function GetPrecio() As Long

' @ Devuelve el precio de un slot.

GetPrecio = Venta_Inventory(Venta_SelectSlot).Precio

End Function
