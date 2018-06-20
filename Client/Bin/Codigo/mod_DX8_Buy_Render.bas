Attribute VB_Name = "mod_DX8_Buy"
'
' @ maTih.-

Option Explicit

Public Buy_Active       As Boolean

Type structBuy
     LineStr()          As String
     NumLines           As Integer
End Type

Public Buy_Data         As structBuy

Sub BuyStateChange(ByRef arrMsjs() As String)

' '
' Cambia los mensajes de la compra.

Dim LoopC As Long

    With Buy_Data
    
         .NumLines = (UBound(arrMsjs()) - 1)
         
         ReDim .LineStr(LBound(arrMsjs()) To UBound(arrMsjs()) - 1) As String
         
         For LoopC = LBound(arrMsjs()) To (UBound(arrMsjs()) - 1)
             .LineStr(LoopC) = arrMsjs(LoopC)
         Next LoopC
                  
         Buy_Active = True
         RenderBuyMenu
                  
    End With

End Sub

Sub RenderBuyMenu()

' '
' Renderiza la venta.

Dim i   As Long

    With Buy_Data
    
         If Not Buy_Active Then Exit Sub
            
         If (Not .NumLines <> 0) Then Exit Sub
         
         For i = LBound(.LineStr()) To UBound(.LineStr())
             Engine_RenderText 30, 30 + (i * 20), CStr(i + 1) & ")" & .LineStr(i), D3DColorARGB(255, 255, 255, 255)
         Next i
         
    End With

End Sub
