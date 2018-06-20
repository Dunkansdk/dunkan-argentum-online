Attribute VB_Name = "mod_DunkanLimpieza"
' programado por maTih.-

Option Explicit

Const MAXBUFFER             As Byte = 255

Type Objeto
     EnPos                  As WorldPos
     SlotUsed               As Boolean
End Type

Public CW(1 To MAXBUFFER)   As Objeto

Sub Agregar(ByRef Pos As WorldPos)

' @ Agrega un objeto a X pos.

Dim N_Slot  As Byte

N_Slot = ProximoSlot()

'No slot.
If Not N_Slot <> 0 Then Exit Sub

With CW(N_Slot)
     .EnPos = Pos
     .SlotUsed = True
End With

End Sub

Sub Limpiar(Optional ByVal DesdeMap As Integer = 0, Optional ByVal HastaMap As Integer = 0)

' @ Limpia el mundo.

Dim OriginalStart   As Integer
Dim OriginalEnd     As Integer

'Si no inicia desde un mapa por defecto es 1.
If Not DesdeMap <> 0 Then DesdeMap = 1: OriginalStart = 1

'Si no termina en un mapa termina en el ultimo.
If Not HastaMap <> 0 Then HastaMap = NumMaps: OriginalEnd = HastaMap

Dim N_Loop      As Long
Dim N_TickStart As Long
Dim N_Avistaged As Boolean
Dim N_ClearPos  As WorldPos

'Setea el tiempo de inicio.
N_TickStart = GetTickCount()

'Avisa a su lugar.
If (Not DesdeMap <> 1) And (Not HastaMap <> NumMaps) Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpiado del mundo.", FontTypeNames.FONTTYPE_DIOS))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    N_Avistaged = True
End If
    
For N_Loop = 1 To MAXBUFFER
    With CW(N_Loop)
         'Si está usado es por que hay un objeto.
         If .SlotUsed Then
            'Está en el rango d mapas?
            If AnalizePos(.EnPos, DesdeMap, HastaMap) Then
                'borra.
                Call EraseObj(10000, .EnPos.map, .EnPos.X, .EnPos.Y)
                .EnPos = N_ClearPos
                .SlotUsed = False
            End If
         End If
    End With
Next N_Loop

'Avisa el tiempo y despausea.
If N_Avistaged Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpiado del mundo finalizado, duración: " & CSng(((GetTickCount() - N_TickStart) / 1000)) & " segundos.", FontTypeNames.FONTTYPE_DIOS))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
End If

End Sub

Function AnalizePos(ByRef Pos As WorldPos, ByVal desde As Integer, ByVal hasta As Integer) As Boolean

' @ Se fija si está en el rango de mapas.

With Pos
     AnalizePos = (.map >= desde) And (.map <= hasta)
End With

End Function

Function ProximoSlot() As Byte

' @ Devuelve un slot para un obj.

Dim N_Loop  As Long

For N_Loop = 1 To MAXBUFFER
    If Not CW(N_Loop).SlotUsed Then ProximoSlot = CByte(N_Loop): Exit Function
Next N_Loop

'Not found.
ProximoSlot = 0

End Function
