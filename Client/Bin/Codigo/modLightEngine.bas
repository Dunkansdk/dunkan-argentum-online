Attribute VB_Name = "modLightEngine"
Option Explicit

Public Type LightVertex
    type As Byte
    affected As Byte
End Type

Public DayLightByte As Byte
Public TwinkLightByteHandle As Long

Public LightMap(1 To 100, 1 To 100) As LightVertex

Public Function LightValue(value As Integer) As Long

' Author: Emanuel Matías 'Dunkan'
' Note: RGB de la luz.

If value > 255 Then value = 255
value = value - TwinkLightByteHandle
LightValue = D3DColorXRGB(value, value, value)

End Function

Public Function DayLight() As Long

' Author: Emanuel Matías 'Dunkan'
' Note: RGB default del AO.

DayLightByte = 200
DayLight = D3DColorXRGB(DayLightByte, DayLightByte, DayLightByte)

End Function

Public Function GetLightValue(ByVal X As Byte, ByVal Y As Byte, Vertice As Byte) As Long

' Author: Emanuel Matías 'Dunkan'
' Note: Setea el RGB del mapa, según el vertice.

Select Case Vertice
    Case 0: 'DN LT VERTEX
        If Y > 99 Then Exit Function
        With LightMap(X, Y + 1)
            If .affected Then
                GetLightValue = LightValue(DayLightByte + .affected * (255 - DayLightByte) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
    Case 1: 'UP LT VERTEX
        With LightMap(X, Y)
            If .affected Then
                GetLightValue = LightValue(DayLightByte + .affected * (255 - DayLightByte) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
    Case 2: 'DN RT VERTEX
        If X > 99 Or Y > 99 Then Exit Function
        With LightMap(X + 1, Y + 1)
            If .affected Then
                GetLightValue = LightValue(DayLightByte + .affected * (255 - DayLightByte) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
    Case 3: 'UP RT VERTEX
        If X > 99 Then Exit Function
        With LightMap(X + 1, Y)
            If .affected Then
                GetLightValue = LightValue(DayLightByte + .affected * (255 - DayLightByte) / 4)
            Else
                GetLightValue = DayLight
            End If
        End With
End Select

End Function

Public Sub SetLight(X As Byte, Y As Byte)

' Author: Emanuel Matías 'Dunkan'
' Note: Setea la luz en el mapa.

With LightMap(X, Y)
    .affected = 4
    .type = 1
End With

AffectVertex X + 1, Y, 2
AffectVertex X - 1, Y, 2
AffectVertex X, Y - 1, 2
AffectVertex X, Y + 1, 2

AffectVertex X - 1, Y - 1, 1
AffectVertex X + 1, Y - 1, 1
AffectVertex X - 1, Y + 1, 1
AffectVertex X + 1, Y + 1, 1

End Sub

Public Sub AffectVertex(X As Byte, Y As Byte, value As Byte)

' Author: Emanuel Matías 'Dunkan'
' Note: Selecciona los vértices

With LightMap(X, Y)
    If .affected < value Then .affected = value
    .type = 1
End With

End Sub
