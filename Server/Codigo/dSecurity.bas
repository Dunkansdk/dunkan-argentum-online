Attribute VB_Name = "dSecurity"
' / Author      : maTih.-
' / Description : Seguridad de Boskorcha AO

Option Explicit

#If SeguridadDunkan Then

'Tipos de seguridad (USUARIOS!!)

Public Type uSecurity
  'Esta es la clave que tendrá que llegar al final del paquete
  KeyUse          As Integer
  'Esta es la clave para manejar los mapas.
  KeyMapEnc       As Long
  'Contador de paquetes , para cambiar la clave
  PackageCount    As Byte
End Type

'Actual llave de seguridad usada por el cliente

Public KeyUsate          As Integer

'Contador de paquetes enviados(usado para cambiar la clave)
 
Public PackageCounts     As Byte

'Llave para "desenvolver" el numero del mapa real.

Public KeyMapChange      As Byte

'Constante para los máximos paquetes sin cambiar la llave

Public Const MAX_PACKETS As Byte = 4

'FIN DECLARACIONES

Public Function KeyMapReNew(ByVal Index As Integer) As Integer

' / Author : maTih.-
' / Note   : Esta funcion le asigna una clave a un slot(userIndex).

'Re-Hice el algoritmo, ahora es más simple, rápido y eficaz ;D (maTih.-)

KeyMapReNew = (Index * (LastUser + 1)) + RandomNumber(alG ^ 2, ALGB ^ RandomNumber(1, 2))

End Function

Public Function KeyMapUsate(ByVal Arg As Long) As Integer

' / Author : maTih.-
' / Note   : Con esto desencriptamos el mapa y lo convertimos en un numero real

KeyMapUsate = (Arg / KeyMapChange)

End Function

Public Function KeyMapEnc(ByVal Arg As Integer, ByVal Index As Integer) As Long

' / Author : maTih.-
' / Note   : Con esto Encriptamos el mapa y lo convertimos en un numero "indiscreto"

KeyMapEnc = (Arg * UserList(Index).SECURITY.KeyMapEnc)

End Function

Public Function KeyUserUsate(ByVal Index As Integer) As Integer

' / Author : maTih.-
' / Note   : Con esto preparamos una key Random para enviar paquetes

KeyUserUsate = (RandomNumber(IIf(Index < 100, Index, 50), 100) * Index)

End Function

Public Function NewKeyUserUsate(ByVal Index As Integer) As Integer

' / Author : maTih.-
' / Note   : Con esto vamos cambiando la key

Dim LastKey   As Integer

'en lastkey nos guardamos la anterior =)

LastKey = UserList(Index).SECURITY.KeyUse

'randomizamos la clave =)

NewKeyUserUsate = (LastKey * RandomNumber(1, 15))

End Function

Public Function CheckValidKey(ByVal Index As Integer, ByVal rKey As Integer) As Boolean

' / Author : maTih.-
' / Note   : Esta funcion Checkea si la clave recibida al final del paquete es la del usuario

Dim KeyUser      As Integer

KeyUser = UserList(Index).SECURITY.KeyUse

CheckValidKey = (KeyUser = rKey)

End Function

#End If

