Attribute VB_Name = "Mod_DSecurity"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' / Author      : maTih.-
' / Description : Seguridad de Boskorcha AO

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

#If SeguridadDunkan Then

Public KeyUsate          As Integer
Public PackageCounts     As Byte
Public KeyMapChange      As Byte

'Public Const MAX_PACKETS As Byte = 4

Public Function KeyMapUsate(ByVal Arg As Long) As Integer

' / Author : maTih.-
' / Note   : Con esto desencriptamos el mapa y lo convertimos en un numero real

    KeyMapUsate = (Arg / KeyMapChange)

End Function

Public Function KeyMapEnc(ByVal Arg As Integer, ByVal Index As Integer) As Long

' / Author: maTih.-
' / Note: Con esto Encriptamos el mapa y lo convertimos en un numero "indiscreto"

    KeyMapEnc = (Arg * userList(Index).SECURITY.KeyMapEnc)

End Function

Public Function KeyUserUsate(ByVal Index As Integer) As Integer

' / Author: maTih.-
' / Note: Con esto preparamos una key Random para enviar paquetes

    KeyUserUsate = (RandomNumber(IIf(Index < 100, Index, 50), 100) * Index)

End Function

Public Function NewKeyUserUsate(ByVal Index As Integer) As Integer

' / Author: maTih.-
' / Note: Con esto vamos cambiando la key

    Dim LastKey   As Integer

    LastKey = userList(Index).SECURITY.KEYUSE
    NewKeyUserUsate = (LastKey * RandomNumber(1, 15)) ' Clave RND

End Function

Public Function CheckValidKey(ByVal Index As Integer, ByVal rKey As Integer) As Boolean

' / Author: maTih.-
' / Note: Esta funcion Checkea si la clave recibida al final del paquete es la del usuario

    Dim KeyUser      As Integer
    
        KeyUser = userList(Index).SECURITY.KEYUSE
        CheckValidKey = (KeyUser = rKey)
    
    End Function

#End If
