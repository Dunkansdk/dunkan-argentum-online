Attribute VB_Name = "Mod_Cuentas"
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

' / Author: maTih
' / Note: Manager de las cuentas. INCOMPLETO.

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


Public Type CharData    'Define como se veran los pjs
    Head    As Integer
    Body    As Integer
    Weapon  As Byte
    Shield  As Byte
    Helmet  As Byte
    Name    As String
    Nivel   As Byte
End Type

Public Type Accounts
    charInfo()          As CharData
    CantidadPersonajes  As Byte
    NameAccount         As String   'Nombre cuenta
End Type

Public Cuentas As Accounts
