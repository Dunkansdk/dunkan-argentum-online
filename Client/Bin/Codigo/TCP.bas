Attribute VB_Name = "Mod_TCP"
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

' / Organizado y tabulado por Dunkan
' - Note, este módulo está AL PEDO.

Public Warping          As Boolean
Public LlegaronSkills   As Boolean
Public LlegaronAtrib    As Boolean
Public LlegoFama        As Boolean


Public Function PuedoQuitarFoco() As Boolean

    PuedoQuitarFoco = True

End Function

Sub Login()

    If EstadoLogin = E_MODO.Normal Then
        Call WriteLogPj
    ElseIf EstadoLogin = E_MODO.CrearNuevoPj Then
        Call WriteLoginNewChar
    ElseIf EstadoLogin = E_MODO.LoginCuenta Then
        Call WriteLogCuenta
    End If
    
    DoEvents
    
Call FlushBuffer
    
End Sub
