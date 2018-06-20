Attribute VB_Name = "Mod_Client_Settings"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' / Author: Emanuel Matías (Dunkansdk)
' / Note: Configuración de los usuarios.

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

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


Private Type structClientCFG
    
    'Engine
    Light_Radius    As Byte
    WaterMovement   As Byte
    Minimap         As Byte
    Ambient         As Byte
    Weater          As Byte
    Projectiles     As Byte
    FogIntensity    As Byte
    RainIntensity   As Byte
    SnowIntensity   As Byte
    Damage          As Byte
    
    'Video
    BufferSize          As Byte
    Aceleration         As Byte
    videoMemory         As Long
    VSynchronization    As Byte
        
End Type
    
Public Settings As structClientCFG

Public Function LoadClientDefaultSettings()

' / Author: Dunkansdk
    
    With Settings
    
        'Engine
        .Light_Radius = 0
        .WaterMovement = 0
        .Minimap = 1
        .Ambient = 1
        .Weater = 1
        .Projectiles = 1
        .FogIntensity = 60
        .RainIntensity = 150
        .SnowIntensity = 150
        .Damage = 1
    
        ' Video
        .Aceleration = 1
        .BufferSize = 10
        .videoMemory = 256
        .VSynchronization = 0
    
    End With
    
End Function
