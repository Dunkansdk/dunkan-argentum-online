Attribute VB_Name = "Mod_DX8_Minimap"
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

' / Módulo creado por Emanuel Matías D'Urso 'Dunkan'
' / Note: Crea el minimapa (DirectX8)

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Type structMinimap

    Pos         As Position
    Color       As Long
    UserIndex   As Position
    Rango       As Byte
    
End Type

Private Minimap As structMinimap

Public ShowMinimap As Boolean

Public Function Engine_Minimap_Build(ByVal Show_Minimap As Byte)

' / Author: Dunkansdk
' / Note: Render minimap tiles, el rango depende de la supervivencia del usuario.

If Not CBool(Show_Minimap) = True Then Exit Function

Dim loopCSkills As Long

Const vS As Byte = 2

  '  For loopCSkills = 1 To NUMSKILLS
    
  '      If UserSkills(loopCSkills) = "0" Then
  '          Call WriteRequestAtributes
  '          Call WriteRequestSkills
  '          Call WriteRequestMiniStats
  '          Call WriteRequestFame
  '          Call FlushBuffer
  '      End If
        
  '      Exit For
        
  '  Next loopCSkills
    
    For Minimap.Pos.Y = 8 To 92
        For Minimap.Pos.X = 10 To 92
        
            With MapData(Minimap.Pos.X, Minimap.Pos.Y)
                
                If .Graphic(1).GrhIndex > "1" Then _
                If .Blocked = 1 Then Engine_Draw_Box Minimap.Pos.X * vS, Minimap.Pos.Y * vS, vS, vS, D3DColorARGB(100, 220, 220, 220)
    
            End With
            
        Next Minimap.Pos.X
    Next Minimap.Pos.Y
    
    Engine_Draw_Box UserPos.X * vS, UserPos.Y * vS, vS, vS, D3DColorARGB(150, 0, 255, 0)

End Function
