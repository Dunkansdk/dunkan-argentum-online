Attribute VB_Name = "Mod_DX8_Light"
Option Explicit

Private Type Light
    Active  As Boolean  'Do we ignore this light?
    Id      As Long
    Map_X   As Integer  'Coordinates
    Map_Y   As Integer
    Color   As Long     'Start colour
    Range   As Byte
    Red     As Byte
    Green   As Byte
    Blue    As Byte
    Type    As Byte
End Type

Public LightRound As Boolean

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

Public Function Light_Remove(ByVal light_index As Long) As Boolean

    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
    
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef Color_value As Long) As Boolean

    If Light_Check(light_index) Then
        Color_value = light_list(light_index).Color
        Light_Color_Value_Get = True
    End If
    
End Function

Public Function Light_Create(ByVal Map_X As Integer, ByVal Map_Y As Integer, ByVal Red As Byte, _
                         ByVal Green As Byte, ByVal Blue As Byte, _
                        Optional ByVal Range As Byte = 1, Optional Id As Byte, Optional ByVal LType As Byte = 1) As Long

    If InMapBounds(Map_X, Map_Y) Then
        Light_Create = Light_Next_Open
        Call Light_Make(Light_Create, Map_X, Map_Y, Range, Id, Red, Green, Blue, LType)
    End If
    
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal Map_X As Integer, ByVal Map_Y As Integer, _
                        ByVal Range As Long, ByVal Id As Long, ByVal Red As Byte, _
                         ByVal Green As Byte, ByVal Blue As Byte, ByVal LType As Byte)

    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).Active = True
    light_list(light_index).Map_X = Map_X
    light_list(light_index).Map_Y = Map_Y
    light_list(light_index).Red = Red
    light_list(light_index).Green = Green
    light_list(light_index).Blue = Blue
    light_list(light_index).Range = Range
    light_list(light_index).Id = Id
    light_list(light_index).Type = LType

End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean

    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).Active Then
            Light_Check = True
        End If
    End If
    
End Function

Public Sub Light_Render_Area()

'   Author: Dunkan
'   Note: Las luces redondas son pesadisimas >.< mejor renderizar solo el area.
'   OPTIMIZAR SUB CUANDO SE PUEDA!!!!!

    Dim i As Long
            
    For i = 1 To light_count
        If light_list(i).Map_X > UserPos.X - TileBufferSize And light_list(i).Map_X < UserPos.X + TileBufferSize Then
            If light_list(i).Map_Y > UserPos.Y - TileBufferSize And light_list(i).Map_Y < UserPos.Y + TileBufferSize Then
                'If Client_Setup.Effect_LightType = 1 Then ' 1 = Redondas
                '    If Light_Check(i) Then Light_Render i
                'Else
                '    If Light_Check(i) Then Map_LightRender_Square i
                'End If
            End If
        End If
    
    Next i
End Sub

Public Sub Light_Render_All()
   
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        'If light_list(loop_counter).Type = 1 Then '1 = Redondas
        '    If Light_Check(loop_counter) Then Light_Render loop_counter
        'Else '0 = Cuadradas FEAAAA xD
            If Light_Check(loop_counter) Then Map_LightRender_Square loop_counter
        'End If
    Next loop_counter
    
End Sub

Private Sub Map_LightRender_Square(ByVal light_index As Long)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim Color(3) As Long
    Dim Ya As Integer
    Dim Xa As Integer
    Dim XCoord As Integer
    Dim YCoord As Integer
    
    Color(0) = D3DColorARGB(255, light_list(light_index).Red, light_list(light_index).Green, light_list(light_index).Blue)
    Color(1) = Color(0)
    Color(2) = Color(0)
    Color(3) = Color(0)

    'Set up light borders
    min_x = light_list(light_index).Map_X - light_list(light_index).Range
    min_y = light_list(light_index).Map_Y - light_list(light_index).Range
    max_x = light_list(light_index).Map_X + light_list(light_index).Range
    max_y = light_list(light_index).Map_Y + light_list(light_index).Range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = Color(2)
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = Color(0)
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = Color(1)
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = Color(3)
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = Color(0)
            MapData(X, min_y).light_value(2) = Color(2)
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = Color(1)
            MapData(X, max_y).light_value(3) = Color(3)
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = Color(2)
            MapData(min_x, Y).light_value(3) = Color(3)
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = Color(0)
            MapData(max_x, Y).light_value(1) = Color(1)
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = Color(0)
                MapData(X, Y).light_value(1) = Color(1)
                MapData(X, Y).light_value(2) = Color(2)
                MapData(X, Y).light_value(3) = Color(3)
            End If
        Next Y
    Next X
    
End Sub

Private Function CalcularRadio(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoordenadas As Integer, ByVal YCoordenadas As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim DistanciaX As Single
    Dim DistanciaY As Single
    Dim DistanciaVertex As Single
    Dim Radio As Integer
    
    Dim CurrentColor As D3DCOLORVALUE
    
    Radio = cRadio * 32
    
    DistanciaX = LightX + 16 - XCoordenadas
    DistanciaY = LightY + 16 - YCoordenadas
    
    DistanciaVertex = Sqr(DistanciaX * DistanciaX + DistanciaY * DistanciaY)
    
    If DistanciaVertex <= Radio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, DistanciaVertex / Radio)
        CalcularRadio = D3DColorXRGB(CurrentColor.R, CurrentColor.G, CurrentColor.B)
        If TileLight > CalcularRadio Then CalcularRadio = TileLight
    Else
        CalcularRadio = TileLight
    End If
End Function

Private Sub Light_Render(ByVal light_index As Long)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim Color As Long
    Dim Ya As Integer
    Dim Xa As Integer
    Dim TileLight As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
    Dim ColorAmbiente As D3DCOLORVALUE
    
    ColorAmbiente.a = 255
    ColorAmbiente.R = ColorActual.R
    ColorAmbiente.G = ColorActual.G
    ColorAmbiente.B = ColorActual.B
    
    Dim XCoord As Integer
    Dim YCoord As Integer
    
    LightColor.a = 255
    LightColor.R = light_list(light_index).Red
    LightColor.G = light_list(light_index).Green
    LightColor.B = light_list(light_index).Blue
    
    'Set up light borders
    min_x = light_list(light_index).Map_X - light_list(light_index).Range
    min_y = light_list(light_index).Map_Y - light_list(light_index).Range
    max_x = light_list(light_index).Map_X + light_list(light_index).Range
    max_y = light_list(light_index).Map_Y + light_list(light_index).Range
    
            For Ya = min_y To max_y
            For Xa = min_x To max_x
                If InMapBounds(Xa, Ya) Then
                    XCoord = Xa * 32
                    YCoord = Ya * 32
                    MapData(Xa, Ya).light_value(1) = CalcularRadio(light_list(light_index).Range, _
                    light_list(light_index).Map_X * 32, light_list(light_index).Map_Y * 32, XCoord, _
                    YCoord, MapData(Xa, Ya).light_value(1), LightColor, ColorAmbiente)

                    XCoord = Xa * 32 + 32
                    YCoord = Ya * 32
                    MapData(Xa, Ya).light_value(3) = CalcularRadio(light_list(light_index).Range, _
                    light_list(light_index).Map_X * 32, light_list(light_index).Map_Y * 32, XCoord, _
                    YCoord, MapData(Xa, Ya).light_value(3), LightColor, ColorAmbiente)
                       
                    XCoord = Xa * 32
                    YCoord = Ya * 32 + 32
                    MapData(Xa, Ya).light_value(0) = CalcularRadio(light_list(light_index).Range, _
                    light_list(light_index).Map_X * 32, light_list(light_index).Map_Y * 32, XCoord, _
                    YCoord, MapData(Xa, Ya).light_value(0), LightColor, ColorAmbiente)
                    
                    XCoord = Xa * 32 + 32
                    YCoord = Ya * 32 + 32
                    MapData(Xa, Ya).light_value(2) = CalcularRadio(light_list(light_index).Range, _
                    light_list(light_index).Map_X * 32, light_list(light_index).Map_Y * 32, XCoord, _
                    YCoord, MapData(Xa, Ya).light_value(2), LightColor, ColorAmbiente)
                End If
            Next Xa
            
        Next Ya

End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    
    LoopC = 1
    Do Until light_list(LoopC).Active = False
        If LoopC = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Light_Next_Open = LoopC
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal Id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    
    LoopC = 1
    Do Until light_list(LoopC).Id = Id
        If LoopC = light_last Then
            Light_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
    
    Light_Find = LoopC
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean

    Dim Index As Long
    
    For Index = 1 To light_last
        If Light_Check(Index) Then
         light_list(Index).Red = 150
         light_list(Index).Blue = 150
         light_list(Index).Green = 150
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function
Public Sub Light_Destroy_ToMap(ByVal X As Byte, ByVal Y As Byte)
    Dim Index As Long
    
    For Index = 1 To light_last
        If light_list(Index).Map_X = X And light_list(Index).Map_Y = Y Then
           light_list(Index).Active = False
           Call Light_Destroy(Index)
        End If
    Next Index
End Sub
Private Sub Light_Destroy(ByVal light_index As Long)
Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).Active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim colorz As Long
    colorz = D3DColorXRGB(150, 150, 150)
    
    'Set up light borders
    min_x = light_list(light_index).Map_X - light_list(light_index).Range
    min_y = light_list(light_index).Map_Y - light_list(light_index).Range
    max_x = light_list(light_index).Map_X + light_list(light_index).Range
    max_y = light_list(light_index).Map_Y + light_list(light_index).Range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = colorz
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = colorz
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = colorz
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = colorz
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = colorz
            MapData(X, min_y).light_value(2) = colorz
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = colorz
            MapData(X, max_y).light_value(3) = colorz
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = colorz
            MapData(min_x, Y).light_value(3) = colorz
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = colorz
            MapData(max_x, Y).light_value(1) = colorz
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = colorz
                MapData(X, Y).light_value(1) = colorz
                MapData(X, Y).light_value(2) = colorz
                MapData(X, Y).light_value(3) = colorz
            End If
        Next Y
    Next X
    
End Sub

