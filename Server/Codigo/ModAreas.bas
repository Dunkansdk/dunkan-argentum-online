Attribute VB_Name = "ModAreas"
'**************************************************************
' ModAreas.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Original Idea by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
' Implemented by Lucio N. Tourrilhes (DuNga)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

' Modulo de envio por areas compatible con la versión 9.10.x ... By DuNga

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    MinX As Integer '-!!!
    MinY As Integer '-!!!
    
    AreaID As Long
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type

Public Const USER_NUEVO As Byte = 255

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte

Private AreasRecive(12) As Integer

Public ConnGroups() As ConnGroup

Public Sub InitAreas()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim loopX As Long

' Setup areas...
    For loopC = 0 To 11
        AreasRecive(loopC) = (2 ^ loopC) Or IIf(loopC <> 0, 2 ^ (loopC - 1), 0) Or IIf(loopC <> 11, 2 ^ (loopC + 1), 0)
    Next loopC
    
    For loopC = 1 To 100
        PosToArea(loopC) = loopC \ 9
    Next loopC
    
    For loopC = 1 To 100
        For loopX = 1 To 100
            'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
            AreasInfo(loopC, loopX) = (loopC \ 9 + 1) * (loopX \ 9 + 1)
        Next loopX
    Next loopC

'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For loopC = 1 To NumMaps
        ConnGroups(loopC).OptValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour))
        
        If ConnGroups(loopC).OptValue = 0 Then ConnGroups(loopC).OptValue = 1
        ReDim ConnGroups(loopC).UserEntrys(1 To ConnGroups(loopC).OptValue) As Long
    Next loopC
End Sub

Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim loopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        
        For loopC = 1 To NumMaps
            EntryValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(loopC).OptValue) \ 2))
            
            ConnGroups(loopC).OptValue = Val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, tCurDay & "-" & tCurHour))
            If ConnGroups(loopC).OptValue = 0 Then ConnGroups(loopC).OptValue = 1
            If ConnGroups(loopC).OptValue >= MapInfo(loopC).NumUsers Then ReDim Preserve ConnGroups(loopC).UserEntrys(1 To ConnGroups(loopC).OptValue) As Long
        Next loopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 15/07/2009
'Es la función clave del sistema de areas... Es llamada al mover un user
'15/07/2009: ZaMa - Now it doesn't send an invisible admin char info
'**************************************************************
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long, map As Long
    
    With UserList(UserIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        
        map = .Pos.map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY

                '<<< User >>>
                If MapData(map, X, Y).UserIndex Then
                    
                    TempInt = MapData(map, X, Y).UserIndex
                    
                    If UserIndex <> TempInt Then
                        
                        ' Solo avisa al otro cliente si no es un admin invisible
                        If Not (UserList(TempInt).flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, UserIndex, TempInt, map, X, Y)
                            
                            'Si el user estaba invisible le avisamos al nuevo cliente de eso
                            If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                                    Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                                End If
                            End If
                        End If
                        
                        ' Solo avisa al otro cliente si no es un admin invisible
                        If Not (.flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, TempInt, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
                            
                            If .flags.invisible Or .flags.Oculto Then
                                If UserList(TempInt).flags.Privilegios And PlayerType.User Then
                                    Call WriteSetInvisible(TempInt, .Char.CharIndex, True)
                                End If
                            End If
                        End If
                        
                        Call FlushBuffer(TempInt)
                    
                    ElseIf Head = USER_NUEVO Then
                        If Not ButIndex Then
                            Call MakeUserChar(False, UserIndex, UserIndex, map, X, Y)
                        End If
                    End If
                End If
                
                'Bot : P
                If MapData(map, X, Y).botIndex <> 0 Then
                    Debug.Print "ASE? "
                   'Si no está invocado no creo nada.
                   If ia_Bot(MapData(map, X, Y).botIndex).Invocado Then
                      ia_EnviarChar UserIndex, MapData(map, X, Y).botIndex
                   End If
                End If
            
            Next Y
        Next X
        
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' Se llama cuando se mueve un Npc
'**************************************************************
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long
    
    With Npclist(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100


        'Precalculados :P
        TempInt = .Pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    Dim loopC As Long
    
    'Search for the user
    For loopC = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(loopC) = UserIndex Then Exit For
    Next loopC
    
    'Char not found
    If loopC > ConnGroups(map).CountEntrys Then Exit Sub
    
    'Remove from old map
    ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys - 1
    TempVal = ConnGroups(map).CountEntrys
    
    'Move list back
    For loopC = loopC To TempVal
        ConnGroups(map).UserEntrys(loopC) = ConnGroups(map).UserEntrys(loopC + 1)
    Next loopC
    
    If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
    End If
End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal map As Integer, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************
On Error GoTo errak
    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
    
1    If Not MapaValido(map) Then Exit Sub
    
2     EsNuevo = True
    
    'Prevent adding repeated users
3    For i = 1 To ConnGroups(map).CountEntrys
4        If ConnGroups(map).UserEntrys(i) = UserIndex Then
5            EsNuevo = False
6            Exit For
7        End If
8    Next i
    
9    If EsNuevo Then
        'Update map and connection groups data
11        ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys + 1
12        TempVal = ConnGroups(map).CountEntrys
        
13        If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim
14            ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
15        End If
        
16        ConnGroups(map).UserEntrys(TempVal) = UserIndex
    End If
    
17    With UserList(UserIndex)
        'Update user
        .AreasInfo.AreaID = 0
        
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
        .AreasInfo.AreaReciveX = 0
        .AreasInfo.AreaReciveY = 0
    End With
    
18    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, ButIndex)

Exit Sub
errak:
    Debug.Print "Linea " & Erl()
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    With Npclist(NpcIndex)
        .AreasInfo.AreaID = 0
        
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
        .AreasInfo.AreaReciveX = 0
        .AreasInfo.AreaReciveY = 0
    End With
    
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub
