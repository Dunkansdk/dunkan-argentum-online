Attribute VB_Name = "mod_Embarcaciones"
' programado por maTih.-

#If Barcos <> 0 Then

Option Explicit

Public Const NUMEMBARCACIONES As Byte = 5
Public Const NUMPASAJEROS     As Byte = 4
Public Const NUMCHARBARCOS    As Integer = (10000 - NUMEMBARCACIONES)

Type Embarcacion
     Destino        As WorldPos         'Destino del barco.
     Salida         As WorldPos         'Pos inicial del barco.
     Actual         As WorldPos         'Pos actual.
     Tripulantes    As Byte             'Cuantos lleva.
     UserIndex()    As Integer          'Usuarios de la embarcación.
     Activada       As Boolean          'Activada esta embarcación?
     Zarpo          As Boolean          'Si ya salió a destino.
     Char           As Char             'Char del barco.
     VelMovimiento  As Byte             'Movimienot.
End Type

Public Embarcaciones(1 To NUMEMBARCACIONES) As Embarcacion

Sub Pasajero(ByVal UserIndex As Integer, ByVal eIndex As Byte)

' @ Anota un nuevo pasajero a un barco

With Embarcaciones(eIndex)

     'No está activada.
     If Not .Activada Then Exit Sub
     
     'Ya sarpó.
     If .Zarpo Then Exit Sub
     
     'No hay mas lugar.
     If .Tripulantes >= NUMPASAJEROS Then Exit Sub
     
     If YaInscripto(UserIndex, eIndex) Then Exit Sub
     
     .Tripulantes = .Tripulantes + 1
     
     .Zarpo = True
     
     'Se anota.
     ReDim Preserve .UserIndex(1 To .Tripulantes) As Integer
     
     'guarda el UI.
     .UserIndex(.Tripulantes) = UserIndex
     
     'Warp sobre el barco.
     WarpUserChar UserIndex, .Actual.map, (.Actual.X - .Tripulantes), .Actual.Y, False
     
     'Avisa
     WriteConsoleMsg UserIndex, "Estás en la tripulación del barco hacia " & MapInfo(.Destino.map).Name & ".", FontTypeNames.FONTTYPE_DIOS
     
     WritePosUpdate UserIndex
     
End With

End Sub

Sub Crear(ByRef Inicio As WorldPos, ByRef Destino As WorldPos)

' @ Crea el char.

Dim eIndex      As Byte         '< Embarcación index.

eIndex = Indice

If Not eIndex <> 0 Then Exit Sub

With Embarcaciones(eIndex)
     
     .Actual = Inicio
     .Salida = Inicio
     .Destino = Destino
     
     .Activada = True
     .Tripulantes = 0
     
     With .Actual
          MapData(.map, .X, .Y).EmbarcacionIndex = eIndex
     End With
     
     'Llena el char.
     With .Char
          .body = iGaleonCiuda
          .Head = 0
          .heading = eHeading.SOUTH
          .CascoAnim = NingunCasco
          .FX = 0
          .loops = 0
          .WeaponAnim = NingunArma
          .ShieldAnim = NingunEscudo
          
          'Crea.
          EnviarDatos eIndex, PrepareMessageCharacterCreate(.body, 0, eHeading.SOUTH, NUMCHARBARCOS + eIndex, Inicio.X, Inicio.Y, 0, 0, 0, 0, 0, "Embarcación hacia <" & MapInfo(Destino.map).Name & ">", 0, 0)
          
          EnviarDatos eIndex, PrepareMessageChatOverHead("Esperando tripulantes para una embarcación hacia Dungeon Veril.", NUMCHARBARCOS + eIndex, vbRed)
          
     End With
     
End With

End Sub

Sub Mover(ByVal eIndex As Byte)

' @ Mueve una embarcación.

With Embarcaciones(eIndex)
     
     Dim eMoviment  As eHeading

     'Velocidad de movimiento.
     If .VelMovimiento <> 0 Then .VelMovimiento = .VelMovimiento - 1: Exit Sub
    
     'Setea el contador.
     .VelMovimiento = 3
        
     'Busca una ruta.
     eMoviment = ruta(eIndex)
     
     'Se mueve hacia la posición
     'No hay ruta?
     
     'Llegó a destino?
     If .Actual.map = .Destino.map Then
        If .Actual.X = .Destino.X Then
           If .Actual.Y = .Destino.Y Then
              Call DescargarPasajeros(eIndex)
           End If
        End If
     End If
     
     If Not eMoviment <> 0 Then Exit Sub
     
     MoverDireccion eIndex, eMoviment
    
End With

End Sub

Sub DescargarPasajeros(ByVal eIndex As Byte)

' @ Descarga los pasajeros de un barco

Dim i   As Long

With Embarcaciones(eIndex)

    For i = 1 To .Tripulantes
        If .UserIndex(i) <> 0 Then
           If UserList(.UserIndex(i)).ConnID <> -1 Then
              WarpUserChar .UserIndex(i), .Destino.map, .Destino.X, .Destino.Y - (i + 3), True
           End If
        End If
    Next i
    
End With
End Sub


Sub MoverUsuarios(ByVal eIndex As Byte, ByVal EMov As eHeading)

On Error GoTo HError

' @ Mueve los usuariso d un barco.

Dim i   As Long

With Embarcaciones(eIndex)
     
     'No hay tripulantes?
     If Not .Tripulantes <> 0 Then Exit Sub
     
     For i = 1 To .Tripulantes
         'Mueve los chars.
         'Si están logeados.
         If UserList(.UserIndex(i)).ConnID <> -1 Then
            'Call WarpUserChar(.userIndex(i), .Actual.map, .Actual.X, (.Actual.Y - i), False)
            'Call mod_DunkanProtocol.WriteChangeScreen(.userIndex(i), EMov)
            MapData(UserList(.UserIndex(i)).Pos.map, UserList(.UserIndex(i)).Pos.X, UserList(.UserIndex(i)).Pos.Y).UserIndex = 0
            UserList(.UserIndex(i)).Pos.map = .Actual.map
            UserList(.UserIndex(i)).Pos.X = .Actual.X
            UserList(.UserIndex(i)).Pos.Y = .Actual.Y - i
            MapData(.Actual.map, .Actual.X, .Actual.Y - i).UserIndex = .UserIndex(i)
            Call SendData(SendTarget.ToPCArea, .UserIndex(i), PrepareMessageCharacterMove(UserList(.UserIndex(i)).Char.CharIndex, UserList(.UserIndex(i)).Pos.X, UserList(.UserIndex(i)).Pos.Y))
            Call ModAreas.CheckUpdateNeededUser(.UserIndex(i), EMov)
            'Call WriteChangeScreen(.userIndex(i), EMov)
            'Call WriteForceCharMove(.userIndex(i), EMov)
         End If
     Next i

End With

HError:

End Sub

Sub NuevaPos(ByVal UserIndex As Integer)



End Sub

Sub CambiarMapaUsuarios(ByVal eIndex As Byte)

' @ Cambia los mapas de los usuarios-

Dim i   As Long

With Embarcaciones(eIndex)
     
     'No hay tripulantes?
     If Not .Tripulantes <> 0 Then Exit Sub
     
     For i = 1 To .Tripulantes
         'Mueve los chars.
         'Si están logeados.
         If UserList(.UserIndex(i)).ConnID <> -1 Then
            Call WarpUserChar(.UserIndex(i), .Actual.map, .Actual.X - (i + 1), .Actual.Y, False)
         End If
     Next i

End With


End Sub

Sub MoverDireccion(ByVal eIndex As Byte, ByVal eMovimiento As eHeading)

' @ Mueve una embarcación hacia un head.

With Embarcaciones(eIndex)

    Select Case eMovimiento

           Case eHeading.NORTH      '< Norte.
                 MapData(.Actual.map, .Actual.X, .Actual.Y).EmbarcacionIndex = 0
                .Actual.Y = .Actual.Y - 1
                
           Case eHeading.EAST       '< Este.
                 MapData(.Actual.map, .Actual.X, .Actual.Y).EmbarcacionIndex = 0
                .Actual.X = .Actual.X + 1
                
           Case eHeading.SOUTH      '< Sur.
                 MapData(.Actual.map, .Actual.X, .Actual.Y).EmbarcacionIndex = 0
                .Actual.Y = .Actual.Y + 1
                
           Case eHeading.WEST       '< Oeste.
                 MapData(.Actual.map, .Actual.X, .Actual.Y).EmbarcacionIndex = 0
                .Actual.X = .Actual.X - 1
                
    End Select
    
    EnviarDatos eIndex, PrepareMessageCharacterMove(NUMCHARBARCOS + eIndex, .Actual.X, .Actual.Y)
    
    MapData(.Actual.map, .Actual.X, .Actual.Y).EmbarcacionIndex = eIndex
    
    'Movemos los clientes de los usuarios
    MoverUsuarios eIndex, eMovimiento
    
    'Se movió y habia un char?
    If MapData(.Actual.map, .Actual.X, .Actual.Y).UserIndex <> 0 Then
       'EL usuario muere instantaneamente.
       'Si no estaba muerto .
       If Not UserList(MapData(.Actual.map, .Actual.X, .Actual.Y).UserIndex).flags.Muerto <> 0 Then
          Call UsUaRiOs.UserDie(MapData(.Actual.map, .Actual.X, .Actual.Y).UserIndex)
          Call Protocol.WriteConsoleMsg(MapData(.Actual.map, .Actual.X, .Actual.Y).UserIndex, "Has sido aplastado por una embarcación!", FontTypeNames.FONTTYPE_DIOS)
       End If
    End If
    
    If MapData(.Actual.map, .Actual.X, .Actual.Y).NpcIndex <> 0 Then
       'Si habia un npc muere.
       Call MuereNpc(MapData(.Actual.map, .Actual.X, .Actual.Y).NpcIndex, 0)
    End If
    
    'Salida?
    If MapData(.Actual.map, .Actual.X, .Actual.Y).TileExit.map <> 0 Then
       'Nuevo map.
       .Actual.map = MapData(.Actual.map, .Actual.X, .Actual.Y).TileExit.map
       
       'Nuevo X?
       If MapData(.Actual.map, .Actual.X, .Actual.Y).TileExit.X <> 0 Then
          .Actual.X = MapData(.Actual.map, .Actual.X, .Actual.Y).TileExit.X
       End If
       
       'Nuevo Y?
       If MapData(.Actual.map, .Actual.X, .Actual.Y).TileExit.Y <> 0 Then
          .Actual.Y = MapData(.Actual.map, .Actual.X, .Actual.Y).TileExit.Y
       End If
       
       MapData(.Actual.map, .Actual.X, .Actual.Y).EmbarcacionIndex = eIndex
       
       'Cambia de mapas los usuarios
       CambiarMapaUsuarios eIndex
       
       'Crea el char.
       EnviarDatos eIndex, PrepareMessageCharacterCreate(.Char.body, 0, eMovimiento, NUMCHARBARCOS + eIndex, .Actual.X, .Actual.Y, 0, 0, 0, 0, 0, "Embarcación hacia <" & MapInfo(.Destino.map).Name & ">", 0, 0)
  
    End If
    
    
    
End With
End Sub

Sub EnviarDatos(ByVal eIndex As Byte, ByRef sData As String)

' @ Envia datos a ela rea de un barco.

With Embarcaciones(eIndex).Actual

    Call modSendData.SendToAreaByPos(.map, .X, .Y, sData)

End With

End Sub

Sub CargarRutas(ByRef MAPFILE As String, ByVal MapIndex As Integer)

' @ Carga las rutas de los barcos.

Dim loopX   As Long
Dim loopY   As Long
Dim NowRuta As String

For loopX = 1 To 100
    For loopY = 1 To 100
        With MapData(MapIndex, loopX, loopY)
             NowRuta = GetVar(App.Path & "\Maps\Mapa" & MapIndex & "Rutas.ini", CStr(loopX) & "," & CStr(loopY), "Direccion")
             
             'Hay?
             If Val(NowRuta) <> 0 Then .Rutas(2) = Val(NowRuta)
             
        End With
    Next loopY
Next loopX
End Sub


Function ruta(ByVal eIndex As Byte) As eHeading

' @ Devuelve una ruta para el barco.

With MapData(Embarcaciones(eIndex).Actual.map, Embarcaciones(eIndex).Actual.X, Embarcaciones(eIndex).Actual.Y)

        ruta = .Rutas(2)

End With

End Function

Function Indice() As Byte

' @ Devuelve slot para nuevo barco.

Dim i   As Long

Indice = 0

For i = 1 To NUMEMBARCACIONES
    If Not Embarcaciones(i).Activada Then
       Indice = CByte(i)
       Exit Function
    End If
Next i

End Function

Function YaInscripto(ByVal UserIndex As Integer, ByVal eIndex As Byte) As Boolean

' @ Ya inscripto en la tripulacion.

Dim i   As Long

YaInscripto = False

If Embarcaciones(eIndex).Tripulantes = 0 Then Exit Function

For i = 1 To Embarcaciones(eIndex).Tripulantes

    If Embarcaciones(eIndex).UserIndex(i) = UserIndex Then
       YaInscripto = True
       Exit Function
    End If

Next i

End Function

#End If
