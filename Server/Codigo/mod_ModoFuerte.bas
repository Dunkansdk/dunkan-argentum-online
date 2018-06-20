Attribute VB_Name = "mod_ModoFuerte"
'programado por maTih.- el 11 de mayo de 2012

Option Explicit

Const MAPA_EVENTO       As Integer = 171    'Mapa del evento.

'Base crimis.
Const RESPAWN_CRIMIS_X  As Byte = 59        'PosX del respawn de crimis
Const RESPAWN_CRIMIS_Y  As Byte = 50        'PosY del repawn de criminales

'Base ciudas.
Const RESPAWN_CIUDAS_X  As Byte = 50        'PosX del respawn de ciudas
Const RESPAWN_CIUDAS_Y  As Byte = 50        'PosY del repawn de ciudas

'Indices de facciones.
Const INDEX_CRIMINAL    As Byte = 2
Const INDEX_CIUDADANO   As Byte = 1

'Maximos jugadores por equipo.
Const MAX_JUGADORES     As Byte = 5

Public FTRIGGER(1 To 2) As Byte

Type tFaccionesEvento
     Contador       As Single       'Contador del %.
     Users(1 To 5)  As Integer      'UserIndexs.
     Muertos        As Byte         'Usuarios muertos.
     Ingresaron     As Byte         'Cuantos usuarios entraron a esta faccion.
     JugandoAhora   As Byte         'Usuarios jugando (contador)
End Type

Type tFuerte
     Fuerte(1 To 2) As tFaccionesEvento
     EventoEnabled  As Boolean      'Si hay evento.
End Type

Public Modo_Fuerte  As tFuerte

Sub Activar(Optional ByVal OrganizadoPor As String = "Automaticamente")

' @ Activa un nuevo evento.

Dim i           As Long
Dim clearFuerte As tFaccionesEvento

FTRIGGER(INDEX_CIUDADANO) = 7     '<Numero del trigger para que suba el contador de % ciudas
FTRIGGER(INDEX_CRIMINAL) = 8     '<Numero del trigger para que suba el contador de % criminal

With Modo_Fuerte
     .EventoEnabled = True
     
     For i = 1 To 2
         .Fuerte(i) = clearFuerte
     Next i
     
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento Fuerte> Organizado " & OrganizadoPor & " 5 cupos por facción, para ingresar /FUERTE.", FontTypeNames.FONTTYPE_EJECUCION)
    
End With

End Sub

Sub Ingresar(ByVal userIndex As Integer)

' @ Usuario se anota al evento.

Dim FIndex      As Byte         '< Indice de la facción al que ingresa
Dim EIndex      As Byte         '< Indice del usuario en el evento.

If criminal(userIndex) Then
   FIndex = INDEX_CRIMINAL
Else
   FIndex = INDEX_CIUDADANO
End If

With Modo_Fuerte.Fuerte(FIndex)

     EIndex = BuscarSlot(FIndex)
     
     'NO hay slot?
     If Not EIndex <> 0 Then
        Call Protocol.WriteConsoleMsg(userIndex, "No hay más cupos para tu facción en el evento.", FontTypeNames.FONTTYPE_EJECUCION)
        Exit Sub
     End If
     
     .Users(EIndex) = userIndex
     
     'Warp a la base.
     WarpUserChar userIndex, MAPA_EVENTO, BuscarX(FIndex), BuscarY(FIndex), True
     
     'Envia mensaje.
     Call EnviarMensaje(FIndex, UserList(userIndex).Name & " Se inscribió para el equipo.")
                
     'Si llena el cupo.
     If .Ingresaron >= MAX_JUGADORES Then
        'Si el usuario era criminal, AND el team ciudadano estaba lleno, comienza.
        If criminal(userIndex) And (Modo_Fuerte.Fuerte(INDEX_CIUDADANO).Ingresaron >= MAX_JUGADORES) Then
        End If
     End If
End With

End Sub

Sub MuereUsuario(ByVal userIndex As Integer)

' @ Muere un usuario en el evento.

With UserList(userIndex)

     Dim F_Index    As Byte     '< Faccion Index.
     
     If criminal(userIndex) Then
        F_Index = INDEX_CRIMINAL
     Else
        F_Index = INDEX_CIUDADANO
     End If
     
     'Actualiza el contador de jugadores.
     Modo_Fuerte.Fuerte(F_Index).JugandoAhora = ActualizarJugandoAhora(F_Index)
    
     'Suma el contador de usuarios muertos.
     Modo_Fuerte.Fuerte(F_Index).Muertos = Modo_Fuerte.Fuerte(F_Index).Muertos + 1
     
     'Murieron todos? resetea el contador.
     If Modo_Fuerte.Fuerte(F_Index).Muertos >= Modo_Fuerte.Fuerte(F_Index).JugandoAhora Then
        Modo_Fuerte.Fuerte(F_Index).Contador = 0
        EnviarMensaje F_Index, "Fuerte " & NombreFaccion(F_Index) & "0%"
     End If
     
End With

End Sub

Sub CheckearTriggers(ByVal FACCION_INDEX As Byte)

' @ Controla los Trigger de un team

Dim i   As Long

For i = 1 To MAX_JUGADORES
    With Modo_Fuerte.Fuerte(FACCION_INDEX)
         'Hay usuario?
         If .Users(i) <> 0 Then
            'Está logeado?
            If UserList(.Users(i)).ConnID <> -1 Then
               'Está en posición del trigger?
               If Not MapData(UserList(.Users(i)).Pos.Map, UserList(.Users(i)).Pos.Y, UserList(.Users(i)).Pos.X).trigger <> FTRIGGER(FACCION_INDEX) Then
                  'Sube el contador y cierra.
                  SendData SendTarget.toMap, MAPA_EVENTO, PrepareMessageConsoleMsg("Fuerte " & NombreFaccion(FACCION_INDEX) & CStr(.Contador) & "%", FonttypeFaccion(FACCION_INDEX))
                  'Llegó a 100%?
                  If (CInt(.Contador) >= 100) Then
                     Call GanaTeam(FACCION_INDEX)
                  End If
                  Exit Sub
               End If
            End If
         End If
    End With
Next i

'Estamos aca, no habia usuarios en los triggers, resetea.
Modo_Fuerte.Fuerte(FACCION_INDEX).Contador = CSng(0)

End Sub

Sub EnviarMensaje(ByVal INDEX_FACCION As Byte, ByVal Message As String)

' @ Envia el mensaje a un team.

Dim i   As Long

    For i = 1 To MAX_JUGADORES
        'Si hay usuario y esta logeado enviamos
        If Modo_Fuerte.Fuerte(INDEX_FACCION).Users(i) <> 0 Then
           If UserList(Modo_Fuerte.Fuerte(INDEX_FACCION).Users(i)).ConnID <> -1 Then
              Call Protocol.WriteConsoleMsg(Modo_Fuerte.Fuerte(INDEX_FACCION).Users(i), Message, FontTypeNames.FONTTYPE_CONSEJOVesA)
           End If
        End If
    Next i

End Sub

Sub GanaTeam(ByVal INDEX_FACCION As Byte)

' @ Declara ganador una facción.

Dim i           As Long
Dim userIndex   As Integer

With Modo_Fuerte
     .EventoEnabled = False         '< Desactiva el evento.
     
     'Se lleva a los perdedores.
     For i = 1 To MAX_JUGADORES
         'Guarda el UI.
         userIndex = .Fuerte(IIf(INDEX_FACCION > 1, 1, 2)).Users(i)
         'Está logeado?
         If userIndex <> 0 Then
            If UserList(userIndex).ConnID <> -1 Then
               'Lleva a ulla.
               WarpUserChar userIndex, 1, (50 + CInt(i)), 50, True
            End If
         End If
     Next i
     
     'Ahora setea los ganadores.
     For i = 1 To MAX_JUGADORES
         userIndex = .Fuerte(INDEX_FACCION).Users(i)
         
         'Hay user
         If userIndex <> 0 Then
            'Está logeado
            If (UserList(userIndex).ConnID <> -1) And (UserList(userIndex).ConnIDValida) Then
               'Lleva a ulla.
               WarpUserChar userIndex, 1, (50 + CInt(i)), 50, True
               'DARLE PREMIO ***************************
               UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + 1000000
               Call Protocol.WriteUpdateGold(userIndex)
            End If
         End If
     Next i
     
     SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Modo Fuerte> El bando " & NombreFaccion(INDEX_FACCION) & " ganó la batalla!", FonttypeFaccion(INDEX_FACCION))
     
     Dim clearFuerte    As tFaccionesEvento
     
     .Fuerte(1) = clearFuerte
     .Fuerte(2) = clearFuerte
     
End With

End Sub

Function BuscarSlot(ByVal INDEX_FACCION As Byte) As Byte

' @ Devuelve slot en el evento para un usuario.

Dim loopX   As Long

For loopX = 1 To MAX_JUGADORES
    With Modo_Fuerte.Fuerte(INDEX_FACCION)
         If Not .Users(loopX) <> 0 Then
            BuscarSlot = CByte(loopX)
            .Ingresaron = .Ingresaron + 1
         End If
    End With
Next loopX

BuscarSlot = 0

End Function

Function BuscarX(ByVal INDEX_FACCION As Byte) As Byte

' @ Busca una posición X para un usuario de una facción.

Select Case INDEX_FACCION
       Case INDEX_CIUDADANO         '< Ciudadano.
            BuscarX = RESPAWN_CIUDAS_X
            
       Case INDEX_CRIMINAL          '< Criminal
            BuscarX = RESPAWN_CRIMIS_X
End Select

End Function

Function BuscarY(ByVal INDEX_FACCION As Byte) As Byte

' @ Busca una posición Y para un usuario de una facción.

Select Case INDEX_FACCION
       Case INDEX_CIUDADANO         '< Ciudadano.
            BuscarY = RESPAWN_CIUDAS_Y
            
       Case INDEX_CRIMINAL          '< Criminal
            BuscarY = RESPAWN_CRIMIS_Y
End Select

End Function

Function FonttypeFaccion(ByVal INDEX_FACCION As Byte) As FontTypeNames

' @ Devuelve la font para una faccion

Select Case INDEX_FACCION
       Case INDEX_CIUDADANO         '< Ciudadano.
            FonttypeFaccion = FontTypeNames.FONTTYPE_CONSEJOVesA
            
       Case INDEX_CRIMINAL          '< Criminal
            FonttypeFaccion = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
End Select

End Function

Function NombreFaccion(ByVal INDEX_FACCION As Byte) As String

' @ Devuelve el nombre para una faccion.

Select Case INDEX_FACCION
       Case INDEX_CIUDADANO         '< Ciudadano.
            NombreFaccion "Armada"
            
       Case INDEX_CRIMINAL          '< Criminal
            NombreFaccion = "Caos"
End Select

End Function

Function ActualizarJugandoAhora(ByVal F_Index As Byte) As Byte

' @ Devuelve los jugadores que tiene la faccion

With Modo_Fuerte.Fuerte(F_Index)
     Dim i  As Long
     
         For i = 1 To MAX_JUGADORES
             'Hay un index.
             If .Users(i) <> 0 Then
                'Si está logeado suma
                If (UserList(.Users(i)).ConnIDValida) And (UserList(.Users(i)).ConnID <> -1) Then
                    ActualizarJugandoAhora = ActualizarJugandoAhora + 1
                Else
                    'Hay un INDEX pero no está logeado, lo borra.
                    .Users(i) = 0
                End If
             End If
         Next i
End With

End Function

Function EnEvento(ByVal userIndex As Integer) As Boolean

' @ Devuelve si está en el evento

Dim F_Index As Byte     '< Faccion Index.
Dim F_Loop  As Long     '< Bucle de los usuarios.
Dim U_Index As Integer  '< UserIndex.

'Busca el INDICE segun su status.
If criminal(userIndex) Then
   F_Index = INDEX_CRIMINAL
Else
   F_Index = INDEX_CIUDADANO
End If

For F_Loop = 1 To MAX_JUGADORES
    With Modo_Fuerte.Fuerte(F_Index)
         U_Index = .Users(F_Loop)
         'Si no es diferente al que buscamos
         If Not U_Index <> userIndex Then
            EnEvento = True
            Exit Function
         End If
    End With
Next F_Loop

EnEvento = False

End Function
