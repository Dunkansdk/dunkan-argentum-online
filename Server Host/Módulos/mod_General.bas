Attribute VB_Name = "mod_General"
' modulo de usos generales - maTih.-

Type estructuraReglas
     Vale_Invisibilidad     As Boolean              'Regla invi.
     Vale_Estupidez         As Boolean              'Regla turbación.
     Vale_Paralizar         As Boolean              'Regla inmo.
     Clase_Valida(1 To 12)  As Boolean              'Clases válidas.
End Type

Type estructuraServers
     Nombre                 As String               'Nombre del server.
     Internet_Protocol      As String               'IP.
     Reglas                 As estructuraReglas     'Reglas del server.
     Online                 As Boolean              'Servidor online?
     NumUsers               As Byte
     MaxUsers               As Byte
End Type

Public Server_List_Port     As Integer
Public Server_List()        As estructuraServers

Public Sub Handle_Incoming_Data(ByRef sData As String)

' Maneja la data entrante al servidor.
            
Select Case Left$(UCase$(sData), 1)
            
            'Requiere lista de servers.
            Case CStr(IncomingPackets.GetServerList)
                 Call Handle_GetServerList
                             
            'Requiere info de un server.
            Case CStr(IncomingPackets.GetServerData)
                 Call Handle_GetServerData(sData)
                 
            'Agrega un servidor a la lista.
            Case CStr(IncomingPackets.AddServerToList)
                 Call Handle_AddServerToList(sData)
            
            Case CStr(IncomingPackets.updateUsers)
                Call Handle_UpdateUsers(sData)
                 
            Case CStr(IncomingPackets.ChangeRulesOfServer)
                 Call Handle_ChangeRulesOfServer(sData)

End Select

End Sub

Public Sub Handle_GetServerList()

' @ Requiere lista de servidores.

On Error GoTo errhandler:

Dim i               As Long
Dim lastServer      As Integer
Dim endString       As String

'Agrega a la lista.
For i = 1 To UBound(Server_List())
    'Online?
    If Server_List(i).Online Then
       endString = endString & Server_List(i).Nombre & ","
    Else
       endString = endString & "Nada,"
    End If
Next i

Outgoing_SendServerList (endString)

'Envia la lista.

errhandler:

Debug.Print "Error en 'handle_GetServerList()'"
End Sub

Public Sub Handle_GetServerData(ByRef sData As String)

' @ Requiere IP de un servidor.

Dim server_Index    As Integer  '< Indice de server.

'Encuentra el svIndex
server_Index = Found_Server_Index(sData)

'No hay
If (Not server_Index <> 0) Or (server_Index > UBound(Server_List())) Then Exit Sub

'Prepara la data.
Outgoing_SendServerData Server_List(server_Index).Internet_Protocol, Server_List(server_Index).Nombre, Server_List(server_Index).NumUsers, Server_List(server_Index).MaxUsers

End Sub

Public Sub Handle_AddServerToList(ByRef sData As String)

' @ Agrega un servidor a la lista.

Dim server_Index    As Integer      '< Servidor al que entra.
Dim tmp_IP          As Single

server_Index = mod_Declares.Found_Clear_Server

'No hay slot ? wtf.
If Not server_Index <> 0 Then Exit Sub

Dim tmp() As String

tmp = Split(sData, "@")

If Search_Index_By_Name(tmp(1)) <> 0 Then Exit Sub

With Server_List(server_Index)
     .Nombre = tmp(1)
     .Internet_Protocol = tmp(2)
     .MaxUsers = Val(tmp(3))
     .Online = True
     Call mod_Declares.Actualizar_Lista
     frm_Main.wskData.Close
     frm_Main.wskData.LocalPort = 555
     frm_Main.wskData.Listen
End With

End Sub

Public Sub Handle_ChangeRulesOfServer(ByRef sData As String)

' @ Cambia las reglas del servidor

Dim server_Index    As Integer
Dim Rule_Index      As Byte         '< Indice de la regla que cambia.

server_Index = mod_Declares.Found_Server_Index(sData)

'NO sv.
If Not server_Index <> 0 Then Exit Sub

'set the rules
With Server_List(server_Index).Reglas

     Select Case Rule_Index
            Case 1   '< Desactiva invi.
                 .Vale_Invisibilidad = Not .Vale_Invisibilidad
                 
     End Select
     
End With

End Sub

Public Sub Handle_UpdateUsers(ByVal sData As String)

' @ Actualiza la cantidad de usuarios.

Dim tmp()   As String
Dim svIndex As Integer

tmp = Split(sData, "@")

svIndex = Search_Index_By_Name(tmp(1))

MsgBox UBound(tmp())

MsgBox tmp(1)

MsgBox (tmp(2))

MsgBox tmp(0)

If svIndex <> 0 Then
   Server_List(svIndex).NumUsers = Val(tmp(2))
   MsgBox Server_List(svIndex).NumUsers
End If

frm_Main.wskData.Close
frm_Main.wskData.LocalPort = 555
frm_Main.wskData.Listen

End Sub

Public Sub Outgoing_SendServerList(ByRef listSvs As String)

' @ Envia la lista de servidores.

Dim endString As String

endString = CStr(OutgoingPackets.SendServerList)

endString = endString & "@" & listSvs

Outgoing_SendData endString

End Sub

Public Sub Outgoing_SendServerData(ByRef IP_Server As String, ByRef SV_Name As String, ByRef SV_Users As Byte, ByVal SV_MaxUser As Byte)

' @ Envia la info de un server

Dim endString As String

endString = CStr(OutgoingPackets.SendServerData)

endString = endString & "@" & IP_Server

endString = endString & "@" & SV_Name

endString = endString & "@" & CStr(SV_Users) & "/" & CStr(SV_MaxUser)

Outgoing_SendData endString

End Sub

Public Sub Outgoing_SendData(ByRef sData As String)

' @ Envia info a un cliente.

With frm_Main.wskData
     If frm_Main.wskData.State = sckConnected Then
        .SendData sData
     End If
End With

End Sub


