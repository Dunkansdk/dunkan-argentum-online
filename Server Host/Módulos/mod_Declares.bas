Attribute VB_Name = "mod_Declares"
' @ Modulo de declaraciones - maTih.-

Option Explicit

Public Enum IncomingPackets
       GetServerList        'Busca lista de servers.
       GetServerData        'Busca info de un server.
       AddServerToList      'Agrega un nuevo server.
       ChangeRulesOfServer  'Cambia reglas del server.
       updateUsers          'Actualiza la cantidad de usuarios.
End Enum

Public Enum OutgoingPackets
       SendServerList       'Envia lista de servers.
       SendServerData       'envia info de un server.
End Enum

Public Waiting_Close    As Boolean

Public Sub Inicializar(ByRef tmp_Arr() As estructuraServers)

' @ Inicializa un tippo de array.

ReDim Preserve tmp_Arr(1 To 1) As estructuraServers

End Sub

Public Sub Actualizar_Lista()

' @ Muestra la lista de servers.

With frm_Main.lst_Svs

     .Clear
     
     Dim i  As Long
     
     For i = 1 To UBound(Server_List())
         'Si está online lo agrega a la lista.
         If Server_List(i).Online Then
            .AddItem Server_List(i).Nombre
         End If
     Next i

End With

End Sub

Public Function Found_Server_Index(ByVal sData As String) As Integer

' @ Encuentra el índice del servidor.

Dim arrTmp() As String

arrTmp = Split(sData, "@")

If arrTmp(1) <> vbNullString Then
   Found_Server_Index = Val(arrTmp(1))
End If

End Function

Public Function Found_Clear_Server() As Integer

' @ Busca un slot libre.

Dim i   As Long

For i = 1 To UBound(Server_List())
    'Busca uno que no esté online.
    If Not Server_List(i).Online Then
       'Guarda el index
       Found_Clear_Server = CInt(i)
       Exit Function
    End If
Next i

'Si llegó aca entonces en la lista no encontró slot, da uno nuevo.
ReDim Preserve Server_List(1 To (UBound(Server_List()) + 1)) As estructuraServers

Found_Clear_Server = UBound(Server_List())

End Function

Public Function Found_Server_IP(ByVal sData As String) As String

' @ Devuelve la ip.

Dim tmpArr() As String

tmpArr = Split(sData, "@")

If UBound(tmpArr()) > 1 Then
   Found_Server_IP = tmpArr(2)
End If

End Function

Public Function Found_Server_Name(ByVal sData As String) As String

' @ Devuelve el name.

Dim tmpArr() As String

tmpArr = Split(sData, "@")

If UBound(tmpArr()) > 1 Then
   Found_Server_Name = tmpArr(1)
End If

End Function

Function Search_Index_By_Name(ByVal sName As String) As Integer

'
' @ Devuelve el índice de un servidor by el nombre.

Dim i   As Long

For i = 1 To UBound(Server_List())
    If UCase$(Server_List(i).Nombre) = UCase$(sName) Then
       Search_Index_By_Name = CInt(i)
       Exit Function
    End If
Next i

Search_Index_By_Name = 0

End Function
