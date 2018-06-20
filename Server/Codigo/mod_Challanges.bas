Attribute VB_Name = "mod_Challanges"
' @ Programado por maTih.-

Option Explicit



Type ChallangeTypes
     EquipoUno(1 To 2)  As Integer      'UI de los del primer team.
     EquipoDos(1 To 2)  As Integer      '************* segundo team.
     Ocupado            As Boolean      'Hay reto en esta arena?
     CuentaRegresiva    As Byte         'Cuenta regresiva.
End Type

Type ChallangeDatas
     MapArenas          As Integer      'Mapa donde están las arenas.
     EquipoUno(1 To 10) As WorldPos     'Posiciónes de las esquinas de los primeros equipos.
     EquipoDos(1 To 10) As WorldPos     '***************************** los segundos equipos.
End Type

Public Retos            As ChallangeTypes
Public RetosData        As ChallangeDatas
Private RetosFile       As String

Sub Inicializar()

' @ Inicializa todo, las arenas y esquinas.

With RetosData

Dim tempLoader  As New clsIniReader
Dim LoopX       As Long
Dim TempString  As String

     RetosFile = App.Path & "\Dat\Retos.txt"
     
     'Initialize reader.
     tempLoader.Initialize RetosFile
     
     'Load the map.
     .MapArenas = val(tempLoader.GetValue("INIT", "Mapa"))

     'Carga las esquinas.
     
     'Esquinas equipo UNO.
     For LoopX = 1 To 10
         With .EquipoUno(LoopX)
              .Map = RetosData.MapArenas
              TempString = tempLoader.GetValue("ESQUINAS", "Uno" & LoopX)
              .X = val(ReadField(1, TempString, Asc("-")))
              .Y = val(ReadField(2, TempString, Asc("-")))
         End With
     Next LoopX
     
     'Esquinas equipo DOS.
     For LoopX = 1 To 10
         With .EquipoDos(LoopX)
              .Map = RetosData.MapArenas
              TempString = tempLoader.GetValue("ESQUINAS", "Dos" & LoopX)
              .X = val(ReadField(1, TempString, Asc("-")))
              .Y = val(ReadField(2, TempString, Asc("-")))
         End With
     Next LoopX
     
End With

End Sub

Sub Enviar(ByVal sendIndex As Integer, ByVal Oponente As String, ByVal cOponente As String, ByVal Compañero As String)

' @ Envia un reto.

Dim ArrUsers(1 To 3) As Integer
Dim LoopX            As Long

ArrUsers(1) = NameIndex(Oponente)
ArrUsers(2) = NameIndex(cOponente)
ArrUsers(3) = NameIndex(Compañero)

With UserList(sendIndex).UserReto
     
     .Envio = True
     .CContrincante = ArrUsers(2)
     .Compañero = ArrUsers(3)
     .Contrincante = ArrUsers(1)
     
     Call Protocol.WriteConsoleMsg(sendIndex, "Reto enviado!", FontTypeNames.FONTTYPE_DIOS)
     
End With

For LoopX = 1 To 3
     
    'El tercero del array es el compañero de quien envia reto
    Select Case LoopX
        
       Case 1           'Oponente uno.
            Call Protocol.WriteConsoleMsg(ArrUsers(LoopX), UserList(sendIndex).name & " y " & Compañero & " los desafian a ti y a " & UserList(ArrUsers(LoopX + 1)).name & " si aceptas tipea /AceptarReto, Si no, /RechazarReto", FontTypeNames.FONTTYPE_DIOS)
        
       Case 2           'Oponente dos.
            Call Protocol.WriteConsoleMsg(ArrUsers(LoopX), UserList(sendIndex).name & " y " & Compañero & " los desafian a ti y a " & UserList(ArrUsers(LoopX - 1)).name & " si aceptas tipea /AceptarReto, Si no, /RechazarReto", FontTypeNames.FONTTYPE_DIOS)
            
       Case 3           'Compañero de sendIndex.
            Call Protocol.WriteConsoleMsg(ArrUsers(LoopX), UserList(sendIndex).name & " Quiere que seas su compañero en un reto 2vs2 contra " & Oponente & " y " & cOponente & " si aceptas tipea /AceptarReto, si no, /RechazarReto", FontTypeNames.FONTTYPE_DIOS)
            
    End Select
    
    UserList(ArrUsers(LoopX)).UserReto.MeEnvio = sendIndex
    
Next LoopX

End Sub
