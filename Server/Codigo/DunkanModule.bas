Attribute VB_Name = "Dunkan_Module"
'HERE WE GO TO WORK THE DIFFERENT SYSTEMS, ALL IN ONE UNIT :D

' / Author : maTih.-
' / Note   : Manejo de todos los sistemas en este módulo

' - Tabulado y organizado por Dunkan

' / Constantes           |||
Public Const CONST_RANDOM As Byte = 4 'Maximo items de "Jugada arriesgada"
Public Const MAX_SOPORTES As Integer = 300

Public Mapa_Duelo As Integer
Public EsperaX As Byte
Public EsperaY As Byte

Public X1_Duelos(1 To 5) As Byte
Public Y1_Duelos(1 To 5) As Byte
Public x2_Duelos(1 To 5) As Byte
Public Y2_Duelos(1 To 5) As Byte

' / Types                |||

Public Type AutoDuelos
   SalasLibres            As Byte
   UserEsperando          As Integer
   UsersDueleando1()      As Integer
   UsersDueleando2()      As Integer
   Cuenta()               As Byte
   ContadorSalas          As Byte
End Type

Public Type Ranking
    UserNames(1 To 10)    As String
    UserLevels(1 To 10)   As Byte
    UserFrags(1 To 10)    As Integer
    UserTag(1 To 10)      As String
    UserClases(1 To 10)   As String
End Type



Public Type GuildvsGuild
    Count1          As Byte     'Contador usuarios guild1
    Count2          As Byte     'Contador usuarios guild2
    Contador        As Byte     'Contador de usuarios máximos
    GuildIndex1     As Integer  'Puntero para el array de clanes
    GuildIndex2     As Integer  'Puntero para el array de clanes 2
    UserIndexs1()   As Integer  'Usuarios Guild1
    UserIndexs2()   As Integer  'Usuarios Guild2
    Ocuped          As Boolean  'Ocupado el cvc??
    Caspers1        As Byte     'Usuarios muertos del guild1
    Caspers2        As Byte     'Usuarios muertos del guild2
End Type

Public Type Torneos
    Cupos           As Byte     'Cupos máximos
    CountCupos      As Byte     'Con esto vamos contando..
    UserIndexs()    As Integer  'Punteros del array Userlist.
    precioInsc      As Long     'Valor de inscripcion
    Hay             As Boolean  '¿Hay un torneo?
    PozoRecaudado   As Long     'Pozo recaudado del oro.
End Type

Public Type ColaSoportes
    Mensajes()  As String
    Usuarios()  As String
    LastMensaje As Byte
End Type

Public Type SoportUser
    Esperando As Boolean  '¿Está esperando una respuesta?
    Respuesta As String   'La respuesta del Game Master.
End Type

Public Type Partys
    ActualExpUno        As Byte     'Actual Exp del Creador
    ActualExpDos        As Byte     'Actual Exp del segundo
    PartyArray(1 To 2)  As String   'Party array de los nombres, para usar el NameIndex
End Type

'Tipo de retos de los usuarios
Public Type RetosUser
    IndexReto       As Byte     'Puntero para el array de retos
    ThereIsReto     As Boolean  'Tiene un reto ?
    tmpGold         As Long     'Variable temporal para el oro
    tmpDrop         As Byte     'Variable temportal si es por items o no
    VuelvePos       As Byte     'Tiempo restante para volver a la ciudad
    TeamReto        As Byte     'TeamReto , usado para 2v2.
    RoundsGanados   As Byte     'Rounds ganados, usado para 1v1.
    Pareja          As Integer  'Pareja, usado para 2v2
End Type

'Tipo de retos para 1vs1
Public Type Retos1v1
    UserIndex1    As Integer  'Puntero para el array de userlist
    UserIndex2    As Integer  ' " "
    Hay           As Boolean  'Si este puntero tiene un reto
    Gold          As Long     'Oro
    Drop          As Byte     'Caen o no los objetos
    CountDown     As Byte     'Cuenta regresiva
End Type

Public Type Retos2v2
    Team1(1 To 2)    As Integer  'Array de los users del equipo1
    Team2(1 To 2)    As Integer  'Array de los users del equipo2
    Hay              As Boolean  'sSi este puntero tiene un reto 2v2
    Oro              As Long     'Oro.
    Drop             As Byte     'Drop
    CountDown        As Byte     'Cuenta regresiva.
    Rounds1          As Byte     'Control de rounds team1
    Rounds2          As Byte     'Lo mismo..
End Type

'Tipo de retos globales
Public Type RetosGlobal
    Retos1v1()    As Retos1v1
    Retos2vs2()   As Retos2v2
    Pointer1vs1   As Byte 'Cantidad de retos 1vs1
    Pointer2vs2   As Byte 'Cantidad de retos 2vs2
End Type

Public Type TypeSubasta
    PresentSale      As Integer  'Actual objIndex .
    PresentAmount    As Integer  'Actual Cantidad.
    ThereIs          As Boolean  'Hay una subasta?
    MiUserIndex      As Integer  'UserIndex del que inicio
    LastUserIndex    As Integer  'UserIndex del ultimo en ofertar.
    LastSale         As Long     'Ultima oferta .
    MiUserName       As String   'Nombre del que inició la subasta, usado por si deslogea.
    MiLastName       As String   'Nombre del Ultimo en ofertar, usado por si cierra.
End Type

'DECLARES            |||
Public tCvC         As GuildvsGuild
Public tSubasta     As TypeSubasta
Public tRetos       As RetosGlobal
Public tSoportes    As ColaSoportes
Public tTorneos     As Torneos
Public tRanking     As Ranking
Public Duelos       As AutoDuelos
Public RankPathFile As String


' / Sistema de subastas -

Public Sub Subasta_Inicia(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Amount As Integer, ByVal minime As Long)

' / Author: maTih
' / Note: Start one sale, by userindex slot & amount.

With tSubasta
    .ThereIs = True
    .LastSale = minime
    .LastUserIndex = UserIndex
    .MiUserIndex = UserIndex
    .PresentAmount = Amount
    .PresentSale = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    .MiUserName = UserList(UserIndex).name
End With

With UserList(UserIndex)
    QuitarUserInvItem UserIndex, Slot, Amount 'Nos llevamos el item
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(" Subasta " & .name & " Está subastando " & Amount & " & objdata(tsubasta.PresentSale) & " & " Con un valor inicial de " & minime & " Para ofertar, tipea /OFERTAR cantidad. La subasta terminará en 3 minutos", FontTypeNames.FONTTYPE_GM)
End With

End Sub

Public Sub Subasta_Termina(Optional ByVal isDesconnection As Boolean = False)

' / Author: maTih
' / Note: In this subroutine terminates the current auction
' / Parameters optional: isDesconnection, used by closesocket of user

Dim obIndex As Obj

With tSubasta
    'Seteamos el objeto
    obIndex.Amount = .PresentAmount
    obIndex.ObjIndex = .PresentSale

    'FIXED.
    'Si lastuserindex = miuserindex, nadie ofreció, te devolvemos el item viejo.
    If .LastUserIndex = .MiUserIndex Then
        MeterItemEnInventario .MiUserIndex, obIndex
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta > terminó sin ninguna oferta", FontTypeNames.FONTTYPE_GUILD)
        Call ResetFlagsSubasta
        Exit Sub
    End If
    
    'FIXED.
    'Nos fijamos si está online
    If .LastUserIndex > 0 Then 'Si lo está, buscamos un slot , y si no lo encontramos, al piso.
    
        If MeterItemEnInventario(.LastUserIndex, obIndex) = False Then Call TirarItemAlPiso(UserList(.LastUserIndex).Pos, obIndex)
    
    Else 'Si está offline, buscamos un puntero temportal en su inv
    
        Dim tmpCharPath As String
        Dim tmpCharSlot As Byte
        
        tmpCharPath = CharPath & UCase$(.MiLastName) & ".chr"
        tmpCharSlot = val(GetVar(tmpCharPath, "Inventory", "CantidadItems"))
    
        If tmpCharSlot < MAX_INVENTORY_SLOTS Then
        
            'Si es menor a MAX_INVENROY_SLOT (Osea 20 creo), lo escribimos.
            WriteVar tmpCharPath, "Inventory", "Obj" & tmpCharSlot + 1, obIndex.ObjIndex & "-" & obIndex.Amount & "-" & 0
    
        Else 'Bien, tenias el inventario lleno, te lo guardo en la boveda.
             'Lo qe hago por los users eh!
            tmpCharSlot = val(GetVar(tmpCharPath, "BancoInventory", "CantidadItems"))
    
            If tmpCharSlot < MAX_BANCOINVENTORY_SLOTS Then 'Ya si no tenes espacio ni en la boveda ni en el inv, los siento.
                WriteVar tmpCharPath, "BancoInventory", "CantidadItems", tmpCharSlot + 1
                WriteVar tmpCharPath, "BancoInventory", "Obj" & tmpCharSlot + 1, obIndex.ObjIndex & "-" & obIndex.Amount
            End If
            
    'Le damos la $$$ al que inició la subasta.
    'FIXED.
    'Nos fijamos si está online (MiUserIndex), si no, escribimos su charfile.
    
    End If
    
        If .MiUserIndex > 0 Then
        
            UserList(.MiUserIndex).Stats.GLD = UserList(.MiUserIndex).Stats.GLD + .LastSale
            WriteUpdateGold .MiUserIndex
            
        Else
        
            Dim tmpCharPaths    As String
            Dim tmpCharGLD      As Long
        
            tmpCharPaths = CharPath & UCase$(.MiUserName) & ".chr"
            tmpCharGLD = val(GetVar(tmpCharPath, "STATS", "GLD"))
        
            'Tenemos todo? Write!
            WriteVar tmpCharPaths, "STATS", "GLD", tmpCharGLD + .LastSale
        End If
        
    End If
    
    ResetFlagsSubasta
    
End With

End Sub

Public Sub ResetFlagsSubasta()

' / Author: maTih
' / Note: Resets all flags of type subasta

With tSubasta
    .LastSale = 0
    .LastUserIndex = 0
    .MiUserIndex = 0
    .PresentSale = 0
    .PresentAmount = 0
    .ThereIs = False
    .MiUserName = vbNullString
    .MiLastName = vbNullString
End With

End Sub

Public Sub Subasta_UsuarioPideInventario(ByVal UserIndex As Integer)
' / Author: maTih
' / Note: User requests the list of objects from your inventory

'Nos ahorramos el envio, pero si hay una subasta.

If tSubasta.ThereIs Then
    WriteConsoleMsg UserIndex, " Ya hay una subasta.", FontTypeNames.FONTTYPE_GUILD
    Exit Sub
End If

'UserIndex está muerto ? si es asi no puede subastar.
If UserList(UserIndex).flags.Muerto Then
    WriteConsoleMsg UserIndex, " Muerto no puedes subastar ningún item.", FontTypeNames.FONTTYPE_GUILD
    Exit Sub
End If

'UserIndex es newbie? entonces no puede.

'TODO: Esto está un poco mal, por los pjs mochilas level1, pero bué - Modificar
If EsNewbie(UserIndex) Then
    WriteConsoleMsg UserIndex, " Eres newbie, no puedes..", FontTypeNames.FONTTYPE_GUILD
    Exit Sub
End If

    ' MsgConsola - no dá.
    If UserList(UserIndex).Invent.NroItems = 0 Then Exit Sub
    
    'Le enviamos el paquete
    Call WriteDaoSendUserInventory(UserIndex)

End Sub

Public Function Subasta_UsuarioOferta(ByVal UserIndex As Integer, ByVal Oferta As Long, ByRef ErrorF As String) As Boolean
' / Author: maTih
' / Note: Function is by send new ofert

' - maTih : pase esta rutina afuncion para usar ErrorF.

'No hay ninguna subasta y el pibe quiere ofertar.
If tSubasta.ThereIs = False Then
    ErrorF = "No hay ninguna subasta."
        Subasta_UsuarioOferta = False
    Exit Function
End If

'Quiso ofertar un valor negativo? logeamos.
If Oferta < 0 Then
    ErrorF = "No puedes ofertar numeros negativos."
    LogHackAttemp UserList(UserIndex).name & " Ofertó un numero negativo :|"
        Subasta_UsuarioOferta = False
    Exit Function
End If

'No puede ofertarse su propia subasta.... creo
If UserIndex = tSubasta.MiUserIndex Then
    ErrorF = "No puedes ofertar tu propia oferta."
        Subasta_UsuarioOferta = False
    Exit Function
End If


'No puede ofertar si no tiene la cantidad que ofreció..
If UserList(UserIndex).Stats.GLD < Oferta Then
    ErrorF = " No tienes " & Oferta & " monedas de oro."
        Subasta_UsuarioOferta = False
    Exit Function
End If


'Quiere ofertar y ya está ganando ? es una joda.
If UserIndex = tSubasta.LastUserIndex Then
    ErrorF = "Ya vas ganando la subasta!!."
        Subasta_UsuarioOferta = False
    Exit Function
End If

'Quiere ofertar, pero su oferta es < MENOR a la ultima. O bien, es menor a la inicial
If Oferta < tSubasta.LastSale Then
    ErrorF = "Tu humilde oferta no es superior a la anterior."
        Subasta_UsuarioOferta = False
    Exit Function
End If

'llegamos hasta acá, el usuario puede ofertar y la funcion se da true
Subasta_UsuarioOferta = True

    'Le devolvemos la moneda al anterior, SOLO si no es igual a miuserindex
    If tSubasta.MiUserIndex <> tSubasta.LastUserIndex Then
        UserList(tSubasta.LastUserIndex).Stats.GLD = UserList(tSubasta.LastUserIndex).Stats.GLD + tSubasta.LastSale
        WriteUpdateGold tSubasta.LastUserIndex
        WriteConsoleMsg tSubasta.LastUserIndex, UserList(UserIndex).name & " ofertó : " & Oferta & " se te devolvió tu oferta anterior.", FontTypeNames.FONTTYPE_GUILD
    End If

    'Listo, ahora limpiamos el anterior LastUserIndex y lo llenamos con el mio.
    tSubasta.LastUserIndex = UserIndex
    
    'Updateamos LastSale, con esta ultima oferta.
    tSubasta.LastSale = Oferta
    tSubasta.MiLastName = UserList(UserIndex).name
    
    'FIXED.
    'Le restamos el oro y le updateamos el cliente.
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Oferta
    WriteUpdateGold UserIndex

End Function


Public Function Subasta_UserPuede(ByVal UserIndex As Integer, ByVal tmpSlot As Byte, ByVal tmpAmount As Integer, ByVal minime As Long, ByRef ErrorS As String) As Boolean

' / Author: maTih
' / Note: Checks by subast.

If tSubasta.ThereIs Then
    Subasta_UserPuede = False
    ErrorS = "Ya hay una subasta, espere para iniciar otra."
    Exit Function
End If

'Si no es un slot válido, entonces no puede.
If tmpSlot < 0 Or tmpSlot > UserList(UserIndex).Invent.NroItems Then
    ErrorS = " Ha ocurrido un error, si crees que estás leyendo esto por error avisale a un administrador."
    Subasta_UserPuede = False
    Exit Function
End If

'Quitado hasta hacer la funcion
'if subasta_itemessubastable(objdata(userlist(userindex).invent.object(tmpslot).objindex) then
'   ErrorS = " No puedes subastar este item."
'   Subasta_UserPuede = False
'   Exit Function
'End If

'Esto se check via cliente, pero pueden pegar con ctrl+ v
If Amount > 10000 Then
    ErrorS = " No puedes subastar mas de 10000 objetos."
    'LogCriticEvent UserList(UserIndex).name & " Intentó hackear el sistema de subasta."
        Subasta_UserPuede = False
    Exit Function
End If

'Si llegó hasta acá, entonces puede
Subasta_UserPuede = True

End Function

'Sistema de retos 1VS1/2VS2


Public Function Retos_ObtenerSlot(ByVal RetoTipe As Byte) As Byte

' / Author: maTih
' / Note: Esta funcion obtiene un slot libre, sea 1vs1 o 2vs2. Se le pasa como parametro si es 1v1/2v2.

Dim loopC   As Long

Select Case RetoTipe

    Case 1  '1 es para 1vs1

        If tRetos.Pointer1vs1 < 10 Then
            Retos_ObtenerSlot = tRetos.Pointer1vs1 + 1
            tRetos.Pointer1vs1 = tRetos.Pointer1vs1 + 1
            Exit Function
        End If

        'Si llegamos acá, es por qe el puntero esta en 10, buscamos un slot vacio
        
        For loopC = 1 To 10
            If tRetos.Retos1v1(loopC).Hay = False Then
                    Retos_ObtenerSlot = loopC
                Exit For
                Exit Function
            End If
        Next loopC

        'Si llegamos acá, osea no se cerró la funcion antes, ponemos
        'retos_obtenerslot en 15, se usará para tirar el mensaje
        '"Todas las salas de retos estan ocupadas"

        Retos_ObtenerSlot = 15

    Case 2  'Es para 2vs2

        If tRetos.Pointer2vs2 < 10 Then
            Retos_ObtenerSlot = tRetos.Pointer2vs2 + 1
            tRetos.Pointer2vs2 = tRetos.Pointer2vs2 + 1
            Exit Function
        End If

        For loopC = 1 To 10
            If tRetos.Retos2vs2(loopC).Hay = False Then
                    Retos_ObtenerSlot = i
                Exit For
                Exit Function
            End If
        Next loopC

        Retos_ObtenerSlot = 15

End Select

End Function

'En esta funcion se va a hacer los checks para poder
'Verificar si puede iniciarse un reto.

Public Function Retos_PuedeIniciar(ByVal RetoMode As Byte, ByVal UserIndex As Integer, ByVal Oponente As String, ByVal Oponente2 As String, ByVal compañero As String, ByVal retogold As Long, ByVal isdrop As Byte, ByRef tmpError As String) As Boolean

' / Author: maTih
' / Note: En esta funcion se va a hacer los checks para poder. Verificar si puede iniciarse un reto.

Select Case RetoMode

    Case 1   'Ssado para 1vs1

    Dim tmpOponente As Integer  'Puntero del oponente
    
        tmpOponente = NameIndex(Oponente)
    
        If tmpOponente <= 0 Then
            tmpError = "Usuario offline"
                Retos_PuedeIniciar = False
            Exit Function
        End If
    
        With UserList(UserIndex)
        
            If .flags.Muerto > 0 Then
                tmpError = "Estás muerto.."
                Exit Function
            End If
        
            If .Counters.Pena > 0 Then
                tmpError = "Estás preso.."
                Exit Function
            End If
            
            If .Reto.ThereIsReto Then
                tmpError = "Ya estás en un reto.."
                Exit Function
            End If
            
        End With

        With UserList(tmpOponente)
        
            If .flags.Muerto > 0 Then
                tmpError = "Está muerto.."
                Exit Function
            End If
            
            If .Counters.Pena > 0 Then
                tmpError = "Está preso.."
                Exit Function
            End If
            
            If .Reto.ThereIsReto Then
                tmpError = "Ya está en un reto.."
                Exit Function
            End If

        End With

        Retos_PuedeIniciar = True

    Case 2 'Para 2vs2

    'Punteros para el array de userlist()
    Dim Oponentetmp As Integer
    Dim Oponentes2 As Integer
    Dim Pareja As Integer

    'Obtenemos sus indices
        Oponentetmp = NameIndex(Oponente)
        Oponentes2 = NameIndex(Oponente2)
        Pareja = NameIndex(compañero)

        If tmpOponente <= 0 Then
            tmpError = "Usuario offline"
            Retos_PuedeIniciar = False
            Exit Function
        End If

        With UserList(UserIndex)
        
            If .flags.Muerto > 0 Then
                tmpError = "Estás muerto.."
                Exit Function
            End If

            If .Counters.Pena > 0 Then
                tmpError = "Estás preso.."
                Exit Function
            End If
            
            If .Reto.ThereIsReto Then
                tmpError = "Ya estás en un reto.."
                Exit Function
            End If

        End With

        With UserList(Oponentetmp)
        
            If .flags.Muerto > 0 Then
                tmpError = .name & " Está muerto.."
                Exit Function
            End If
            
            If .Counters.Pena > 0 Then
                tmpError = .name & " Está preso.."
                Exit Function
            End If
            
            If .Reto.ThereIsReto Then
                tmpError = .name & " Ya está en un reto.."
                Exit Function
            End If
        End With

        With UserList(Oponente2)
        
            If .flags.Muerto > 0 Then
                tmpError = .name & " Está muerto.."
                Exit Function
            End If
            
            If .Counters.Pena > 0 Then
                tmpError = .name & " Está preso.."
                Exit Function
            End If
            
            If .Reto.ThereIsReto Then
                tmpError = .name & " Ya está en un reto.."
                Exit Function
            End If
            
        End With

        With UserList(Pareja)
        
            If .flags.Muerto > 0 Then
                tmpError = .name & " Está muerto.."
                Exit Function
            End If
            
            If .Counters.Pena > 0 Then
                tmpError = .name & " Está preso.."
                Exit Function
            End If

            If .Reto.ThereIsReto Then
                tmpError = .name & " Ya está en un reto.."
                Exit Function
            End If

        End With

    Retos_PuedeIniciar = True

End Select

End Function


Public Sub Retos_Arranca2v2(ByVal UserIndex As Integer, ByVal Compa As Integer, ByVal Opon1 As Integer, ByVal Opon2 As Integer, ByVal Amount As Long, ByVal itemDrop As Byte)

' / Author: maTih
' / Note: Esta rutina da inicio a un reto..

Dim tmpSlot     As Byte

tmpSlot = Retos_ObtenerSlot(2)

    With UserList(UserIndex)
        .Reto.ThereIsReto = True
        .Reto.IndexReto = tmpSlot
    End With
    
    With UserList(Compa)
        .Reto.ThereIsReto = True
        .Reto.IndexReto = tmpSlot
    End With
    
    With UserList(Opon1)
        .Reto.ThereIsReto = True
        .Reto.IndexReto = tmpSlot
    End With
    
    With UserList(Opon2)
        .Reto.ThereIsReto = True
        .Reto.IndexReto = tmpSlot
    End With
    
    tRetos.Retos2vs2(tmpSlot).Team1(1) = UserIndex
    tRetos.Retos2vs2(tmpSlot).Team1(2) = Compa
    tRetos.Retos2vs2(tmpSlot).Team2(1) = Opon1
    tRetos.Retos2vs2(tmpSlot).Team2(2) = Opon2
    tRetos.Retos2vs2(tmpSlot).Hay = True
    tRetos.Retos2vs2(tmpSlot).CountDown = 10
    tRetos.Retos2vs2(tmpSlot).Oro = Amount
    tRetos.Retos2vs2(tmpSlot).Drop = itemDrop

    Dim MiName As String, MiCompa As String, OponName As String, Opon2Name As String

    MiName = UserList(UserIndex).name
    MiCompa = UserList(Compa).name
    OponName = UserList(Opon1).name
    Opon2Name = UserList(Opon2).name

    'TODO : no usar senddata tomap, por qe se usa 1 solo mapa.
    WritePauseToggle UserIndex
    WritePauseToggle Compa
    WritePauseToggle Opon1
    WritePauseToggle Opon2

    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos 2vs2 > " & MiName & " & " & MiCompa & " versus " & OponName & " & " & Opon2Name & " ha iniciado", FontTypeNames.FONTTYPE_GUILD)

End Sub

Public Sub Retos_MuereUsuario(ByVal RetoMode As Byte, ByVal Muerto As Integer, ByVal Atacante As Integer)
With UserList(Muerto)

' / Author: maTih
' / Note: Esta rutina controla la muerte de usuarios, controla rounds & demases.

Select Case RetoMode

    Case 1  'Es 1v1?

        If UserList(Atacante).Reto.RoundsGanados <= 1 Then ' si todavia no tiene 2 ganados

            RevivirUsuario Muerto
            WritePauseToggle Muerto
            WritePauseToggle Atacante
            
            tRetos.Retos1v1(.Reto.IndexReto).CountDown = 5

        Else   'Si es mayor/igual a 2 , entonces ya gano el duelo

            Retos_Gana 1, Muerto, Atacante, tRetos.Retos1v1(.Reto.IndexReto).Gold, tRetos.Retos1v1(.Reto.IndexReto).Drop

        End If

    Case 2  'es 2v2, entonces check aca.

    Dim Pointer As Byte
    Dim i       As Long
    
        Pointer = .Reto.IndexReto

        If .Reto.TeamReto = 1 Then
        
        'Checkeamos si es del team1.
        'Si la pareja está muerta..
        
            If tRetos.Retos2vs2(Pointer).Rounds2 <= 1 And UserList(.Reto.Pareja).flags.Muerto = 1 Then
    
                RevivirUsuario Muerto
                RevivirUsuario .Reto.Pareja
    
                'Aca hay qe warpear usuarios, pausado hasta mapear toodo
                'Metemos la cuenta regresiva
    
                tRetos.Retos2vs2(Pointer).CountDown = 10
                
                tRetos.Retos2vs2(Pointer).Rounds2 = tRetos.Retos2vs2(Pointer).Rounds2 + 1
            
            ElseIf tRetos.Retos2vs2(Pointer).Rounds2 >= 2 And UserList(.Reto.Pareja).flags.Muerto = 1 Then
            ' Ganaron
                Retos_Gana 2, Muerto, Atacante, tRetos.Retos2vs2(Pointer).Oro, tRetos.Retos2vs2(Pointer).Drop, .Reto.Pareja, UserList(Atacante).Reto.Pareja
            
            End If

        Else  'si teamreto no es 1, entonces es 2.

        'Si la pareja está muerta..
            If tRetos.Retos2vs2(Pointer).Rounds1 <= 1 And UserList(.Reto.Pareja).flags.Muerto = 1 Then
            
                RevivirUsuario Muerto
                RevivirUsuario .Reto.Pareja
                
                tRetos.Retos2vs2(Pointer).CountDown = 10
                
                tRetos.Retos2vs2(Pointer).Rounds1 = tRetos.Retos2vs2(Pointer).Rounds1 + 1

            ElseIf tRetos.Retos2vs2(Pointer).Rounds1 >= 2 And UserList(.Reto.Pareja).flags.Muerto = 1 Then
            'Ganaron
                Retos_Gana 2, Muerto, Atacante, tRetos.Retos2vs2(Pointer).Oro, tRetos.Retos2vs2(Pointer).Drop, .Reto.Pareja, UserList(Atacante).Reto.Pareja

            End If
            
        End If
        
    End Select
    
End With

End Sub

Public Sub Retos_Gana(ByVal RetoMode As Byte, ByVal Muerto As Integer, ByVal Atacker As Integer, ByVal Gold As Long, ByVal Item As Byte, Optional ByVal ParejaMuerto As Integer, Optional ByVal ParejaAtack As Integer)

' / Note: Pendiente de terminar

Select Case RetoMode



End Select

End Sub

Public Function Soportes_ObtenerSlot() As Integer

' / Author: maTih
' / Note: Sistema de soporte Reescrito

If tSoportes.LastMensaje < MAX_SOPORTES Then
    tSoportes.LastMensaje = tSoportes.LastMensaje + 1
    Soportes_ObtenerSlot = tSoportes.LastMensaje
    Exit Function
End If

'llegamos acá ? list soportes es 300....
Dim loopC As Long

For loopC = 1 To 300
    If tSoportes.Usuarios(i) = "NINGUNSOPORTE" Then
            Soportes_ObtenerSlot = loopC
        Exit For
        Exit Function
    End If
Next loopC

Soportes_ObtenerSlot = 302

End Function

Public Sub Soportes_Send(ByVal name As String, ByVal Soport As String)

' / Author: maTih

Dim tmpSlot As Byte

    tmpSlot = Soportes_ObtenerSlot

    If tmpSlot = 302 Then Exit Sub
    
    ReDim Preserve tSoportes.Mensajes(1 To tmpSlot) As String
    ReDim Preserve tSoportes.Usuarios(1 To tmpSlot) As String
    
    tSoportes.Mensajes(tmpSlot) = Soport
    tSoportes.Usuarios(tmpSlot) = name
    
    WriteConsoleMsg NameIndex(name), "Su soporte ha sido enviado, será respondido en brevedad.", FontTypeNames.FONTTYPE_GUILD

End Sub

Public Function Soportes_Puede(ByVal UserIndex As Integer, ByRef SiErr As String) As Boolean

' / Author: maTih

With UserList(UserIndex)

    If .Soportes.Esperando Then
        SiErr = " Ya has enviado un soporte, espere a que sea respondido.."
            Soportes_Puede = False
        Exit Function
    End If
    
    If .Soportes.Respuesta <> vbNullString Then
        SiErr = " Un administrador ha respondido tu soporte, deberás leerlo antes de poder enviar otro.."
            Soportes_Puede = False
        Exit Function
    End If
    
    If tSoportes.LastMensaje >= MAX_SOPORTES Then
        SiErr = " Un error ha ocurrido, puede que los slots estén llenos, notifica a un administrador por favor."
            Soportes_Puede = False
        Exit Function
    End If

    Soportes_Puede = True

End With

End Function

Public Sub Soportes_GMRead(ByVal Slot As Byte, ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)

    If Slot < 0 Or Slot > tSoportes.LastMensaje Then Exit Sub
    
    WriteConsoleMsg UserIndex, tSoportes.Usuarios(Slot) & " envió > " & tSoportes.Mensajes(Slot), FontTypeNames.FONTTYPE_GUILD
    
    tSoportes.Usuarios(Slot) = "NINGUNSOPORTE"
    tSoportes.Mensajes(Slot) = "NINGUNSOPORTE"

End With

End Sub

Public Sub Soportes_USERRead(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)
    
    If .Soportes.Respuesta <> vbNullString Then
        WriteConsoleMsg UserIndex, .Soportes.Respuesta, FontTypeNames.FONTTYPE_GUILD
        .Soportes.Respuesta = vbNullString
    End If

End With

End Sub

' - Partys
Public Sub NewParty_ChangeExp(ByVal Slot1 As Byte, ByVal slot2 As Byte, ByVal NewExp1 As Byte, ByVal NewExp2 As Byte, ByVal UserChange As Integer)

' / Author: maTih
' / Note: Partys 90/10

With UserList(UserChange)

    If Slot1 < 0 Or slot2 < 0 Then Exit Sub
    
    If (NewExp1 + NewExp2) > 100 Then
        WriteConsoleMsg UserChange, "No puedes asignar mas de 100 puntos.", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If
    
    If NameIndex(.Partys.PartyArray(Slot1)) Or NameIndex(.Partys.PartyArray(slot2)) <= 0 Then
        WriteConsoleMsg UserChange, "Usuario de la party offline ;s", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If

    WriteConsoleMsg NameIndex(.Partys.PartyArray(slot2)), .name & " Ha cambiado la experiencia de la party!! " & NewExp1 & " para " & .name & " y , " & NewExp2 & " para ti.", FontTypeNames.FONTTYPE_GUILD

    .Partys.ActualExpUno = NewExp1
    .Partys.ActualExpDos = NewExp2

End With

End Sub

Public Sub NewParty_Create(ByVal UserIndex As Integer, ByVal targetUI As Integer, ByVal ExpInicial1 As Byte, ByVal ExpInicial2 As Byte)

' / Author: maTih

With UserList(UserIndex)

    .Partys.ActualExpUno = ExpInicial1
    .Partys.ActualExpUno = ExpInicial2
    .Partys.PartyArray(1) = .name
    .Partys.PartyArray(2) = UserList(targetUI).name
    
    WriteConsoleMsg targetUI, .name & " ha iniciado una party contigo!! , Experiencia para " & .name & " : " & ExpInicial1 & " , Experiencia para ti : " & ExpInicial2, FontTypeNames.FONTTYPE_GUILD
    
    UserList(targetUI).Partys.ActualExpUno = ExpInicial1
    UserList(targetUI).Partys.ActualExpUno = ExpInicial2
    UserList(targetUI).Partys.PartyArray(1) = .name
    UserList(targetUI).Partys.PartyArray(2) = UserList(targetUI).name

End With

End Sub

Public Sub NewExp_CierraUsuario(ByVal UsuarioClose As Integer)

' / Author: maTih

With UserList(UsuarioClose)

    .Partys.ActualExpDos = 0
    .Partys.ActualExpUno = 0

    If .Partys.PartyArray(1) = .name Then
    
        WriteConsoleMsg NameIndex(.Partys.PartyArray(2)), .name & " Ha deslogeado, se cierra la party", FontTypeNames.FONTTYPE_GUILD
        UserList(NameIndex(.Partys.PartyArray(2))).Partys.ActualExpDos = 0
        UserList(NameIndex(.Partys.PartyArray(2))).Partys.ActualExpUno = 0
        .Partys.PartyArray(1) = vbNullString
        .Partys.PartyArray(2) = vbNullString
        .Partys.ActualExpDos = 0
        .Partys.ActualExpUno = 0
        
    Else
    
        WriteConsoleMsg NameIndex(.Partys.PartyArray(1)), .name & " Ha deslogeado, se cierra la party", FontTypeNames.FONTTYPE_GUILD
        UserList(NameIndex(.Partys.PartyArray(1))).Partys.ActualExpDos = 0
        UserList(NameIndex(.Partys.PartyArray(1))).Partys.ActualExpUno = 0
        .Partys.PartyArray(1) = vbNullString
        .Partys.PartyArray(2) = vbNullString
        
    End If

End With

End Sub

' - Sistema de ScreenShots, tanto como /Foto nick, como Foro-Denuncias.
Public Sub Screens_Request(ByVal targetName As String, ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)

    If NameIndex(targetName) > 0 Then
    
    ' / ...

    Else

        WriteConsoleMsg UserIndex, "Usuario offline..", FontTypeNames.FONTTYPE_GUILD

    End If

End With

End Sub

Public Sub Screens_Denounce(ByVal UserIndex As Integer)

' / Author: maTih

' / ...

End Sub

' - Sistema de torneos.
Public Function Torneo_PuedeCrear(ByVal Cupos As Byte, ByVal Tipe As Byte, ByVal Precio As Long, ByRef Error As String) As Boolean

' / Author: maTih
' / Note: Funcion para checkear si un usuario puede crear torneo.

    If Cupos > NumUsers Then
        Error = " No hay tantos usuarios conectados."
            Torneo_PuedeCrear = False
        Exit Function
    End If
    
    If Tipe > 2 Or Tipe < 0 Then
        Error = " Tipo de torneo inválido...."
            Torneo_PuedeCrear = False
        Exit Function
    End If
    
    If Precio < 0 Or Precio > 150000 Then
        Error = " Ingresaste un valor de inscripción inválido , o mayor a 150k."
            Torneo_PuedeCrear = False
        Exit Function
    End If
    
    If tTorneos.Hay Then
        Error = " Hay un torneo en marcha, cierrelo é inicie otro."
            Torneo_PuedeCrear = False
        Exit Function
    End If

Torneo_PuedeCrear = True

End Function

Public Sub Torneo_Crear(ByVal UI As Integer, ByVal Cupos As Byte, ByVal tipos As Byte, ByVal Precio As Long)

' / Author: maTih
' / Note: Rutina para crear un torneo.

Dim Tipo(1 To 2)    As String
Dim Mensaje         As String

    With UserList(UI)
    
        Tipo(1) = "Torneo 1vs1"
        Tipo(2) = "DeathMatch"
        
        Mensaje = "Torneo> Cupos: " & Cupos & " de tipo " & Tipo(tipos)
        
        If Precio > 0 Then
            Mensaje = Mensaje & " Precio de inscripción " & Precio
        End If
        
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "Creó un " & Mensaje, FontTypeNames.FONTTYPE_GUILD)
        
        tTorneos.Hay = True
        tTorneos.Cupos = Cupos
        tTorneos.CountCupos = 0
        tTorneos.precioInsc = Precio
        
        ReDim tTorneos.UserIndexs(1 To Cupos) As Integer
    
    End With

End Sub

Public Function Torneo_UsuarioEntra(ByVal User As Integer, ByRef tmpM As String) As Boolean

' / Author: maTih
' / Note: Rutina por la cual un usuario ingresa a un torneo
'         Pasado a funcion

With UserList(User)

    'Si está muerto no...
    If .flags.Muerto > 0 Then
        tmpM = "Estás muerto!!!"
            Torneo_UsuarioEntra = False
        Exit Function
    End If
    
    If .Counters.Pena > 0 Then
        tmpM = "Estás preso!!!"
            Torneo_UsuarioEntra = False
        Exit Function
    End If
    
    If .Reto.ThereIsReto Then
        tmpM = "Estás en un reto!!!"
            Torneo_UsuarioEntra = False
        Exit Function
    End If
    
    If tTorneos.Hay = False Then
        tmpM = "No hay ningún torneo actualmente!!!"
            Torneo_UsuarioEntra = False
        Exit Function
    End If
    
    If tTorneos.precioInsc > .Stats.GLD Then
        tmpM = " El precio de inscripción es " & tTorneos.precioInsc & "!!!"
            Torneo_UsuarioEntra = False
        Exit Function
    End If
    
    If tTorneos.CountCupos >= tTorneos.Cupos Then
        tmpM = "Cupos alcanzados.!!!"
            Torneo_UsuarioEntra = False
        Exit Function
    End If
    
    Torneo_UsuarioEntra = True
    
    tTorneos.UserIndexs(tTorneos.CountCupos + 1) = User
    tTorneos.CountCupos = tTorneos.CountCupos + 1
    tTorneos.PozoRecaudado = tTorneos.PozoRecaudado + tTorneos.precioInsc
    .Stats.GLD = .Stats.GLD - tTorneos.precioInsc
    
    tmpM = "Has ingresado al torneo."
    
    If tTorneos.CountCupos >= tTorneos.Cupos Then
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo > CUPOS LLENOS!.", FontTypeNames.FONTTYPE_GUILD)
        tTorneos.Hay = False
    End If

End With

End Function

' - Ranking, al pedo como cenicero de moto.
Public Function Ranking_Puede(ByVal UserIndex As Integer, ByRef posiTion As Byte) As Boolean

' / Author: maTih
' / Note: Checks for same user entring to ranking

With UserList(UserIndex)

    Dim tmpNivel    As Byte
    Dim tmpFrags    As Integer
    Dim loopC       As Long
    
        tmpNivel = .Stats.ELV
        tmpFrags = .Stats.UsuariosMatados
        
        If Ranking_Esta(UCase$(.name)) Then Exit Function
        
        'Este bucle va a recorrer los 10 usuarios en el ranking y sus niveles
        For loopC = 1 To 10
            
            If tmpNivel > tRanking.UserLevels(loopC) Then
                posiTion = loopC
                Exit For
                    Ranking_Puede = True
                Exit Function
            End If
        
        Next loopC
    
        Ranking_Puede = False

End With

End Function

Public Function Ranking_Esta(ByVal Namee As String) As Boolean

' / Author: maTih
' / Note: Esta funcion devuelve True ó False segun si está en el ranking o NO.

Dim loopC As Long

    For loopC = 1 To 10
        If tRanking.UserNames(loopC) = Namee Then
                Ranking_Esta = True
            Exit For
            Exit Function
        End If
    Next loopC

    Ranking_Esta = False
    
End Function

Public Sub Ranking_Cargar()

' / Author : maTih
' / Note : Esta rutina carga los ranking.

RankPathFile = App.Path & "\Ranking.ini"

Dim loopC   As Long

For loopC = 1 To 10

    tRanking.UserNames(loopC) = GetVar(RankPathFile, "RANKING" & loopC, "Nombre")
    tRanking.UserLevels(loopC) = val(GetVar(RankPathFile, "RANKING" & loopC, "Nivel"))
    tRanking.UserFrags(loopC) = val(GetVar(RankPathFile, "RANKING" & loopC, "Frags"))
    tRanking.UserTag(loopC) = GetVar(RankPathFile, "RANKING" & loopC, "Clan")

Next loopC

End Sub

Public Sub Ranking_AgregarUsuario(ByVal UserIndex As Integer, ByVal Posicion As Byte)

' / Author: maTih
' / Note: Esta rutina agrega un usuario al ranking, y hace un intercambio de posiciones si es necesario.

'Variables de uso temporal para guardar la posicion , y para hacer el intercambio
Dim NamePos     As String
Dim tagPos      As String
Dim FragPos     As Integer
Dim NivelPos    As Byte
Dim ClasePos    As String

With UserList(UserIndex)

    'si la posicion es 10, entonces no hay qe hacer ningun intercambio.
    If Posicion < 10 Then
    
        NamePos = GetVar(RankPathFile, "RANKING" & Posicion, "Nombre")
        tagPos = GetVar(RankPathFile, "RANKING" & Posicion, "Clan")
        FragPos = val(GetVar(RankPathFile, "RANKING" & Posicion, "Frags"))
        NivelPos = val(GetVar(RankPathFile, "RANKING" & Posicion, "Nivel"))
        ClasePos = GetVar(RankPathFile, "RANKING" & Posicion, "Clase")
        
        'Ya tenemos los datos, intercambiamos.
        WriteVar RankPathFile, "RANKING" & Posicion, "Nombre", .name
        WriteVar RankPathFile, "RANKING" & Posicion, "Clase", ListaClases(.clase)
        WriteVar RankPathFile, "RANKING" & Posicion, "Clan", modGuilds.GuildName(.GuildIndex)
        WriteVar RankPathFile, "RANKING" & Posicion, "Nivel", .Stats.ELV
        WriteVar RankPathFile, "RANKING" & Posicion, "Frags", .Stats.UsuariosMatados
        
        tRanking.UserFrags(Posicion) = .Stats.UsuariosMatados
        tRanking.UserLevels(Posicion) = .Stats.ELV
        tRanking.UserNames(Posicion) = .name
        tRanking.UserTag(Posicion) = modGuilds.GuildName(.GuildIndex)
        tRanking.UserClases(Posicion) = ListaClases(.clase)
        
        'Ahora escribimos los datos del usuario qe bajo 1 posicion
        WriteVar RankPathFile, "RANKING" & Posicion + 1, "Nombre", NamePos
        WriteVar RankPathFile, "RANKING" & Posicion + 1, "Frags", FragPos
        WriteVar RankPathFile, "RANKING" & Posicion + 1, "Nivel", NivelPos
        WriteVar RankPathFile, "RANKING" & Posicion + 1, "Clan", tagPos
        WriteVar RankPathFile, "RANKING" & Posicion + 1, "Clase", ClasePos
        
        tRanking.UserFrags(Posicion + 1) = FragPos
        tRanking.UserLevels(Posicion + 1) = NivelPos
        tRanking.UserNames(Posicion + 1) = NamePos
        tRanking.UserTag(Posicion + 1) = tagPos
        tRanking.UserClases(Posicion + 1) = ClasePos

    Else
        
        WriteVar RankPathFile, "RANKING" & 10, "Nombre", .name
        WriteVar RankPathFile, "RANKING" & 10, "Clan", modGuilds.GuildName(.GuildIndex)
        WriteVar RankPathFile, "RANKING" & 10, "Clase", ListaClases(.clase)
        WriteVar RankPathFile, "RANKING" & 10, "Nivel", .Stats.ELV
        WriteVar RankPathFile, "RANKING" & 10, "Frags", .Stats.UsuariosMatados
        
        tRanking.UserFrags(10) = .Stats.UsuariosMatados
        tRanking.UserLevels(10) = .Stats.ELV
        tRanking.UserNames(10) = .name
        tRanking.UserTag(10) = modGuilds.GuildName(.GuildIndex)
        tRanking.UserClases(10) = ListaClases(.clase)
        
    End If
    
End With

End Sub

'SISTEMA DE GUILD VS GUILD.

Public Sub WarGuild_Start(ByVal guildUno As Integer, ByVal GuildDos As Integer, ByVal Contador As Byte)
'Author : maTih
' Note  : War Guild vs Guild start by GuildsIndexs.
With tCvC

.GuildIndex1 = guildUno
.GuildIndex2 = GuildDos
.Caspers1 = 0
.Caspers2 = 0
.Ocuped = True
ReDim .UserIndexs1(1 To Contador) As Integer
ReDim .UserIndexs2(1 To Contador) As Integer

SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Guild vs Guild > " & clan1 & " vs " & clan2, FontTypeNames.FONTTYPE_GUILD)


End With

End Sub

Public Sub WarGuild_UserAccept(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)

    If .flags.Muerto = 1 Then
        WriteConsoleMsg UserIndex, "Estás muerto.", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If

    If .Counters.Pena > 0 Then
        WriteConsoleMsg UserIndex, "Estás en la carcel -.- .", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If
    
    If .TieneCvc = 1 Then
        WriteConsoleMsg UserIndex, "Ya estás en clan vs Clan!!.", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If
    
    If .GuildIndex <= 0 Then
        WriteConsoleMsg UserIndex, "No tienes clan!!.", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If
    
    If tCvC.GuildIndex1 = .GuildIndex Then
    
        If tCvC.Count1 >= tCvC.Contador Then
            WriteConsoleMsg UserIndex, "No puedes ingresar, ya ha iniciado!!.", FontTypeNames.FONTTYPE_GUILD
            Exit Sub
        End If
        
        tCvC.Count1 = tCvC.Count1 + 1
        
        If tCvC.Count1 >= tCvC.Contador Then
            SendData SendTarget.ToDiosesYclan, tCvC.GuildIndex2, PrepareMessageConsoleMsg("El clan " & modGuilds.GuildName(tCvC.GuildIndex1) & " ha aceptado .", FontTypeNames.FONTTYPE_GUILD)
        End If
        
    ElseIf tCvC.GuildIndex2 = .GuildIndex Then
        
        If tCvC.Count2 >= tCvC.Contador Then
            WriteConsoleMsg UserIndex, "No puedes ingresar, ya ha iniciado!!.", FontTypeNames.FONTTYPE_GUILD
            Exit Sub
        End If
        
        tCvC.Count2 = tCvC.Count2 + 1
        
        If tCvC.Count2 >= tCvC.Contador Then
            SendData SendTarget.ToDiosesYclan, tCvC.GuildIndex1, PrepareMessageConsoleMsg("El clan " & modGuilds.GuildName(tCvC.GuildIndex2) & " ha aceptado .", FontTypeNames.FONTTYPE_GUILD)
        End If
    
    End If

End With

End Sub

'Sistema de durabilidad en items
Public Sub CheckDurability(ByVal UserIndex As Integer, ByVal VictimIndex As Integer, ByVal Caso As Byte, ByVal Quien As Byte)

' / Author: maTih
' / Note: Checks of durabiltyItems by slot and Pointers in array Userlist()

Select Case Caso

    Case 1           'Armour
    
        With UserList(VictimIndex)
        
            .Invent.Object(.Invent.ArmourEqpSlot).Durabilidad = .Invent.Object(.Invent.ArmourEqpSlot).Durabilidad - 1
            
            If .Invent.Object(.Invent.ArmourEqpSlot).Durabilidad <= 0 Then
            'Desequipamos su armoreqpslot y se lo quitamos
                Call Desequipar(VictimIndex, .Invent.ArmourEqpSlot)
                WriteConsoleMsg VictimIndex, "Tu armadura se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
            End If
            
        End With

        If Quien = 1 Then ' le pegó un usuario ?

            With UserList(UserIndex)
            
            .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad - 1
            
                If .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad <= 0 Then
                
                'Desequipamos su arma y se la quitamos
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    WriteConsoleMsg UserIndex, "Tu arma se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
                End If
                
            End With
            
        End If

    Case 2            'Shield
    
        With UserList(VictimIndex)
        
            .Invent.Object(.Invent.EscudoEqpSlot).Durabilidad = .Invent.Object(.Invent.EscudoEqpSlot).Durabilidad - 1
            
            If .Invent.Object(.Invent.EscudoEqpSlot).Durabilidad <= 0 Then
                'Desequipamos su escudo y se lo quitamos
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                WriteConsoleMsg VictimIndex, "Tu Escudo se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
            End If
            
        End With
    
        If Quien = 1 Then ' le pegó un usuario ?
    
            With UserList(UserIndex)
            
            .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad - 1
                
                If .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad <= 0 Then
                'Desequipamos su arma y se la quitamos
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    WriteConsoleMsg UserIndex, "Tu arma se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
                End If
            
            End With
    
        End If

    Case 3            'Casco
    
        With UserList(VictimIndex)
        
        .Invent.Object(.Invent.CascoEqpSlot).Durabilidad = .Invent.Object(.Invent.CascoEqpSlot).Durabilidad - 1
        
            If .Invent.Object(.Invent.CascoEqpSlot).Durabilidad <= 0 Then
            'Desequipamos su casco y se lo quitamos
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                WriteConsoleMsg VictimIndex, "Tu casco se ha auto-destruido por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
            End If
        
        End With

        If Quien = 1 Then ' le pegó un usuario ?

            With UserList(UserIndex)
            
            .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad - 1
            
                If .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad <= 0 Then
                'Desequipamos su arma y se la quitamos
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    WriteConsoleMsg UserIndex, "Tu arma se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
                End If
            
            End With

        End If

    Case 4            'Anillos !
 
        With UserList(VictimIndex)
        
        .Invent.Object(.Invent.AnilloEqpSlot).Durabilidad = .Invent.Object(.Invent.AnilloEqpSlot).Durabilidad - 1
        
            If .Invent.Object(.Invent.AnilloEqpSlot).Durabilidad <= 0 Then
            'Desequipamos su anillo y se lo quitamos
                Call Desequipar(VictimIndex, .Invent.AnilloEqpSlot)
                WriteConsoleMsg VictimIndex, "Tu anillo se ha auto-destruido por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
            End If
        
        End With

        If Quien = 1 Then ' le pegó un usuario ?

            With UserList(UserIndex)
            
            .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad - 1
            
                If .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad <= 0 Then
                'Desequipamos su aarma y se la quitamos
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    WriteConsoleMsg UserIndex, "Tu arma se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
                End If
            
            End With

        End If

    Case 5                'Armas (usado para hit contra NPC's)
 
        With UserList(UserIndex)
    
        .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad - 1
        
            If .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad <= 0 Then
            'desequipamos su aarma y se la qitamos
                Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                WriteConsoleMsg UserIndex, "Tu arma se ha auto-destruida por que su durabilidad llegó a 0!!", FontTypeNames.FONTTYPE_GUILD
            End If
        
        End With

End Select

End Sub

Public Sub Load_DurabilityItems(ByVal UserIndex As Integer)

' / Author: maTih

Dim charUser As String      'Obtenemos su nickname en la app.path charfiles

With UserList(useridex)

    charUser = CharPath & .name & ".chr"
    
    If .Invent.WeaponEqpSlot > 0 Then
        .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = val(GetVar(charUser, "Durabilidad", "Arma"))
    End If
    
    If .Invent.EscudoEqpSlot > 0 Then
        .Invent.Object(.Invent.EscudoEqpSlot).Durabilidad = val(GetVar(charUser, "Durabilidad", "Escudo"))
    End If
    
    If .Invent.AnilloEqpSlot > 0 Then
        .Invent.Object(.Invent.AnilloEqpSlot).Durabilidad = val(GetVar(charUser, "Durabilidad", "Anillo"))
    End If
    
    If .Invent.CascoEqpSlot > 0 Then
        .Invent.Object(.Invent.CascoEqpSlot).Durabilidad = val(GetVar(charUser, "Durabilidad", "Casco"))
    End If
    
    If .Invent.ArmourEqpSlot > 0 Then
        .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = val(GetVar(charUser, "Durabilidad", "Armadura"))
    End If

End With

End Sub

Public Sub Save_DurabilityItems(ByVal UserIndex As Integer)

' / Author : maTih
' / Note   : Save durability eqipSlot by userindex in charFile

Dim tmpChar As String

With UserList(UserIndex)

    tmpChar = CharPath & .name & ".chr"
    
    If .Invent.WeaponEqpSlot > 0 Then
        WriteVar tmpChar, "Durabilidad", "Arma", .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad
    End If
    
    If .Invent.EscudoEqpSlot > 0 Then
        WriteVar tmpChar, "Durabilidad", "Escudo", .Invent.Object(.Invent.EscudoEqpSlot).Durabilidad
    End If
    
    If .Invent.AnilloEqpSlot > 0 Then
        WriteVar tmpChar, "Durabilidad", "Anillo", .Invent.Object(.Invent.AnilloEqpSlot).Durabilidad
    End If
    
    If .Invent.CascoEqpSlot > 0 Then
        WriteVar tmpChar, "Durabilidad", "Casco", .Invent.Object(.Invent.CascoEqpSlot).Durabilidad
    End If
    
    If .Invent.ArmourEqpSlot > 0 Then
        WriteVar tmpChar, "Durabilidad", "Armadura", .Invent.Object(.Invent.ArmourEqpSlot).Durabilidad
    End If

End With

End Sub

Public Sub Reset_DurabilityItems(ByVal UserIndex As Integer)

' / Author : maTih
' / Note   : Reset flags inventory items durability

With UserList(UserIndex)
    
    .Invent.Object(.Invent.EscudoEqpSlot).Durabilidad = 0
    .Invent.Object(.Invent.ArmourEqpSlot).Durabilidad = 0
    .Invent.Object(.Invent.WeaponEqpSlot).Durabilidad = 0
    .Invent.Object(.Invent.CascoEqpSlot).Durabilidad = 0
    .Invent.Object(.Invent.AnilloEqpSlot).Durabilidad = 0

End With

End Sub

Public Sub Duelos_Ingreso(ByVal Ingresa As Integer)

' / Author : maTih
' / Note   : Ingresa is parameter to ingress fight!

With UserList(Ingresa)

.Stats.GLD = .Stats.GLD - 100000

WriteUpdateGold Ingresa

'one user in waiting ??

If Duelos.UserEsperando > 0 Then

Duelos.ContadorSalas = Duelos.ContadorSalas + 1
Duelos.SalasLibres = Duelos.SalasLibres - 1
ReDim Preserve Duelos.Cuenta(1 To Duelos.ContadorSalas) As Byte
ReDim Preserve Duelos.UsersDueleando1(1 To Duelos.ContadorSalas) As Integer

Duelos.UsersDueleando1(Duelos.ContadorSalas) = Ingresa
Duelos.UsersDueleando2(Duelos.ContadorSalas) = Duelos.UserEsperando

Duelos.Cuenta(Duelos.ContadorSalas) = 5

WarpUserChar Ingresa, Mapa_Duelo, X1_Duelos(Duelos.ContadorSalas), Y2_Duelos(Duelos.ContadorSalas), True, False
WarpUserChar Duelos.UserEsperando, Mapa_Duelo, X1_Duelos(Duelos.ContadorSalas), Y2_Duelos(Duelos.ContadorSalas), True, False



WritePauseToggle Ingresa
WritePauseToggle Duelos.UserEsperando

Duelos.UserEsperando = 0

Else        'no waiting user, useresperando = ingresa.

Duelos.UserEsperando = Ingresa

.Stats.GLD = .Stats.GLD - 100000

WriteUpdateGold Ingresa

WarpUserChar Ingresa, mapa_duelos, EsperaX, EsperaY, True, False
WriteConsoleMsg Ingresa, "Te has inscripto, espera a que alguien ingrese.", FontTypeNames.FONTTYPE_GUILD

End If

End With

End Sub

Public Function Duelos_PIngresar(ByVal iUser As Integer, ByRef Error As String) As Boolean

Dim tNpc As Integer


With UserList(iUser)

tNpc = .flags.targetNPC

Duelos_PIngresar = False

If .Stats.GLD < 100000 Then Error = "Debes tener 100.000 monedas": Exit Function

If .flags.Muerto = 1 Then Error = "Estás muerto!!": Exit Function

If MapInfo(.Pos.Map).Pk Then Error = "Debes estar en una ciudad!": Exit Function

If Npclist(tNpc).Numero <> DUELOS_NPC Then Error = "Debes clickear el npc de duelos.": Exit Function

If Duelos.ContadorSalas >= 5 Then Error = "Todas las salas están ocupadas": Exit Function

Duelos_PIngresar = True

Duelos_Ingreso iUser

End With

End Function

Public Sub Duelos_LoadArrays()

On Error GoTo LastError

Mapa_Duelo = 1
EsperaX = val(GetVar(DatPath & "Duelos.ini", "Espera", "EsperaX"))
EsperaY = val(GetVar(DatPath & "Duelos.ini", "Espera", "EsperaY"))

Dim i As Long

For i = 1 To 5

X1_Duelos(i) = val(GetVar(DatPath & "Duelos.ini", "Esquinas1", "EsquinaX" & i))
Y1_Duelos(i) = val(GetVar(DatPath & "Duelos.ini", "Esquinas1", "EsquinaY" & i))

x2_Duelos(i) = val(GetVar(DatPath & "Duelos.ini", "Esquinas2", "EsquinaX" & i))
Y2_Duelos(i) = val(GetVar(DatPath & "Duelos.ini", "Esquinas2", "EsquinaY" & i))

Next i

Exit Sub

LastError:

LogError "Error cargando arrays duelos."

End Sub

Public Sub Duelo_MuereU(ByVal iUser As Integer, ByVal Attacker As Integer)

Dim tmpP As Byte

With UserList(iUser)

tmpP = .DueloArray

UserList(Attacker).Stats.GLD = UserList(Attacker).Stats.GLD + 100000

WriteUpdateGold Attacker

WarpUserChar Attacker, 1, 50, 55, True, False
WarpUserChar iUser, 1, 58, 45, False, False

Duelo_ResetArray tmpP

End With


End Sub

Public Sub Duelo_ResetArray(ByVal Pointer As Byte)

With Duelos

.ContadorSalas = .ContadorSalas - 1
.SalasLibres = .SalasLibres + 1

.UsersDueleando1(Pointer) = 0
.UsersDueleando2(Pointer) = 0

End With
End Sub

Public Function CalcularUsuarios(ByVal Testeo As Boolean) As Integer

' / Author : maTih
' / Note   : Testeo es para probar, esta funcion calcula los usuarios online.

  If Testeo Then CalcularUsuarios = 1: Exit Function
  
  Dim i As Long
  Dim Contando As Integer
  'vamos uno por uno.
  For i = 1 To LastUser
   If UserList(i).flags.Privilegios = PlayerType.User Then
     Contando = Contando + 1
   End If
  Next i

  'variable final
  
  CalcularUsuarios = Contando
  
End Function

'Drag & Drop.

Public Sub DragDrop_AlPiso(ByVal Usuario As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Slot As Byte)

' / Author       : maTih
' / Note         : Drag & Drop userObjs to targetPosition
' / Modification : Agrego slot para evitar bucles innecesarios

With UserList(Usuario)

    If MapInfo(.Pos.Map).Pk = True Then
        WriteConsoleMsg Usuario, "No está permitido arrojar items en zona segura.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
     
    If MapData(.Pos.Map, X, Y).ObjInfo.Amount > 9999 Then
        WriteConsoleMsg Usuario, "No hay espacio en el piso!.", FontTypeNames.FONTTYPE_INFO
        Exit Sub
    End If
      
    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = Objeto Or MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Then
        If Cantidad + MapData(.Pos.Map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
            Cantidad = MAX_INVENTORY_OBJS - MapData(.Pos.Map, X, Y).ObjInfo.Amount
        End If
    End If
    
    Dim tObj As Obj
        tObj.Amount = Cantidad
        tObj.ObjIndex = Objeto
         
        MakeObj tObj, .Pos.Map, X, Y
        
        QuitarUserInvItem Usuario, Slot, Cantidad
        UpdateUserInv False, Usuario, Slot
    
    Exit Sub

End With

End Sub

Public Sub DragDrop_DragNpc(ByVal Usuario As Integer, ByVal Npc As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer, ByVal Slot As Byte)

With UserList(Usuario)

'Variables de uso temporal

    Dim NpcDragea         As Boolean
    Dim NpcEsBove         As Boolean
     
    NpcDragea = (Npclist(Npc).Comercia = 1)
    NpcEsBove = (Npclist(Npc).name = "Banquero")
     
    .flags.targetNPC = Npc
     
    If NpcEsBove Then
        UserDejaObj Usuario, CInt(Slot), Cantidad
        UpdateUserInv False, Usuario, Slot
        Exit Sub
    End If
    
    If NpcDragea Then
        Comercio eModoComercio.Venta, Usuario, Npc, CInt(Slot), Cantidad
        WriteUpdateGold Usuario
        Exit Sub
    End If

    Call WriteConsoleMsg(Usuario, "Ese npc no es un comerciante!!", FontTypeNames.FONTTYPE_INFO)
    
End With

Exit Sub

End Sub

Public Sub DragDrop_DragUsuario(ByVal UserSend As Integer, ByVal targetUser As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal Slot As Byte)

Dim tObj         As Obj
Dim tmpTxt       As String

tObj.ObjIndex = ObjIndex
tObj.Amount = Amount

QuitarUserInvItem UserSend, Slot, Amount

UpdateUserInv False, usersed, Slot

If Not MeterItemEnInventario(targetUser, tObj) Then
TirarItemAlPiso UserList(targetUser).Pos, tObj
End If

If Amount = 1 Then
tmpTxt = "Tu " & ObjData(ObjIndex).name
Else
tmpTxt = Amount & " - " & ObjData(ObjIndex).name
End If

WriteConsoleMsg UserSend, "Le has arrojado a " & UserList(targetUser).name & " " & tmpTxt, FontTypeNames.FONTTYPE_GUILD

If Amount = 1 Then
tmpTxt = "Su " & ObjData(ObjIndex).name
Else
tmpTxt = Amount & " - " & ObjData(ObjIndex).name
End If

WriteConsoleMsg targetUser, UserList(UserSend).name & " Te ha arrojado " & tmpTxt, FontTypeNames.FONTTYPE_GUILD

End Sub

'SISTEMA DE TARGETS

Public Function GetTargets(ByVal Target As Integer, ByVal targetFichado As Integer) As String

' \   Author  :  maTih
' \   Note    :  Calculate targets

Dim pIndex       As Integer       'Index del array de partyes
Dim gIndex       As Integer       'Index del array de clanes

'MANEJO DE PARTY

If UserList(Target).PartyIndex > 0 Then

pIndex = UserList(Target).PartyIndex

'MISMO PARTYINDEX ??

If UserList(Target).PartyIndex = UserList(targetFichado).PartyIndex Then

GetTargets = GetTargetParty(Target, targetFichado)

Else

GetTargets = GetTargetPartyWithOutIndex(Target, targetFichado)

End If

End If

'MANEJO DE CLAN

If UserList(Target).GuildIndex > 0 Then

'MISMO GUILDINDEX ?

If UserList(Target).GuildIndex = UserList(targetFichado).GuildIndex Then

GetTargets = GetTargetGuild(Target, targetFichado)

Else

GetTargets = GetTargetGuildWithOutIndex(Target, targetFichado)

End If

End If


End Function

Public Function GetTargetParty(ByVal TI As Integer, ByVal OTI As Integer) As String

' \  Author  : maTih.-
' \  Note    : Prepare targets by party index

End Function

Public Function GetTargetPartyWithOutIndex(ByVal TI As Integer, ByVal OTI As Integer) As String

' \  Author  : maTih.-
' \  Note    : Prepare targets with out party index

End Function

Public Function GetTargetGuild(ByVal TI As Integer, ByVal OTI As Integer) As String

' \  Author  : maTih.-
' \  Note    : Prepare targets by guildindex

End Function

Public Function GetTargetGuildWithOutIndex(ByVal TI As Integer, ByVal OTI As Integer) As String

' \  Author  : maTih.-
' \  Note    : Prepare targets with out guildindex

End Function
