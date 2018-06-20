Attribute VB_Name = "mod_IA"
#If ConBots Then

Option Explicit
 
'Defensa del bot jeje
Private Const IA_MINDEF  As Integer = 10
Private Const IA_MAXDEF  As Integer = 12
 
'Charindex reservado.
 Private Const IA_CHAR    As Integer = 9994 'Hardcodeado : P
 
'Datos del char
 
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
 
'ATENCION : Acá van los números de objetos!!!
 
Private Const IA_HEAD    As Integer = 4
Private Const IA_BODY    As Integer = 986
Public Const MAX_BOTS   As Byte = 25
 
'ATENCION : Acá van los números de objetos!!!
 
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
 
'Cantidad de hechizos que lanza
 
Private Const IA_M_SPELL As Byte = 3
Private Const IA_NUMCHAT As Byte = 5
 
'Constantes de intervalos.
 
Private Const IA_SINT   As Integer = 800    'Intervalo entre hechizo-hechizo.
Private Const IA_SREMO  As Integer = 500    'Intervalo remo.
Private Const IA_MOVINT As Integer = 240    'Intervalo caminta.
Private Const IA_USEOBJ As Integer = 200    'Intervalo usar potas.
Private Const IA_HITINT As Integer = 200    'Intervalo para golpe
Private Const IA_PROINT As Integer = 700    'Intervalo de flecha
Private Const IA_TALKIN As Integer = 4000   'Intervalo de hablAR :P

'Probabilidades de que te pegue
 
Private Const IA_CASTEO As Byte = 77
 
Private Const IA_PROBEV As Byte = 160
Private Const IA_PROBEX As Byte = 220

Private Const IA_SLOTS  As Byte = 20
 
Type ia_Interval
     SpellCount         As Byte         'Intervalo para tirar hechizos.
     UseItemCount       As Byte         'Intervalo para usar pociones.
     MoveCharCount      As Byte         'Intervalo para mover el char.
     ParalizisCount     As Byte         'Intervalo para removerse.
     HitCount           As Byte         'Intervalo para pegar golpesito.
     ArrowCount         As Byte         'Intervalo para flechas
     ChatCount          As Byte         'INtervalo para hablar XD
End Type
 
Type ia_Spells
     DamageMin          As Byte         'Minimo daño que hace.
     DamageMax          As Byte         'Maximo daño que hace.
     spellIndex         As Byte         'Lo usamos para el fx.
End Type

Enum eIASupportActions
     SRemover = 1                       'Remueve.
     SCurar = 2                         'Cura.
End Enum
 
Enum eIAClase
     Clerigo = 1                        'Bot Clero
     Mago = 2                           'Bot Mago
     Cazador = 3                        'Bot kza
End Enum

Enum eIAactions
     ePegar = 1                          'accion pegar.
     eMagia = 2                          'atacar con hechizo
End Enum

Enum eIAMoviments
     SeguirVictima = 1                   'Si seguia la victima
     MoverRandom = 2                     'Random moviment :P
End Enum
 
Type Bot
     EsCriminal           As Boolean
     Pos                As WorldPos     'Posicion en el mundo.
     maxVida            As Integer      'Maxima vida.
     Vida               As Integer      'Vida del bot.
     Clase              As eIAClase     'Clases de bot.
     Tag                As String       'Tag del bot.
     Mana               As Integer      'Mana del bot.
     maxMana            As Integer      'Maxima mana
     Char               As Char         'Apariencia.
     Invocado           As Boolean      'Si existe.
     Paralizado         As Boolean      'Si está inmo.
     Intervalos         As ia_Interval  'Intervalos de acciones.
     Viajante           As Boolean      'Bot Viajante :P
     ViajanteUser       As Integer      'Usuario que atacó al viajante.
     UltimaAccion       As eIAactions   'ULTIMA ACCION/ATAQUE.
     UltimoMovimiento   As eIAMoviments 'ULTIMO MOVIMIENTO
     Navegando          As Boolean      'Navegando?
     ViajanteAntes      As WorldPos     'Pos cuando un viajante ataca un usuario.
     Inv(1 To IA_SLOTS) As Obj          'Inventario del bot.
     UltimaIdaObjeto    As Boolean      'Ultimo movimiento fue buscar objs?
End Type
 
Public ia_Bot(1 To MAX_BOTS)           As Bot
Public ia_spell(1 To IA_M_SPELL)       As ia_Spells
Public ia_Chats(1 To IA_NUMCHAT)      As String

'Cantidad de bots invocados.
Public NumInvocados                    As Byte

Function ia_CascoByClase(ByVal botIndex As Byte) As Integer

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve el casco/gorro según la clase del bot

Select Case ia_Bot(botIndex).Clase

       Case eIAClase.Clerigo        'Bot clero
            ia_CascoByClase = 131   'Completo : P
        
       Case eIAClase.Mago           'Bot mago.
            ia_CascoByClase = 662   'Vara
            
       Case eIAClase.Cazador        'Bot kza
            ia_CascoByClase = 405   'de plata
        
End Select

End Function

Function ia_EscudoByClase(ByVal botIndex As Byte) As Integer

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve el escudo según la clase del bot

Select Case ia_Bot(botIndex).Clase

       Case eIAClase.Clerigo        'Bot clero
            ia_EscudoByClase = 130  'De plata.
        
       Case eIAClase.Mago           'Bot mago.
            ia_EscudoByClase = -1   'Nada
            
       Case eIAClase.Cazador        'bot kaza
            ia_EscudoByClase = 404  'escudo d tortu
        
End Select

End Function

Function ia_ArmaByClase(ByVal botIndex As Byte) As Integer

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve el arma según la clase del bot

Select Case ia_Bot(botIndex).Clase

       Case eIAClase.Clerigo        'Bot clero
            ia_ArmaByClase = 129    'Dos filos : P
        
       Case eIAClase.Mago           'Bot mago.
            ia_ArmaByClase = 400    'Vara
            
       Case eIAClase.Cazador        'bot cazador
            ia_ArmaByClase = 665    'arko de kza
        
End Select

End Function

Function ia_VidaByClase(ByVal botIndex As Byte) As Integer

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve la vida según la clase.

Select Case ia_Bot(botIndex).Clase
       Case eIAClase.Clerigo        '<Clerigo.
            'Vida random. (de clerigo 41)
            ia_VidaByClase = 21 + (RandomNumber(8, 10) * 41)
        
       Case eIAClase.Mago           '<Mago
            'Vida random (de mago 39)
            ia_VidaByClase = 21 + (RandomNumber(7, 9) * 39)
            
       Case eIAClase.Cazador        '<Kza
            'Vida random de cazador humano 42
            ia_VidaByClase = 21 + (RandomNumber(8, 11) * 42)
            
End Select

End Function

Function ia_ManaByClase(ByVal botIndex As Byte) As Integer

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve maná según la clase.

Select Case ia_Bot(botIndex).Clase
       Case eIAClase.Clerigo        '<Clerigo.
            'Mana de clero 41 : P
            ia_ManaByClase = 1480
        
       Case eIAClase.Mago           '<Mago
            'Mana de mago 39 : P
            ia_ManaByClase = 1954
            
       Case eIAClase.Cazador        'caza sin mana
            ia_ManaByClase = 0
            
End Select

End Function

Function ia_CalcularGolpe(ByVal victimIndex As Integer) As Integer

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Calcula el golpe (daño) q hace el bot al user.

Dim ParteCuerpo     As Integer
Dim DañoAbsorvido   As Integer

ParteCuerpo = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)

'Si pega en la cabeza.
If ParteCuerpo = PartesCuerpo.bCabeza Then
   'Si tiene casco baja el golpe
       If UserList(victimIndex).Invent.CascoEqpObjIndex <> 0 Then
          DañoAbsorvido = RandomNumber(ObjData(UserList(victimIndex).Invent.CascoEqpObjIndex).MinDef, ObjData(UserList(victimIndex).Invent.CascoEqpObjIndex).MaxDef)
       End If
Else
    'Se fija por la armadura.
       If UserList(victimIndex).Invent.ArmourEqpObjIndex <> 0 Then
          DañoAbsorvido = RandomNumber(ObjData(UserList(victimIndex).Invent.ArmourEqpObjIndex).MinDef, ObjData(UserList(victimIndex).Invent.ArmourEqpObjIndex).MaxDef)
       End If
End If
       
'DEVUELVE.
ia_CalcularGolpe = (RandomNumber(150, 180) - DañoAbsorvido)
       
End Function

Function ia_AciertaGolpe(ByVal victimIndex As Integer) As Boolean

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Evasión del usuario aquí.

Dim tempEvasion     As Long
Dim tempEvasionEsc  As Long
Dim tempResultado   As Long

'Evasión del usuario.
tempEvasion = PoderEvasion(victimIndex)

'Evasión del usuario con escudos.
'Tiene escudo?
If UserList(victimIndex).Invent.EscudoEqpObjIndex <> 0 Then
    tempEvasionEsc = PoderEvasionEscudo(victimIndex)
    tempEvasionEsc = tempEvasion + tempEvasionEsc
Else
    tempEvasionEsc = 0
End If

'Acierta?
tempResultado = MaximoInt(10, MinimoInt(90, 50 + (IA_PROBEX - tempEvasion) * 0.4))

'Random.
ia_AciertaGolpe = (RandomNumber(1, 100) <= tempResultado)

End Function

Function ia_PuedeMeele(ByRef PosBot As WorldPos, ByRef PosVictim As WorldPos, ByRef NewHeading As eHeading) As Boolean

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Se fija si está al lado, y guarda el heading.

With PosVictim
    
    'Mirando hacia la derecha lo tiene ?
    If PosBot.X + 1 = .X Then
       ia_PuedeMeele = (.Y = PosBot.Y)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.EAST
       End If
       
       Exit Function
    End If
    
    'mirando hacia izq?
    If PosBot.X - 1 = .X Then
       ia_PuedeMeele = (.Y = PosBot.Y)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.WEST
       End If
       
       Exit Function
    End If
    
    'mirando arriba
    If PosBot.Y - 1 = .Y Then
       ia_PuedeMeele = (.X = PosBot.X)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.NORTH
       End If
       
       Exit Function
    End If
    
    'Abajo.
    If PosBot.Y + 1 = .Y Then
       ia_PuedeMeele = (PosBot.X = .X)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.SOUTH
       End If
       
       Exit Function
    End If
    
End With

End Function

Sub ia_CreateChar(ByVal ProximoBot As Byte)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Crea el char.

Dim PackageToSend   As String

With ia_Bot(ProximoBot).Char
 
    .body = ObjData(IA_BODY).Ropaje
    .Head = IA_HEAD
    
    'Siempre tienen arma.
    .WeaponAnim = ObjData(ia_ArmaByClase(ProximoBot)).WeaponAnim
    
    'Escudo no, me fijo si tienen..
    If ia_EscudoByClase(ProximoBot) <> -1 Then
        .ShieldAnim = ObjData(ia_EscudoByClase(ProximoBot)).ShieldAnim
    End If
    
    'Casco si..
    .CascoAnim = ObjData(ia_CascoByClase(ProximoBot)).CascoAnim
    
    'Precalculado : P
    .CharIndex = IA_CHAR + ProximoBot
    
    'Preparo el paquete de datos.
    
            Dim tmp_Color   As eNickColor
            
            If ia_Bot(ProximoBot).EsCriminal Then
               tmp_Color = eNickColor.ieCriminal
            Else
               tmp_Color = eNickColor.ieCiudadano
            End If
    
    PackageToSend = PrepareMessageCharacterCreate(.body, .Head, eHeading.SOUTH, .CharIndex, ia_Bot(ProximoBot).Pos.X, ia_Bot(ProximoBot).Pos.Y, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, ia_Bot(ProximoBot).Tag, tmp_Color, 0)
    
    'Actualizo el area.
    ia_SendToBotArea ProximoBot, PackageToSend
    
End With
 
End Sub
 
Sub ia_Spawn(ByVal iaClase As eIAClase, ByRef PosToSpawn As WorldPos, ByRef BotTag As String, ByVal Viajante As Boolean, ByVal esPk As Boolean)
 
' @designer     :  maTih.-
' @date         :  2012/02/01

Dim ProximoBot  As Byte
Dim PackageSend As String

ProximoBot = IA_GetNextSlot

If Not ProximoBot <> 0 Then Exit Sub

With ia_Bot(ProximoBot)
    
    .Invocado = True
    
    .Clase = iaClase
    
    .Mana = ia_ManaByClase(ProximoBot)
    .Vida = ia_VidaByClase(ProximoBot)
    .maxMana = .Mana
    .maxVida = .Vida
    
    .EsCriminal = esPk
    
    'Si es "viajante"..
    .Viajante = Viajante
    
    .Tag = BotTag
    
    .Paralizado = False
    
    'Seteo la posición.
    .Pos = PosToSpawn
    
    'Creo el char.
    ia_CreateChar ProximoBot
   
    'Primer action ! : D
    ia_Action ProximoBot
   
    frmMain.timerIA.Enabled = True
   
    PackageSend = PrepareMessageChatOverHead("VeNGan PutOs xD!", .Char.CharIndex, vbCyan)
   
    ia_SendToBotArea ProximoBot, PackageSend
   
    .Intervalos.SpellCount = 100
   
    NumInvocados = NumInvocados + 1
    
    MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = ProximoBot
   
End With
 
End Sub
 
Public Sub ia_Spells()
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
'Un poco hardcodeado pero bueno :D
 
'Hechizo 1 : descarga.
ia_spell(1).DamageMax = 120
ia_spell(1).DamageMax = 177
ia_spell(1).spellIndex = 23
 
'Hechizo 2 : apoca
 
ia_spell(2).DamageMin = 190
ia_spell(2).DamageMax = 220
ia_spell(2).spellIndex = 25

'Paralizar.
ia_spell(3).DamageMax = 0
ia_spell(3).DamageMin = 0
ia_spell(3).spellIndex = 9

ia_Chats(1) = "Te Voy A hAcER mIErDa xD!!"
ia_Chats(2) = "TE KABE NW"
ia_Chats(3) = "NO SALDRÁS CON VIDA, MALDITO!"
ia_Chats(4) = "FRACASADO, TE HARÉ TRIZAS"
ia_Chats(5) = "TE VAS A IR AL RESU XD"
 
End Sub
 
Sub ia_RandomMoveChar(ByVal botIndex As Byte, ByVal siguiendoIndex As Integer, ByRef HError As Boolean)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
With ia_Bot(botIndex)
 
    Dim nRandom     As Byte
   
    '25% De probabilidades de moverse a
    'cualquiera de las cuatro direcciones.
   
    nRandom = RandomNumber(1, 4)
   
    Select Case nRandom
   
           Case 1
           
           If ia_LegalPos(.Pos.X + 1, .Pos.Y, botIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
           'Borro el BotIndex del tile anterior.
           MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
           .Pos.X = .Pos.X + 1
           
           Case 2
           
           If ia_LegalPos(.Pos.X - 1, .Pos.Y, botIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
                .Pos.X = .Pos.X - 1
           
           Case 3
           
           If ia_LegalPos(.Pos.X, .Pos.Y + 1, botIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
                .Pos.Y = .Pos.Y + 1
           
           Case 4
           
           If ia_LegalPos(.Pos.X, .Pos.Y - 1, botIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
           'Borro el BotIndex del tile anterior.
           MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
           .Pos.Y = .Pos.Y - 1
   
    End Select
 
End With
 
End Sub

Sub ia_CargarRutas(ByRef MAPFILE As String, ByVal MapIndex As Integer)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @modificated  :  Carga las rutas de un mapa : D

Dim loopX   As Long
Dim loopY   As Long
Dim tmpVal  As eHeading

For loopX = 1 To 100
    For loopY = 1 To 100
        
        tmpVal = Val(GetVar(MAPFILE, CStr(loopX) & "," & CStr(loopY), "Direccion"))
        
        If tmpVal <> 0 Then
           MapData(MapIndex, loopX, loopY).Rutas(1) = tmpVal
        End If
        
    Next loopY
Next loopX

End Sub
 
Function ia_LegalPos(ByVal X As Byte, ByVal Y As Byte, ByVal botIndex As Byte, Optional ByVal siguiendoUser As Integer = 0) As Boolean
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @modificated  :  Esta función ya no trabaja con la pos del npc si no que ahora usa los parámetros.
 
ia_LegalPos = False
 
With MapData(ia_Bot(botIndex).Pos.map, X, Y)
 
     '¿Es un mapa valido?
    If (ia_Bot(botIndex).Pos.map <= 0 Or ia_Bot(botIndex).Pos.map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then Exit Function
 
     'Tile bloqueado?
     If .Blocked <> 0 Then Exit Function
   
     'Hay un usuario?
     If .UserIndex > 0 Then
        'Si no es un adminInvisible entonces nos vamos.
        If UserList(.UserIndex).flags.AdminInvisible <> 1 Then Exit Function
    End If
 
    'Hay un NPC?
    If .NpcIndex <> 0 Then Exit Function
     
    'Hay un bot?
    If .botIndex <> 0 Then Exit Function
    
    'Siguiendo Index?
    If siguiendoUser <> 0 Then
        'Válido para evitar el rango Y pero no su eje X.
        If Abs(Y - UserList(siguiendoUser).Pos.Y) > RANGO_VISION_Y Then Exit Function
   
        If Abs(X - UserList(siguiendoUser).Pos.X) > RANGO_VISION_X Then Exit Function
    End If
    
     ia_LegalPos = True
   
End With
 
End Function
 
Sub ia_SearchPath(ByVal botIndex As Byte, ByRef tPos As WorldPos, ByRef findHeading As eHeading)

' @designer     :  maTih.-
' @date         :  2012/03/13
' @                Buscá una ruta y guarda el heading.

findHeading = FindDirection(ia_Bot(botIndex).Pos, tPos)

End Sub

Sub ia_MoveToHeading(ByVal botIndex As Byte, ByVal toHeading As eHeading, ByRef FoundErr As Boolean)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Mueve el char del npc hacia una posición.

FoundErr = True

Select Case toHeading

       Case eHeading.NORTH  '<Move norte.
            'No legal pos.
            If Not ia_LegalPos(ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y - 1, botIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(botIndex).Pos.map, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y).botIndex = 0
            'Set la nueva posición
            ia_Bot(botIndex).Pos.Y = ia_Bot(botIndex).Pos.Y - 1
            
       Case eHeading.EAST   '<Move este.
            'Si hay posición inválida no se peude mover.
            If Not ia_LegalPos(ia_Bot(botIndex).Pos.X + 1, ia_Bot(botIndex).Pos.Y, botIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(botIndex).Pos.map, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y).botIndex = 0
            
            'Set la nueva posición
            ia_Bot(botIndex).Pos.X = ia_Bot(botIndex).Pos.X + 1
            
       Case eHeading.SOUTH  '<Move sur.
            'Si hay posición inválida no se peude mover.
            If Not ia_LegalPos(ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y + 1, botIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(botIndex).Pos.map, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y).botIndex = 0
            
            'Set la nueva posición
            ia_Bot(botIndex).Pos.Y = ia_Bot(botIndex).Pos.Y + 1
            
       Case eHeading.WEST   '<Move oeste.
            'Si hay posición inválida no se peude mover.
            If Not ia_LegalPos(ia_Bot(botIndex).Pos.X - 1, ia_Bot(botIndex).Pos.Y, botIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(botIndex).Pos.map, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y).botIndex = 0
            
            'Set la nueva posición
            ia_Bot(botIndex).Pos.X = ia_Bot(botIndex).Pos.X - 1
            
End Select

FoundErr = False

End Sub


Sub ia_MoveViajante(ByVal botIndex As Byte, ByVal Direccion As eHeading)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Move el viajante hacia una posición

Dim HabiaAgua As Boolean

With ia_Bot(botIndex)

     'Hacia donde se mueve..
     Select Case Direccion
            
            Case eHeading.NORTH     'Norte.
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
                 .Pos.Y = .Pos.Y - 1
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = botIndex
                 
            Case eHeading.EAST      'Este.
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
                 .Pos.X = .Pos.X + 1
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = botIndex
            
            Case eHeading.SOUTH     'Sur.
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
                 .Pos.Y = .Pos.Y + 1
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = botIndex
                 
            Case eHeading.WEST      'Oeste.
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
                 .Pos.X = .Pos.X - 1
                 MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = botIndex
     End Select
     
     HabiaAgua = HayAgua(.Pos.map, .Pos.X, .Pos.Y)
     
     If HabiaAgua Then
        'Si hay agua cambio el cuerpo.
        ia_SendToBotArea botIndex, PrepareMessageCharacterChange(395, 0, Direccion, .Char.CharIndex, 0, 0, 0, 0, 0)
        .Navegando = True
     Else
        'No habia agua, y... estaba navegando?
        If .Navegando Then
           'cambio el body y demas.
           ia_SendToBotArea botIndex, PrepareMessageCharacterChange(.Char.body, .Char.Head, Direccion, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim)
           .Navegando = False
        End If
    End If
    
     'Actualizamso
     ia_SendToBotArea botIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y)
        
End With

End Sub

Function ia_HeadingToMolestNpc(ByVal NpcIndex As Integer) As eHeading


' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve un heading para un npc que está molestando el paso.

Dim nPos    As WorldPos

nPos = Npclist(NpcIndex).Pos

With MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)

     'Pos legal hacia arriba?
     If LegalPosNPC(nPos.map, nPos.X, nPos.Y - 1, 0) Then
        'Mientras no halla bot.
        If Not .botIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.NORTH
        End If
     End If
     
     'Pos legal hacia abajo?
     If LegalPosNPC(nPos.map, nPos.X, nPos.Y + 1, 0) Then
        'Mientras no halla bot.
        If Not .botIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.SOUTH
        End If
     End If
     
     'Pos legal hacia la izquierda?
     If LegalPosNPC(nPos.map, nPos.X - 1, nPos.Y, 0) Then
        'Mientras no halla bot.
        If Not .botIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.WEST
        End If
     End If
     
     'Pos legal hacia la derecha?
     If LegalPosNPC(nPos.map, nPos.X + 1, nPos.Y, 0) Then
        'Mientras no halla bot.
        If Not .botIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.EAST
        End If
     End If
     
End With

End Function

Function ia_Objetos(ByVal botIndex As Byte) As WorldPos

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Busca objetos valiosos en el area.

Dim loopX   As Long
Dim loopY   As Long
Dim BotPos  As WorldPos

BotPos = ia_Bot(botIndex).Pos

For loopY = BotPos.Y - MinYBorder + 1 To BotPos.Y + MinYBorder - 1
        For loopX = BotPos.X - MinXBorder + 1 To BotPos.X + MinXBorder - 1
            'Hay un objeto.
            If MapData(BotPos.map, loopX, loopY).ObjInfo.objIndex <> 0 Then
               If ObjData(MapData(BotPos.map, loopX, loopY).ObjInfo.objIndex).Valioso <> 0 Then
                  ia_Objetos.map = BotPos.map
                  ia_Objetos.X = loopX
                  ia_Objetos.Y = loopY
                  Exit Function
                End If
            End If
        Next loopX
Next loopY

ia_Objetos.map = 0

End Function

Function ia_SlotInventario(ByVal botIndex As Byte) As Byte

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Busca un slot libre.

Dim loopX   As Long

For loopX = 1 To IA_SLOTS
    With ia_Bot(botIndex).Inv(loopX)
         'No hay objeto.
         If Not .objIndex <> 0 Then
            ia_SlotInventario = CByte(loopX)
            Exit Function
         End If
    End With
Next loopX

ia_SlotInventario = 0

End Function

Sub ia_ActionViajante(ByVal botIndex As Byte)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Acciones de los bots que viajan hacia mapas.

Dim RutaDir     As eHeading
Dim molestNpc   As Integer
Dim ObjetoPos   As WorldPos

With ia_Bot(botIndex)

     'Está paralizado?
     If .Paralizado Then
        'Puede tirar hechizos.
        If .Intervalos.SpellCount = 0 Then
           'se remueve
           ia_SendToBotArea botIndex, PrepareMessageChatOverHead("AN HOAX VORP", .Char.CharIndex, vbCyan)
           .Paralizado = False
           .Intervalos.SpellCount = (IA_SINT / 30)
        End If
    End If
        
     'Se puede mover?
     If Not .Intervalos.MoveCharCount = 0 Then Exit Sub
        
     .Intervalos.MoveCharCount = (IA_MOVINT / 50)
     
     'Tiene una ruta?
     RutaDir = ia_HayRuta(.Pos)
    
     'Ve un objeto valioso?
     ObjetoPos = ia_Objetos(botIndex)
     
     If ObjetoPos.map <> 0 Then
        'Lo va a buscar, pero antes , setea su vieja pos.
        If Not .UltimaIdaObjeto Then
            .ViajanteAntes = .Pos
        End If
        
        ia_SearchPath botIndex, ObjetoPos, RutaDir
        .UltimaIdaObjeto = True
     End If
     
     'No hay ruta?
     If Not RutaDir <> 0 Then
        'habia atacado un usuario? si es así volvemos a la pos.
        ia_SearchPath botIndex, .ViajanteAntes, RutaDir
     End If
     
     If RutaDir <> 0 Then
        
        'Hacia donde mueve?
        Select Case RutaDir
                
               Case eHeading.NORTH      '<Mueve norte.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map, .Pos.X, .Pos.Y - 1).NpcIndex
                    
                    #If Barcos <> 0 Then
                        If molestNpc <> 0 Then
                            ia_SendToBotArea botIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                            Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                        End If
                    #End If
                    
               Case eHeading.SOUTH      '<Mueve sur.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map, .Pos.X, .Pos.Y + 1).NpcIndex
                    
                    If molestNpc <> 0 Then
                       ia_SendToBotArea botIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                       'Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
                       
               Case eHeading.EAST       '<Mueve este.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map, .Pos.X + 1, .Pos.Y).NpcIndex
                    
                    If molestNpc <> 0 Then
                       ia_SendToBotArea botIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                       'Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
                    
               Case eHeading.WEST       '<Mueve oeste.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map, .Pos.X - 1, .Pos.Y).NpcIndex
                    
                    If molestNpc <> 0 Then
                       ia_SendToBotArea botIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                       'Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
        End Select
        
        'Move:p
        ia_MoveViajante botIndex, RutaDir
        'Set el heading.
        .Char.heading = RutaDir
     End If
     
     'Está en una pos que hay un objeto?
     If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objIndex <> 0 Then
        'Es valioso?
        If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objIndex).Valioso <> 0 Then
           'Busca un slot y lo guarda en su inventario
           Dim FreeSlotInInvent     As Byte
           Dim ObjInPosition        As Obj
           FreeSlotInInvent = ia_SlotInventario(botIndex)
           'Lo agarra siempre que alla slot.
           If FreeSlotInInvent <> 0 Then
              'Usamos este buffer.
              ObjInPosition.objIndex = MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.objIndex
              ObjInPosition.Amount = MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.Amount
              'Lo guardamos.
              .Inv(FreeSlotInInvent) = ObjInPosition
              'Borra.
              EraseObj 10000, .Pos.map, .Pos.X, .Pos.Y
              .UltimaIdaObjeto = False
           End If
        End If
    End If
     
     'Encontramos una salida? - translados.
     If MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.map <> 0 Then
        'Mapa válido?
        If MapaValido(MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.map) Then
            'Asignamos nuevas posiciones, borramos el char anterior.
            ia_SendToBotArea botIndex, PrepareMessageCharacterRemove(.Char.CharIndex)
            'Pos del npc.
            .Pos.map = MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.map
            
            'Por si no tiene heading.
            If Not .Char.heading <> 0 Then .Char.heading = eHeading.SOUTH
            
            'Nueva X?
            If MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.X <> 0 Then
                .Pos.X = MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.X
            End If
            
            'Nueva Y?
            If MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.Y <> 0 Then
                .Pos.Y = MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.Y
            End If
            
             MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = botIndex
            'Creamos.
            
            Dim tmp_Color   As eNickColor
            
            If .EsCriminal Then
               tmp_Color = eNickColor.ieCriminal
            Else
               tmp_Color = eNickColor.ieCiudadano
            End If
            
            ia_SendToBotArea botIndex, PrepareMessageCharacterCreate(.Char.body, .Char.Head, .Char.heading, .Char.CharIndex, .Pos.X, .Pos.Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, .Tag, tmp_Color, 0)
        End If
     End If
     
End With

End Sub

Function ia_HayRuta(ByRef InPos As WorldPos) As eHeading

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Devuelve la dircción de la ruta en una pos.

With MapData(InPos.map, InPos.X, InPos.Y)
     
     ia_HayRuta = .Rutas(1)
     
End With

End Function

Sub ia_SupportOthers(ByVal botIndex As Byte, ByRef Supported As Boolean)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Un bot supportea otro.

Dim botIndexToSupport   As Byte
Dim supportAction       As eIASupportActions

'Si no tiene intervalo..
If ia_Bot(botIndex).Intervalos.SpellCount <> 0 Then Exit Sub

'Busca un bot a ayudar.
botIndexToSupport = ia_GetSupportBot(botIndex, supportAction)

'No encontró, no supportea..
If Not botIndexToSupport <> 0 Then Supported = False: Exit Sub

'Que acción debe realizar?
Select Case supportAction

       Case eIASupportActions.SCurar        '<Cura!
            'Lanza graves.
            'Crea fx.
            ia_SendToBotArea botIndexToSupport, mod_DunkanProtocol.Send_CreateSpell(ia_Bot(botIndex).Char.CharIndex, ia_Bot(botIndexToSupport).Char.CharIndex, Hechizos(5).EffectIndex, Hechizos(5).loops)
            
            'Cartel.
            ia_SendToBotArea botIndex, PrepareMessageChatOverHead("EN CORP SANCTIS", ia_Bot(botIndex).Char.CharIndex, vbCyan)
            
            'Suma un random de vida.
            ia_Bot(botIndexToSupport).Vida = ia_Bot(botIndexToSupport).maxVida + RandomNumber(55, 77)
            
            'PARA QUE NO PASE LA VIDA MAXIMA
            If ia_Bot(botIndexToSupport).Vida > ia_Bot(botIndexToSupport).maxVida Then ia_Bot(botIndexToSupport).Vida = ia_Bot(botIndexToSupport).maxVida
       
            Supported = True
       
      Case eIASupportActions.SRemover       '<Remueve paralizis.
            'Crea el fx, el remo no tiene fx.
            'ia_sendtobotarea botindextosupport
            
            'Paralizis count.
            If ia_Bot(botIndexToSupport).Intervalos.ParalizisCount > 6 Then Exit Sub
            
            'Cartel
            ia_SendToBotArea botIndex, PrepareMessageChatOverHead("AN HOAX VORP", ia_Bot(botIndex).Char.CharIndex, vbCyan)
            
            'Saca el flag
            ia_Bot(botIndexToSupport).Paralizado = False
            
            Supported = True
            
End Select

End Sub

Function ia_BotEnArea(ByVal botIndex As Byte, ByVal otherBotIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim BotIndexPos As WorldPos

BotIndexPos = ia_Bot(botIndex).Pos

Dim loopX   As Long
Dim loopY   As Long

For loopY = BotIndexPos.Y - MinYBorder + 1 To BotIndexPos.Y + MinYBorder - 1
        For loopX = BotIndexPos.X - MinXBorder + 1 To BotIndexPos.X + MinXBorder - 1

            If MapData(BotIndexPos.map, loopX, loopY).botIndex = otherBotIndex Then
                ia_BotEnArea = True
                Exit Function
            End If
        
        Next loopX
Next loopY

ia_BotEnArea = False

End Function

Function ia_GetSupportBot(ByVal botIndex As Byte, ByRef SAction As eIASupportActions) As Byte

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Busca un bot a ayudar.

Dim loopX   As Long

For loopX = 1 To MAX_BOTS
    
    'Si no es mi BotIndex
    If loopX <> botIndex Then
        
       'Está invocado?
       If ia_Bot(loopX).Invocado Then
          'Está en el area?
          If ia_BotEnArea(botIndex, loopX) Then
             'Está paralizado/tiene poca vida?
             If ia_Bot(loopX).Vida <> ia_Bot(loopX).maxVida Or ia_Bot(loopX).Paralizado Then
                'Encontrado.
                ia_GetSupportBot = CByte(loopX)
                'Devuelve la acción.
                SAction = IIf(ia_Bot(loopX).Vida <> ia_Bot(loopX).maxVida, eIASupportActions.SCurar, eIASupportActions.SRemover)
                Exit Function
             End If
          End If
       End If
       
    End If
    
Next loopX

ia_GetSupportBot = 0
End Function

Sub ia_Action(ByVal botIndex As Byte)
 
On Error GoTo Errhandler        '< maTih XD
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Acciones de los bots.
 
Dim pIndex      As Integer
Dim sRandom     As Integer
Dim rMan        As Integer
Dim FoundErr    As Boolean
Dim moveHeading As eHeading
Dim AyudoBot    As Boolean
 
If EnPausa Then Exit Sub
 
With ia_Bot(botIndex)
 
    'Es un bot viajante?
    If .Viajante Then
          'Mientras no esté contra ningún pibe
          If Not .ViajanteUser <> 0 Then
             ia_CheckInts botIndex
             ia_ActionViajante botIndex
             Exit Sub
          End If
    End If
    
    If Not .ViajanteUser <> 0 Then
        pIndex = ia_FindTarget(.Pos, .EsCriminal)
    Else
        pIndex = .ViajanteUser
    End If
    
    'No hay usuario.
    If pIndex <= 0 Then Exit Sub
 
    'Contadores de intervalo.
    ia_CheckInts botIndex
   
    'EL bot boquea!
    If Not .Intervalos.ChatCount <> 0 Then
       .Intervalos.ChatCount = (IA_TALKIN / 40)
        
       'Envia msj random
       ia_SendToBotArea botIndex, PrepareMessageChatOverHead(ia_Chats(RandomNumber(1, 5)), .Char.CharIndex, vbRed)
       .Intervalos.SpellCount = (IA_SINT / 100)
    End If
    
    'Si se puede mover AND no está inmo se mueve al azar.
    If .Intervalos.MoveCharCount = 0 And .Paralizado = False Then
        
        'Tiene target?
        If pIndex <> 0 Then
           'busco un path.
           ia_SearchPath botIndex, UserList(pIndex).Pos, moveHeading
        End If
        
        'Es clero?
        If Not .Clase <> eIAClase.Clerigo Then
           'Si tiene la vida llena lo persigue.
           If .Vida = .maxVida Then
              ia_MoveToHeading botIndex, moveHeading, FoundErr
           Else
            'Si no , se mueve al azar.
              ia_RandomMoveChar botIndex, pIndex, FoundErr
           End If
         End If
                   
         'Es mago?
        If .Clase = eIAClase.Mago Or .Clase = eIAClase.Cazador Then
           'Si no tiene la vida llena se mueve al azar.
           If Not .Vida = .maxVida Then
              ia_RandomMoveChar botIndex, pIndex, FoundErr
           Else
              'Tiene la vida llena, que fue el ultimo movimiento?
              'Siguio la victima?
              If .UltimoMovimiento = eIAMoviments.SeguirVictima Then
                 'Mueve random.
                 ia_RandomMoveChar botIndex, pIndex, FoundErr
                 'Seteo.
                 .UltimoMovimiento = eIAMoviments.MoverRandom
              Else
                 'Se movió al azar, sigue su victima.
                 ia_MoveToHeading botIndex, moveHeading, FoundErr
                 'Seteo el nuevo flag.
                 .UltimoMovimiento = eIAMoviments.SeguirVictima
             End If
        End If
       End If
       
       
        If Not FoundErr Then
                'Se movió, guardo el BotIndex.
                MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = botIndex
       
                'NEW--------
                'Checkeo si es una posición válida.
 
                'Actualizamos.
                ia_SendToBotArea botIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y)
       
                .Intervalos.MoveCharCount = (IA_MOVINT / 40)
        End If
        
    End If
   
   
    'STATS..
   
        'Prioriza la vida ante todo
       
        If .Vida < .maxVida Then
           
            'Checkeo el intervalo.
            If .Intervalos.UseItemCount > 0 Then Exit Sub
           
            'Recupera 20 cada 200 ms.
            .Vida = .Vida + 20
           
            If .Vida > .maxVida Then .Vida = .maxVida
           
            'Uso la poción, seteo el interval
            .Intervalos.UseItemCount = (IA_USEOBJ / 40)
           
            Exit Sub
        End If
       
        'Si tenia la vida llena usa azules.
       
        If .Mana < .maxMana Then
       
            'Checkeo el intervalo.
           
            If .Intervalos.UseItemCount = 0 Then
       
                'Recupera un % de la mana.
                If .Clase <> eIAClase.Mago Then
                Dim recuperoMana    As Long
                    recuperoMana = Porcentaje(.maxMana, 5)
                Else
                    recuperoMana = Porcentaje(.maxMana, 3)
                End If
                
                .Mana = .Mana + recuperoMana
           
                If .Mana > .maxMana Then .Mana = .maxMana
           
            .Intervalos.UseItemCount = (IA_USEOBJ / 40)
 
            End If
           
            'Hacer una constante después, con esto hacemos un random
            'Para que azulee y combee a la ves.
            If RandomNumber(1, 4) < 4 Then Exit Sub
        End If
   
    'Bueno si está acá es por que tenia la vida y mana llenas.
     
    'Es cazador??
    If .Clase = eIAClase.Cazador Then
       'Intervalo permite?
       If Not .Intervalos.ArrowCount = 0 Then Exit Sub
       'Kza manqea ! XD - 25% de prob fallar
       If RandomNumber(1, 100) > 65 Then Exit Sub
       'Probabilidad de evadir.
       If Not RandomNumber(1, 100) <= MaximoInt(10, MinimoInt(90, 50 + ((220 - PoderEvasion(pIndex)) * 0.4))) Then
          'Atacó y falló!!
          Call WriteConsoleMsg(pIndex, .Tag & " Te lanzó un flechazo pero falló!", FontTypeNames.FONTTYPE_FIGHT)
          'setea intervalo
          .Intervalos.ArrowCount = (IA_PROINT / 25)
          Exit Sub
       End If
       
       Dim ArrowDamage  As Integer  '<DañoBase.
       Dim ArmourIndex  As Integer  '<ArmaduraObjIndex
       Dim HelmetIndex  As Integer  '<CascoObjIndex
       
       ArrowDamage = RandomNumber(185, 225)
       
       'Restamos si tiene armadura.
       ArmourIndex = UserList(pIndex).Invent.ArmourEqpObjIndex
       HelmetIndex = UserList(pIndex).Invent.CascoEqpObjIndex
       
       'Pega en cabeza?
       If RandomNumber(1, 6) = 6 Then
          'Absorve.
          If HelmetIndex <> 0 Then
             ArrowDamage = ArrowDamage - RandomNumber(ObjData(HelmetIndex).MinDef, ObjData(HelmetIndex).MaxDef)
          End If
       Else
          'Armadura absorce.
          If ArmourIndex <> 0 Then
             ArrowDamage = ArrowDamage - RandomNumber(ObjData(ArmourIndex).MinDef, ObjData(ArmourIndex).MaxDef)
          End If
       End If
       
       'crea fx.
       SendData SendTarget.ToPCArea, pIndex, mod_DunkanProtocol.Send_CreateArrow(.Char.CharIndex, UserList(pIndex).Char.CharIndex, ObjData(553).GrhIndex)
       
       'crea daño
       Call mod_DunkanGeneral.Enviar_DañoAUsuario(pIndex, ArrowDamage)
       
       'Sacude un flechazo.
       UserList(pIndex).Stats.MinHp = UserList(pIndex).Stats.MinHp - ArrowDamage
       
       Call WriteConsoleMsg(pIndex, .Tag & " Te ha pegado un flechazo por " & ArrowDamage, FontTypeNames.FONTTYPE_FIGHT)
       
       'Muere?
       If UserList(pIndex).Stats.MinHp <= 0 Then
          UserDie pIndex
          Call WriteConsoleMsg(pIndex, .Tag & " Te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
       End If
        
       'Intervalo
       .Intervalos.ArrowCount = (IA_PROINT / 20)
        
       'client update
       WriteUpdateHP pIndex
       Exit Sub
    End If
    'Puede castear?
    'Si el usuario no tiene la vida llena ataca
    Dim tmpHP   As Long
    
    tmpHP = (UserList(pIndex).Stats.MinHp)
    
    tmpHP = (tmpHP * 100) / (UserList(pIndex).Stats.MaxHp)
   
    If .Intervalos.SpellCount = 0 Then
    
    'Es clérigo And puedepegar??
    If Not .Clase <> eIAClase.Clerigo And .Intervalos.HitCount = 0 And Not .UltimaAccion = eIAactions.ePegar Then
       'Está al alcance de la víctima para un gole meele?
       Dim newBotHeading   As eHeading
       If ia_PuedeMeele(.Pos, UserList(pIndex).Pos, newBotHeading) Then
            'Acierta el golpe?
            If ia_AciertaGolpe(pIndex) Then
               'Antes que nada cambiamos el heading, si es válido.
               If newBotHeading <> 0 And newBotHeading <> .Char.heading Then
                    'ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_ChangeHeadingChar(.Char.CharIndex, newBotHeading)
               End If
               
               'Calcula el golpe
               Dim GolpeVal     As Integer
               GolpeVal = ia_CalcularGolpe(pIndex)
               
               'Resta.
               UserList(pIndex).Stats.MinHp = UserList(pIndex).Stats.MinHp - GolpeVal
               
               'crea el fx de la sangre.
               SendData SendTarget.ToPCArea, pIndex, PrepareMessageCreateFX(UserList(pIndex).Char.CharIndex, FXSANGRE, 5)
               
               'Avisa.
               Call WriteConsoleMsg(pIndex, .Tag & " Te ha pegado por " & GolpeVal & ".", FontTypeNames.FONTTYPE_FIGHT)
               
               'Setea flag.
               .UltimaAccion = eIAactions.ePegar
               
               'Muere?
               If UserList(pIndex).Stats.MinHp <= 0 Then
                  Call UserDie(pIndex)
               End If
               
               'update hp.
               WriteUpdateHP pIndex
               
               'Intervalo de golpe.
               .Intervalos.HitCount = (IA_HITINT / 40)
               'Intervalo de hechizo.
               .Intervalos.SpellCount = (IA_SINT / 40)
               'Intervalo de golpe+pociones.
               .Intervalos.UseItemCount = (IA_USEOBJ / 60)
               Exit Sub
            End If
        End If
    End If
    
       'Feo, aunque digamos que solo hace apoca desc remo
       'Así que va a andar bien.
       
       'Si la mana es < a 300 [gasto del remo] no hacemos nada.
       
       If .Mana < 300 Then Exit Sub
       
       'Si está paralizado AND el usuario no tiene poka vida prioriza removerse.
       
        If .Paralizado And tmpHP > 60 Then
            
            'Intervalo de remo :@
            If .Intervalos.ParalizisCount <> 0 Then Exit Sub
            
            'Palabras mágicas.
            
            ia_SendToBotArea botIndex, PrepareMessageChatOverHead(Hechizos(10).PalabrasMagicas, .Char.CharIndex, vbCyan)
           
            .Paralizado = False
           
            'Agrego esto por que si no tirarle inmo era al pedo
            'Seguia caminando practicamente :PP
           
            .Intervalos.ParalizisCount = (IA_SREMO / 10)
           
            'Se removió entonces salimos del sub y seteamos el intervalo
           
            .Intervalos.SpellCount = (IA_SINT / 40)
           
            Exit Sub
           
        End If
       
        'No está paralizado entonces castea un hechizo random.
       
        'Si, es un robot, pero puede manquear no?
       
        If RandomNumber(1, 100) > IA_CASTEO Then Exit Sub
       
        sRandom = RandomNumber(1, IA_M_SPELL)
       
        'Ayuda otros bots si es que hay
        If NumInvocados <> 1 Then
           ia_SupportOthers botIndex, AyudoBot
           
           If AyudoBot Then
              'SETEA INTERVALO
              .Intervalos.SpellCount = (IA_SINT / 40)
              Exit Sub
           End If
        End If
           
        'Si el usuario ya estaba paralizado AND el random es paralizar, entonces buscamos de nuevo
        If UserList(pIndex).flags.Paralizado = 1 And sRandom = 3 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
        
        'Si soy mago y el usuario es mago también no paraliza.
        If UserList(pIndex).Clase = eClass.Mage And .Clase = eIAClase.Mago Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
        
        'Si el usuario tiene menos del 75% de vida juega al ataque.
        
        If tmpHP < 75 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
        
        'Si no llega con la mana del hechizo AND la del otro
        'tampoco entonces no hacemos nada
       
        If sRandom = 1 Then
           
            'Si no llega a la mana del spell 1 (descarga)
            'No hacemos nada ya que tampoco llega
            'al apocalipsis.
           
            rMan = Hechizos(ia_spell(1).spellIndex).ManaRequerido
           
            If rMan > .Mana Then Exit Sub
           
        ElseIf sRandom = 2 Then
       
            rMan = Hechizos(ia_spell(2).spellIndex).ManaRequerido
               
            'Pero si es spell 2 (apoca) AND llegamos
            'con la mana para descarga, entonces
            'Seteamos sRandom como 1 y casteamos
            'descarga.
           
            If rMan > .Mana Then
               
                'Modifico la formula y hago un random
                'Dado a que una ves que queda en -1000 de mana
                'Nunca más tira apoca y castea puras descargas.
               
                If .Mana > 460 And RandomNumber(1, 100) < 30 Then
                    sRandom = 1
                Else
                    Exit Sub
                End If
            End If
       End If
       
        rMan = Hechizos(ia_spell(sRandom).spellIndex).ManaRequerido
       
        'Descontamos la maná y seteamos el intervalo.
        .Mana = .Mana - rMan
       
        'Set última action.
        .UltimaAccion = eIAactions.eMagia
        
        .Intervalos.SpellCount = (IA_SINT / 20) 'Se chekea cada 40 ms.
       
        'Creamos el fx y le descontamos la vida al usuario.
       
        ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_CreateSpell(.Char.CharIndex, UserList(pIndex).Char.CharIndex, Hechizos(ia_spell(sRandom).spellIndex).EffectIndex, Hechizos(ia_spell(sRandom).spellIndex).loops)
       
        ia_SendToBotArea botIndex, PrepareMessageChatOverHead(Hechizos(ia_spell(sRandom).spellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan)
       
        'Paralizar?
        If sRandom = 3 Then
           'Paralizado : P
           UserList(pIndex).flags.Paralizado = 1
           UserList(pIndex).Counters.Paralisis = IntervaloParalizado
           WriteParalizeOK pIndex
           WriteConsoleMsg pIndex, .Tag & " Te ha paralizado", FontTypeNames.FONTTYPE_FIGHT
        End If
       
        'Random damage :D
       
        sRandom = RandomNumber(ia_spell(sRandom).DamageMin, ia_spell(sRandom).DamageMax)
       
        'Al daño le restamos , si el usuario tiene, defensa mágica.
        If UserList(pIndex).Invent.AnilloEqpObjIndex <> 0 Then
           sRandom = sRandom - RandomNumber(ObjData(UserList(pIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(pIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        'NO numeros negativos.
        If sRandom < 0 Then sRandom = 0
       
        'Quitamos daño.
    
        UserList(pIndex).Stats.MinHp = UserList(pIndex).Stats.MinHp - sRandom
            
        If sRandom <> 0 Then
            'AVISO AL USUARIO!
            Call WriteConsoleMsg(pIndex, .Tag & " Te ha quitado " & sRandom & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        'Check si muere.
       
        If UserList(pIndex).Stats.MinHp <= 0 Then
             UserDie pIndex
            'Era viajante y mató el usuario?
             If Not pIndex <> .ViajanteUser Then
                .ViajanteUser = 0
             End If
             
             'aviso que murio.
             Call WriteConsoleMsg(pIndex, .Tag & " Te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
             
        End If
       
        'Actualizamos el cliente.
       
        WriteUpdateHP pIndex
       
    End If
End With
 
Exit Sub

Errhandler:
 
End Sub

Sub ia_EnviarChar(ByVal UserIndex As Integer, ByVal botIndex As Byte)

' @designer     :  maTih.-
' @date         :  2012/03/13
' @                Envia el char del bot a un usuario (sistema de areas!!)

With ia_Bot(botIndex).Char
            Dim tmp_Color   As eNickColor
            
            If ia_Bot(botIndex).EsCriminal Then
               tmp_Color = eNickColor.ieCriminal
            Else
               tmp_Color = eNickColor.ieCiudadano
            End If
            
     Call Protocol.WriteCharacterCreate(UserIndex, .body, .Head, eHeading.SOUTH, .CharIndex, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, ia_Bot(botIndex).Tag, tmp_Color, 0)
End With

End Sub
 
Sub ia_UserDamage(ByVal Spell As Byte, ByVal botIndex As Byte, ByVal UserIndex As Integer)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
Dim rMan     As Integer
Dim Damage   As Integer
Dim usedFont As FontTypeNames
 
usedFont = FontTypeNames.FONTTYPE_FIGHT
 
'Checkeo que el hechizo no sea 0.
If Not Spell <> 0 Then Exit Sub
 
With UserList(UserIndex)
 
    rMan = Hechizos(Spell).ManaRequerido
   
    'Llega con la mana?
   
    If rMan > .Stats.MinMAN Then
        WriteConsoleMsg UserIndex, "No tienes suficiente mana!", usedFont
        Exit Sub
    End If
    
    'Skills?
    
    If Hechizos(Spell).MinSkill > .Stats.UserSkills(eSkill.Magia) Then
       WriteConsoleMsg UserIndex, "No tienes suficientes skills en magia", usedFont
       Exit Sub
    End If
   
    'Soy ciudadano y el target es un bot viajante?
    
    If Not criminal(UserIndex) And ia_Bot(botIndex).Viajante And .flags.Seguro Then
        WriteConsoleMsg UserIndex, "Para atacar bots viajantes debes desactivar el seguro", usedFont
        Exit Sub
    End If
    
    If Hechizos(Spell).Inmoviliza Or Hechizos(Spell).Paraliza Then
       
        'Le pongo el flag en verdadero.
        ia_Bot(botIndex).Paralizado = True
       
        'Mensaje informando.
        WriteConsoleMsg UserIndex, "Has paralizado ah " & ia_Bot(botIndex).Tag, usedFont
        
        'Creo la animacion sobre el char.
        ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_CreateSpell(.Char.CharIndex, ia_Bot(botIndex).Char.CharIndex, Hechizos(Spell).EffectIndex, Hechizos(Spell).loops)
        
        'SpellWorlds.
        DecirPalabrasMagicas Hechizos(Spell).PalabrasMagicas, UserIndex
       
        'Quito mana y energia
        .Stats.MinMAN = .Stats.MinMAN - rMan
       
        'le doy intervalo
       
        ia_Bot(botIndex).Intervalos.ParalizisCount = (IA_SREMO / 10)
       
        WriteUpdateMana UserIndex
       
        Exit Sub
    End If
   
    'Era un Viajante
   
    Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
    Damage = Damage + Porcentaje(Damage, 3 * .Stats.ELV)
   
    If Not Damage <> 0 Then Exit Sub
   
   If ia_Bot(botIndex).Viajante Then
        Dim eraPK   As Boolean
        
        If Not ia_Bot(botIndex).ViajanteAntes.map Then
            ia_Bot(botIndex).ViajanteAntes = ia_Bot(botIndex).Pos
        End If
        
        'No era criminal.
        eraPK = criminal(UserIndex)
    
        'No era criminal y atacó un viajante, es criminal.
        If Not eraPK Then VolverCriminal UserIndex
    
        'Ahora el bot se enojó viejo..
        ia_Bot(botIndex).ViajanteUser = UserIndex
    
        UserList(UserIndex).AtacoViajante = botIndex
    
        WriteConsoleMsg UserIndex, "Has atacado un viajante!! ahora eres un criminal, y además el viajante te atacará!", usedFont
   End If
   
        'Quitamos vida
    
    If Hechizos(Spell).StaffAffected Then
       If UserList(UserIndex).Clase = eClass.Mage Then
          If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
             Damage = (Damage * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
          Else
             Damage = Damage * 0.7 'Baja damage a 70% del original
          End If
        End If
     End If
        
     If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDELFICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
        Damage = Damage * 1.04  'laud magico de los bardos
     End If
    
    ia_Bot(botIndex).Vida = ia_Bot(botIndex).Vida - Damage
    
    'No está paralizado.
    If Not ia_Bot(botIndex).Paralizado Then
        'Le pegaron, se cagó todo y se mueve random.
        Dim keepMoving  As Boolean
    
        ia_RandomMoveChar botIndex, UserIndex, keepMoving
    
        'No hubo error, por ende se movió.
        If Not keepMoving Then
           'Guardo la nueva pos.
           MapData(ia_Bot(botIndex).Pos.map, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y).botIndex = botIndex
       
           'Actualizo el area del bot.
           ia_SendToBotArea botIndex, PrepareMessageCharacterMove(ia_Bot(botIndex).Char.CharIndex, ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y)
       
           'Intervalo de caminata.
           ia_Bot(botIndex).Intervalos.MoveCharCount = (IA_MOVINT / 40)
        End If
    End If
    
    'Aviso al usuario.
    WriteConsoleMsg UserIndex, "Le has quitado " & Damage & " puntos de vida a " & ia_Bot(botIndex).Tag, usedFont
   
    'Tiro las spell worlds
    DecirPalabrasMagicas Hechizos(Spell).PalabrasMagicas, UserIndex
   
    'Creo el fx.
    ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_CreateSpell(.Char.CharIndex, ia_Bot(botIndex).Char.CharIndex, Hechizos(Spell).EffectIndex, Hechizos(Spell).loops)
   
    'saco mana y energia y actualizo el cliente
    .Stats.MinMAN = .Stats.MinMAN - rMan
       
    WriteUpdateMana UserIndex
   
    If ia_Bot(botIndex).Vida <= 0 Then
        'Murió?
        ia_EraseChar botIndex, True
        WriteConsoleMsg UserIndex, "Has matado ah " & ia_Bot(botIndex).Tag & ".", usedFont
    End If
   
End With
 
End Sub
 
Sub ia_DamageHit(ByVal botIndex As Byte, ByVal UserIndex As Integer)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
Dim nDamage      As Integer
 
'Calculo el daño.
nDamage = CalcularDaño(UserIndex)
 
'Resto la defensa del bot.
nDamage = nDamage - (RandomNumber(IA_MINDEF, IA_MAXDEF))
 
'Aviso al usuario.
WriteConsoleMsg UserIndex, "Le has pegado ah " & ia_Bot(botIndex).Tag & " por " & nDamage, FontTypeNames.FONTTYPE_FIGHT
 
'Creo daño :)
ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_CreateDamage(ia_Bot(botIndex).Pos.X, ia_Bot(botIndex).Pos.Y, nDamage)

'Resto vida.
ia_Bot(botIndex).Vida = ia_Bot(botIndex).Vida - nDamage
 
'seteo el flag.
UserList(UserIndex).AtacoViajante = botIndex

'Murio?
If ia_Bot(botIndex).Vida <= 0 Then
    'Era viajante?
    If ia_Bot(botIndex).Viajante Then
       'Reset el flag.
       UserList(UserIndex).AtacoViajante = 0
    End If
    ia_EraseChar botIndex, True
End If
 
End Sub

Sub ia_SendToBotArea(ByVal botIndex As Byte, ByVal PackData As String)

' @designer     :  maTih.-
' @date         :  2012/03/13
' @                Envia paquetes al area de un bot.

'Nueva versión del sub, más simple y diría que más práctica : P

With ia_Bot(botIndex)
    Call modSendData.SendToAreaByPos(.Pos.map, .Pos.X, .Pos.Y, PackData)
End With

End Sub

Sub ia_TirarInventario(ByVal botIndex As Byte)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Pincha el inventario de un bot.

Dim loopX   As Long
Dim iObjs() As Integer
Dim iObj    As Obj
Dim tmpPos  As WorldPos

'Arma array de objetos
ia_ArrayObjetos iObjs, botIndex

For loopX = 1 To UBound(iObjs())

    'Crea el objeto.
    iObj.objIndex = iObjs(loopX)

    'Si el objIndex es >= 36 and <=30  , son pociones
    If iObjs(loopX) >= 36 And iObjs(loopX) <= 39 Then
       iObj.Amount = RandomNumber(1000, 1200)
    Else
       'No eran pociones, son flechas?
       If Not iObjs(loopX) <> 553 Then
          iObj.Amount = RandomNumber(500, 900)
       Else
          iObj.Amount = 1
       End If
    End If
    
    'Si eran pociones azules y el bot era caza..
    If iObj.Amount = 37 And ia_Bot(botIndex).Clase = eIAClase.Cazador Then iObj.Amount = 0
    
    'si hay objIndex.
    If iObj.objIndex Then
        'Busca un tile libre.
        Call Tilelibre(ia_Bot(botIndex).Pos, tmpPos, iObj, True, True)
    
        'Si encontró (raro que no encuentre)
        If tmpPos.X <> 0 And tmpPos.Y <> 0 Then
           'Crea el objeto
           MakeObj iObj, tmpPos.map, tmpPos.X, tmpPos.Y
        End If
    End If
    
Next loopX

'Ya tiro los objetos de su equipo, ahora , si era viajante, tira los que lukeo, si es que tiene.
If ia_Bot(botIndex).Viajante Then
   For loopX = 1 To IA_SLOTS
       With ia_Bot(botIndex).Inv(loopX)
            
            iObj.objIndex = .objIndex
            iObj.Amount = .Amount
            
            Call Tilelibre(ia_Bot(botIndex).Pos, tmpPos, iObj, True, True)
            
            'Si encontró posición.
            If tmpPos.X <> 0 And tmpPos.Y <> 0 Then
               MakeObj iObj, tmpPos.map, tmpPos.X, tmpPos.Y
            End If
       End With
   Next loopX
End If

End Sub

Sub ia_ArrayObjetos(ByRef arrayObjs() As Integer, ByVal botIndex As Byte)

' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Arma un array de objetos.

'Set primeras dimensiones. (potas,arma y casco)

ReDim arrayObjs(1 To 4) As Integer

'Pociones.
arrayObjs(1) = 38
arrayObjs(2) = 37

'Arma
arrayObjs(3) = ia_ArmaByClase(botIndex)

'Casco
arrayObjs(4) = ia_CascoByClase(botIndex)

'Si no es mago, tiene escudo y dopas.
If ia_Bot(botIndex).Clase <> eIAClase.Mago Then
   'redim
   ReDim Preserve arrayObjs(1 To 7) As Integer
   arrayObjs(5) = ia_EscudoByClase(botIndex)
   arrayObjs(6) = 36
   arrayObjs(7) = 39
End If

'Si es caza, tira flechas.
'No sabemos el ultimo elemento que tenemos!! no jugarsela a tirar 5.

If ia_Bot(botIndex).Clase = eIAClase.Cazador Then
   ReDim Preserve arrayObjs(1 To UBound(arrayObjs()) + 1) As Integer
   arrayObjs(UBound(arrayObjs())) = 553
End If

End Sub

Sub ia_EraseChar(ByVal botIndex As Byte, Optional ByVal killedbyUSER As Boolean = False)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @note         :  Borra el char y los datos del bot.
 
With ia_Bot(botIndex)
    'Borro el char.
    ia_SendToBotArea botIndex, PrepareMessageCharacterRemove(.Char.CharIndex)
    
    'Borro el botIndex
    MapData(.Pos.map, .Pos.X, .Pos.Y).botIndex = 0
    
    Dim dummyPos    As WorldPos
    
    .ViajanteAntes = dummyPos
    
    'Mató un usuario? pincha inventario!
    If killedbyUSER Then
       ia_TirarInventario botIndex
    End If
    
    'Reset char,
    With .Char
         .body = 0
         .CascoAnim = 0
         .FX = 0
         .loops = 0
         .Head = 0
         .heading = 0
         .ShieldAnim = 0
         .WeaponAnim = 0
    End With
    
    'Reset STATS
    .Vida = 0
    .Mana = 0
    
    'Reset pos.
    With .Pos
         .map = 0
         .X = 0
         .Y = 0
    End With
    
    'Reset flags.
    .Invocado = False
    .Paralizado = False
   
    'Reset intervalos.
    With .Intervalos
         .MoveCharCount = 0
         .SpellCount = 0
         .UseItemCount = 0
         .ParalizisCount = 0
    End With
    
    'Reset viajante flag.
    .Viajante = False
    .ViajanteUser = 0
    
    'Resta el contador
    NumInvocados = NumInvocados - 1
    
End With
 
End Sub
 
Sub ia_CheckInts(ByVal botIndex As Byte)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
With ia_Bot(botIndex).Intervalos
     
    If .ArrowCount > 0 Then .ArrowCount = .ArrowCount - 1
    If .MoveCharCount > 0 Then .MoveCharCount = .MoveCharCount - 1
    If .SpellCount > 0 Then .SpellCount = .SpellCount - 1
    If .UseItemCount > 0 Then .UseItemCount = .UseItemCount - 1
    If .ParalizisCount > 0 Then .ParalizisCount = .ParalizisCount - 1
    If .HitCount > 0 Then .HitCount = .HitCount - 1
    If .ChatCount > 0 Then .ChatCount = .ChatCount - 1
    
End With
 
End Sub

Function ia_FindTarget(Pos As WorldPos, Optional ByVal esPk As Boolean = False) As Integer

' @designer     :  maTih.-
' @date         :  2012/03/13
' @note         :  Busca alguien a quien cagar a trompadas!!

Dim loopX       As Long         '< Bucle del tileX.
Dim loopY       As Long         '< Bucle del tileY.
Dim tmpIndex    As Integer

For loopY = Pos.Y - (MinYBorder + 1) To Pos.Y + (MinYBorder - 1)
        For loopX = Pos.X - (MinXBorder + 1) To Pos.X + (MinXBorder - 1)
            'Hay usuario?
            If MapData(Pos.map, loopX, loopY).UserIndex > 0 Then
               'No está muerto
               If UserList(MapData(Pos.map, loopX, loopY).UserIndex).flags.Muerto = 0 Then
                  'Es ciuda el bot y el usuario?
                  If Not esPk Or Server_Info.DeathMathc Then
                     'Devuelve.
                     ia_FindTarget = MapData(Pos.map, loopX, loopY).UserIndex
                  Else
                     tmpIndex = MapData(Pos.map, loopX, loopY).UserIndex
                     If Not esPk And criminal(tmpIndex) Then
                         ia_FindTarget = tmpIndex
                     Else
                        If esPk And Not criminal(tmpIndex) Then
                           ia_FindTarget = tmpIndex
                        End If
                    End If
                 End If
                  Exit Function
               End If
            End If
        Next loopX
Next loopY

ia_FindTarget = 0
End Function

Function IA_GetNextSlot() As Byte

' @ Devuelve un slot para bots.

Dim loopX   As Long

For loopX = 1 To MAX_BOTS
    If Not ia_Bot(loopX).Invocado Then
       IA_GetNextSlot = CByte(loopX)
       Exit Function
    End If
Next loopX

IA_GetNextSlot = 0
End Function

#End If


