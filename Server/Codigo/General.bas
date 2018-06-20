Attribute VB_Name = "General"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'maTih : Por ahora dejo esto acá hasta que se programe.
Public ArrayTorneoList(1 To 8) As String
'Damos a los NPCs el mismo rango de visiï¿½n que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6


Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************

Dim CuerpoDesnudo As Integer

With UserList(UserIndex)
    Select Case .Genero
        Case eGenero.Hombre
            Select Case .Raza
                Case eRaza.Humano
                    CuerpoDesnudo = 21
                Case eRaza.Drow
                    CuerpoDesnudo = 32
                Case eRaza.Elfo
                    CuerpoDesnudo = 210
                Case eRaza.Gnomo
                    CuerpoDesnudo = 222
                Case eRaza.Enano
                    CuerpoDesnudo = 53
            End Select
        Case eGenero.Mujer
            Select Case .Raza
                Case eRaza.Humano
                    CuerpoDesnudo = 39
                Case eRaza.Drow
                    CuerpoDesnudo = 40
                Case eRaza.Elfo
                    CuerpoDesnudo = 259
                Case eRaza.Gnomo
                    CuerpoDesnudo = 260
                Case eRaza.Enano
                    CuerpoDesnudo = 60
            End Select
    End Select
    
    If Mimetizado Then
        .CharMimetizado.body = CuerpoDesnudo
    Else
        .Char.body = CuerpoDesnudo
    End If
    
    .flags.Desnudo = 1
End With

End Sub



Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
'***************************************************

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)
    End If

End Sub


Function HayAgua(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        With MapData(map, X, Y)
            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
            (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
            (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
               .Graphic(2) = 0 Then
                    HayAgua = True
            Else
                    HayAgua = False
            End If
        End With
    Else
      HayAgua = False
    End If

End Function

Private Function HayLava(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
    If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(map, X, Y).Graphic(1) >= 5837 And MapData(map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
      HayLava = False
    End If

End Function


Sub LimpiarMundo()
'***************************************************
'Author: Unknow
'Last Modification: 04/15/2008
'01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
'04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
'***************************************************
On Error GoTo Errhandler

    Dim i As Integer
    Dim d As New cGarbage
    
    For i = TrashCollector.Count To 1 Step -1
        Set d = TrashCollector(i)
        Call EraseObj(1, d.map, d.X, d.Y)
        Call TrashCollector.Remove(i)
        Set d = Nothing
    Next i
    
    Call SecurityIp.IpSecurityMantenimientoLista
    
    Exit Sub

Errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.Description)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

#If UsarQueSocket = 0 Then

    Obj.AddressFamily = AF_INET
    Obj.Protocol = IPPROTO_IP
    Obj.SocketType = SOCK_STREAM
    Obj.Binary = False
    Obj.Blocking = False
    Obj.BufferSize = 1024
    Obj.LocalPort = Port
    Obj.backlog = 5
    Obj.listen

#End If

End Sub

Sub Main()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    Dim f As Date
    NumUsers = 1
    ChDir App.Path
    ChDrive App.Path

    '*** IA ****
    
    #If ConBots Then
        mod_IA.ia_Spells
    #End If
    
    Set Memory = New clsMemoryManage
    Memory.Initialize
    
    Call LoadMotd
    Call BanIpCargar
    
    Prision.map = 66
    Libertad.map = 66
    
    Prision.X = 75
    Prision.Y = 47
    Libertad.X = 75
    Libertad.Y = 65
    
    'Genera el array de gms.
    Call mod_DunkanGMs.GenerarArray
    
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
    
    Call mod_DunkanRankings.CargarTodos
    
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    Dim i       As Long
    
    Call LoadMapData
    
    mod_DunkanGeneral.Load_MapInfo
    
    For i = 1 To NUMCLASES
        Call mod_DunkanCs.Initialize(i)
    Next i
    
    frmCargando.Show
    
    'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\"
    CharPath = App.Path & "\Charfile\"
    AccountPath = App.Path & "\Accounts\"
    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    DoEvents
    
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."

    Call CargarSpawnList
    Call CargarForbidenWords
    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
    MaxUsers = 0
    Call LoadSini
    
    '*************************************************
    
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    'Call LoadOBJData

    Call mod_DunkanGeneral.Setear_Objs
        
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call mod_DunkanGeneral.Setear_Hechizos
    
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance    '4/01/08 Pablo ToxicWaste
 

    Call generateMatrix(MATRIX_INITIAL_MAP)
    
    'Comentado porque hay worldsave en ese mapa!
    'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Dim loopC As Integer
    
    frmMain.Internet_Protocol = frmMain.Inet1.OpenURL("http://ip1.dynupdate.no-ip.com:8245/")
    
    'Resetea las conexiones de los usuarios
    For loopC = 1 To MaxUsers
        UserList(loopC).ConnID = -1
        UserList(loopC).ConnIDValida = False
        Set UserList(loopC).incomingData = New clsByteQueue
        Set UserList(loopC).outgoingData = New clsByteQueue
    Next loopC
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    With frmMain
        .GameTimer.Enabled = True
        .Auditoria.Enabled = True
        
#If SeguridadAlkon Then
        .securityTimer.Enabled = True
#End If
    End With
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Configuracion de los sockets
    
    Call SecurityIp.InitIpTables(1000)
    
#If UsarQueSocket = 1 Then
    
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
    
#ElseIf UsarQueSocket = 0 Then
    
    frmCargando.Label1(2).Caption = "Configurando Sockets"
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Binary = False
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then
    
    frmMain.Serv.Iniciar Puerto
    
#ElseIf UsarQueSocket = 3 Then
    
    frmMain.TCPServ.Encolar True
    frmMain.TCPServ.IniciarTabla 1009
    frmMain.TCPServ.SetQueueLim 51200
    frmMain.TCPServ.Iniciar Puerto
    
#End If
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Unload frmCargando
    
    'Log
    Dim N As Integer
    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
    Call InicializaEstadisticas

    frmMain.ESPERANDO_MAPS = True

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

    FileExist = LenB(dir$(file, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************

    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function MapaValido(ByVal map As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    MapaValido = map >= 1 And map <= NumMaps
End Function

Sub MostrarNumUsers()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    frmMain.lblUser.Caption = "Número de usuarios jugando: " & NumUsers

    frmMain.EnviarNumUsers

End Sub


Public Sub LogCriticEvent(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile

Exit Sub

Errhandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub


Public Sub LogError(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Public Sub LogStatic(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogTarea(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile

Exit Sub

Errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************ç

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Public Sub LogAsesinato(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub
Public Sub LogHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Public Sub LogCheating(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Arg As String
    Dim i As Integer
    
    
    For i = 1 To 33
    
    Arg = ReadField(i, cad, 44)
    
    If LenB(Arg) = 0 Then Exit Function
    
    Next i
    
    ValidInputNP = True

End Function


Sub Restart()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
    Dim loopC As Long
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

    For loopC = 1 To MaxUsers
        Call CloseSocket(loopC)
    Next
    
    
    For loopC = 1 To UBound(UserList())
        Set UserList(loopC).incomingData = Nothing
        Set UserList(loopC).outgoingData = Nothing
    Next loopC
    
    ReDim UserList(1 To MaxUsers) As User
    
    For loopC = 1 To MaxUsers
        UserList(loopC).ConnID = -1
        UserList(loopC).ConnIDValida = False
        Set UserList(loopC).incomingData = New clsByteQueue
        Set UserList(loopC).outgoingData = New clsByteQueue
    Next loopC
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    
    
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos

#If UsarQueSocket = 0 Then

    '*****************Setup socket
    frmMain.Socket1.AddressFamily = AF_INET
    frmMain.Socket1.Protocol = IPPROTO_IP
    frmMain.Socket1.SocketType = SOCK_STREAM
    frmMain.Socket1.Binary = False
    frmMain.Socket1.Blocking = False
    frmMain.Socket1.BufferSize = 1024
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    'Escucha
    frmMain.Socket1.LocalPort = Val(Puerto)
    frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
    'Log it
    Dim N As Integer
    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " servidor reiniciado."
    Close #N
    
    'Ocultar
    
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If

  
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/11/2009
'15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
'23/11/2009: ZaMa - Optimizacion de codigo.
'**************************************************************

    With UserList(UserIndex)
        If MapInfo(.Pos.map).Zona <> "DUNGEON" Then
            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 1 And _
               MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 2 And _
               MapData(.Pos.map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = True
        Else
            Intemperie = False
        End If
    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'Makes an admin invisible o visible.
'13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'***************************************************
    
    With UserList(UserIndex)
        If .flags.AdminInvisible = 0 Then
            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False
            End If
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            .Char.body = 0
            .Char.Head = 0
            
            ' Solo el admin sabe que se hace invi
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True, 0))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            ' Solo el admin sabe que se hace visible
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.heading, _
            .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim))
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False, 0))
             
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y, True)
        End If
    End With
    
End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        .Counters.IdleCount = 0
        
        Dim Suerte As Integer
        Dim res As Integer
        Dim cant As Integer
        Dim MeditarSkill As Byte
    
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual As Long
        TActual = GetTickCount() And &H7FFFFFFF
        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub
        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True
        End If
            
        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(UserIndex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, mod_DunkanProtocol.Send_CreateMeditation(.Char.CharIndex, 0, 0))
            Exit Sub
        End If
        
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 And MeditarSkill >= -1 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 And MeditarSkill >= 11 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 And MeditarSkill >= 21 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 And MeditarSkill >= 31 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 And MeditarSkill >= 41 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 And MeditarSkill >= 51 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 And MeditarSkill >= 61 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 And MeditarSkill >= 71 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 And MeditarSkill >= 81 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 And MeditarSkill >= 91 Then
            Suerte = 7
        ElseIf MeditarSkill = 100 Then
            Suerte = 5
        End If
        res = RandomNumber(1, Suerte)
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)
            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant
            If .Stats.MinMAN > .Stats.MaxMAN Then _
                .Stats.MinMAN = .Stats.MaxMAN
            
            If Not .flags.UltimoMensaje = 22 Then
                Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de maná!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 22
            End If
            
            Call WriteUpdateMana(UserIndex)

        Else

        End If
    End With
End Sub

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 11/19/2009
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'11/19/2009: Pato - Ahora el bandido se oculta la mitad del tiempo de las demás clases.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
'********************************************************
On Error GoTo Errhandler
    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        If .Counters.TiempoOculto <= 0 Then
            
                .Counters.TiempoOculto = IntervaloOculto
            If .Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    Exit Sub
                End If
            End If
            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
            End If
        End If
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub


Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(UserIndex).Clase
    Case eClass.Assasin
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
    Case eClass.Cleric, eClass.Paladin
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
    Case eClass.Bard
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).Clase = eClass.Assasin Then
            Daño = Round(Daño * 1.4, 0)
        Else
            Daño = Round(Daño * 1.5, 0)
        End If
        
        UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - Daño
        
        'Envia daño.
        Call mod_DunkanGeneral.Enviar_DañoAUsuario(VictimUserIndex, Daño)
        
        Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Int(Daño * 2)
        
        'Daño del npc.
        Call mod_DunkanGeneral.Enviar_DañoANpc(VictimNpcIndex, Daño)
        
        Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(Daño * 2), FontTypeNames.FONTTYPE_FIGHT)

    End If
    
Else
    Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)

End If

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(UserIndex).Clase <> eClass.Bandit Then Exit Sub
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)

If RandomNumber(0, 100) < Suerte Then
    Daño = Int(Daño * 0.75)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - Daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & Daño & ".", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & Daño & ".", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & Daño & ".", FontTypeNames.FONTTYPE_FIGHT)

    End If
End If

End Sub

Public Sub DoAcuchillar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 12/01/2010
'***************************************************

    Dim CuchillaObjIndex    As Integer
    
    If UserList(UserIndex).Clase <> eClass.Pirat Then Exit Sub
    If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub

    If RandomNumber(0, 100) < PROB_ACUCHILLAR Then
        Daño = Int(Daño * DAÑO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - Daño
            
            'Envia flechas, va, cuchillas : D - maTih.-
            CuchillaObjIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
            
            'Tiene arma?
            If CuchillaObjIndex <> 0 Then
                'Acuchilla? seguro q si
                If ObjData(CuchillaObjIndex).Acuchilla = 1 Then
                   'Envia..
                   Call mod_DunkanGeneral.Enviar_FlechaAUser(UserList(UserIndex).Char.CharIndex, VictimUserIndex, ObjData(CuchillaObjIndex).GrhIndex)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & UserList(VictimUserIndex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha acuchillado por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        Else
            'Envia flechas, va, cuchillas : D - maTih.-
            CuchillaObjIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
            
            'Tiene arma?
            If CuchillaObjIndex <> 0 Then
                'Acuchilla? seguro q si
                If ObjData(CuchillaObjIndex).Acuchilla = 1 Then
                   'Envia..
                   Call mod_DunkanGeneral.Enviar_FlechaANpc(UserList(UserIndex).Char.CharIndex, VictimNpcIndex, ObjData(CuchillaObjIndex).GrhIndex)
                End If
            End If
            
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Daño
            Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & Daño, FontTypeNames.FONTTYPE_FIGHT)

        End If
    End If
    
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
'13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
'***************************************************

On Error GoTo Errhandler

    Dim Suerte As Double
    Dim res As Integer
    Dim Skill As Integer
    
    With UserList(UserIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            .Counters.TiempoOculto = Suerte
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
            ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                    NingunEscudo, NingunCasco)
            End If

        Else
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4
            End If
            '[/CDT]

        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")


End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    If .Faccion.ArmadaReal = 1 Then
                        .Char.body = iFragataReal
                    ElseIf .Faccion.FuerzasCaos = 1 Then
                        .Char.body = iFragataCaos
                    Else
                        Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
                        If criminal(UserIndex) Then
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
                        Else
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
                        End If
                    End If
                Else
                    .Char.body = iFragataFantasmal
                End If
                
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
            
            With .Char
                Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = RandomNumber(-100, 100) ' Invi variable :D
            .flags.invisible = 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                Call SetInvisible(UserIndex, .Char.CharIndex, False)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            End If
        End If
    End With

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
        End If
    End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else
            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)
            End If
        
        End If
    End With

End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If .Counters.Paralisis > 0 Then
            .Counters.Paralisis = .Counters.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
            '.Flags.AdministrativeParalisis = 0
            Call WriteParalizeOK(UserIndex)
        End If
    End With

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
        
        Dim massta As Integer
        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
            Else
                EnviarStats = True
                .Counters.STACounter = 0
                If .flags.Desnudo Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
               
                massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
                .Stats.MinSta = .Stats.MinSta + massta
                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta
                End If
            End If
        End If
    End With
    
End Sub

Public Sub CargaNpcsDat()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
End Sub

Sub PasarSegundo()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            'Tiempo auto-respawn
            With UserList(i)
                 If .Auto_Respawn <> 0 Then
                    .Auto_Respawn = .Auto_Respawn - 1
                    'Llega a0
                    If (Not .Auto_Respawn <> 0) Then
                        RevivirUsuario i
                        .Stats.MinHp = .Stats.MaxHp
                        WriteUpdateHP i
                        WarpUserChar i, .Pos.map, RandomNumber(25, 75), RandomNumber(25, 75), False
                        .Auto_Respawn = 0
                    End If
                 End If
            End With
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por jugar Argentum Online", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i)
                    Call FlushBuffer(i)
                    
                    Call CloseSocket(i)
                End If
            End If
        End If
    Next i
Exit Sub

Errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.Description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'WorldSave
    Call ES.DoBackUp

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Ta As Long
    Ta = GetTickCount() And &H7FFFFFFF
    
    Call EstadisticasWeb.Inicializa(frmMain.hWnd)
    Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim loopC As Long
    
    ' Free all NPC indexes
    For loopC = 1 To MAXNPCS
        Npclist(loopC).flags.NPCActive = False
    Next loopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub
