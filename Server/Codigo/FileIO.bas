Attribute VB_Name = "ES"
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



Public Sub CargarSpawnList()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, loopC As Integer
    N = Val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For loopC = 1 To N
        SpawnList(loopC).NpcIndex = Val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopC))
        SpawnList(loopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopC)
    Next loopC
    
End Sub

Function EsAdmin(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))
    
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsAdmin = True
            Exit Function
        End If
    Next WizNum
    EsAdmin = False

End Function

Function EsDios(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsDios = True
            Exit Function
        End If
    Next WizNum
    EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsSemiDios = True
            Exit Function
        End If
    Next WizNum
    EsSemiDios = False

End Function

Function EsConsejero(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsConsejero = True
            Exit Function
        End If
    Next WizNum
    EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsRolesMaster = True
            Exit Function
        End If
    Next WizNum
    EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, cad As String, Tam As Long
    N = FreeFile(1)
    Open Name For Input As #N
    Tam = 0
    Do While Not EOF(N)
        Tam = Tam + 1
        Line Input #N, cad
    Loop
    Close N
    TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim N As Integer, i As Integer
    N = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #N
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #N, ForbidenNames(i)
    Next i
    
    Close N

End Sub

Public Sub CargarHechizos()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    
    Dim Hechizo As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = Val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .tipo = Val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = Val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = Val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            ' maTih.-  *   Unificados FXGrh y ParticleID aca.
            .EffectIndex = Val(Leer.GetValue("Hechizo" & Hechizo, "EffectIndex"))
            
            'Si no tiene EI , por defecto dejamos el FX.
            If Not .EffectIndex <> 0 Then .EffectIndex = .FXgrh
            
            .isParticle = Val(Leer.GetValue("Hechizo" & Hechizo, "isParticle"))
            
            .loops = Val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
        '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            
            .Invisibilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = Val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            
            .CuraVeneno = Val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = Val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = Val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = Val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = Val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = Val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = Val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = Val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = Val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
        '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = Val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
            .NeedStaff = Val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
        End With
    Next Hechizo
    
    Set Leer = Nothing
    
    Exit Sub

Errhandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Sub LoadMotd()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    MaxLines = Val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)
    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i

End Sub

Public Sub DoBackUp()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    haciendoBK = True
    Dim i As Integer
    
    
    
    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind
    
    
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    
    Call LimpiarMundo
  
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)
    
    haciendoBK = False
    
    'Log
    On Error Resume Next
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile
End Sub

Public Sub GrabarMapa(ByVal map As Long, ByVal MAPFILE As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim TempInt As Integer
    Dim loopC As Long
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
            
    Put FreeFileMap, , MapInfo(map).MapVersion
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                If .Particle Then ByFlags = ByFlags Or 32
                If .WaterEffect Then ByFlags = ByFlags Or 64
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , .Graphic(1)
                
                For loopC = 2 To 4
                    If .Graphic(loopC) Then _
                        Put FreeFileMap, , .Graphic(loopC)
                Next loopC
                
                If .trigger Then _
                    Put FreeFileMap, , CInt(.trigger)
                    
                If .Particle Then _
                    Put FreeFileMap, , .Particle
                    
                If .WaterEffect Then _
                    Put FreeFileMap, , .WaterEffect
                
                '.inf file
                
                ByFlags = 0
                
                If .ObjInfo.objIndex > 0 Then
                   If ObjData(.ObjInfo.objIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.objIndex = 0
                        .ObjInfo.Amount = 0
                    End If
                End If
    
                If .TileExit.map Then ByFlags = ByFlags Or 1
                If .NpcIndex Then ByFlags = ByFlags Or 2
                If .ObjInfo.objIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If .TileExit.map Then
                    Put FreeFileInf, , .TileExit.map
                    Put FreeFileInf, , .TileExit.X
                    Put FreeFileInf, , .TileExit.Y
                End If
                
                If .NpcIndex Then _
                    Put FreeFileInf, , Npclist(.NpcIndex).Numero
                
                If .ObjInfo.objIndex Then
                    Put FreeFileInf, , .ObjInfo.objIndex
                    Put FreeFileInf, , .ObjInfo.Amount
                End If
            End With
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    With MapInfo(map)
    
        'write .dat file
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Name", .Name)
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "MusicNum", .Music)
        Call WriteVar(MAPFILE & ".dat", "mapa" & map, "MagiaSinefecto", .MagiaSinEfecto)
        Call WriteVar(MAPFILE & ".dat", "mapa" & map, "InviSinEfecto", .InviSinEfecto)
        Call WriteVar(MAPFILE & ".dat", "mapa" & map, "ResuSinEfecto", .ResuSinEfecto)
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "StartPos", .StartPos.map & "-" & .StartPos.X & "-" & .StartPos.Y)
        
    
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Terreno", .Terreno)
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Zona", .Zona)
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Restringir", .Restringir)
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "BackUp", str(.BackUp))
    
        If .Pk Then
            Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Pk", "0")
        Else
            Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Pk", "1")
        End If
    End With

End Sub
Sub LoadArmasHerreria()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, lc As Integer
    
    N = Val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmasHerrero(lc) = Val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadArmadurasHerreria()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, lc As Integer
    
    N = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmadurasHerrero(lc) = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadBalance()
'***************************************************
'Author: Unknown
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Agrego recompensas faccionarias.
'***************************************************

    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        With ModClase(i)
            .Evasion = Val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .DañoArmas = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
            .DañoProyectiles = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
            .DañoWrestling = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
            .Escudo = Val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
        End With
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        With ModRaza(i)
            .Fuerza = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = Val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
        End With
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = Val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = Val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = Val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = Val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))
   

    
End Sub

Sub LoadObjCarpintero()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, lc As Integer
    
    N = Val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To N) As Integer
    
    For lc = 1 To N
        ObjCarpintero(lc) = Val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub



Sub LoadOBJData()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo Errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    
    Dim Leer    As New clsIniReader
    Dim tmpPath As String
    
    tmpPath = DatPath & "Obj.dat"

    Call Leer.Initialize(tmpPath)
    
    'obtiene el numero de obj
    NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0
    
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    
    'Llena la lista
    For Object = 1 To NumObjDatas
        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = Val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = Val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex
            End If
            
            .OBJType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
                    
            'maTih.- objetos que buscan los bot
            .Valioso = Val(Leer.GetValue("OBJ" & Object, "BotBusca"))
                        
            .Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            Select Case .OBJType
                Case eOBJType.otArmadura
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    
                Case eOBJType.otESCUDO
                    .ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                Case eOBJType.otCASCO
                    .CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                Case eOBJType.otWeapon
                    .WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                    .Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .Minhit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .StaffPower = Val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                    .StaffDamageBonus = Val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                    .Refuerzo = Val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = Val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))

                Case eOBJType.otInstrumentos
                    .Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otBarcos
                    .MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .Minhit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                Case eOBJType.otFlechas
                    .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .Minhit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    
                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .Minhit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    
                Case eOBJType.otTeleport
                    .Radio = Val(Leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otMochilas
                    .MochilaType = Val(Leer.GetValue("OBJ" & Object, "MochilaType"))

            End Select
            
            .Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = Val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = Val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = Val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = Val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
            
            .Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))
            If .Cerrada = 1 Then
                .Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
                .clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
            End If
            
            'Puertas y llaves
            .clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = Val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
            .Guante = Val(Leer.GetValue("OBJ" & Object, "Guante"))
            
            .DefensaMagicaMax = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
            .SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then _
                .Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
                .MaderaElfica = Val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
            
            'Bebidas
            .MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            .Upgrade = Val(Leer.GetValue("OBJ" & Object, "Upgrade"))
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        End With
    Next Object
    
    
    Set Leer = Nothing

    
    Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 11/19/2009
'11/19/2009: Pato - Load the EluSkills and ExpSkills
'*************************************************
Dim loopC As Long

With UserList(UserIndex)
    With .Stats
        For loopC = 1 To NUMATRIBUTOS
            .UserAtributos(loopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & loopC))
            .UserAtributosBackUP(loopC) = .UserAtributos(loopC)
        Next loopC
        
        For loopC = 1 To NUMSKILLS
            .UserSkills(loopC) = CInt(UserFile.GetValue("SKILLS", "SK" & loopC))
            .EluSkills(loopC) = CInt(UserFile.GetValue("SKILLS", "ELUSK" & loopC))
            .ExpSkills(loopC) = CInt(UserFile.GetValue("SKILLS", "EXPSK" & loopC))
        Next loopC
        
        For loopC = 1 To MAXUSERHECHIZOS
            .UserHechizos(loopC) = CInt(UserFile.GetValue("Hechizos", "H" & loopC))
        Next loopC
        
        .GLD = CLng(UserFile.GetValue("STATS", "GLD"))
        .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
        .MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
        .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
        
        .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
        .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
        .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
        .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
        .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
        .Minhit = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
        .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
        .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
        
        .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
        .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
        
        .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
        
        .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
        .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
        .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
        
        
        .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
        .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))
    End With
    
    With .flags
        If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
            .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
        If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
            .Privilegios = .Privilegios Or PlayerType.ChaosCouncil
    End With
End With
End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex).Reputacion
        .AsesinoRep = Val(UserFile.GetValue("REP", "Asesino"))
        .BandidoRep = Val(UserFile.GetValue("REP", "Bandido"))
        .BurguesRep = Val(UserFile.GetValue("REP", "Burguesia"))
        .LadronesRep = Val(UserFile.GetValue("REP", "Ladrones"))
        .NobleRep = Val(UserFile.GetValue("REP", "Nobles"))
        .PlebeRep = Val(UserFile.GetValue("REP", "Plebe"))
        .Promedio = Val(UserFile.GetValue("REP", "Promedio"))
    End With
    
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
    Dim loopC As Long
    Dim ln As String
    
    With UserList(UserIndex)
        With .Faccion
            .ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
            .FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
            .CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
            .CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
            .RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
            .RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
            .RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
            .RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
            .RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
            .RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
            .Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
            .NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
            .FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
            .MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
            .NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))
        End With
        
        With .flags
            .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
            .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            
            .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
            .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
            .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
            .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
            .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
            .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
            
            'Matrix
            .lastMap = CInt(UserFile.GetValue("FLAGS", "LastMap"))
        End With
        
        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado
        End If
        
        
        .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
        .Counters.AsignedSkills = CByte(Val(UserFile.GetValue("COUNTERS", "SkillsAsignados")))
        
       ' .email = UserFile.GetValue("CONTACTO", "Email")
        
               
        .Genero = UserFile.GetValue("INIT", "Genero")
        .Clase = UserFile.GetValue("INIT", "Clase")
        .Raza = UserFile.GetValue("INIT", "Raza")
        .Hogar = UserFile.GetValue("INIT", "Hogar")
        .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
        
        With .OrigChar
            .Head = CInt(UserFile.GetValue("INIT", "Head"))
            .body = CInt(UserFile.GetValue("INIT", "Body"))
            .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
            .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
            .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
            
            .heading = eHeading.SOUTH
        End With
        
        #If ConUpTime Then
            .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If
        
        If .flags.Muerto = 0 Then
            .Char = .OrigChar
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
        End If
        
        
        .desc = UserFile.GetValue("INIT", "Desc")
        
        .Pos.map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
        .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
        .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
        'Lista de objetos del banco
        For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = UserFile.GetValue("BancoInventory", "Obj" & loopC)
            .BancoInvent.Object(loopC).objIndex = CInt(ReadField(1, ln, 45))
            .BancoInvent.Object(loopC).Amount = CInt(ReadField(2, ln, 45))
        Next loopC
        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************
        
        
        'Lista de objetos
        For loopC = 1 To MAX_INVENTORY_SLOTS
            ln = UserFile.GetValue("Inventory", "Obj" & loopC)
            .Invent.Object(loopC).objIndex = CInt(ReadField(1, ln, 45))
            .Invent.Object(loopC).Amount = CInt(ReadField(2, ln, 45))
            .Invent.Object(loopC).Equipped = CByte(ReadField(3, ln, 45))
        Next loopC
        
        'Obtiene el indice-objeto del arma
        .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).objIndex
        End If
        
        'Obtiene el indice-objeto del armadura
        .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).objIndex
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1
        End If
        
        'Obtiene el indice-objeto del escudo
        .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).objIndex
        End If
        
        'Obtiene el indice-objeto del casco
        .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).objIndex
        End If
        
        'Obtiene el indice-objeto barco
        .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).objIndex
        End If
        
        'Obtiene el indice-objeto municion
        .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).objIndex
        End If
        
        '[Alejo]
        'Obtiene el indice-objeto anilo
        .Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
        If .Invent.AnilloEqpSlot > 0 Then
            .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).objIndex
        End If
        
        .Invent.MochilaEqpSlot = CByte(UserFile.GetValue("Inventory", "MochilaSlot"))
        If .Invent.MochilaEqpSlot > 0 Then
            .Invent.MochilaEqpObjIndex = .Invent.Object(.Invent.MochilaEqpSlot).objIndex
        End If
        
        .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
        Dim NpcIndex As Integer
        For loopC = 1 To MAXMASCOTAS
            .MascotasType(loopC) = Val(UserFile.GetValue("MASCOTAS", "MAS" & loopC))
        Next loopC
        
        ln = UserFile.GetValue("Guild", "GUILDINDEX")
        If IsNumeric(ln) Then
            .GuildIndex = CInt(ln)
        Else
            .GuildIndex = 0
        End If
    End With

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
    
    Dim map As Integer
    Dim TempInt As Integer
    Dim tFileName As String
    Dim npcfile As String
    
    On Error GoTo man
        
        NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        
       
        Call InitAreas
        
        frmCargando.cargar.min = 0
        frmCargando.cargar.max = NumMaps
        frmCargando.cargar.Value = 0
        
        MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
        
       
        For map = 1 To NumMaps
            If Val(GetVar(App.Path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
                tFileName = App.Path & "\WorldBackUp\Mapa" & map
                
                If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                    tFileName = App.Path & MapPath & "Mapa" & map
                End If
            Else
                tFileName = App.Path & MapPath & "Mapa" & map
            End If
            
            Call CargarMapa(map, tFileName)
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            DoEvents
        Next map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub



Sub LoadMapData()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
    
    Dim map As Integer
    Dim TempInt As Integer
    Dim tFileName As String
    Dim npcfile As String
    
    On Error GoTo man
        
        NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        Call InitAreas
        
        frmCargando.cargar.min = 0
        frmCargando.cargar.max = NumMaps
        frmCargando.cargar.Value = 0
        
        MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
        
        Dim MapReader   As New clsIniReader
         
        For map = 1 To NumMaps
            
            tFileName = App.Path & MapPath & "Mapa" & map
            Call CargarMapa(map, tFileName)
            
            'Carga las rutas.
            'Existe?
            If FileExist(App.Path & "\Maps\Mapa" & map & "Rutas.txt") Then
                MapReader.Initialize App.Path & "\Maps\Mapa" & map & "Rutas.txt"
                Call mod_DunkanGeneral.Cargar_RutasMap(MapReader, map)
            End If
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            DoEvents
        Next map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer

    FreeFileMap = FreeFile

    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1

    FreeFileInf = FreeFile

    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(map).MapVersion
    Get #FreeFileMap, , MiCabecera
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt

    'inf Header
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(map, X, Y)

                '.dat file
                Get FreeFileMap, , ByFlags

                If ByFlags And 1 Then
                    .Blocked = 1
                End If

                Get FreeFileMap, , .Graphic(1)

                'Layer 2 used?
                If ByFlags And 2 Then Get FreeFileMap, , .Graphic(2)

                'Layer 3 used?
                If ByFlags And 4 Then Get FreeFileMap, , .Graphic(3)

                'Layer 4 used?
                If ByFlags And 8 Then Get FreeFileMap, , .Graphic(4)

                'Trigger used?
                If ByFlags And 16 Then
                    'Enums are 4 byte long in VB, so we make sure we only read 2
                    Get FreeFileMap, , TempInt
                    .trigger = TempInt
                End If

                Get FreeFileInf, , ByFlags

                If ByFlags And 1 Then
                    Get FreeFileInf, , .TileExit.map
                    Get FreeFileInf, , .TileExit.X
                    Get FreeFileInf, , .TileExit.Y
                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                    Get FreeFileInf, , .NpcIndex


                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    Get FreeFileInf, , .ObjInfo.objIndex
                    Get FreeFileInf, , .ObjInfo.Amount
                End If
            End With
        Next X
    Next Y


    Close FreeFileMap
    Close FreeFileInf

    With MapInfo(map)
        .Name = GetVar(MAPFl & ".dat", "Mapa" & map, "Name")
        .Music = GetVar(MAPFl & ".dat", "Mapa" & map, "MusicNum")
        .StartPos.map = Val(ReadField(1, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
        .StartPos.X = Val(ReadField(2, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
        .StartPos.Y = Val(ReadField(3, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
        .MagiaSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "MagiaSinEfecto"))
        .InviSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "InviSinEfecto"))
        .ResuSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "ResuSinEfecto"))
        .NoEncriptarMP = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "NoEncriptarMP"))

        .RoboNpcsPermitido = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "RoboNpcsPermitido"))
        
        If Val(GetVar(MAPFl & ".dat", "Mapa" & map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False
        End If

        
        .Terreno = GetVar(MAPFl & ".dat", "Mapa" & map, "Terreno")
        .Zona = GetVar(MAPFl & ".dat", "Mapa" & map, "Zona")
        .Restringir = GetVar(MAPFl & ".dat", "Mapa" & map, "Restringir")
        .BackUp = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "BACKUP"))
    End With
   
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " - Pos: " & X & "," & Y & "." & Err.Description)
End Sub

Sub LoadSini()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Temporal As Long
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
    BootDelBackUp = Val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))
    
    'Misc
    #If SeguridadAlkon Then
    
    Call SECURITY.SetServerIp(GetVar(IniPath & "Server.ini", "INIT", "ServerIp"))
    
    #End If
    
    
    Puerto = Val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))

    IdleLimit = Val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))

    'Intervalos
    SanaIntervaloSinDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
    StaminaIntervaloSinDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
    SanaIntervaloDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
    StaminaIntervaloDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
    IntervaloSed = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
    IntervaloHambre = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
    IntervaloVeneno = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
    IntervaloParalizado = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
    IntervaloInvisible = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
    IntervaloFrio = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
    IntervaloWavFx = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
    IntervaloInvocacion = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
    IntervaloParaConexion = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
    FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
    IntervaloAtacable = 60000 ' Cargar desde balance.dat
    IntervaloOwnedNpc = 18000 ' Cargar desde balance.dat
    
    IntervaloUserPuedeCastear = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
    IntervaloUserPuedeTrabajar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
    IntervaloUserPuedeAtacar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloGolpeUsar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))
    
    
    IntervaloCerrarConexion = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&

    'Max users
    Temporal = Val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User
    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agregó en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    
    Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
    Nix.map = GetVar(DatPath & "Ciudades.dat", "Nix", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "Nix", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "Nix", "Y")
    
    Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
    Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
    Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
    Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
    Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    
    Ciudades(eCiudad.cUllathorpe) = Ullathorpe
    Ciudades(eCiudad.cNix) = Nix
    Ciudades(eCiudad.cBanderbill) = Banderbill
    Ciudades(eCiudad.cLindos) = Lindos
    Ciudades(eCiudad.cArghal) = Arghal
    
    Call MD5sCarga
    
    Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
    Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'Escribe VAR en un archivo
'***************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)
'*************************************************
'Author: Unknown
'Last modified: 12/01/2010 (ZaMa)
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'11/19/2009: Pato - Save the EluSkills and ExpSkills
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'*************************************************

On Error GoTo Errhandler

Dim OldUserHead As Long

With UserList(UserIndex)

    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If .Clase = 0 Or .Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
        Exit Sub
    End If
    
    
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
    
    If FileExist(UserFile, vbNormal) Then
        If .flags.Muerto = 1 Then
            OldUserHead = .Char.Head
            .Char.Head = GetVar(UserFile, "INIT", "Head")
        End If
    '       Kill UserFile
    End If
    
    Dim loopC As Integer
    
    
    Call WriteVar(UserFile, "FLAGS", "Muerto", CStr(.flags.Muerto))
    Call WriteVar(UserFile, "FLAGS", "Escondido", CStr(.flags.Escondido))
    Call WriteVar(UserFile, "FLAGS", "Hambre", CStr(.flags.Hambre))
    Call WriteVar(UserFile, "FLAGS", "Sed", CStr(.flags.Sed))
    Call WriteVar(UserFile, "FLAGS", "Desnudo", CStr(.flags.Desnudo))
    Call WriteVar(UserFile, "FLAGS", "Ban", CStr(.flags.Ban))
    Call WriteVar(UserFile, "FLAGS", "Navegando", CStr(.flags.Navegando))
    Call WriteVar(UserFile, "FLAGS", "Envenenado", CStr(.flags.Envenenado))
    Call WriteVar(UserFile, "FLAGS", "Paralizado", CStr(.flags.Paralizado))
    'Matrix
    Call WriteVar(UserFile, "FLAGS", "LastMap", CStr(.flags.lastMap))
    
    Call WriteVar(UserFile, "CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
    Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
    
    
    Call WriteVar(UserFile, "COUNTERS", "Pena", CStr(.Counters.Pena))
    Call WriteVar(UserFile, "COUNTERS", "SkillsAsignados", CStr(.Counters.AsignedSkills))
    
    Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", CStr(.Faccion.ArmadaReal))
    Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", CStr(.Faccion.FuerzasCaos))
    Call WriteVar(UserFile, "FACCIONES", "CiudMatados", CStr(.Faccion.CiudadanosMatados))
    Call WriteVar(UserFile, "FACCIONES", "CrimMatados", CStr(.Faccion.CriminalesMatados))
    Call WriteVar(UserFile, "FACCIONES", "rArCaos", CStr(.Faccion.RecibioArmaduraCaos))
    Call WriteVar(UserFile, "FACCIONES", "rArReal", CStr(.Faccion.RecibioArmaduraReal))
    Call WriteVar(UserFile, "FACCIONES", "rExCaos", CStr(.Faccion.RecibioExpInicialCaos))
    Call WriteVar(UserFile, "FACCIONES", "rExReal", CStr(.Faccion.RecibioExpInicialReal))
    Call WriteVar(UserFile, "FACCIONES", "recCaos", CStr(.Faccion.RecompensasCaos))
    Call WriteVar(UserFile, "FACCIONES", "recReal", CStr(.Faccion.RecompensasReal))
    Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", CStr(.Faccion.Reenlistadas))
    Call WriteVar(UserFile, "FACCIONES", "NivelIngreso", CStr(.Faccion.NivelIngreso))
    Call WriteVar(UserFile, "FACCIONES", "FechaIngreso", .Faccion.FechaIngreso)
    Call WriteVar(UserFile, "FACCIONES", "MatadosIngreso", CStr(.Faccion.MatadosIngreso))
    Call WriteVar(UserFile, "FACCIONES", "NextRecompensa", CStr(.Faccion.NextRecompensa))
    
    
    '¿Fueron modificados los atributos del usuario?
    If Not .flags.TomoPocion Then
        For loopC = 1 To UBound(.Stats.UserAtributos)
            Call WriteVar(UserFile, "ATRIBUTOS", "AT" & loopC, CStr(.Stats.UserAtributos(loopC)))
        Next loopC
    Else
        For loopC = 1 To UBound(.Stats.UserAtributos)
            '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
            Call WriteVar(UserFile, "ATRIBUTOS", "AT" & loopC, CStr(.Stats.UserAtributosBackUP(loopC)))
        Next loopC
    End If
    
    For loopC = 1 To UBound(.Stats.UserSkills)
        Call WriteVar(UserFile, "SKILLS", "SK" & loopC, CStr(.Stats.UserSkills(loopC)))
        Call WriteVar(UserFile, "SKILLS", "ELUSK" & loopC, CStr(.Stats.EluSkills(loopC)))
        Call WriteVar(UserFile, "SKILLS", "EXPSK" & loopC, CStr(.Stats.ExpSkills(loopC)))
    Next loopC
    
    
    'Call WriteVar(UserFile, "CONTACTO", "Email", .email)
    
    Call WriteVar(UserFile, "INIT", "Genero", .Genero)
    Call WriteVar(UserFile, "INIT", "Raza", .Raza)
    Call WriteVar(UserFile, "INIT", "Hogar", .Hogar)
    Call WriteVar(UserFile, "INIT", "Clase", .Clase)
    Call WriteVar(UserFile, "INIT", "Desc", .desc)
    
    Call WriteVar(UserFile, "INIT", "Heading", CStr(.Char.heading))
    
    Call WriteVar(UserFile, "INIT", "Head", CStr(.OrigChar.Head))
    
    If .flags.Muerto = 0 Then
        Call WriteVar(UserFile, "INIT", "Body", CStr(.Char.body))
    End If
    
    Call WriteVar(UserFile, "INIT", "Arma", CStr(.Char.WeaponAnim))
    Call WriteVar(UserFile, "INIT", "Escudo", CStr(.Char.ShieldAnim))
    Call WriteVar(UserFile, "INIT", "Casco", CStr(.Char.CascoAnim))
    
    #If ConUpTime Then
        Dim TempDate As Date
        TempDate = Now - .LogOnTime
        .LogOnTime = Now
        .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        .UpTime = .UpTime
        Call WriteVar(UserFile, "INIT", "UpTime", .UpTime)
    #End If
    
    'First time around?
    If GetVar(UserFile, "INIT", "LastIP1") = vbNullString Then
        Call WriteVar(UserFile, "INIT", "LastIP1", .ip & " - " & Date & ":" & Time)
    'Is it a different ip from last time?
    ElseIf .ip <> Left$(GetVar(UserFile, "INIT", "LastIP1"), InStr(1, GetVar(UserFile, "INIT", "LastIP1"), " ") - 1) Then
        Dim i As Integer
        For i = 5 To 2 Step -1
            Call WriteVar(UserFile, "INIT", "LastIP" & i, GetVar(UserFile, "INIT", "LastIP" & CStr(i - 1)))
        Next i
        Call WriteVar(UserFile, "INIT", "LastIP1", .ip & " - " & Date & ":" & Time)
    'Same ip, just update the date
    Else
        Call WriteVar(UserFile, "INIT", "LastIP1", .ip & " - " & Date & ":" & Time)
    End If
    
    
    
    Call WriteVar(UserFile, "INIT", "Position", .Pos.map & "-" & .Pos.X & "-" & .Pos.Y)
    
    
    Call WriteVar(UserFile, "STATS", "GLD", CStr(.Stats.GLD))
    Call WriteVar(UserFile, "STATS", "BANCO", CStr(.Stats.Banco))
    
    Call WriteVar(UserFile, "STATS", "MaxHP", CStr(.Stats.MaxHp))
    Call WriteVar(UserFile, "STATS", "MinHP", CStr(.Stats.MinHp))
    
    Call WriteVar(UserFile, "STATS", "MaxSTA", CStr(.Stats.MaxSta))
    Call WriteVar(UserFile, "STATS", "MinSTA", CStr(.Stats.MinSta))
    
    Call WriteVar(UserFile, "STATS", "MaxMAN", CStr(.Stats.MaxMAN))
    Call WriteVar(UserFile, "STATS", "MinMAN", CStr(.Stats.MinMAN))
    
    Call WriteVar(UserFile, "STATS", "MaxHIT", CStr(.Stats.MaxHIT))
    Call WriteVar(UserFile, "STATS", "MinHIT", CStr(.Stats.Minhit))
    
    Call WriteVar(UserFile, "STATS", "MaxAGU", CStr(.Stats.MaxAGU))
    Call WriteVar(UserFile, "STATS", "MinAGU", CStr(.Stats.MinAGU))
    
    Call WriteVar(UserFile, "STATS", "MaxHAM", CStr(.Stats.MaxHam))
    Call WriteVar(UserFile, "STATS", "MinHAM", CStr(.Stats.MinHam))
    
    Call WriteVar(UserFile, "STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
      
    Call WriteVar(UserFile, "STATS", "EXP", CStr(.Stats.Exp))
    Call WriteVar(UserFile, "STATS", "ELV", CStr(.Stats.ELV))
    
    
    Call WriteVar(UserFile, "STATS", "ELU", CStr(.Stats.ELU))
    Call WriteVar(UserFile, "MUERTES", "UserMuertes", CStr(.Stats.UsuariosMatados))
    'Call WriteVar(UserFile, "MUERTES", "CrimMuertes", CStr(.Stats.CriminalesMatados))
    Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
      
    '[KEVIN]----------------------------------------------------------------------------
    '*******************************************************************************************
    Call WriteVar(UserFile, "BancoInventory", "CantidadItems", Val(.BancoInvent.NroItems))
    Dim loopd As Integer
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, .BancoInvent.Object(loopd).objIndex & "-" & .BancoInvent.Object(loopd).Amount)
    Next loopd
    '*******************************************************************************************
    '[/KEVIN]-----------
      
    'Save Inv
    Call WriteVar(UserFile, "Inventory", "CantidadItems", Val(.Invent.NroItems))
    
    For loopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(UserFile, "Inventory", "Obj" & loopC, .Invent.Object(loopC).objIndex & "-" & .Invent.Object(loopC).Amount & "-" & .Invent.Object(loopC).Equipped)
    Next loopC
    
    Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
    Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
    Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
    Call WriteVar(UserFile, "Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
    Call WriteVar(UserFile, "Inventory", "MochilaSlot", CStr(.Invent.MochilaEqpSlot))
    '/Nacho
    
    Call WriteVar(UserFile, "Inventory", "AnilloSlot", CStr(.Invent.AnilloEqpSlot))
    
    'Reputacion
    Call WriteVar(UserFile, "REP", "Asesino", CStr(.Reputacion.AsesinoRep))
    Call WriteVar(UserFile, "REP", "Bandido", CStr(.Reputacion.BandidoRep))
    Call WriteVar(UserFile, "REP", "Burguesia", CStr(.Reputacion.BurguesRep))
    Call WriteVar(UserFile, "REP", "Ladrones", CStr(.Reputacion.LadronesRep))
    Call WriteVar(UserFile, "REP", "Nobles", CStr(.Reputacion.NobleRep))
    Call WriteVar(UserFile, "REP", "Plebe", CStr(.Reputacion.PlebeRep))
    
    Dim L As Long
    L = (-.Reputacion.AsesinoRep) + _
        (-.Reputacion.BandidoRep) + _
        .Reputacion.BurguesRep + _
        (-.Reputacion.LadronesRep) + _
        .Reputacion.NobleRep + _
        .Reputacion.PlebeRep
    L = L / 6
    Call WriteVar(UserFile, "REP", "Promedio", CStr(L))
    
    Dim cad As String
    
    For loopC = 1 To MAXUSERHECHIZOS
        cad = .Stats.UserHechizos(loopC)
        Call WriteVar(UserFile, "HECHIZOS", "H" & loopC, cad)
    Next
    
    Dim NroMascotas As Long
    NroMascotas = .NroMascotas
    
    For loopC = 1 To MAXMASCOTAS
        ' Mascota valida?
        If .MascotasIndex(loopC) > 0 Then
            ' Nos aseguramos que la criatura no fue invocada
            If Npclist(.MascotasIndex(loopC)).Contadores.TiempoExistencia = 0 Then
                cad = .MascotasType(loopC)
            Else 'Si fue invocada no la guardamos
                cad = "0"
                NroMascotas = NroMascotas - 1
            End If
            Call WriteVar(UserFile, "MASCOTAS", "MAS" & loopC, cad)
        Else
            cad = .MascotasType(loopC)
            Call WriteVar(UserFile, "MASCOTAS", "MAS" & loopC, cad)
        End If
    
    Next
    
    Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", CStr(NroMascotas))
    
    'Devuelve el head de muerto
    If .flags.Muerto = 1 Then
        .Char.Head = iCabezaMuerto
    End If
End With

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + _
            (-.BandidoRep) + _
            .BurguesRep + _
            (-.LadronesRep) + _
            .NobleRep + _
            .PlebeRep
        L = L / 6
        criminal = (L < 0)
    End With

End Function


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)
    
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Apuestas.Ganancias = Val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

Public Sub generateMatrix(ByVal Mapa As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Integer
Dim j As Integer
Dim X As Integer
Dim Y As Integer

ReDim distanceToCities(1 To NumMaps) As HomeDistance

For j = 1 To NUMCIUDADES
    For i = 1 To NumMaps
        distanceToCities(i).distanceToCity(j) = -1
    Next i
Next j

For j = 1 To NUMCIUDADES
    For i = 1 To 4
        Select Case i
            Case eHeading.NORTH
                Call setDistance(getLimit(Ciudades(j).map, eHeading.NORTH), j, i, 0, 1)
            Case eHeading.EAST
                Call setDistance(getLimit(Ciudades(j).map, eHeading.EAST), j, i, 1, 0)
            Case eHeading.SOUTH
                Call setDistance(getLimit(Ciudades(j).map, eHeading.SOUTH), j, i, 0, 1)
            Case eHeading.WEST
                Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, -1, 0)
            Case eHeading.NorthEast
                Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, 1, -1)
            Case eHeading.SouthEast
                Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, 1, 1)
            Case eHeading.NorthWest
                Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, -1, 1)
            Case eHeading.SouthWest
                Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, -1, -1)
        End Select
    Next i
Next j

End Sub

Public Sub setDistance(ByVal Mapa As Integer, ByVal city As Byte, ByVal side As Integer, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Integer
Dim lim As Integer

If Mapa <= 0 Or Mapa > NumMaps Then Exit Sub

If distanceToCities(Mapa).distanceToCity(city) >= 0 Then Exit Sub

If Mapa = Ciudades(city).map Then
    distanceToCities(Mapa).distanceToCity(city) = 0
Else
    distanceToCities(Mapa).distanceToCity(city) = Abs(X) + Abs(Y)
End If

For i = 1 To 4
    lim = getLimit(Mapa, i)
    If lim > 0 Then
        Select Case i
            Case eHeading.NORTH
                Call setDistance(lim, city, i, X, Y + 1)
            Case eHeading.EAST
                Call setDistance(lim, city, i, X + 1, Y)
            Case eHeading.SOUTH
                Call setDistance(lim, city, i, X, Y - 1)
            Case eHeading.WEST
                Call setDistance(lim, city, i, X - 1, Y)
            Case eHeading.NorthEast
                Call setDistance(lim, city, i, X + 1, Y - 1)
            Case eHeading.SouthEast
                Call setDistance(lim, city, i, X + 1, Y + 1)
            Case eHeading.NorthWest
                Call setDistance(lim, city, i, X - 1, Y + 1)
            Case eHeading.SouthWest
                Call setDistance(lim, city, i, X - 1, Y - 1)
        End Select
    End If
Next i
End Sub

Public Function getLimit(ByVal Mapa As Integer, ByVal side As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modification: 31/01/2010
'Retrieves the limit in the given side in the given map.
'TODO: This should be set in the .inf map file.
'***************************************************
Dim i, X, Y As Integer

If Mapa <= 0 Then Exit Function

For X = 15 To 87
    For Y = 0 To 3
        Select Case side
            Case eHeading.NORTH
                getLimit = MapData(Mapa, X, 7 + Y).TileExit.map
            Case eHeading.EAST
                getLimit = MapData(Mapa, 92 - Y, X).TileExit.map
            Case eHeading.SOUTH
                getLimit = MapData(Mapa, X, 94 - Y).TileExit.map
            Case eHeading.WEST
                getLimit = MapData(Mapa, 9 + Y, X).TileExit.map
        End Select
        If getLimit > 0 Then Exit Function
    Next Y
Next X
End Function
