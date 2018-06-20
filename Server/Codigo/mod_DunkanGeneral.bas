Attribute VB_Name = "mod_DunkanGeneral"
' @ Dise�ado & Implementado por maTih.-
' @ Dunkan AO Funciones/SubRutinas Generales. (HAY DE TODO)

Option Explicit

Public RutasFile    As String

Public Type TInfo
       Mapa          As WorldPos     '< Mapa del servidor.
       Resucitar     As Boolean      '< Vale resucitaR?
       Invisibilidad As Boolean      '< Vale invi?
       AutoRespawn   As Boolean      '< Auto resucitar?
       DeathMathc    As Boolean      '< Es deathmatch?
       TiempoRespawn As Integer      '< Tiempo de resu.
       Clase(1 To 8) As Boolean      '< Clases permitidas.
End Type

Public Server_Info   As TInfo

Sub Load_MapData(ByVal map_Server As Integer)

'
' @ Carga el mapData() - 1 solo mapa.

Server_Info.Mapa.map = map_Server

Call ModAreas.InitAreas

ReDim MapData(map_Server To map_Server, 1 To 100, 1 To 100) As MapBlock

Call ES.CargarMapa(map_Server, App.Path & "\Maps\Mapa" & CStr(map_Server))

End Sub

Sub Load_MapInfo()

'
' @ Carga el mapInfo()

Dim map_Path    As String
Dim num_Maps    As String
Dim Loop_Map    As Long

num_Maps = GetVar(App.Path & "\Dat\Map.dat", "INIT", "NumMaps")

num_Maps = 1

ReDim MapInfo(1 To Val(num_Maps)) As MapInfo

For Loop_Map = 1 To Val(num_Maps)
    map_Path = App.Path & "\Maps\Mapa" & CStr(Loop_Map) & ".dat"
    
    With MapInfo(Loop_Map)
         .Name = GetVar(map_Path, "Mapa" & CStr(Loop_Map), "Name")
         .Terreno = GetVar(map_Path, "Mapa" & CStr(Loop_Map), "TERRENO")
         .Zona = GetVar(map_Path, "Mapa" & CStr(Loop_Map), "ZONA")
         
         frmMain.c_Map.AddItem .Name
         
    End With
    
Next Loop_Map

End Sub

Public Sub Informar_Muerte(ByVal userIndex As Integer)

'
'  @ Checkea si termina la ronda

Dim i           As Long
Dim Pk_Killeds  As Byte
Dim C_Killeds   As Byte
Dim Pk_Total    As Byte
Dim C_Total     As Byte
Dim Termino     As Boolean
Dim End_Msg     As String
Dim to_User_X   As Byte
Dim to_User_Y   As Byte

For i = 1 To LastUser
    With UserList(i)
         
         If .ConnID <> -1 Then
         
            If criminal(i) Then
               If UserList(i).flags.Muerto <> 0 Then
                  Pk_Killeds = Pk_Killeds + 1
               End If
               
               Pk_Total = Pk_Total + 1
            Else
               If UserList(i).flags.Muerto <> 0 Then
                  C_Killeds = C_Killeds + 1
               End If
               
               C_Total = C_Total + 1
            End If
        
        End If
    End With
Next i

If criminal(userIndex) Then
   If Pk_Killeds >= Pk_Total Then
       Termino = True
       End_Msg = "Los ciudadanos ganan la ronda!"
   End If
Else
    If C_Killeds >= C_Total Then
       Termino = True
       End_Msg = "Los criminales ganan la ronda!"
    End If
End If



If (Termino = True) Then
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(End_Msg, FontTypeNames.FONTTYPE_GUILD)
    For i = 1 To LastUser
        RevivirUsuario i
        UserList(i).Stats.MinHp = UserList(i).Stats.MaxHp
        WriteUpdateHP i
        to_User_X = RandomNumber(25, 77)
        to_User_Y = RandomNumber(25, 77)
        'FindLegalPos i, Server_Info.Mapa.map, CInt(to_User_X), CInt(to_User_Y)
        WarpUserChar i, Server_Info.Mapa.map, to_User_X, to_User_Y, True
    Next i
End If


End Sub

Public Sub Setear_Objs()

'
' @ Setea los objetos - HARDCODEADO

ObjData(1).MaxHIT = 0
ObjData(1).Minhit = 0
ObjData(1).MinDef = 0
ObjData(1).MaxDef = 0
ObjData(1).OBJType = 11
ObjData(1).DefensaMagicaMin = 0
ObjData(1).DefensaMagicaMax = 0
ObjData(1).DuracionEfecto = 0
ObjData(1).TipoPocion = 4
ObjData(1).WeaponAnim = 0
ObjData(1).ShieldAnim = 0
ObjData(1).CascoAnim = 0
ObjData(1).WeaponRazaEnanaAnim = 0
ObjData(1).Ropaje = 0
ObjData(1).Apu�ala = 0
ObjData(1).GrhIndex = 541
ObjData(1).GrhSecundario = 0
ObjData(1).MinModificador = 12
ObjData(1).MaxModificador = 20
ObjData(1).Municion = 0
ObjData(1).Name = "Poci�n Azul"
ObjData(1).StaffDamageBonus = 0
ObjData(1).StaffPower = 0


ObjData(2).MaxHIT = 0
ObjData(2).Minhit = 0
ObjData(2).MinDef = 0
ObjData(2).MaxDef = 0
ObjData(2).OBJType = 11
ObjData(2).DefensaMagicaMin = 0
ObjData(2).DefensaMagicaMax = 0
ObjData(2).DuracionEfecto = 0
ObjData(2).TipoPocion = 3
ObjData(2).WeaponAnim = 0
ObjData(2).ShieldAnim = 0
ObjData(2).CascoAnim = 0
ObjData(2).WeaponRazaEnanaAnim = 0
ObjData(2).Ropaje = 0
ObjData(2).Apu�ala = 0
ObjData(2).GrhIndex = 542
ObjData(2).GrhSecundario = 0
ObjData(2).MinModificador = 30
ObjData(2).MaxModificador = 30
ObjData(2).Municion = 0
ObjData(2).Name = "Poci�n Roja"
ObjData(2).StaffDamageBonus = 0
ObjData(2).StaffPower = 0


ObjData(3).MaxHIT = 20
ObjData(3).Minhit = 7
ObjData(3).MinDef = 0
ObjData(3).MaxDef = 0
ObjData(3).OBJType = 2
ObjData(3).DefensaMagicaMin = 0
ObjData(3).DefensaMagicaMax = 0
ObjData(3).DuracionEfecto = 0
ObjData(3).TipoPocion = 0
ObjData(3).WeaponAnim = 24
ObjData(3).ShieldAnim = 0
ObjData(3).CascoAnim = 0
ObjData(3).WeaponRazaEnanaAnim = 0
ObjData(3).Ropaje = 0
ObjData(3).Apu�ala = 0
ObjData(3).GrhIndex = 566
ObjData(3).GrhSecundario = 0
ObjData(3).MinModificador = 0
ObjData(3).MaxModificador = 0
ObjData(3).Municion = 0
ObjData(3).Name = "Hacha de Guerra Dos Filos"
ObjData(3).StaffDamageBonus = 0
ObjData(3).StaffPower = 0


ObjData(4).MaxHIT = 0
ObjData(4).Minhit = 0
ObjData(4).MinDef = 2
ObjData(4).MaxDef = 5
ObjData(4).OBJType = 16
ObjData(4).DefensaMagicaMin = 0
ObjData(4).DefensaMagicaMax = 0
ObjData(4).DuracionEfecto = 0
ObjData(4).TipoPocion = 0
ObjData(4).WeaponAnim = 0
ObjData(4).ShieldAnim = 6
ObjData(4).CascoAnim = 0
ObjData(4).WeaponRazaEnanaAnim = 0
ObjData(4).Ropaje = 2
ObjData(4).Apu�ala = 0
ObjData(4).GrhIndex = 4860
ObjData(4).GrhSecundario = 0
ObjData(4).MinModificador = 0
ObjData(4).MaxModificador = 0
ObjData(4).Municion = 0
ObjData(4).Name = "Escudo de Plata"
ObjData(4).StaffDamageBonus = 0
ObjData(4).StaffPower = 0


ObjData(5).MaxHIT = 0
ObjData(5).Minhit = 0
ObjData(5).MinDef = 3
ObjData(5).MaxDef = 8
ObjData(5).OBJType = 17
ObjData(5).DefensaMagicaMin = 0
ObjData(5).DefensaMagicaMax = 0
ObjData(5).DuracionEfecto = 0
ObjData(5).TipoPocion = 0
ObjData(5).WeaponAnim = 0
ObjData(5).ShieldAnim = 0
ObjData(5).CascoAnim = 1
ObjData(5).WeaponRazaEnanaAnim = 0
ObjData(5).Ropaje = 2
ObjData(5).Apu�ala = 0
ObjData(5).GrhIndex = 559
ObjData(5).GrhSecundario = 0
ObjData(5).MinModificador = 0
ObjData(5).MaxModificador = 0
ObjData(5).Municion = 0
ObjData(5).Name = "Casco de Hierro"
ObjData(5).StaffDamageBonus = 0
ObjData(5).StaffPower = 0


ObjData(6).MaxHIT = 8
ObjData(6).Minhit = 5
ObjData(6).MinDef = 0
ObjData(6).MaxDef = 0
ObjData(6).OBJType = 2
ObjData(6).DefensaMagicaMin = 0
ObjData(6).DefensaMagicaMax = 0
ObjData(6).DuracionEfecto = 0
ObjData(6).TipoPocion = 0
ObjData(6).WeaponAnim = 52
ObjData(6).ShieldAnim = 0
ObjData(6).CascoAnim = 0
ObjData(6).WeaponRazaEnanaAnim = 53
ObjData(6).Ropaje = 0
ObjData(6).Apu�ala = 1
ObjData(6).GrhIndex = 5595
ObjData(6).GrhSecundario = 0
ObjData(6).MinModificador = 0
ObjData(6).MaxModificador = 0
ObjData(6).Municion = 0
ObjData(6).Name = "DAGA + 4"
ObjData(6).StaffDamageBonus = 0
ObjData(6).StaffPower = 0


ObjData(7).MaxHIT = 0
ObjData(7).Minhit = 0
ObjData(7).MinDef = 40
ObjData(7).MaxDef = 45
ObjData(7).OBJType = 3
ObjData(7).DefensaMagicaMin = 0
ObjData(7).DefensaMagicaMax = 0
ObjData(7).DuracionEfecto = 0
ObjData(7).TipoPocion = 0
ObjData(7).WeaponAnim = 0
ObjData(7).ShieldAnim = 0
ObjData(7).CascoAnim = 0
ObjData(7).WeaponRazaEnanaAnim = 0
ObjData(7).Ropaje = 192
ObjData(7).Apu�ala = 0
ObjData(7).GrhIndex = 3156
ObjData(7).GrhSecundario = 0
ObjData(7).MinModificador = 0
ObjData(7).MaxModificador = 0
ObjData(7).Municion = 0
ObjData(7).Name = "Armadura de Placas Completa +1 (E/G)"
ObjData(7).StaffDamageBonus = 0
ObjData(7).StaffPower = 0


ObjData(8).MaxHIT = 100
ObjData(8).Minhit = 100
ObjData(8).MinDef = 0
ObjData(8).MaxDef = 0
ObjData(8).OBJType = 2
ObjData(8).DefensaMagicaMin = 0
ObjData(8).DefensaMagicaMax = 0
ObjData(8).DuracionEfecto = 0
ObjData(8).TipoPocion = 0
ObjData(8).WeaponAnim = 15
ObjData(8).ShieldAnim = 0
ObjData(8).CascoAnim = 0
ObjData(8).WeaponRazaEnanaAnim = 65
ObjData(8).Ropaje = 0
ObjData(8).Apu�ala = 0
ObjData(8).GrhIndex = 716
ObjData(8).GrhSecundario = 0
ObjData(8).MinModificador = 0
ObjData(8).MaxModificador = 0
ObjData(8).Municion = 0
ObjData(8).Name = "Espada Mata Dragones"
ObjData(8).StaffDamageBonus = 0
ObjData(8).StaffPower = 0


ObjData(9).MaxHIT = 20
ObjData(9).Minhit = 13
ObjData(9).MinDef = 0
ObjData(9).MaxDef = 0
ObjData(9).OBJType = 2
ObjData(9).DefensaMagicaMin = 0
ObjData(9).DefensaMagicaMax = 0
ObjData(9).DuracionEfecto = 0
ObjData(9).TipoPocion = 0
ObjData(9).WeaponAnim = 13
ObjData(9).ShieldAnim = 0
ObjData(9).CascoAnim = 0
ObjData(9).WeaponRazaEnanaAnim = 63
ObjData(9).Ropaje = 0
ObjData(9).Apu�ala = 0
ObjData(9).GrhIndex = 713
ObjData(9).GrhSecundario = 0
ObjData(9).MinModificador = 0
ObjData(9).MaxModificador = 0
ObjData(9).Municion = 0
ObjData(9).Name = "Espada de Plata"
ObjData(9).StaffDamageBonus = 0
ObjData(9).StaffPower = 0


ObjData(10).MaxHIT = 0
ObjData(10).Minhit = 0
ObjData(10).MinDef = 1
ObjData(10).MaxDef = 1
ObjData(10).OBJType = 16
ObjData(10).DefensaMagicaMin = 0
ObjData(10).DefensaMagicaMax = 0
ObjData(10).DuracionEfecto = 0
ObjData(10).TipoPocion = 0
ObjData(10).WeaponAnim = 0
ObjData(10).ShieldAnim = 3
ObjData(10).CascoAnim = 0
ObjData(10).WeaponRazaEnanaAnim = 0
ObjData(10).Ropaje = 2
ObjData(10).Apu�ala = 0
ObjData(10).GrhIndex = 712
ObjData(10).GrhSecundario = 0
ObjData(10).MinModificador = 0
ObjData(10).MaxModificador = 0
ObjData(10).Municion = 0
ObjData(10).Name = "Escudo de Tortuga"
ObjData(10).StaffDamageBonus = 0
ObjData(10).StaffPower = 0


ObjData(11).MaxHIT = 0
ObjData(11).Minhit = 0
ObjData(11).MinDef = 20
ObjData(11).MaxDef = 25
ObjData(11).OBJType = 17
ObjData(11).DefensaMagicaMin = 0
ObjData(11).DefensaMagicaMax = 0
ObjData(11).DuracionEfecto = 0
ObjData(11).TipoPocion = 0
ObjData(11).WeaponAnim = 0
ObjData(11).ShieldAnim = 0
ObjData(11).CascoAnim = 8
ObjData(11).WeaponRazaEnanaAnim = 0
ObjData(11).Ropaje = 2
ObjData(11).Apu�ala = 0
ObjData(11).GrhIndex = 717
ObjData(11).GrhSecundario = 0
ObjData(11).MinModificador = 0
ObjData(11).MaxModificador = 0
ObjData(11).Municion = 0
ObjData(11).Name = "Casco de Plata"
ObjData(11).StaffDamageBonus = 0
ObjData(11).StaffPower = 0


ObjData(12).MaxHIT = 1
ObjData(12).Minhit = 1
ObjData(12).MinDef = 0
ObjData(12).MaxDef = 0
ObjData(12).OBJType = 2
ObjData(12).DefensaMagicaMin = 0
ObjData(12).DefensaMagicaMax = 0
ObjData(12).DuracionEfecto = 0
ObjData(12).TipoPocion = 0
ObjData(12).WeaponAnim = 10
ObjData(12).ShieldAnim = 0
ObjData(12).CascoAnim = 0
ObjData(12).WeaponRazaEnanaAnim = 62
ObjData(12).Ropaje = 0
ObjData(12).Apu�ala = 0
ObjData(12).GrhIndex = 986
ObjData(12).GrhSecundario = 0
ObjData(12).MinModificador = 0
ObjData(12).MaxModificador = 0
ObjData(12).Municion = 0
ObjData(12).Name = "B�culo Engarzado"
ObjData(12).StaffDamageBonus = 34
ObjData(12).StaffPower = 2


ObjData(13).MaxHIT = 0
ObjData(13).Minhit = 0
ObjData(13).MinDef = 1
ObjData(13).MaxDef = 1
ObjData(13).OBJType = 17
ObjData(13).DefensaMagicaMin = 10
ObjData(13).DefensaMagicaMax = 15
ObjData(13).DuracionEfecto = 0
ObjData(13).TipoPocion = 0
ObjData(13).WeaponAnim = 0
ObjData(13).ShieldAnim = 0
ObjData(13).CascoAnim = 4
ObjData(13).WeaponRazaEnanaAnim = 0
ObjData(13).Ropaje = 2
ObjData(13).Apu�ala = 0
ObjData(13).GrhIndex = 1018
ObjData(13).GrhSecundario = 0
ObjData(13).MinModificador = 0
ObjData(13).MaxModificador = 0
ObjData(13).Municion = 0
ObjData(13).Name = "Sombrero de Mago"
ObjData(13).StaffDamageBonus = 0
ObjData(13).StaffPower = 0


ObjData(14).MaxHIT = 0
ObjData(14).Minhit = 0
ObjData(14).MinDef = 0
ObjData(14).MaxDef = 0
ObjData(14).OBJType = 18
ObjData(14).DefensaMagicaMin = 18
ObjData(14).DefensaMagicaMax = 22
ObjData(14).DuracionEfecto = 0
ObjData(14).TipoPocion = 0
ObjData(14).WeaponAnim = 0
ObjData(14).ShieldAnim = 0
ObjData(14).CascoAnim = 0
ObjData(14).WeaponRazaEnanaAnim = 0
ObjData(14).Ropaje = 0
ObjData(14).Apu�ala = 0
ObjData(14).GrhIndex = 4902
ObjData(14).GrhSecundario = 0
ObjData(14).MinModificador = 0
ObjData(14).MaxModificador = 0
ObjData(14).Municion = 0
ObjData(14).Name = "Anillo de Disoluci�n M�gica"
ObjData(14).StaffDamageBonus = 0
ObjData(14).StaffPower = 0


ObjData(15).MaxHIT = 0
ObjData(15).Minhit = 0
ObjData(15).MinDef = 20
ObjData(15).MaxDef = 25
ObjData(15).OBJType = 3
ObjData(15).DefensaMagicaMin = 0
ObjData(15).DefensaMagicaMax = 0
ObjData(15).DuracionEfecto = 0
ObjData(15).TipoPocion = 0
ObjData(15).WeaponAnim = 0
ObjData(15).ShieldAnim = 0
ObjData(15).CascoAnim = 0
ObjData(15).WeaponRazaEnanaAnim = 0
ObjData(15).Ropaje = 56
ObjData(15).Apu�ala = 0
ObjData(15).GrhIndex = 681
ObjData(15).GrhSecundario = 0
ObjData(15).MinModificador = 0
ObjData(15).MaxModificador = 0
ObjData(15).Municion = 0
ObjData(15).Name = "T�nica de Druida"
ObjData(15).StaffDamageBonus = 0
ObjData(15).StaffPower = 0


ObjData(16).MaxHIT = 0
ObjData(16).Minhit = 0
ObjData(16).MinDef = 10
ObjData(16).MaxDef = 15
ObjData(16).OBJType = 17
ObjData(16).DefensaMagicaMin = 8
ObjData(16).DefensaMagicaMax = 12
ObjData(16).DuracionEfecto = 0
ObjData(16).TipoPocion = 0
ObjData(16).WeaponAnim = 0
ObjData(16).ShieldAnim = 0
ObjData(16).CascoAnim = 13
ObjData(16).WeaponRazaEnanaAnim = 0
ObjData(16).Ropaje = 2
ObjData(16).Apu�ala = 0
ObjData(16).GrhIndex = 20068
ObjData(16).GrhSecundario = 0
ObjData(16).MinModificador = 0
ObjData(16).MaxModificador = 0
ObjData(16).Municion = 0
ObjData(16).Name = "Casco de Tigre"
ObjData(16).StaffDamageBonus = 0
ObjData(16).StaffPower = 0


ObjData(17).MaxHIT = 1
ObjData(17).Minhit = 1
ObjData(17).MinDef = 0
ObjData(17).MaxDef = 0
ObjData(17).OBJType = 18
ObjData(17).DefensaMagicaMin = 0
ObjData(17).DefensaMagicaMax = 0
ObjData(17).DuracionEfecto = 0
ObjData(17).TipoPocion = 0
ObjData(17).WeaponAnim = 0
ObjData(17).ShieldAnim = 0
ObjData(17).CascoAnim = 0
ObjData(17).WeaponRazaEnanaAnim = 0
ObjData(17).Ropaje = 0
ObjData(17).Apu�ala = 0
ObjData(17).GrhIndex = 1523
ObjData(17).GrhSecundario = 0
ObjData(17).MinModificador = 0
ObjData(17).MaxModificador = 0
ObjData(17).Municion = 0
ObjData(17).Name = "La�d �lfico"
ObjData(17).StaffDamageBonus = 0
ObjData(17).StaffPower = 0


ObjData(18).MaxHIT = 1
ObjData(18).Minhit = 1
ObjData(18).MinDef = 0
ObjData(18).MaxDef = 0
ObjData(18).OBJType = 18
ObjData(18).DefensaMagicaMin = 0
ObjData(18).DefensaMagicaMax = 0
ObjData(18).DuracionEfecto = 0
ObjData(18).TipoPocion = 0
ObjData(18).WeaponAnim = 0
ObjData(18).ShieldAnim = 0
ObjData(18).CascoAnim = 0
ObjData(18).WeaponRazaEnanaAnim = 0
ObjData(18).Ropaje = 0
ObjData(18).Apu�ala = 0
ObjData(18).GrhIndex = 1522
ObjData(18).GrhSecundario = 0
ObjData(18).MinModificador = 0
ObjData(18).MaxModificador = 0
ObjData(18).Municion = 0
ObjData(18).Name = "Flauta �lfica"
ObjData(18).StaffDamageBonus = 0
ObjData(18).StaffPower = 0

End Sub

Public Sub Setear_Hechizos()

'
' @ Setea los hechizos - HARDCODEADO

ReDim Hechizos(1 To 46) As tHechizo

Hechizos(1).Nombre = "Ant�doto M�gico"
Hechizos(1).MaxHp = 0
Hechizos(1).MinHp = 0
Hechizos(1).Inmoviliza = 0
Hechizos(1).Invisibilidad = 0
Hechizos(1).Revivir = 0
Hechizos(1).FXgrh = 2
Hechizos(1).loops = 2
Hechizos(1).EffectIndex = 2
Hechizos(1).PalabrasMagicas = "NIHIL VED"
Hechizos(1).desc = "Con este conjuro podr�s mutar los fluidos t�xicos de tu cuerpo para detener su vil accionar, contrarrestando los efectos nocivos de todo tipo de ponzo�as y venenos. Se trata de un conjuro tan simple como necesario, puesto que, adem�s de los magos que provocan envenenamiento, son varias las criaturas que lanzan veneno en estas tierras."
Hechizos(1).HechizeroMsg = "Le has detenido el envenenamiento a"
Hechizos(1).TargetMsg = "te ha detenido el envenenamiento."
Hechizos(1).tipo = 2
Hechizos(1).RemoverParalisis = 0
Hechizos(1).Paraliza = 0
Hechizos(1).ManaRequerido = 12
Hechizos(1).Target = 1
Hechizos(1).NeedStaff = 0
Hechizos(1).StaffAffected = False

Hechizos(2).Nombre = "Dardo M�gico"
Hechizos(2).MaxHp = 5
Hechizos(2).MinHp = 3
Hechizos(2).Inmoviliza = 0
Hechizos(2).Invisibilidad = 0
Hechizos(2).Revivir = 0
Hechizos(2).FXgrh = 15
Hechizos(2).loops = 2
Hechizos(2).EffectIndex = 15
Hechizos(2).PalabrasMagicas = "OHL VOR PEK"
Hechizos(2).desc = "�ste es uno de los m�s elementales y sencillos hechizos de ataque que podr�s aprender. No causa un gran da�o a la v�ctima, pero al consumir muy pocos recursos, resulta una herramienta muy �til para el entrenamiento de los j�venes hechiceros."
Hechizos(2).HechizeroMsg = "Has lanzado un dardo m�gico sobre"
Hechizos(2).TargetMsg = "ha lanzado un dardo m�gico sobre ti."
Hechizos(2).tipo = 1
Hechizos(2).RemoverParalisis = 0
Hechizos(2).Paraliza = 0
Hechizos(2).ManaRequerido = 10
Hechizos(2).Target = 3
Hechizos(2).NeedStaff = 0
Hechizos(2).StaffAffected = False

Hechizos(3).Nombre = "Curar Heridas Leves"
Hechizos(3).MaxHp = 5
Hechizos(3).MinHp = 1
Hechizos(3).Inmoviliza = 0
Hechizos(3).Invisibilidad = 0
Hechizos(3).Revivir = 0
Hechizos(3).FXgrh = 0
Hechizos(3).loops = 0
Hechizos(3).EffectIndex = 0
Hechizos(3).PalabrasMagicas = "CORP SANC"
Hechizos(3).desc = "Con este sencillo conjuro podr�s efectuar la sanaci�n inmediata de peque�as heridas y devolver lentamente la salud a quien se lo lances. Suele ser de gran utilidad para los j�venes aprendices de magia, para poder entrenar al resguardo del ataque de las fieras salvajes."
Hechizos(3).HechizeroMsg = "Le has curado algunas peque�as heridas a"
Hechizos(3).TargetMsg = "te ha curado algunas peque�as heridas."
Hechizos(3).tipo = 1
Hechizos(3).RemoverParalisis = 0
Hechizos(3).Paraliza = 0
Hechizos(3).ManaRequerido = 10
Hechizos(3).Target = 3
Hechizos(3).NeedStaff = 0
Hechizos(3).StaffAffected = False

Hechizos(4).Nombre = "Toxina"
Hechizos(4).MaxHp = 0
Hechizos(4).MinHp = 0
Hechizos(4).Inmoviliza = 0
Hechizos(4).Invisibilidad = 0
Hechizos(4).Revivir = 0
Hechizos(4).FXgrh = 3
Hechizos(4).loops = 2
Hechizos(4).EffectIndex = 3
Hechizos(4).PalabrasMagicas = "SERP XON IN"
Hechizos(4).desc = "Con este hechizo de aprendizaje simple y veloz, podr�s inundar el cuerpo de la v�ctima de un mortal veneno, con el que lentamente ir� perdiendo la vida hasta sucumbir. Por ser un hechizo que requiere poco conocimiento, resulta de gran utilidad en las primeras etapas, fundamentalmente como modo de defensa personal e incluso para el entrenamiento."
Hechizos(4).HechizeroMsg = "Has envenenado a"
Hechizos(4).TargetMsg = "te ha envenenado."
Hechizos(4).tipo = 2
Hechizos(4).RemoverParalisis = 0
Hechizos(4).Paraliza = 0
Hechizos(4).ManaRequerido = 24
Hechizos(4).Target = 3
Hechizos(4).NeedStaff = 0
Hechizos(4).StaffAffected = False

Hechizos(5).Nombre = "Curar Heridas Graves"
Hechizos(5).MaxHp = 35
Hechizos(5).MinHp = 12
Hechizos(5).Inmoviliza = 0
Hechizos(5).Invisibilidad = 0
Hechizos(5).Revivir = 0
Hechizos(5).FXgrh = 9
Hechizos(5).loops = 0
Hechizos(5).EffectIndex = 9
Hechizos(5).PalabrasMagicas = "EN CORP SANCTIS"
Hechizos(5).desc = "Con esta milagrosa sanaci�n para las heridas m�s profundas y dolorosas, podr�s r�pidamente devolver la salud a quienes agonizan. Se trata de un conjuro de una complejidad intermedia, pero que suele resultar fundamental para los j�venes aventureros que desean recorrer el mundo y enfrentarse con sus peligros."
Hechizos(5).HechizeroMsg = "Le has sanado graves heridas a"
Hechizos(5).TargetMsg = "te ha sanado graves heridas."
Hechizos(5).tipo = 1
Hechizos(5).RemoverParalisis = 0
Hechizos(5).Paraliza = 0
Hechizos(5).ManaRequerido = 40
Hechizos(5).Target = 3
Hechizos(5).NeedStaff = 0
Hechizos(5).StaffAffected = False

Hechizos(6).Nombre = "Flecha M�gica"
Hechizos(6).MaxHp = 12
Hechizos(6).MinHp = 6
Hechizos(6).Inmoviliza = 0
Hechizos(6).Invisibilidad = 0
Hechizos(6).Revivir = 0
Hechizos(6).FXgrh = 33
Hechizos(6).loops = 0
Hechizos(6).EffectIndex = 33
Hechizos(6).PalabrasMagicas = "VAX PER"
Hechizos(6).desc = "Con este sortilegio lograr�s que una inmaterial flecha impacte en la v�ctima produci�ndole heridas de mediana gravedad. Su aprendizaje es medianamente f�cil, y puede resultar de mucha utilidad para combatir contra las criaturas menores que suelen habitar los suburbios de las ciudades."
Hechizos(6).HechizeroMsg = "Has lanzado una flecha m�gica sobre"
Hechizos(6).TargetMsg = "ha lanzado una flecha m�gica sobre ti."
Hechizos(6).tipo = 1
Hechizos(6).RemoverParalisis = 0
Hechizos(6).Paraliza = 0
Hechizos(6).ManaRequerido = 20
Hechizos(6).Target = 3
Hechizos(6).NeedStaff = 0
Hechizos(6).StaffAffected = False

Hechizos(7).Nombre = "Flecha El�ctrica"
Hechizos(7).MaxHp = 20
Hechizos(7).MinHp = 12
Hechizos(7).Inmoviliza = 0
Hechizos(7).Invisibilidad = 0
Hechizos(7).Revivir = 0
Hechizos(7).FXgrh = 32
Hechizos(7).loops = 0
Hechizos(7).EffectIndex = 32
Hechizos(7).PalabrasMagicas = "SUN VAP"
Hechizos(7).desc = "�ste es un interesante hechizo que te permitir� canalizar las elerg�as el�ctricas del mundo en ataques direccionados a un determinado objetivo. Si bien el da�o que causa no es muy grande, es seguramente la mejor herramienta para el entrenamiento y la defensa de los hechiceros menos experimentados."
Hechizos(7).HechizeroMsg = "Has lanzado una flecha el�ctrica sobre"
Hechizos(7).TargetMsg = "ha lanzado una flecha el�ctrica sobre ti."
Hechizos(7).tipo = 1
Hechizos(7).RemoverParalisis = 0
Hechizos(7).Paraliza = 0
Hechizos(7).ManaRequerido = 40
Hechizos(7).Target = 3
Hechizos(7).NeedStaff = 0
Hechizos(7).StaffAffected = False

Hechizos(8).Nombre = "Proyectil M�gico"
Hechizos(8).MaxHp = 35
Hechizos(8).MinHp = 30
Hechizos(8).Inmoviliza = 0
Hechizos(8).Invisibilidad = 0
Hechizos(8).Revivir = 0
Hechizos(8).FXgrh = 10
Hechizos(8).loops = 0
Hechizos(8).EffectIndex = 10
Hechizos(8).PalabrasMagicas = "VAX IN TAR"
Hechizos(8).desc = "Este rentable hechizo de ataque resulta el ideal para los niveles intermedios, pues les permitir� un gran poder de ataque que no s�lo facilitar� entrenar velozmente, si no que adem�s, puede ser utilizado como una respetable arma en los combates contra tus adversarios."
Hechizos(8).HechizeroMsg = "Has lanzado un proyectil m�gico sobre"
Hechizos(8).TargetMsg = "ha lanzado un proyectil m�gico sobre ti."
Hechizos(8).tipo = 1
Hechizos(8).RemoverParalisis = 0
Hechizos(8).Paraliza = 0
Hechizos(8).ManaRequerido = 45
Hechizos(8).Target = 3
Hechizos(8).NeedStaff = 0
Hechizos(8).StaffAffected = False

Hechizos(9).Nombre = "Paralizar"
Hechizos(9).MaxHp = 0
Hechizos(9).MinHp = 0
Hechizos(9).Inmoviliza = 0
Hechizos(9).Invisibilidad = 0
Hechizos(9).Revivir = 0
Hechizos(9).FXgrh = 8
Hechizos(9).loops = 0
Hechizos(9).EffectIndex = 8
Hechizos(9).PalabrasMagicas = "HOAX VORP"
Hechizos(9).desc = "Con este arcano encantamiento podr�s dejar completamente petrificado a la v�ctima durante un determinado lapso de tiempo. Se trata de uno de los hechizos m�s utilizado en estas tierras por su gran utilidad tanto en el combate como en el entrenamiento."
Hechizos(9).HechizeroMsg = "Has paralizado a"
Hechizos(9).TargetMsg = "te ha paralizado."
Hechizos(9).tipo = 2
Hechizos(9).RemoverParalisis = 0
Hechizos(9).Paraliza = 1
Hechizos(9).ManaRequerido = 450
Hechizos(9).Target = 3
Hechizos(9).NeedStaff = 0
Hechizos(9).StaffAffected = False

Hechizos(10).Nombre = "Devolver Movilidad"
Hechizos(10).MaxHp = 0
Hechizos(10).MinHp = 0
Hechizos(10).Inmoviliza = 0
Hechizos(10).Invisibilidad = 0
Hechizos(10).Revivir = 0
Hechizos(10).FXgrh = 0
Hechizos(10).loops = 0
Hechizos(10).EffectIndex = 0
Hechizos(10).PalabrasMagicas = "AN HOAX VORP"
Hechizos(10).desc = "Conjuro con el cual podr�s contrarrestar los efectos tanto de la par�lisis como de la inmovilizaci�n. En virtud de que los hechizos de estatismo son usados con mucha frecuencia por todos las clases m�gicas que habitan estas tierras, �ste conjuro se vuelve pr�cticamente vital para cualquier valiente que est� dispuesto a combatir."
Hechizos(10).HechizeroMsg = "Le has devuelto la movilidad a"
Hechizos(10).TargetMsg = "te ha devuelto la movilidad."
Hechizos(10).tipo = 2
Hechizos(10).RemoverParalisis = 1
Hechizos(10).Paraliza = 0
Hechizos(10).ManaRequerido = 300
Hechizos(10).Target = 3
Hechizos(10).NeedStaff = 0
Hechizos(10).StaffAffected = False

Hechizos(11).Nombre = "Resucitar"
Hechizos(11).MaxHp = 0
Hechizos(11).MinHp = 0
Hechizos(11).Inmoviliza = 0
Hechizos(11).Invisibilidad = 0
Hechizos(11).Revivir = 1
Hechizos(11).FXgrh = 0
Hechizos(11).loops = 0
Hechizos(11).EffectIndex = 0
Hechizos(11).PalabrasMagicas = "AHIL KN� X�R"
Hechizos(11).desc = "Encanto sagrado que te permitir� devolver la vida a los difuntos. Esta incre�ble habilidad fue ense�ada personalmente por el mismo Gulfas Morgolock al otrora Rey de los hombres en las remotas �pocas del �rbol Blanco. Su aprendizaje no es f�cil y su utilizaci�n resulta muy desgastante, pero sin dudas es una de las piezas m�s importantes de la magia de las tierras del Argentum."
Hechizos(11).HechizeroMsg = "Has resucitado a"
Hechizos(11).TargetMsg = "te ha resucitado."
Hechizos(11).tipo = 2
Hechizos(11).RemoverParalisis = 0
Hechizos(11).Paraliza = 0
Hechizos(11).ManaRequerido = 550
Hechizos(11).Target = 1
Hechizos(11).NeedStaff = 2
Hechizos(11).StaffAffected = False

Hechizos(12).Nombre = "Ataque de Hambre"
Hechizos(12).MaxHp = 0
Hechizos(12).MinHp = 0
Hechizos(12).Inmoviliza = 0
Hechizos(12).Invisibilidad = 0
Hechizos(12).Revivir = 0
Hechizos(12).FXgrh = 0
Hechizos(12).loops = 0
Hechizos(12).EffectIndex = 0
Hechizos(12).PalabrasMagicas = "�L AEX"
Hechizos(12).desc = "Embrujo que tiene por objeto provocar un considerable desbalance nutricional en el destinatario, lo que conlleva a la p�rdida de energ�a y de concentraci�n. Ideal para lograr la debilitaci�n del adversario durante los combates."
Hechizos(12).HechizeroMsg = "Le has generado un fuerte hambre a"
Hechizos(12).TargetMsg = "te ha generado un fuerte hambre."
Hechizos(12).tipo = 1
Hechizos(12).RemoverParalisis = 0
Hechizos(12).Paraliza = 0
Hechizos(12).ManaRequerido = 20
Hechizos(12).Target = 1
Hechizos(12).NeedStaff = 0
Hechizos(12).StaffAffected = False

Hechizos(13).Nombre = "Terrible hambre de Ig�r"
Hechizos(13).MaxHp = 0
Hechizos(13).MinHp = 0
Hechizos(13).Inmoviliza = 0
Hechizos(13).Invisibilidad = 0
Hechizos(13).Revivir = 0
Hechizos(13).FXgrh = 0
Hechizos(13).loops = 0
Hechizos(13).EffectIndex = 0
Hechizos(13).PalabrasMagicas = "�X '�L AEX"
Hechizos(13).desc = "Abominable hechizo que provoca en la v�ctima un estado de desnutrici�n absoluto. Este encantamiento fue creado por el mago Nimbul mientras estudiaba la conducta de los gigantes de las monta�as, la primer v�ctima de este hechizo fue un gigante llamado Ig�r, de ah� el nombre del mismo."
Hechizos(13).HechizeroMsg = "Le has lanzado terrible hambre de Ig�r a"
Hechizos(13).TargetMsg = "te ha lanzado terrible hambre de Ig�r."
Hechizos(13).tipo = 1
Hechizos(13).RemoverParalisis = 0
Hechizos(13).Paraliza = 0
Hechizos(13).ManaRequerido = 55
Hechizos(13).Target = 1
Hechizos(13).NeedStaff = 0
Hechizos(13).StaffAffected = False

Hechizos(14).Nombre = "Invisibilidad"
Hechizos(14).MaxHp = 0
Hechizos(14).MinHp = 0
Hechizos(14).Inmoviliza = 0
Hechizos(14).Invisibilidad = 1
Hechizos(14).Revivir = 0
Hechizos(14).FXgrh = 0
Hechizos(14).loops = 0
Hechizos(14).EffectIndex = 0
Hechizos(14).PalabrasMagicas = vbNullString
Hechizos(14).desc = "Se trata de uno de los recursos m�gicos m�s importante de estas tierras. Con �l podr�s volverte (t� o a quien se lo lances) completamente invisible a la vista de los dem�s. Esto lo convierte en una de las herramientas m�s importantes utilizadas por los combatientes guerreros."
Hechizos(14).HechizeroMsg = "Has vuelto invisible a"
Hechizos(14).TargetMsg = "te ha vuelto invisible."
Hechizos(14).tipo = 2
Hechizos(14).RemoverParalisis = 0
Hechizos(14).Paraliza = 0
Hechizos(14).ManaRequerido = 500
Hechizos(14).Target = 1
Hechizos(14).NeedStaff = 0
Hechizos(14).StaffAffected = False

Hechizos(15).Nombre = "Tormenta de Fuego"
Hechizos(15).MaxHp = 55
Hechizos(15).MinHp = 45
Hechizos(15).Inmoviliza = 0
Hechizos(15).Invisibilidad = 0
Hechizos(15).Revivir = 0
Hechizos(15).FXgrh = 7
Hechizos(15).loops = 0
Hechizos(15).EffectIndex = 7
Hechizos(15).PalabrasMagicas = "EN VAX ON TAR"
Hechizos(15).desc = "Poderoso hechizo de ataque. Su costo en relaci�n al da�o que provoca lo convierte en un arma m�s propia para el combate que para el entrenamiento. Es ideal para aquellos aprendices de magia que todav�a no alcanzaron a conocer el secreto de los hechizos mayores y necesitan estar bien equipados para la guerra."
Hechizos(15).HechizeroMsg = "Has lanzado una tormenta de fuego sobre"
Hechizos(15).TargetMsg = "lanz� una tormenta de fuego sobre ti."
Hechizos(15).tipo = 1
Hechizos(15).RemoverParalisis = 0
Hechizos(15).Paraliza = 0
Hechizos(15).ManaRequerido = 250
Hechizos(15).Target = 3
Hechizos(15).NeedStaff = 0
Hechizos(15).StaffAffected = True

Hechizos(16).Nombre = "Llamado a la Naturaleza"
Hechizos(16).MaxHp = 0
Hechizos(16).MinHp = 0
Hechizos(16).Inmoviliza = 0
Hechizos(16).Invisibilidad = 0
Hechizos(16).Revivir = 0
Hechizos(16).FXgrh = 0
Hechizos(16).loops = 0
Hechizos(16).EffectIndex = 0
Hechizos(16).PalabrasMagicas = "Nature et worg"
Hechizos(16).desc = "El secreto dru�dico de la invocaci�n se concentra en este embrujo, que te permitir� implorar por ayuda a la madre naturaleza. Al hacerlo, dos temibles lobos acudir�n en tu auxilio, para socorrerte de cualquier peligro que pudieses estar padeciendo."
Hechizos(16).HechizeroMsg = "Has Llamado a la naturaleza contra"
Hechizos(16).TargetMsg = "ha llamado a la naturaleza sobre ti."
Hechizos(16).tipo = 4
Hechizos(16).RemoverParalisis = 0
Hechizos(16).Paraliza = 0
Hechizos(16).ManaRequerido = 120
Hechizos(16).Target = 4
Hechizos(16).NeedStaff = 0
Hechizos(16).StaffAffected = False

Hechizos(17).Nombre = "Llamado Nigromante"
Hechizos(17).MaxHp = 0
Hechizos(17).MinHp = 0
Hechizos(17).Inmoviliza = 0
Hechizos(17).Invisibilidad = 0
Hechizos(17).Revivir = 0
Hechizos(17).FXgrh = 0
Hechizos(17).loops = 0
Hechizos(17).EffectIndex = 0
Hechizos(17).PalabrasMagicas = "Mo� c�mus"
Hechizos(17).desc = "Con este espantoso ritual de magia negra podr�s lograr que los restos mortales de quienes descansan en paz, se alcen en tu ayuda. Al hacerlo, dos tenebrosos zombies emerger�n de la tierra para colaborar contigo."
Hechizos(17).HechizeroMsg = "Has invocado la ayuda de los muertos contra"
Hechizos(17).TargetMsg = "ha utilizado la Nigromancia sobre ti."
Hechizos(17).tipo = 4
Hechizos(17).RemoverParalisis = 0
Hechizos(17).Paraliza = 0
Hechizos(17).ManaRequerido = 400
Hechizos(17).Target = 4
Hechizos(17).NeedStaff = 1
Hechizos(17).StaffAffected = False

Hechizos(18).Nombre = "Celeridad"
Hechizos(18).MaxHp = 0
Hechizos(18).MinHp = 0
Hechizos(18).Inmoviliza = 0
Hechizos(18).Invisibilidad = 0
Hechizos(18).Revivir = 0
Hechizos(18).FXgrh = 20
Hechizos(18).loops = 0
Hechizos(18).EffectIndex = 20
Hechizos(18).PalabrasMagicas = "YUP A'INC"
Hechizos(18).desc = "Podr�s aumentar la agilidad del destinatario a trav�s de este muy �til sortilegio. Con �l devolver�s al cuerpo toda la ligereza y velocidad que necesita para adquirir una verdadera destreza guerrera. Su facilidad y los pocos recursos que insume su pr�ctica, lo convierten en una herramienta sumamente �til que no puede faltarle a ning�n buen combatiente."
Hechizos(18).HechizeroMsg = "Has aumentado la agilidad de"
Hechizos(18).TargetMsg = "ha aumentado tu agilidad."
Hechizos(18).tipo = 1
Hechizos(18).RemoverParalisis = 0
Hechizos(18).Paraliza = 0
Hechizos(18).ManaRequerido = 40
Hechizos(18).Target = 1
Hechizos(18).NeedStaff = 0
Hechizos(18).StaffAffected = False

Hechizos(19).Nombre = "Torpeza"
Hechizos(19).MaxHp = 0
Hechizos(19).MinHp = 0
Hechizos(19).Inmoviliza = 0
Hechizos(19).Invisibilidad = 0
Hechizos(19).Revivir = 0
Hechizos(19).FXgrh = 0
Hechizos(19).loops = 0
Hechizos(19).EffectIndex = 0
Hechizos(19).PalabrasMagicas = "ASYNC YUP A'INC"
Hechizos(19).desc = "Este conjuro tiene por finalidad contrarrestar o anular los efectos del hechizo de celeridad, reduciendo la agilidad y destreza que hayan adquirido quienes fueron hechizados con aqu�l."
Hechizos(19).HechizeroMsg = "Has lanzado torpeza sobre"
Hechizos(19).TargetMsg = "ha lanzado torpeza sobre ti."
Hechizos(19).tipo = 1
Hechizos(19).RemoverParalisis = 0
Hechizos(19).Paraliza = 0
Hechizos(19).ManaRequerido = 20
Hechizos(19).Target = 1
Hechizos(19).NeedStaff = 0
Hechizos(19).StaffAffected = False

Hechizos(20).Nombre = "Fuerza"
Hechizos(20).MaxHp = 0
Hechizos(20).MinHp = 0
Hechizos(20).Inmoviliza = 0
Hechizos(20).Invisibilidad = 0
Hechizos(20).Revivir = 0
Hechizos(20).FXgrh = 17
Hechizos(20).loops = 0
Hechizos(20).EffectIndex = 17
Hechizos(20).PalabrasMagicas = "Ar A'kron"
Hechizos(20).desc = " El poder m�gico que encierra este hechizo te permitir� aumentar notoriamente la fuerza y el poder�o de la persona sobre quien lo invoques. Al igual que el hechizo de celeridad, su facilidad y los pocos recursos que insume su pr�ctica, lo convierten en una herramienta sumamente �til que no puede faltarle a ning�n buen combatiente."
Hechizos(20).HechizeroMsg = "Has lanzado fuerza sobre"
Hechizos(20).TargetMsg = "ha lanzado fuerza sobre ti."
Hechizos(20).tipo = 1
Hechizos(20).RemoverParalisis = 0
Hechizos(20).Paraliza = 0
Hechizos(20).ManaRequerido = 50
Hechizos(20).Target = 1
Hechizos(20).NeedStaff = 0
Hechizos(20).StaffAffected = False

Hechizos(21).Nombre = "Debilidad"
Hechizos(21).MaxHp = 0
Hechizos(21).MinHp = 0
Hechizos(21).Inmoviliza = 0
Hechizos(21).Invisibilidad = 0
Hechizos(21).Revivir = 0
Hechizos(21).FXgrh = 0
Hechizos(21).loops = 0
Hechizos(21).EffectIndex = 0
Hechizos(21).PalabrasMagicas = "Xoom varp"
Hechizos(21).desc = "Este conjuro tiene por finalidad contrarrestar o anular los efectos del hechizo de Fuerza, reduciendo la fuerza y el poder�o que hayan adquirido quienes fueron hechizados con aqu�l."
Hechizos(21).HechizeroMsg = "Has lanzado Debilidad sobre"
Hechizos(21).TargetMsg = "ha lanzado Debilidad sobre ti."
Hechizos(21).tipo = 1
Hechizos(21).RemoverParalisis = 0
Hechizos(21).Paraliza = 0
Hechizos(21).ManaRequerido = 45
Hechizos(21).Target = 1
Hechizos(21).NeedStaff = 0
Hechizos(21).StaffAffected = False

Hechizos(22).Nombre = "Llamado a Uhkrul"
Hechizos(22).MaxHp = 0
Hechizos(22).MinHp = 0
Hechizos(22).Inmoviliza = 0
Hechizos(22).Invisibilidad = 0
Hechizos(22).Revivir = 0
Hechizos(22).FXgrh = 0
Hechizos(22).loops = 0
Hechizos(22).EffectIndex = 0
Hechizos(22).PalabrasMagicas = "Arg Zaga�arak"
Hechizos(22).desc = "Invoca un Drag�n Rojo."
Hechizos(22).HechizeroMsg = "Llamado a Uhkrul."
Hechizos(22).TargetMsg = vbNullString
Hechizos(22).tipo = 4
Hechizos(22).RemoverParalisis = 0
Hechizos(22).Paraliza = 0
Hechizos(22).ManaRequerido = 1
Hechizos(22).Target = 4
Hechizos(22).NeedStaff = 0
Hechizos(22).StaffAffected = False

Hechizos(23).Nombre = "Descarga El�ctrica"
Hechizos(23).MaxHp = 85
Hechizos(23).MinHp = 55
Hechizos(23).Inmoviliza = 0
Hechizos(23).Invisibilidad = 0
Hechizos(23).Revivir = 0
Hechizos(23).FXgrh = 11
Hechizos(23).loops = 0
Hechizos(23).EffectIndex = 11
Hechizos(23).PalabrasMagicas = "T 'HY KOOOL"
Hechizos(23).desc = "�ste es uno de los hechizos de da�o m�s poderoso de todo el Argentum. Al controlarlo podr�s manipular las fuerzas de la naturaleza de modo tal que del cielo mismo caiga una fuerte descarga de electricidad sobre el objetivo. El da�o que provoca en sus v�citimas puede ser mortal frente a principiantes o aprendices."
Hechizos(23).HechizeroMsg = "Has lanzado una descarga el�ctrica sobre"
Hechizos(23).TargetMsg = "ha lanzado una descarga el�ctrica sobre ti."
Hechizos(23).tipo = 1
Hechizos(23).RemoverParalisis = 0
Hechizos(23).Paraliza = 0
Hechizos(23).ManaRequerido = 460
Hechizos(23).Target = 3
Hechizos(23).NeedStaff = 0
Hechizos(23).StaffAffected = True

Hechizos(24).Nombre = "Inmovilizar"
Hechizos(24).MaxHp = 0
Hechizos(24).MinHp = 0
Hechizos(24).Inmoviliza = 1
Hechizos(24).Invisibilidad = 0
Hechizos(24).Revivir = 0
Hechizos(24).FXgrh = 12
Hechizos(24).loops = 0
Hechizos(24).EffectIndex = 12
Hechizos(24).PalabrasMagicas = "�r Prop s'uo"
Hechizos(24).desc = "Este valioso sortilegio te permitir� dejar a la v�ctima sin la capacidad de desplazarse, tendr� la sensaci�n de que una poderosa fuerza lo atrae hacia el piso, y si bien podr� mover sus miembros, no podr� dejar el lugar en el que se encuentra."
Hechizos(24).HechizeroMsg = "Has inmovilizado a"
Hechizos(24).TargetMsg = "te ha inmovilizado."
Hechizos(24).tipo = 2
Hechizos(24).RemoverParalisis = 0
Hechizos(24).Paraliza = 0
Hechizos(24).ManaRequerido = 300
Hechizos(24).Target = 3
Hechizos(24).NeedStaff = 0
Hechizos(24).StaffAffected = False

Hechizos(25).Nombre = "Apocalipsis"
Hechizos(25).MaxHp = 100
Hechizos(25).MinHp = 85
Hechizos(25).Inmoviliza = 0
Hechizos(25).Invisibilidad = 0
Hechizos(25).Revivir = 0
Hechizos(25).FXgrh = 13
Hechizos(25).loops = 0
Hechizos(25).EffectIndex = 13
Hechizos(25).PalabrasMagicas = "Rahma Na�arak O'al"
Hechizos(25).desc = "El ataque m�gico m�s letal de estas tierras. S�lo aquellos avezados en el arte de la magia pueden aprender este temible hechizo. Su poder es fulminante y suele ser mortal para quien es v�ctima de �l."
Hechizos(25).HechizeroMsg = "Has lanzado Apocalipsis sobre"
Hechizos(25).TargetMsg = "lanz� Apocalipsis sobre ti."
Hechizos(25).tipo = 1
Hechizos(25).RemoverParalisis = 0
Hechizos(25).Paraliza = 0
Hechizos(25).ManaRequerido = 1000
Hechizos(25).Target = 3
Hechizos(25).NeedStaff = 0
Hechizos(25).StaffAffected = True

Hechizos(26).Nombre = "Invocar Elemental de Fuego"
Hechizos(26).MaxHp = 0
Hechizos(26).MinHp = 0
Hechizos(26).Inmoviliza = 0
Hechizos(26).Invisibilidad = 0
Hechizos(26).Revivir = 0
Hechizos(26).FXgrh = 7
Hechizos(26).loops = 0
Hechizos(26).EffectIndex = 7
Hechizos(26).PalabrasMagicas = "Fir Yur'rax"
Hechizos(26).desc = "Encontrar�s la colaboraci�n de uno de los elementos vitales del universo, el coraz�n mismo del ardoroso fuego vendr� en tu auxilio."
Hechizos(26).HechizeroMsg = "Has invocado un elemental de fuego contra"
Hechizos(26).TargetMsg = "ha invocado un elemental de fuego sobre ti."
Hechizos(26).tipo = 4
Hechizos(26).RemoverParalisis = 0
Hechizos(26).Paraliza = 0
Hechizos(26).ManaRequerido = 1100
Hechizos(26).Target = 4
Hechizos(26).NeedStaff = 2
Hechizos(26).StaffAffected = False

Hechizos(27).Nombre = "Invocar Elemental de Agua"
Hechizos(27).MaxHp = 55
Hechizos(27).MinHp = 30
Hechizos(27).Inmoviliza = 0
Hechizos(27).Invisibilidad = 0
Hechizos(27).Revivir = 0
Hechizos(27).FXgrh = 7
Hechizos(27).loops = 0
Hechizos(27).EffectIndex = 7
Hechizos(27).PalabrasMagicas = "Wata Mantra'rax"
Hechizos(27).desc = "Encontrar�s la colaboraci�n de uno de los elementos que conforman el universo, el agua vital vendr� en tu auxilio."
Hechizos(27).HechizeroMsg = "Has invocado un elemental de agua contra"
Hechizos(27).TargetMsg = "ha invocado un elemental de agua sobre ti."
Hechizos(27).tipo = 4
Hechizos(27).RemoverParalisis = 0
Hechizos(27).Paraliza = 0
Hechizos(27).ManaRequerido = 980
Hechizos(27).Target = 4
Hechizos(27).NeedStaff = 1
Hechizos(27).StaffAffected = False

Hechizos(28).Nombre = "Invocar Elemental de Tierra"
Hechizos(28).MaxHp = 0
Hechizos(28).MinHp = 0
Hechizos(28).Inmoviliza = 0
Hechizos(28).Invisibilidad = 0
Hechizos(28).Revivir = 0
Hechizos(28).FXgrh = 7
Hechizos(28).loops = 0
Hechizos(28).EffectIndex = 7
Hechizos(28).PalabrasMagicas = "Mu Mantra'rax"
Hechizos(28).desc = "Encontrar�s la colaboraci�n de uno de los elementos vitales del universo, la sagrada tierra vendr� en tu auxilio."
Hechizos(28).HechizeroMsg = "Has invocado un elemental de tierra contra"
Hechizos(28).TargetMsg = "ha invocado un elemental de tierra sobre ti."
Hechizos(28).tipo = 4
Hechizos(28).RemoverParalisis = 0
Hechizos(28).Paraliza = 0
Hechizos(28).ManaRequerido = 980
Hechizos(28).Target = 4
Hechizos(28).NeedStaff = 1
Hechizos(28).StaffAffected = False

Hechizos(29).Nombre = "Implorar Ayuda"
Hechizos(29).MaxHp = 0
Hechizos(29).MinHp = 0
Hechizos(29).Inmoviliza = 0
Hechizos(29).Invisibilidad = 0
Hechizos(29).Revivir = 0
Hechizos(29).FXgrh = 7
Hechizos(29).loops = 0
Hechizos(29).EffectIndex = 7
Hechizos(29).PalabrasMagicas = "Ar 'Cos Mantra'rax"
Hechizos(29).desc = "Implora la ayuda divina de los dioses"
Hechizos(29).HechizeroMsg = "Has implorado ayuda a los dioses!"

Hechizos(29).tipo = 4
Hechizos(29).RemoverParalisis = 0
Hechizos(29).Paraliza = 0
Hechizos(29).ManaRequerido = 1
Hechizos(29).Target = 4
Hechizos(29).NeedStaff = 0
Hechizos(29).StaffAffected = False

Hechizos(30).Nombre = "Ceguera"
Hechizos(30).MaxHp = 0
Hechizos(30).MinHp = 0
Hechizos(30).Inmoviliza = 0
Hechizos(30).Invisibilidad = 0
Hechizos(30).Revivir = 0
Hechizos(30).FXgrh = 0
Hechizos(30).loops = 0
Hechizos(30).EffectIndex = 0
Hechizos(30).PalabrasMagicas = "CAE ' XitA"
Hechizos(30).desc = "Este embrujo le quitar� el sentido de la vista a vuestro oponente."
Hechizos(30).HechizeroMsg = "Has lanzado ceguera sobre"
Hechizos(30).TargetMsg = "lanz� ceguera vos."
Hechizos(30).tipo = 2
Hechizos(30).RemoverParalisis = 0
Hechizos(30).Paraliza = 0
Hechizos(30).ManaRequerido = 1
Hechizos(30).Target = 1
Hechizos(30).NeedStaff = 0
Hechizos(30).StaffAffected = False

Hechizos(31).Nombre = "Aturdir"
Hechizos(31).MaxHp = 0
Hechizos(31).MinHp = 0
Hechizos(31).Inmoviliza = 0
Hechizos(31).Invisibilidad = 0
Hechizos(31).Revivir = 0
Hechizos(31).FXgrh = 0
Hechizos(31).loops = 0
Hechizos(31).EffectIndex = 0
Hechizos(31).PalabrasMagicas = "ASYNC GAM AL�"
Hechizos(31).desc = "Este sortilegio provocar� que la v�ctima pierda moment�neamente todo tipo de sentido de la orientaci�n."
Hechizos(31).HechizeroMsg = "Has aturdido a"
Hechizos(31).TargetMsg = "te ha aturdido."
Hechizos(31).tipo = 2
Hechizos(31).RemoverParalisis = 0
Hechizos(31).Paraliza = 0
Hechizos(31).ManaRequerido = 800
Hechizos(31).Target = 1
Hechizos(31).NeedStaff = 0
Hechizos(31).StaffAffected = False

Hechizos(32).Nombre = "Ira de Dios"
Hechizos(32).MaxHp = 9999
Hechizos(32).MinHp = 9999
Hechizos(32).Inmoviliza = 0
Hechizos(32).Invisibilidad = 0
Hechizos(32).Revivir = 0
Hechizos(32).FXgrh = 26
Hechizos(32).loops = 0
Hechizos(32).EffectIndex = 26
Hechizos(32).PalabrasMagicas = "La IRA de Dios"
Hechizos(32).desc = "S�lo el encono de los Dioses puede provocar tanto da�o."
Hechizos(32).HechizeroMsg = "Has lanzado la IRA de Dios sobre"
Hechizos(32).TargetMsg = "te lanz� la IRA de Dios."
Hechizos(32).tipo = 1
Hechizos(32).RemoverParalisis = 0
Hechizos(32).Paraliza = 0
Hechizos(32).ManaRequerido = 1
Hechizos(32).Target = 3
Hechizos(32).NeedStaff = 0
Hechizos(32).StaffAffected = False

Hechizos(33).Nombre = "Invocaci�n de Ultratumba"
Hechizos(33).MaxHp = 0
Hechizos(33).MinHp = 0
Hechizos(33).Inmoviliza = 0
Hechizos(33).Invisibilidad = 0
Hechizos(33).Revivir = 0
Hechizos(33).FXgrh = 0
Hechizos(33).loops = 0
Hechizos(33).EffectIndex = 0
Hechizos(33).PalabrasMagicas = "C�rpse D�m Ex"
Hechizos(33).desc = "Utilizando el obscuro y controvertido arte de la nigromancia eres capaz de reanimar los restos de un guerrero y ligarlos a tu voluntad."
Hechizos(33).HechizeroMsg = "Has invocado al guerrero de ultratumba contra"
Hechizos(33).TargetMsg = "ha invocado al guerrero de ultratumba sobre ti."
Hechizos(33).tipo = 4
Hechizos(33).RemoverParalisis = 0
Hechizos(33).Paraliza = 0
Hechizos(33).ManaRequerido = 650
Hechizos(33).Target = 4
Hechizos(33).NeedStaff = 1
Hechizos(33).StaffAffected = False

Hechizos(34).Nombre = "Por el culo"
Hechizos(34).MaxHp = 9999
Hechizos(34).MinHp = 9999
Hechizos(34).Inmoviliza = 0
Hechizos(34).Invisibilidad = 0
Hechizos(34).Revivir = 0
Hechizos(34).FXgrh = 15
Hechizos(34).loops = 2
Hechizos(34).EffectIndex = 15
Hechizos(34).PalabrasMagicas = "P 'ER LE C�|_0"
Hechizos(34).desc = "Le das por el ano a tu contrincante, quien deber� ponerse luego pomadita."
Hechizos(34).HechizeroMsg = "Se la has dado por el culo a"
Hechizos(34).TargetMsg = "te la ha dado por el culo."
Hechizos(34).tipo = 1
Hechizos(34).RemoverParalisis = 0
Hechizos(34).Paraliza = 0
Hechizos(34).ManaRequerido = 1
Hechizos(34).Target = 1
Hechizos(34).NeedStaff = 0
Hechizos(34).StaffAffected = False

Hechizos(35).Nombre = "flecha"
Hechizos(35).MaxHp = 60
Hechizos(35).MinHp = 0
Hechizos(35).Inmoviliza = 0
Hechizos(35).Invisibilidad = 0
Hechizos(35).Revivir = 0
Hechizos(35).FXgrh = 14
Hechizos(35).loops = 0
Hechizos(35).EffectIndex = 14

Hechizos(35).desc = "flecha"

Hechizos(35).TargetMsg = "te ha atacado con su arco y flecha."
Hechizos(35).tipo = 1
Hechizos(35).RemoverParalisis = 0
Hechizos(35).Paraliza = 0
Hechizos(35).ManaRequerido = 3000
Hechizos(35).Target = 1
Hechizos(35).NeedStaff = 0
Hechizos(35).StaffAffected = False

Hechizos(36).Nombre = "Tormenta Pretoriana"
Hechizos(36).MaxHp = 95
Hechizos(36).MinHp = 0
Hechizos(36).Inmoviliza = 0
Hechizos(36).Invisibilidad = 0
Hechizos(36).Revivir = 0
Hechizos(36).FXgrh = 7
Hechizos(36).loops = 0
Hechizos(36).EffectIndex = 7
Hechizos(36).PalabrasMagicas = "EN VAX ON TAR"
Hechizos(36).desc = "Tormenta Pretoriana"
Hechizos(36).HechizeroMsg = "Has lanzado tormenta sobre"
Hechizos(36).TargetMsg = "lanz� tormenta de fuego sobre vos."
Hechizos(36).tipo = 1
Hechizos(36).RemoverParalisis = 0
Hechizos(36).Paraliza = 0
Hechizos(36).ManaRequerido = 520
Hechizos(36).Target = 3
Hechizos(36).NeedStaff = 0
Hechizos(36).StaffAffected = False

Hechizos(37).Nombre = "Flecha Cazador Pretoriano"
Hechizos(37).MaxHp = 160
Hechizos(37).MinHp = 130
Hechizos(37).Inmoviliza = 0
Hechizos(37).Invisibilidad = 0
Hechizos(37).Revivir = 0
Hechizos(37).FXgrh = 14
Hechizos(37).loops = 0
Hechizos(37).EffectIndex = 14

Hechizos(37).desc = "Flecha Cazador Pretoriano"
Hechizos(37).HechizeroMsg = "Has lanzado una flecha sobre"
Hechizos(37).TargetMsg = "lanz� un flechazo sobre vos."
Hechizos(37).tipo = 1
Hechizos(37).RemoverParalisis = 0
Hechizos(37).Paraliza = 0
Hechizos(37).ManaRequerido = 3000
Hechizos(37).Target = 3
Hechizos(37).NeedStaff = 0
Hechizos(37).StaffAffected = False

Hechizos(38).Nombre = "Remover Invisibilidad"
Hechizos(38).MaxHp = 0
Hechizos(38).MinHp = 0
Hechizos(38).Inmoviliza = 0
Hechizos(38).Invisibilidad = 0
Hechizos(38).Revivir = 0
Hechizos(38).FXgrh = 2
Hechizos(38).loops = 2
Hechizos(38).EffectIndex = 2
Hechizos(38).PalabrasMagicas = "AN ROHL �X MA�O"
Hechizos(38).desc = "Hechizo invocado por el mago pretoriano unicamente"

Hechizos(38).tipo = 0
Hechizos(38).RemoverParalisis = 0
Hechizos(38).Paraliza = 0
Hechizos(38).ManaRequerido = 0
Hechizos(38).Target = 0
Hechizos(38).NeedStaff = 0
Hechizos(38).StaffAffected = False

Hechizos(39).Nombre = "Paralizar NPCs"
Hechizos(39).MaxHp = 0
Hechizos(39).MinHp = 0
Hechizos(39).Inmoviliza = 0
Hechizos(39).Invisibilidad = 0
Hechizos(39).Revivir = 0
Hechizos(39).FXgrh = 8
Hechizos(39).loops = 0
Hechizos(39).EffectIndex = 8
Hechizos(39).PalabrasMagicas = "HOAX MANTRA"
Hechizos(39).desc = "Hechizo invocado por el sacerdote pretoriano unicamente"

Hechizos(39).tipo = 0
Hechizos(39).RemoverParalisis = 0
Hechizos(39).Paraliza = 0
Hechizos(39).ManaRequerido = 0
Hechizos(39).Target = 0
Hechizos(39).NeedStaff = 0
Hechizos(39).StaffAffected = False

Hechizos(40).Nombre = "Bendici�n de Sortilego"
Hechizos(40).MaxHp = 9999
Hechizos(40).MinHp = 9999
Hechizos(40).Inmoviliza = 0
Hechizos(40).Invisibilidad = 0
Hechizos(40).Revivir = 0
Hechizos(40).FXgrh = 9
Hechizos(40).loops = 9
Hechizos(40).EffectIndex = 9
Hechizos(40).PalabrasMagicas = "In Nomine Patris et Fili et Spiritus Sancti, Am�n."
Hechizos(40).desc = "Provoca una profunda curaci�n de las heridas m�s terribles. Restaura entre 12 y 35 puntos de salud."
Hechizos(40).HechizeroMsg = "Has lanzado tu bendici�n de Sortilego sobre"
Hechizos(40).TargetMsg = "te ha lanzado su bendici�n de Sortilego."
Hechizos(40).tipo = 1
Hechizos(40).RemoverParalisis = 0
Hechizos(40).Paraliza = 0
Hechizos(40).ManaRequerido = 1
Hechizos(40).Target = 3
Hechizos(40).NeedStaff = 0
Hechizos(40).StaffAffected = False

Hechizos(41).Nombre = "Desaturdir"
Hechizos(41).MaxHp = 0
Hechizos(41).MinHp = 0
Hechizos(41).Inmoviliza = 0
Hechizos(41).Invisibilidad = 0
Hechizos(41).Revivir = 0
Hechizos(41).FXgrh = 0
Hechizos(41).loops = 0
Hechizos(41).EffectIndex = 0
Hechizos(41).PalabrasMagicas = "AN ASYNC GAM AL�"
Hechizos(41).desc = "Con este conjuro podr�s contrarrestar los nefastos efectos del hechizo aturdir."
Hechizos(41).HechizeroMsg = "Le has quitado su aturdimiento a"
Hechizos(41).TargetMsg = "te ha quitado el aturdimiento."
Hechizos(41).tipo = 2
Hechizos(41).RemoverParalisis = 0
Hechizos(41).Paraliza = 0
Hechizos(41).ManaRequerido = 350
Hechizos(41).Target = 1
Hechizos(41).NeedStaff = 0
Hechizos(41).StaffAffected = False

Hechizos(42).Nombre = "Mimetismo"
Hechizos(42).MaxHp = 0
Hechizos(42).MinHp = 0
Hechizos(42).Inmoviliza = 0
Hechizos(42).Invisibilidad = 0
Hechizos(42).Revivir = 0
Hechizos(42).FXgrh = 0
Hechizos(42).loops = 0
Hechizos(42).EffectIndex = 0
Hechizos(42).PalabrasMagicas = "Cimim Ux Ma�o"
Hechizos(42).desc = "Con este encantamiento adquirir�s temporalmente la apariencia f�sica de otra persona."
Hechizos(42).HechizeroMsg = "Haz adquirido la apariencia de"
Hechizos(42).TargetMsg = "ha adquirido magicamente tu apariencia."
Hechizos(42).tipo = 2
Hechizos(42).RemoverParalisis = 0
Hechizos(42).Paraliza = 0
Hechizos(42).ManaRequerido = 800
Hechizos(42).Target = 3
Hechizos(42).NeedStaff = 0
Hechizos(42).StaffAffected = False

Hechizos(43).Nombre = "El Ojo del Demiurgo"
Hechizos(43).MaxHp = 0
Hechizos(43).MinHp = 0
Hechizos(43).Inmoviliza = 0
Hechizos(43).Invisibilidad = 0
Hechizos(43).Revivir = 0
Hechizos(43).FXgrh = 3
Hechizos(43).loops = 6
Hechizos(43).EffectIndex = 3
Hechizos(43).PalabrasMagicas = "An Ma�o naq v�k�"
Hechizos(43).desc = "Un haz de luz divina te permitir� momentaneamente ver aquello que resulta invisible a los ojos mortales."
Hechizos(43).HechizeroMsg = "Has invocado el Ojo del Demiurgo para ver lo invisible."
Hechizos(43).TargetMsg = "te ha invocado el Ojo del Demiurgo para ver lo invisible."
Hechizos(43).tipo = 2
Hechizos(43).RemoverParalisis = 0
Hechizos(43).Paraliza = 0
Hechizos(43).ManaRequerido = 1
Hechizos(43).Target = 4
Hechizos(43).NeedStaff = 0
Hechizos(43).StaffAffected = False

Hechizos(44).Nombre = "Flecha Elfica"
Hechizos(44).MaxHp = 35
Hechizos(44).MinHp = 30
Hechizos(44).Inmoviliza = 0
Hechizos(44).Invisibilidad = 0
Hechizos(44).Revivir = 0
Hechizos(44).FXgrh = 0
Hechizos(44).loops = 2
Hechizos(44).EffectIndex = 0

Hechizos(44).desc = "Una flecha finamente tallada por los antiguos elfos, un suave resplandor la rodea."
Hechizos(44).HechizeroMsg = "Has lanzado una flecha elfica sobre"
Hechizos(44).TargetMsg = "ha lanzado una flecha elfica sobre ti."
Hechizos(44).tipo = 1
Hechizos(44).RemoverParalisis = 0
Hechizos(44).Paraliza = 0
Hechizos(44).ManaRequerido = 10
Hechizos(44).Target = 1
Hechizos(44).NeedStaff = 0
Hechizos(44).StaffAffected = False

Hechizos(45).Nombre = "Drenar"
Hechizos(45).MaxHp = 40
Hechizos(45).MinHp = 35
Hechizos(45).Inmoviliza = 0
Hechizos(45).Invisibilidad = 0
Hechizos(45).Revivir = 0
Hechizos(45).FXgrh = 35
Hechizos(45).loops = 2
Hechizos(45).EffectIndex = 35

Hechizos(45).desc = "Un hechizo maligno que drena del esp�ritu del enemigo su salud, y se la entrega al conjurador."
Hechizos(45).HechizeroMsg = "Has drenado vida de"
Hechizos(45).TargetMsg = "te ha drenado vida."
Hechizos(45).tipo = 1
Hechizos(45).RemoverParalisis = 0
Hechizos(45).Paraliza = 0
Hechizos(45).ManaRequerido = 10
Hechizos(45).Target = 1
Hechizos(45).NeedStaff = 0
Hechizos(45).StaffAffected = False

Hechizos(46).Nombre = "Invocar Mascota"
Hechizos(46).MaxHp = 0
Hechizos(46).MinHp = 0
Hechizos(46).Inmoviliza = 0
Hechizos(46).Invisibilidad = 0
Hechizos(46).Revivir = 0
Hechizos(46).FXgrh = 7
Hechizos(46).loops = 0
Hechizos(46).EffectIndex = 7
Hechizos(46).PalabrasMagicas = "Ts�lo Kai'Tor"
Hechizos(46).desc = "Podr�s lograr que tu m�s lejana mascota regrese a t�."
Hechizos(46).HechizeroMsg = "Has invocado una mascota contra"
Hechizos(46).TargetMsg = "ha invocado una mascota sobre ti."
Hechizos(46).tipo = 4
Hechizos(46).RemoverParalisis = 0
Hechizos(46).Paraliza = 0
Hechizos(46).ManaRequerido = 0
Hechizos(46).Target = 4
Hechizos(46).NeedStaff = 0
Hechizos(46).StaffAffected = False

End Sub

Public Sub Cargar_RutasMap(ByRef MapReader As clsIniReader, ByVal MapIndex As Integer)

' @ Carga las rutas de un mapa.

Dim loopC As Long
Dim loopY As Long
Dim rutaIndex As Byte

For loopC = 1 To 100
    For loopY = 1 To 100
        With MapData(MapIndex, loopC, loopY)
             rutaIndex = Val(MapReader.GetValue(CStr(loopC) & "," & CStr(loopY), "RutaIndex"))
             
             'Hay ruta?
             If rutaIndex <> 0 Then
                .Rutas(rutaIndex) = Val(MapReader.GetValue(CStr(loopC) & "," & CStr(loopY), "Direccion"))
             End If
             
        End With
    Next loopY
Next loopC

End Sub


Public Sub Enviar_Da�oAUsuario(ByVal userIndex As Integer, ByVal Da�o As Integer)

' @ Envia crear da�o a un char.

Dim UserCharIndex   As Integer
Dim PacketToSend    As String

'Obtengo el char.
UserCharIndex = UserList(userIndex).Char.CharIndex

'Preparo el paquete.
PacketToSend = mod_DunkanProtocol.Send_CreateDamage(UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y, Da�o)

'Envio.
SendData SendTarget.ToPCArea, userIndex, PacketToSend

End Sub

Public Sub Enviar_Da�oANpc(ByVal NpcIndex As Integer, ByVal Da�o As Integer)

' @ Envia crear da�o a un NPC.

Dim NpcCharIndex    As Integer
Dim PacketToSend    As String

'Get the character,
NpcCharIndex = Npclist(NpcIndex).Char.CharIndex

'Prepare the outgoing Data
PacketToSend = mod_DunkanProtocol.Send_CreateDamage(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Da�o)

'Send
SendData SendTarget.ToNPCArea, NpcIndex, PacketToSend

End Sub

Public Sub Enviar_HechizoAUser(ByVal Attacker As Integer, ByVal victim As Integer, ByVal EffectIndex As Integer, ByVal FXLoops As Integer)

' @ Envia crear hechizo/particula a un usuario.

Dim ACharIndex      As Integer  '< AttackerCharIndex
Dim VCharIndex      As Integer  '< VictimCharIndex
Dim PackageToSend   As String   '< Paquete a enviar.

'Obtengo los chars.
ACharIndex = UserList(Attacker).Char.CharIndex
VCharIndex = UserList(victim).Char.CharIndex

'Preparo el paquete.
PackageToSend = mod_DunkanProtocol.Send_CreateSpell(ACharIndex, VCharIndex, EffectIndex, FXLoops)

'Y envio :)
SendData SendTarget.ToPCArea, victim, PackageToSend

End Sub

Public Sub Enviar_HechizoANpc(ByVal AttackerUser As Integer, ByVal VictimNpc As Integer, ByVal EffectIndex As Integer, ByVal FXLoops As Integer)

' @ Envia crear hechizo/particula a un usuario.

Dim ACharIndex      As Integer  '< AttackerCharIndex
Dim VCharIndex      As Integer  '< VictimCharIndex
Dim PackageToSend   As String   '< Paquete a enviar.

'Obtengo los chars.
ACharIndex = UserList(AttackerUser).Char.CharIndex
VCharIndex = Npclist(VictimNpc).Char.CharIndex

'Preparo el paquete.
PackageToSend = mod_DunkanProtocol.Send_CreateSpell(ACharIndex, VCharIndex, EffectIndex, FXLoops)

'Y envio :)
SendData SendTarget.ToNPCArea, VictimNpc, PackageToSend

End Sub

Public Sub Enviar_FlechaANpc(ByVal attackerCharIndex As Integer, ByVal VictimNpc As Integer, ByVal GrhIndex As Integer)

' @ Envia crear una flecha sobre NPC.

Dim CharNpc     As Integer
Dim Package     As String

'Obtengo el char del npc
CharNpc = Npclist(VictimNpc).Char.CharIndex

'Prepara el paquete.
Package = mod_DunkanProtocol.Send_CreateArrow(attackerCharIndex, CharNpc, GrhIndex)

SendData SendTarget.ToNPCArea, VictimNpc, Package

End Sub

Public Sub Enviar_FlechaAUser(ByVal attackerCharIndex As Integer, ByVal VictimUserIndex As Integer, ByVal GrhIndex As Integer)

' @ Envia crear una flecha sobre usuario.

Dim CharIndexVictim As Integer
Dim PackageSend     As String

'Obtengo el char de la victima
CharIndexVictim = UserList(VictimUserIndex).Char.CharIndex

'Prepara el paquete.
PackageSend = mod_DunkanProtocol.Send_CreateArrow(attackerCharIndex, CharIndexVictim, GrhIndex)

SendData SendTarget.ToPCArea, VictimUserIndex, PackageSend

End Sub

Public Sub Preparar_CharNpc(ByRef OriginalChar As Char, ByRef CharNpc As Char)

' @ Prepara el equipamiento de un NPC.

With OriginalChar

     'Equipa arma?
     If .WeaponAnim <> 0 Then
        CharNpc.WeaponAnim = ObjData(.WeaponAnim).WeaponAnim
     Else
        CharNpc.WeaponAnim = -1
     End If
     
     'Equipa escudo?
     If .ShieldAnim <> 0 Then
        CharNpc.ShieldAnim = ObjData(.ShieldAnim).ShieldAnim
     Else
        CharNpc.ShieldAnim = -1
     End If
     
     'Equipa casco?
     If .CascoAnim <> 0 Then
        CharNpc.CascoAnim = ObjData(.CascoAnim).CascoAnim
     Else
        CharNpc.CascoAnim = -1
     End If
     
End With

End Sub
