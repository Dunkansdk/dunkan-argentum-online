Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal victimIndex As Integer, ByVal AttackerIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Ahora no te vuelve cirminal por matar un atacable
'***************************************************

    Dim DaExp As Integer
    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(victimIndex).Stats.ELV) * 2
    
    With UserList(AttackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        If TriggerZonaPelea(victimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
        
            ' Es legal matarlo si estaba en atacable
            If UserList(victimIndex).flags.AtacablePor <> AttackerIndex Then
                EraCriminal = criminal(AttackerIndex)
                
                With .Reputacion
                    If Not criminal(victimIndex) Then
                        .AsesinoRep = .AsesinoRep + vlASESINO * 2
                        If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                        .BurguesRep = 0
                        .NobleRep = 0
                        .PlebeRep = 0
                    Else
                        .NobleRep = .NobleRep + vlNoble
                        If .NobleRep > MAXREP Then .NobleRep = MAXREP
                    End If
                End With
                
                If criminal(AttackerIndex) Then
                    If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
                Else
                    If EraCriminal Then Call RefreshCharStatus(AttackerIndex)
                End If
            End If
        End If
        
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
        'Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
        'Call WriteConsoleMsg(VictimIndex, "¡" & .name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, victimIndex, DaExp)
        Call WriteMultiMessage(victimIndex, eMessages.UserKill, AttackerIndex)
        
        'Call UserDie(VictimIndex)
        Call FlushBuffer(victimIndex)
        
        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(victimIndex).Name)
    End With
End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        
        If .flags.Navegando = 1 Then
            Call ToogleBoatBody(UserIndex)
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.Head = .OrigChar.Head
        End If
        
        If .flags.Traveling Then
            .flags.Traveling = 0
            .Counters.goHome = 0
            Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateMana(UserIndex)
    End With
End Sub

Public Sub ToogleBoatBody(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 13/01/2010
'Gives boat body depending on user alignment.
'***************************************************

    Dim Ropaje As Integer
    
    With UserList(UserIndex)
        
 
        .Char.Head = 0
        
        ' Barco de armada
        If .Faccion.ArmadaReal = 1 Then
            .Char.body = iFragataReal
            
        ' Barco de caos
        ElseIf .Faccion.FuerzasCaos = 1 Then
            .Char.body = iFragataCaos
        
        'Barcos neutrales
        Else
            Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
            
            If criminal(UserIndex) Then
                Select Case Ropaje
                    Case iBarca
                        .Char.body = iBarcaPk
                    
                    Case iGalera
                        .Char.body = iGaleraPk
                    
                    Case iGaleon
                        .Char.body = iGaleonPk
                End Select
            Else
                Select Case Ropaje
                    Case iBarca
                        .Char.body = iBarcaCiuda
                    
                    Case iGalera
                        .Char.body = iGaleraCiuda
                    
                    Case iGaleon
                        .Char.body = iGaleonCiuda
                End Select
            End If
        End If
        
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
    End With

End Sub

Public Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
    End With
End Sub

Public Function GetWeaponAnim(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Integer
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/29/10
'
'***************************************************
    Dim Tmp As Integer

    With UserList(UserIndex)
        Tmp = ObjData(objIndex).WeaponRazaEnanaAnim
            
        If Tmp > 0 Then
            If .Raza = eRaza.Enano Or .Raza = eRaza.Gnomo Then
                GetWeaponAnim = Tmp
                Exit Function
            End If
        End If
        
        GetWeaponAnim = ObjData(objIndex).WeaponAnim
    End With
End Function


Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)
'*************************************************
'Author: Unknown
'Last modified: 08/01/2009
'08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
'*************************************************

On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
        CharList(.Char.CharIndex) = 0
        
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        
        ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        End If
        
        Call QuitarUser(UserIndex, .Pos.map)
        
        MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)
End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/07/2009
'Refreshes the status and tag of UserIndex.
'04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
'*************************************************
    Dim ClanTag As String
    Dim NickColor As Byte
    
    With UserList(UserIndex)
        
        NickColor = GetNickColor(UserIndex)
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Call ToogleBoatBody(UserIndex)
            End If
            
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

Public Function GetNickColor(ByVal UserIndex As Integer) As Byte
'*************************************************
'Author: ZaMa
'Last modified: 15/01/2010
'
'*************************************************
    
    With UserList(UserIndex)
        
        If criminal(UserIndex) Then
            GetNickColor = eNickColor.ieCriminal
        Else
            GetNickColor = eNickColor.ieCiudadano
        End If
        
        If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable
    End With
    
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, _
        ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ButIndex As Boolean = False)
'*************************************************
'Author: Unknown
'Last modified: 15/01/2010
'23/07/2009: Budi - Ahora se envía el nick
'15/01/2010: ZaMa - Ahora se envia el color del nick.
'*************************************************

On Error GoTo Errhandler

    Dim CharIndex As Integer
    Dim ClanTag As String
    Dim NickColor As Byte
    Dim UserName As String
    Dim Privileges As Byte
    
1    With UserList(UserIndex)
    
            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
2                CharIndex = NextOpenCharIndex
3                .Char.CharIndex = CharIndex
4                CharList(CharIndex) = UserIndex
            End If
            
            'Place character on map if needed
5            If toMap Then MapData(map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            If Not toMap Then

6                NickColor = GetNickColor(UserIndex)
7                Privileges = .flags.Privilegios
                
                'Preparo el nick
                If .showName Then
8                    UserName = .Name
                    
9                    If .flags.EnConsulta Then
11                        UserName = UserName & " " & TAG_CONSULT_MODE
13                   Else
                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then _
                                UserName = UserName & " <" & ClanTag & ">"
                        Else
                            If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else
                                If LenB(ClanTag) <> 0 Then _
                                    UserName = UserName & " <" & ClanTag & ">"
                            End If
                        End If
                    End If
                End If
            
14                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, _
                            .Char.CharIndex, X, Y, _
                            .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                            UserName, NickColor, Privileges)
            Else
                'Hide the name and clan - set privs as normal user
15                 Call AgregarUser(UserIndex, .Pos.map, ButIndex)
            End If

    End With
Exit Sub

Errhandler:
    LogError ("MakeUserChar: Linea " & Erl() & " num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    'Call CloseSocket(UserIndex)
End Sub



Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 _
                    Or UserList(UserIndex).flags.Vuela = 1
End Function

Sub MoveUserChar2(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
'*************************************************
'Author: Unknown
'Last modified: 13/07/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
'13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
'13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
'*************************************************
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim CasPerPos As WorldPos
    
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
               
            CasperIndex = MapData(UserList(UserIndex).Pos.map, nPos.X, nPos.Y).UserIndex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                ' Los admins invisibles no pueden patear caspers
                If Not (UserList(UserIndex).flags.AdminInvisible = 1) Then
                    
                    If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
                        End If
                    End If
    
                    CasperHeading = InvertHeading(nHeading)
                    CasPerPos = UserList(CasperIndex).Pos
                    Call HeadtoPos(CasperHeading, CasPerPos)
    
                    With UserList(CasperIndex)
                        
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then _
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, CasPerPos.X, CasPerPos.Y))
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                            
                        'Update map and user pos
                        .Pos = CasPerPos
                        .Char.heading = CasperHeading
                        MapData(.Pos.map, CasPerPos.X, CasPerPos.Y).UserIndex = CasperIndex
                    
                    End With
                
                    'Actualizamos las áreas de ser necesario
                    Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If

            
            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not UserList(UserIndex).flags.AdminInvisible = 1 Then _
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not ((UserList(UserIndex).flags.AdminInvisible = 1) And CasperIndex <> 0) Then
            Dim oldUserIndex As Integer
            
            oldUserIndex = MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex
            
            ' Si no hay intercambio de pos con nadie
            If oldUserIndex = UserIndex Then
                MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
            End If
            
            UserList(UserIndex).Pos = nPos
            UserList(UserIndex).Char.heading = nHeading
            MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
            
            'Actualizamos las áreas de ser necesario
            Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
        Else
            Call WritePosUpdate(UserIndex)
        End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

On Error GoTo ErrMove

'*************************************************
'Author: Unknown
'Last modified: 13/07/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
'13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
'13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
'*************************************************
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim CasPerPos As WorldPos
    
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
        
    If MoveToLegalPos(UserList(UserIndex).Pos.map, nPos.X, nPos.Y, sailing, Not sailing) Then
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
               
            CasperIndex = MapData(UserList(UserIndex).Pos.map, nPos.X, nPos.Y).UserIndex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                ' Los admins invisibles no pueden patear caspers
                If Not (UserList(UserIndex).flags.AdminInvisible = 1) Then
                    
                    If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
                        End If
                    End If
    
                    CasperHeading = InvertHeading(nHeading)
                    CasPerPos = UserList(CasperIndex).Pos
                    Call HeadtoPos(CasperHeading, CasPerPos)
    
                    With UserList(CasperIndex)
                        
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then _
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, CasPerPos.X, CasPerPos.Y))
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                            
                        'Update map and user pos
                        .Pos = CasPerPos
                        .Char.heading = CasperHeading
5                        MapData(.Pos.map, CasPerPos.X, CasPerPos.Y).UserIndex = CasperIndex
                    
                    End With
                
                    'Actualizamos las áreas de ser necesario
8                    Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If

            
            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not UserList(UserIndex).flags.AdminInvisible = 1 Then _
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not ((UserList(UserIndex).flags.AdminInvisible = 1) And CasperIndex <> 0) Then
            Dim oldUserIndex As Integer
            
            oldUserIndex = MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex
            
            ' Si no hay intercambio de pos con nadie
            If oldUserIndex = UserIndex Then
                MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
            End If
            
            UserList(UserIndex).Pos = nPos
            UserList(UserIndex).Char.heading = nHeading
            MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
            
            'Actualizamos las áreas de ser necesario
14            Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
        Else
            Call WritePosUpdate(UserIndex)
        End If

    Else
        Call WritePosUpdate(UserIndex)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
        
    ' maTih.-
    Exit Sub
ErrMove:
    Debug.Print "Error en MoveUserChar [LINE:" & Erl() & "]"
End Sub




Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim loopC As Long
    
    For loopC = 1 To MAXCHARS
        If CharList(loopC) = 0 Then
            NextOpenCharIndex = loopC

            NumChars = NumChars + 1
            
            If loopC > LastChar Then _
                LastChar = loopC
            
            Exit Function
        End If
    Next loopC
End Function

Function NextOpenUser() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim loopC As Long
    
    For loopC = 1 To MaxUsers + 1
        If loopC > MaxUsers Then Exit For
        If (UserList(loopC).ConnID = -1 And UserList(loopC).flags.UserLogged = False) Then Exit For
    Next loopC
    
    NextOpenUser = loopC
End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GuildI As Integer
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.Minhit & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).Minhit & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.Minhit & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        

        
#If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim tempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        tempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.GLD & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.Clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        

    End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String
    
    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        End If

        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)

        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To .CurrentInventorySlots
            If .Invent.Object(j).objIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).objIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    Dim j As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub


Public Function PuedeAcuchillar(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 25/01/2010 (ZaMa)
'
'***************************************************
    
    With UserList(UserIndex)
        If .Clase = eClass.Pirat Then
            If .Invent.WeaponEqpObjIndex > 0 Then
                PuedeAcuchillar = (ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1)
            End If
        End If
    End With
    
End Function



''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 12/01/2010 (ZaMa)
'04/15/2008: NicoNZ - Ahora se resetea el counter del invi
'13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
'27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
'21/07/2009: Marco - Al morir se desactiva el comercio seguro.
'16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
'27/11/2009: Budi - Al morir envia los atributos originales.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
'************************************************
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    With UserList(UserIndex)
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1

        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If

        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(UserIndex)
        End If
        
        '<<<< Invisible >>>>
        If Server_Info.Invisibilidad Then
            If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .flags.invisible = 0
                .Counters.TiempoOculto = 0
                .Counters.Invisibilidad = 0
                
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
            End If
        End If

        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
        End If

        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            .Char.body = iFragataFantasmal
        End If
        
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateMana(UserIndex)
        
        'Server Auto-Respawn
        If Server_Info.AutoRespawn Then
           'Setea el tiempo
            .Auto_Respawn = Server_Info.TiempoRespawn
         End If
        
         If (Server_Info.DeathMathc = False) Then
            Call mod_DunkanGeneral.Informar_Muerte(UserIndex)
         End If
        
        '<<Cerramos comercio seguro>>

    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If EsNewbie(Muerto) Then Exit Sub
    
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        If criminal(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name
                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                    .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
            End If
            
            If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                .Faccion.Reenlistadas = 200  'jaja que trucho
                
                'con esto evitamos que se vuelva a reenlistar
            End If
        Else
            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name
                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                    .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            End If
        End If
        
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
            .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
    End With
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
    Dim loopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.map = Pos.map
    nPos.X = 0
    nPos.Y = 0
    
    Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If loopC > 15 Then
            Exit Do
        End If
        
        For tY = Pos.Y - loopC To Pos.Y + loopC
            For tX = Pos.X - loopC To Pos.X + loopC
                
                If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.map, tX, tY).ObjInfo.objIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.objIndex <> Obj.objIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.map, tX, tY).ObjInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = Pos.X + loopC
                        tY = Pos.Y + loopC
                    End If
                End If
            
            Next tX
        Next tY
        
        loopC = loopC + 1
    Loop
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal FX As Boolean, Optional ByVal Teletransported As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 13/11/2009
'15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
'13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
'**************************************************************
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    With UserList(UserIndex)
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        Call WriteRemoveAllDialogs(UserIndex)
        
        OldMap = .Pos.map
        OldX = .Pos.X
        OldY = .Pos.Y

        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
        
        If OldMap <> map Then
             
            'Cambió de mapa AND atacó un viajante.
            .AtacoViajante = 0
            Call WriteChangeMap(UserIndex, map, MapInfo(.Pos.map).MapVersion)
            Call WritePlayMidi(UserIndex, Val(ReadField(1, MapInfo(map).Music, 45)))
            
            'Update new Map Users
            MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
        
            'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
            Dim nextMap, previousMap As Boolean
            nextMap = IIf(distanceToCities(map).distanceToCity(.Hogar) >= 0, True, False)
            previousMap = IIf(distanceToCities(.Pos.map).distanceToCity(.Hogar) >= 0, True, False)

            If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
                .flags.lastMap = .Pos.map
            ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
                .flags.lastMap = 0
            ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
                .flags.lastMap = .flags.lastMap
            End If
        
        End If
        
        .Pos.X = X
        .Pos.Y = Y
        .Pos.map = map
        
        Call MakeUserChar(True, map, UserIndex, map, X, Y)
        Call WriteUserCharIndexInServer(UserIndex)
        
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SetInvisible(UserIndex, .Char.CharIndex, True)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
        
        If Teletransported Then
            If .flags.Traveling = 1 Then
                .flags.Traveling = 0
                .Counters.goHome = 0
                Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
            End If
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If

      
    End With
End Sub




''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/04/08 (NicoNZ)
'
'***************************************************
    Dim isNotVisible As Boolean
    Dim HiddenPirat As Boolean

    With UserList(UserIndex)
        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).Pk, IntervaloCerrarConexion, 0)
            
            isNotVisible = (.flags.Oculto Or .flags.invisible)
            If isNotVisible Then
                .flags.invisible = 0
                
                If .flags.Oculto Then
                    If .flags.Navegando = 1 Then
                        If .Clase = eClass.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
                            Call ToogleBoatBody(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                                NingunEscudo, NingunCasco)
                            HiddenPirat = True
                        End If
                    End If
                End If
                
                .flags.Oculto = 0
                
                ' Para no repetir mensajes
                If Not HiddenPirat Then Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                Call SetInvisible(UserIndex, .Char.CharIndex, False)

            End If
            
            If .flags.Traveling = 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
            End If
            
            Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ViejoNick As String
    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
    End If
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
#If ConUpTime Then
        Dim TempSecs As Long
        Dim tempStr As String
        TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
        tempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
    
    End If
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim CharFile As String
    
On Error Resume Next
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/02/2010
'Nacho: Actualiza el tag al cliente
'21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP

            
            If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0

        End If
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, ByVal UserCharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim sndNick As String

With UserList(UserIndex)
    If invisible Then
    .OldBody = .Char.body
    .OldHead = .Char.Head
    Else
    .Char.body = .OldBody
    .Char.Head = .OldHead
    End If
    
    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(UserCharIndex, invisible, CInt(.Counters.Invisibilidad)))
    
    sndNick = .Name
    
    If invisible Then
        sndNick = sndNick & " " & TAG_USER_INVISIBLE
    Else

    End If
    
    Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, UserIndex, PrepareMessageCharacterChangeNick(UserCharIndex, sndNick))
End With
End Sub

Public Sub SetConsulatMode(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/06/10
'
'***************************************************

Dim sndNick As String

With UserList(UserIndex)
    sndNick = .Name
    
    If .flags.EnConsulta Then
        sndNick = sndNick & " " & TAG_CONSULT_MODE
    Else

    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))
End With
End Sub

Public Function IsArena(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 10/11/2009
'Returns true if the user is in an Arena
'**************************************************************
    IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)
End Function


Public Function GetDireccion(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/11/2009
'Devuelve la direccion hacia donde esta el usuario
'**************************************************************
    Dim X As Integer
    Dim Y As Integer
    
    X = UserList(UserIndex).Pos.X - UserList(OtherUserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    
    If X = 0 And Y > 0 Then
        GetDireccion = "Sur"
    ElseIf X = 0 And Y < 0 Then
        GetDireccion = "Norte"
    ElseIf X > 0 And Y = 0 Then
        GetDireccion = "Este"
    ElseIf X < 0 And Y = 0 Then
        GetDireccion = "Oeste"
    ElseIf X > 0 And Y < 0 Then
        GetDireccion = "NorEste"
    ElseIf X < 0 And Y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf X > 0 And Y > 0 Then
        GetDireccion = "SurEste"
    ElseIf X < 0 And Y > 0 Then
        GetDireccion = "SurOeste"
    End If

End Function

Public Function SameFaccion(ByVal UserIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/11/2009
'Devuelve True si son de la misma faccion
'**************************************************************
    SameFaccion = (esCaos(UserIndex) And esCaos(OtherUserIndex)) Or _
                    (esArmada(UserIndex) And esArmada(OtherUserIndex))
End Function

Public Function FarthestPet(ByVal UserIndex As Integer) As Integer
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/11/2009
'Devuelve el indice de la mascota mas lejana.
'**************************************************************
On Error GoTo Errhandler
    
    Dim PetIndex As Integer
    Dim Distancia As Integer
    Dim OtraDistancia As Integer
    
    With UserList(UserIndex)
        If .NroMascotas = 0 Then Exit Function
    
        For PetIndex = 1 To MAXMASCOTAS
            ' Solo pos invocar criaturas que exitan!
            If .MascotasIndex(PetIndex) > 0 Then
                ' Solo aplica a mascota, nada de elementales..
                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        ' Por si tiene 1 sola mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                    Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                    Else
                        ' La distancia de la proxima mascota
                        OtraDistancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                        Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                        ' Esta mas lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex
                        End If
                    End If
                End If
            End If
        Next PetIndex
    End With

    Exit Function
    
Errhandler:
    Call LogError("Error en FarthestPet")
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal UserIndex As Integer, ByVal Skill As Byte, ByVal Allocation As Boolean)
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 11/20/2009
'
'*************************************************

With UserList(UserIndex).Stats
    If .UserSkills(Skill) < MAXSKILLPOINTS Then
        If Allocation Then
            .ExpSkills(Skill) = 0
        Else
            .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)
        End If
        
        .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
    Else
        .ExpSkills(Skill) = 0
        .EluSkills(Skill) = 0
    End If
End With

End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, ByVal objIndex As Integer, ByVal Amount As Long) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks Wether the user has the required amount of items in the inventory or not
'**************************************************************

    Dim Slot As Long
    Dim ItemInvAmount As Long
    
    For Slot = 1 To UserList(UserIndex).CurrentInventorySlots
        ' Si es el item que busco
        If UserList(UserIndex).Invent.Object(Slot).objIndex = objIndex Then
            ' Lo sumo a la cantidad total
            ItemInvAmount = ItemInvAmount + UserList(UserIndex).Invent.Object(Slot).Amount
        End If
    Next Slot

    HasEnoughItems = Amount <= ItemInvAmount
End Function



Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If UserList(UserIndex).Invent.MochilaEqpObjIndex > 0 Then
    getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(UserIndex).Invent.MochilaEqpObjIndex).MochilaType * 5 '5=slots por fila, hacer constante
Else
    getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
End If
End Function

Public Sub goHome(ByVal UserIndex As Integer)
Dim Distance As Integer
Dim tiempo As Long

With UserList(UserIndex)
    If .flags.Muerto = 1 Then
        If .flags.lastMap = 0 Then
            Distance = distanceToCities(.Pos.map).distanceToCity(.Hogar)
        Else
            Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY
        End If
        
        tiempo = (Distance + 1) * 30 'segundos
        
        .Counters.goHome = tiempo / 6 'Se va a chequear cada 6 segundos.
        
        .flags.Traveling = 1

        Call WriteMultiMessage(UserIndex, eMessages.Home, Distance, tiempo, , MapInfo(Ciudades(.Hogar).map).Name)
    Else
        Call WriteConsoleMsg(UserIndex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End With
End Sub

Public Function ToogleToAtackable(ByVal UserIndex As Integer, ByVal OwnerIndex As Integer, Optional ByVal StealingNpc As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 15/01/2010
'Change to Atackable mode.
'***************************************************
    
    Dim AtacablePor As Integer
    
    With UserList(UserIndex)
    
        AtacablePor = .flags.AtacablePor
            
        If AtacablePor > 0 Then
            ' Intenta robar un npc
            If StealingNpc Then
                ' Puede atacar el mismo npc que ya estaba robando, pero no una nuevo.
                If AtacablePor <> OwnerIndex Then
                    Call WriteConsoleMsg(UserIndex, "No puedes atacar otra criatura con dueño hasta que haya terminado tu castigo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            ' Esta atacando a alguien en estado atacable => Se renueva el timer de atacable
            Else
                ' Renovar el timer
                Call IntervaloEstadoAtacable(UserIndex, True)
                ToogleToAtackable = True
                Exit Function
            End If
        End If
        
        .flags.AtacablePor = OwnerIndex
    
        ' Actualizar clientes
        Call RefreshCharStatus(UserIndex)
        
        ' Inicializar el timer
        Call IntervaloEstadoAtacable(UserIndex, True)
        
        ToogleToAtackable = True
        
    End With
    
End Function

Public Sub setHome(ByVal UserIndex As Integer, ByVal newHome As eCiudad, ByVal NpcIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 30/04/2010
'30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
'***************************************************
    If newHome < eCiudad.cUllathorpe Or newHome > cArghal Then Exit Sub
    UserList(UserIndex).Hogar = newHome
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
End Sub
