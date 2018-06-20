Attribute VB_Name = "Admin"
 'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A clPARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not,  can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'

'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type

Public Apuestas As tAPuestas

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloUserPuedeAtacar As Long
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public Puerto As Integer

Public BootDelBackUp As Byte
Public Lloviendo As Boolean
Public DeNoche As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    VersionOK = (Ver = ULTIMAVERSION)
End Function


Sub WorldSave()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************



End Sub

Public Sub PurgarPenas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    UserList(UserIndex).Counters.Pena = Minutos
    
    
    Call WarpUserChar(UserIndex, Prision.map, Prision.X, Prision.Y, True)
    
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(UserIndex).flags.Traveling = 1 Then
        UserList(UserIndex).flags.Traveling = 0
        UserList(UserIndex).Counters.goHome = 0
        Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
    End If
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"
    End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    BANCheck = (Val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    If MD5ClientesActivado = 1 Then
        For i = 0 To UBound(MD5s)
            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function
            End If
        Next i
        MD5ok = False
    Else
        MD5ok = True
    End If

End Function

Public Sub MD5sCarga()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim loopC As Integer
    
    MD5ClientesActivado = Val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
    
    If MD5ClientesActivado = 1 Then
        ReDim MD5s(Val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
        For loopC = 0 To UBound(MD5s)
            MD5s(loopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (loopC + 1))
            MD5s(loopC) = txtOffset(hexMd52Asc(MD5s(loopC)), 55)
        Next loopC
    End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    BanIps.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Dale As Boolean
    Dim loopC As Long
    
    Dale = True
    loopC = 1
    Do While loopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(loopC) <> ip)
        loopC = loopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = loopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim N As Long
    
    N = BanIpBuscar(ip)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

End Function

Public Sub BanIpGuardar()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ArchivoBanIp As String
    Dim ArchN As Long
    Dim loopC As Long
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For loopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(loopC)
    Next loopC
    
    Close #ArchN

End Sub

Public Sub BanIpCargar()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanIp As String
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    Do While BanIps.Count > 0
        BanIps.Remove 1
    Loop
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop
    
    Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Static Andando As Boolean
    Static Contador As Long
    Dim Tmp As Boolean
    
    Contador = Contador + 1
    
    If Contador >= 10 Then
        Contador = 0
        Tmp = EstadisticasWeb.EstadisticasAndando()
        
        If Andando = False And Tmp = True Then
            Call InicializaEstadisticas
        End If
        
        Andando = Tmp
    End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************

    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/02/07
'
'***************************************************

    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
            
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(reason) & " " & Date & " " & Time)
                        
                        If (userPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call LogBan(tUser, bannerUserIndex, reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
            
            'Ponemos el flag de ban a 1
            UserList(tUser).flags.Ban = 1
            
            If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If
            
            Call LogGM(.Name, "BAN a " & UserName)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            cantPenas = Val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(reason) & " " & Date & " " & Time)
            
            Call CloseSocket(tUser)
        End If
    End With
End Sub

