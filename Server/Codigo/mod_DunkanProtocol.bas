Attribute VB_Name = "mod_DunkanProtocol"
' @ Diseñado & Implementado por maTih.-
' @ Dunkan AO Protocol.

Option Explicit

Public Function Send_CreateSpell(ByVal CharIndex As Integer, ByVal OtherCharIndex As Integer, ByVal EffectIndex As Integer, ByVal FXLoops As Integer) As String

' @ Envia el paquete para crear hechizos en chars.

With auxiliarBuffer
     
     Call .WriteByte(ServerPacketID.CreateSpell)
     Call .WriteInteger(CharIndex)
     Call .WriteInteger(OtherCharIndex)
     Call .WriteInteger(EffectIndex)
     Call .WriteInteger(FXLoops)
     
     Send_CreateSpell = .ReadASCIIStringFixed(.length)
     
End With

End Function

Public Function Send_CreateMeditation(ByVal CharIndex As Integer, ByVal EffectIndex As Integer, ByVal FXLoops As Integer) As String

' @ Envia el paquete para crear meditaciones en chars.

With auxiliarBuffer
     
     Call .WriteByte(ServerPacketID.CreateMeditation)
     Call .WriteInteger(CharIndex)
     Call .WriteInteger(EffectIndex)
     Call .WriteInteger(FXLoops)
     
     Send_CreateMeditation = .ReadASCIIStringFixed(.length)
     
End With

End Function

Public Function Send_CreateArrow(ByVal CharIndex As Integer, ByVal OtherCharIndex As Integer, ByVal GrhIndex As Integer) As String

' @ Envia el paquete para crear flechas en chars.

With auxiliarBuffer
     Call .WriteByte(ServerPacketID.CreateArrow)
     Call .WriteInteger(CharIndex)
     Call .WriteInteger(OtherCharIndex)
     Call .WriteInteger(GrhIndex)
     
     Send_CreateArrow = .ReadASCIIStringFixed(.length)
     
End With

End Function

Public Function Send_CreateDamage(ByVal TileX As Byte, ByVal TileY As Byte, ByVal Damage As Integer) As String

' @ Envia el paquete para crear daño en una posición de X mapa.

With auxiliarBuffer
     Call .WriteByte(ServerPacketID.CreateDamage)
     Call .WriteByte(TileX)
     Call .WriteByte(TileY)
     Call .WriteInteger(Damage)
     
     Send_CreateDamage = .ReadASCIIStringFixed(.length)
     
End With

End Function




Public Sub HandleDragInventory(ByVal UserIndex As Integer)

' @ Author : Amraphen.
'            Drag&Drop de objetos en el inventario.

Dim ObjSlot1 As Byte
Dim ObjSlot2 As Byte

Dim tmpUserObj As UserOBJ
 
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
    With UserList(UserIndex)
        'Leemos el paquete
        Call .incomingData.ReadByte
       
        ObjSlot1 = .incomingData.ReadByte
        ObjSlot2 = .incomingData.ReadByte
       
        'Cambiamos si alguno es un anillo
        If .Invent.AnilloEqpSlot = ObjSlot1 Then
            .Invent.AnilloEqpSlot = ObjSlot2
        ElseIf .Invent.AnilloEqpSlot = ObjSlot2 Then
            .Invent.AnilloEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un armor
        If .Invent.ArmourEqpSlot = ObjSlot1 Then
            .Invent.ArmourEqpSlot = ObjSlot2
        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
            .Invent.ArmourEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un barco
        If .Invent.BarcoSlot = ObjSlot1 Then
            .Invent.BarcoSlot = ObjSlot2
        ElseIf .Invent.BarcoSlot = ObjSlot2 Then
            .Invent.BarcoSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un casco
        If .Invent.CascoEqpSlot = ObjSlot1 Then
            .Invent.CascoEqpSlot = ObjSlot2
        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
            .Invent.CascoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un escudo
        If .Invent.EscudoEqpSlot = ObjSlot1 Then
            .Invent.EscudoEqpSlot = ObjSlot2
        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
            .Invent.EscudoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es munición
        If .Invent.MunicionEqpSlot = ObjSlot1 Then
            .Invent.MunicionEqpSlot = ObjSlot2
        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
            .Invent.MunicionEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un arma
        If .Invent.WeaponEqpSlot = ObjSlot1 Then
            .Invent.WeaponEqpSlot = ObjSlot2
        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
            .Invent.WeaponEqpSlot = ObjSlot1
        End If
       
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, UserIndex, ObjSlot1)
        Call UpdateUserInv(False, UserIndex, ObjSlot2)
    End With

End Sub

Public Sub HandleDragToPos(ByVal UserIndex As Integer)

' @ Author : maTih.-
'            Drag&Drop de objetos en del inventario a una posición.

Dim X       As Byte
Dim Y       As Byte
Dim Slot    As Byte
Dim Amount  As Integer
Dim tUser   As Integer
Dim tNpc    As Integer

Call UserList(UserIndex).incomingData.ReadByte

X = UserList(UserIndex).incomingData.ReadByte()
Y = UserList(UserIndex).incomingData.ReadByte()
Slot = UserList(UserIndex).incomingData.ReadByte()
Amount = UserList(UserIndex).incomingData.ReadInteger()

tUser = MapData(UserList(UserIndex).Pos.map, X, Y).UserIndex
tNpc = MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex


End Sub

Public Sub HandleLogCuenta(ByVal UserIndex As Integer)

'
' @ Login to account.

With UserList(UserIndex)

     Dim tmp_Buffer     As New clsByteQueue
     Dim tmp_AccName    As String
     
     tmp_Buffer.CopyBuffer .incomingData
     
     tmp_Buffer.ReadByte
     
     tmp_AccName = tmp_Buffer.ReadASCIIString()

     WriteSendCuenta UserIndex, tmp_AccName

     .incomingData.CopyBuffer tmp_Buffer
     
End With

End Sub

Public Sub HandleCompra(ByVal UserIndex As Integer)

'
' @ Compra/Requiere lista de items.

Dim C_Index     As Integer
Dim P_Obj       As Obj

With UserList(UserIndex)

    Call .incomingData.ReadByte

    C_Index = .incomingData.ReadInteger()

    'Si es 1 o más arriba
    If C_Index > 0 Then
       If (C_Index <> 0) And (C_Index <= UBound(Buy(UserList(UserIndex).Clase).Buys())) Then
          Call mod_DunkanCs.Comprar(C_Index, UserIndex)
          Exit Sub
       End If
    End If
    
    If C_Index < 0 Then
        If C_Index = -1 Then
           WriteCompra UserIndex, mod_DunkanCs.GetList(.Clase)
        End If
    End If
End With

End Sub

Public Sub WriteCompra(ByVal UserIndex As Integer, ByRef sList As String)

'
' @ Envia una lista

With UserList(UserIndex).outgoingData
     .WriteByte ServerPacketID.Compra
     .WriteASCIIString sList
End With

End Sub

Public Sub HandleBando(ByVal UserIndex As Integer)

'
' @ Escoje un bando.

Call UserList(UserIndex).incomingData.ReadByte

If UserList(UserIndex).incomingData.ReadBoolean() Then
   With UserList(UserIndex).Reputacion
        .PlebeRep = 0
        .NobleRep = 0
        .BurguesRep = 0
        .AsesinoRep = 30000
        .LadronesRep = 30000
        .Promedio = 0
        .BandidoRep = 500

        If criminal(UserIndex) Then
           RefreshCharStatus UserIndex

        End If
   End With
End If

End Sub

Public Sub HandlePotear(ByVal UserIndex As Integer)

'
' @ Usa item

With UserList(UserIndex)
     Dim Slot       As Byte
     Dim objIndex   As Integer
     
     .incomingData.ReadByte
     
     Slot = .incomingData.ReadByte()
     
     If (Not Slot <> 0) Then Exit Sub
     
     objIndex = .Invent.Object(Slot).objIndex
     
     If (Not objIndex <> 0) Then Exit Sub
     
     'Cura vida
     If ObjData(objIndex).TipoPocion = 3 Then
        'Usa el item
        .Stats.MinHp = .Stats.MinHp + RandomNumber(ObjData(objIndex).MinModificador, ObjData(objIndex).MaxModificador)
        If .Stats.MinHp > .Stats.MaxHp Then _
            .Stats.MinHp = .Stats.MaxHp
                      
        WriteUpdateHP UserIndex
     ElseIf ObjData(objIndex).TipoPocion = 4 Then
        'Usa el item
        'nuevo calculo para recargar mana
        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
        If .Stats.MinMAN > .Stats.MaxMAN Then _
            .Stats.MinMAN = .Stats.MaxMAN
        
        WriteUpdateMana UserIndex
     End If
     
End With

End Sub

Public Sub HandleLogPj(ByVal UserIndex As Integer)

'
' @ Login to pjIndex.

Dim classIndex  As Integer
Dim buffer      As New clsByteQueue

With UserList(UserIndex)
    
     Call buffer.CopyBuffer(.incomingData)
        
     Call buffer.ReadByte
        
     classIndex = buffer.ReadByte()

     Call mod_DunkanModos.Crear_Personaje(UserIndex, classIndex, buffer.ReadASCIIString())
     
     Call .incomingData.CopyBuffer(buffer)
     
End With

End Sub

Public Sub WriteSendCuenta(ByVal UserIndex As Integer, ByRef account_Name As String)

'
' @ Envia cuenta ok.

With UserList(UserIndex).outgoingData
     .WriteByte ServerPacketID.SendCuenta
     .WriteASCIIString account_Name
End With

End Sub
Public Sub HandleSpawnBot(ByVal UserIndex As Integer)

' @ Spawnea bot.

With UserList(UserIndex)
     
     Dim botName    As String           '<Nick <CLAN>.
     Dim botClass   As eIAClase         '<Clase.
     Dim botViaja   As Boolean          '<Viajante o no.
     Dim SpawnInPos As WorldPos         '< Pos.
     Dim buffer     As New clsByteQueue
     
     Call buffer.CopyBuffer(.incomingData)
     
     Call buffer.ReadByte
     
     botName = buffer.ReadASCIIString()
     botClass = buffer.ReadByte()
     botViaja = buffer.ReadBoolean()
     
     SpawnInPos.map = buffer.ReadInteger()
     SpawnInPos.X = buffer.ReadByte()
     SpawnInPos.Y = buffer.ReadByte()
     
     'Hardcoded, not viajante all paths reads.
     If botViaja Then
        SpawnInPos = Ullathorpe
     End If
     
     'Only GOODs.
     If Not .flags.Privilegios <> PlayerType.Dios Then
        Call mod_IA.ia_Spawn(botClass, SpawnInPos, botName, botViaja, True)
     End If
     
     Call .incomingData.CopyBuffer(buffer)
End With

End Sub

Public Sub HandleSendChallange(ByVal UserIndex As Integer)

' @ Author : maTih.-
'          : Envia reto 1.1/2.2

Dim tempBuffer          As New clsByteQueue
Dim ChallangeMode       As Byte
Dim Oponente            As String
Dim compañeroOponente   As String
Dim Compañero           As String

With UserList(UserIndex)

     Call tempBuffer.CopyBuffer(.incomingData)
     
     Call tempBuffer.ReadByte
     
     ChallangeMode = tempBuffer.ReadByte()
     
     'Tipo de reto.
     Select Case ChallangeMode
            
            Case 1      '1vs1
                 'Lee el oponente
                 Oponente = tempBuffer.ReadASCIIString()
                 
            Case 2      '2vs2
                 Oponente = tempBuffer.ReadASCIIString()
                 compañeroOponente = tempBuffer.ReadASCIIString()
                 Compañero = tempBuffer.ReadASCIIString()
            
     End Select

End With

End Sub

Public Sub HandleAcceptChallange(ByVal UserIndex As Integer)

' @ Author : maTih.-
'          : Acepta un reto.

With UserList(UserIndex)

     Call .incomingData.ReadByte

End With

End Sub

Public Sub HandleRejectChallange(ByVal UserIndex As Integer)

' @ Author : maTih.-
'          : Rechaza un reto.

With UserList(UserIndex)

     Call .incomingData.ReadByte

End With

End Sub

Public Sub HandleRequestRanking(ByVal UserIndex As Integer)

' @ Author : maTih.-
'          : Requiere un ranking

With UserList(UserIndex)
     
     Call .incomingData.ReadByte
     
     Dim RankingType    As Byte
     Dim classSelect    As eClass
     Dim dataTosend     As String
     
     RankingType = .incomingData.ReadByte()
     classSelect = .incomingData.ReadByte()
     
     Select Case RankingType
            
            Case 1          '< Ranking más oro.
                 dataTosend = mod_DunkanRankings.ListaOro()
            
            Case 2          '< Ranking más nivel
                 dataTosend = mod_DunkanRankings.ListaNivel()
                 
            Case 3          '< Ranking mas vida.
                 dataTosend = mod_DunkanRankings.ListaVidas(classSelect)
     End Select
     
     If dataTosend <> vbNullString Then
        
     End If
     
End With

End Sub
