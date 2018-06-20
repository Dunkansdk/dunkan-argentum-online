Attribute VB_Name = "Mod_DunkanProtocol"
Option Explicit

'No tiene mucho que ver acá, pero para unificar..

Public Config_Particles     As Boolean

Type user_Acc_Chars
     My_Char        As Char
     My_Name        As String
     My_Class       As eClass
End Type

Type user_Acc
     Chars(1 To 8)  As user_Acc_Chars
     acc_Name       As String
End Type

Public Acc_Data     As user_Acc

Public Sub HandleSendCuenta()

'
' @ Recibe la cuenta.

Dim tmp_Buff    As New clsByteQueue
Dim loopX       As Long

'Copia los datos al buffer temportal
tmp_Buff.CopyBuffer incomingData

'Quita el byteID
tmp_Buff.ReadByte

'Obtiene el nombre del a cuenta
Acc_Data.acc_Name = tmp_Buff.ReadASCIIString()

For loopX = 1 To 8
    frmNewCuenta.lst_Pjs.AddItem ListaClases(loopX)
Next loopX

'Muestra el form.
frmNewCuenta.Show

'Borra los datos del buffer temporal
incomingData.CopyBuffer tmp_Buff

End Sub

Public Sub WriteLogCuenta()

'
' @ logea a la cuenta.

With outgoingData
     .WriteByte ClientPacketID.LogCuenta
     .WriteASCIIString Acc_Data.acc_Name
End With

End Sub

Public Sub HandleChangeScreen()

' @ Dispara un moveScreen.

Call incomingData.ReadByte

Call Mod_DX8_Engine.MoveScreen(incomingData.ReadByte())

End Sub

Public Sub HandleCompra()

'
' @ Listas de compras :p

Dim Buffer      As New clsByteQueue
Dim tmp_Arr()   As String

Call Buffer.CopyBuffer(incomingData)

Call Buffer.ReadByte

tmp_Arr = Split(Buffer.ReadASCIIString(), ",")

Call BuyStateChange(tmp_Arr())

Call incomingData.CopyBuffer(Buffer)

End Sub

Public Sub WritePotear(ByVal slot_Used As Byte)

'
' @ Usar item

With outgoingData
     .WriteByte ClientPacketID.Potear
     .WriteByte slot_Used
End With

End Sub

Public Sub WriteCompra(ByVal C_Index As Integer)

'
' @ COMPRA UN ITEM

With outgoingData
     .WriteByte ClientPacketID.Compra
     .WriteInteger C_Index
End With

End Sub


Public Sub WriteLogPj()

'
' @ Conecta un pj

With outgoingData
     .WriteByte ClientPacketID.LogPJ
     .WriteByte frmNewCuenta.Personaje_Index
     .WriteASCIIString Acc_Data.acc_Name
End With

End Sub

Public Sub WriteBando(ByVal Criminal_Select As Boolean)

'
' @ Elige criminal/ciuaddano

With outgoingData
     .WriteByte ClientPacketID.Bando
     .WriteBoolean Criminal_Select
End With

End Sub


' @ Diseñado & Implementado por maTih.-
' @ Dunkan AO Protocol.

Public Sub HandleCreateSpell()

' @ Crea FX/Particula sobre chars.

Dim AttackerCharIndex   As Integer
Dim VictimCharIndex     As Integer
Dim EffectIndex         As Integer
Dim FXLoops             As Integer
Dim ParticleCasteada    As Boolean

With incomingData
     
     'Borra packetID
     .ReadByte
     
     'Carga data.
     AttackerCharIndex = .ReadInteger()
     VictimCharIndex = .ReadInteger()
     EffectIndex = .ReadInteger()
     FXLoops = .ReadInteger()
     
     'No quiere particulas.
     If Not Config_Particles Then
        Call SetCharacterFx(VictimCharIndex, EffectIndex, FXLoops)
     Else
        'Quiere particulas.
         ParticleCasteada = (Mod_DX8_Graphics.Engine_UTOV_Particle(AttackerCharIndex, VictimCharIndex, EffectIndex) <> 0)
         
         'Si quiere particulas, pero el hechizo no tiene particula
         'Mostramos el fx.
         If Not ParticleCasteada Then
            Call SetCharacterFx(VictimCharIndex, EffectIndex, FXLoops)
         End If
         
     End If
     
End With

End Sub

Public Sub HandleCreateMeditation()

' @ Crea FX/Particula de una meditación sobre char.

Dim charIndex           As Integer
Dim EffectIndex         As Integer
Dim FXLoops             As Integer
Dim meditationOk        As Boolean
Dim ParticleIndex       As Integer

With incomingData
     
     'Borra packetID
     .ReadByte
     
     'Carga data.
     charIndex = .ReadInteger()
     EffectIndex = .ReadInteger()
     FXLoops = .ReadInteger()
     
     'El char tenia partícula?
     If charlist(charIndex).iParticle <> 0 Then
        'Borra ...
        Effect(charlist(charIndex).iParticle).Used = False
        charlist(charIndex).iParticle = 0
        Exit Sub
     End If
        
     ParticleIndex = Mod_DX8_Graphics.Engine_UTOV_Particle(charIndex, charIndex, EffectIndex)
     
     meditationOk = (ParticleIndex <> 0)
     
     If Not meditationOk Then
        SetCharacterFx charIndex, EffectIndex, FXLoops
     Else
        charlist(charIndex).iParticle = ParticleIndex
     End If
     
End With

End Sub

Public Sub HandleCreateArrow()

' @ Crea proyectiles en chars.

Dim AttackerCharIndex   As Integer
Dim VictimCharIndex     As Integer
Dim GrhIndexArrow       As Integer

With incomingData
     
     'Borra packetID
     .ReadByte
     
     'Carga data.
     AttackerCharIndex = .ReadInteger()
     VictimCharIndex = .ReadInteger()
     GrhIndexArrow = .ReadInteger()
     
     'Crea : P
     Call Mod_DX8_Gore.Engine_Projectile_Create(AttackerCharIndex, VictimCharIndex, GrhIndexArrow, 0)
     
End With

End Sub

Public Sub HandleCreateDamage()

' @ Crea daño en mapa.

Dim TargetPos       As WorldPos
Dim DamageVal       As Integer

With incomingData
     
     'PacketID.
     .ReadByte
     
     'Default (AL PEDO, NO SE USA)
     TargetPos.Map = UserMap
     
     'Get the Data (:
     TargetPos.X = .ReadByte()
     TargetPos.Y = .ReadByte()
     DamageVal = .ReadInteger()
     
     Call Mod_DX8_Gore.Engine_Damage_Create(CInt(TargetPos.X), CInt(TargetPos.Y), DamageVal, 255, 0, 0)
     
End With

End Sub

Public Sub HandleModifyClimate()

' @ Modifica el clima del cliente.

'Pendiente de hacer que esto funcione : P

Dim NewClientClimate        As Byte

With incomingData
     
     .ReadByte
     
     NewClientClimate = .ReadByte()
     
End With

End Sub

Public Sub HandleSynchronizeTime()

' @ Sincroniza la hora de la pc cliente con el servidor.

'Paquetes contienen strings hay que usar esto !
Dim Buffer      As New clsByteQueue

'Copio el buffer.
Buffer.CopyBuffer incomingData

'Borro el PacketID.
Buffer.ReadByte

'Obtengo el tiempo.
Time = Buffer.ReadASCIIString()

Call ShowConsoleMsg("Sincronización de hora > Terminada!, Ahora la hora está sincronizada con la del servidor!")

'Deleteo el buffer.
incomingData.CopyBuffer Buffer

End Sub

Public Sub HandleChangeHeadingChar()

' @ Forza el cambio del heading de un char (bots!!)

Dim charIndex   As Integer
Dim NewHeading  As E_Heading

With incomingData
     .ReadByte
     
     charIndex = .ReadInteger()
     
     NewHeading = .ReadByte()
     
     If NewHeading <> charlist(charIndex).Heading And NewHeading <> 0 Then
        charlist(charIndex).Heading = NewHeading
     End If
     
End With

End Sub

Public Sub HandleSendRankingLista()

' @ Recibe un ranking

Dim Buffer      As New clsByteQueue
Dim arrUsers()  As String
Dim loopX       As Long

Call Buffer.CopyBuffer(incomingData)

Call Buffer.ReadByte

arrUsers = Split(Buffer.ReadASCIIString(), ",")

'Limpiamos antes.
Call frmRanking.LimpiarUsers

For loopX = LBound(arrUsers()) To UBound(arrUsers())
    ' <<<<<< Hay personaje >>>>>>
    If arrUsers(loopX) <> vbNullString Then
        frmRanking.lblUser(loopX) = arrUsers(loopX)
    Else
        frmRanking.lblUser(loopX) = "Ninguno"
    End If
Next loopX


frmRanking.Show , frmMain

Call incomingData.CopyBuffer(Buffer)

End Sub



Public Sub WriteDragInventory(ByVal originalSlot As Byte, ByVal targetSlot As Byte)

' @ Author : maTih.-
'            Drag&Drop de objetos en el inventario.

    With outgoingData
         .WriteByte ClientPacketID.DragInventory
         .WriteByte originalSlot
         .WriteByte targetSlot
    End With

End Sub

Public Sub WriteDragToPos(ByVal X As Byte, ByVal Y As Byte, ByVal slot As Byte, ByVal amount As Integer)

' @ Author : maTih.-
'            Drag&Drop de objetos en del inventario a una posición.

    With outgoingData
         .WriteByte ClientPacketID.DragToPos
         .WriteByte X
         .WriteByte Y
         .WriteByte slot
         .WriteInteger amount
    End With
    
End Sub

Public Sub WriteSpawnBot(ByRef botNAME As String, ByVal botClase As Byte, ByVal botMap As Integer, ByVal botX As Byte, ByVal botY As Byte, ByVal viajanTe As Boolean)

' @ Author : maTih.-
'          : Spawn bot en posición.

    With outgoingData
         .WriteByte ClientPacketID.SpawnBot
         .WriteASCIIString botNAME
         .WriteByte botClase
         .WriteBoolean viajanTe
         
         'write pos
         .WriteInteger botMap
         .WriteByte botX
         .WriteByte botY
    End With

End Sub
