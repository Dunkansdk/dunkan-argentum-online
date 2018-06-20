Attribute VB_Name = "Mod_Protocol_Dunkansdk"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' Author: maTih

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Handles the packets of Client
Private Enum ClientDaoPacketID
   ComprarVentaItem = 1
   AgregarVentaItem
   QuitarVentaItem
   AcceptQuest
   SendReto        ' - /Retos
   AcceptReto      ' - /Retas
   partystart      ' - Party exp
   ChangeExpParty  ' - Cambia los porcentajes de la party
   SendSoport      ' - Soporte
   SendSoportForm  ' - FrmSoport
   ReadSoport      ' - Leer soporte
   RespondSoport   ' - Responder soporte
   QuestAcept      ' - Aceptar Quest index
   QuestRequestInf ' - Requiere info de la quest actual
   BankGoldTrans   ' - Transfiere oro de un personaje a otro
   EventIngress    ' - Ingresa al evento actual
   DragObj         ' - Drag & Drop (Inventario)
   DragDropObj     ' - Drag & Drop (Target)
   RecoverAccount  ' - Recupera cuenta
   RequestTorneoF  ' - Admin form Torneo.
   CreateTorneo    ' - Crea un torneo.
   ActionTorneo    ' - Admin accion Torneo.
   IngressTorneo   ' - User ingresa al torneo
   
   'Accounts in developed
   CreateNewAccount
   LoginExistingAccount
   LoginCharAccount
   
   ' - - - - - - - - - - - - - - - - -
   
   OffertSubast    ' - Oferta subasta actual
   CreateSubast    ' - User crea una subast
   ScreenDenounce  ' - Foto denuncia area visión
   REQUESTSCREEN   ' - Admin requiere screen de un cliente

End Enum

'Handles the packets off server
Private Enum ServerDaoPacketID
   UpdateVentaSlot = 1
   UpdateVenta
   ShowVentaForm
   UpdateTargets
   QuestListLVL
   ParchamentMsg    ' - Muestra un mensaje en el formulario del pergamino
   CreateParticle   ' - Crea particula sobre el char y victimchar
   CreateProyectile ' - Crea Proyectiles (flechas) sobre chars
   SendQuestForm    ' - Envía formulario de quest
   SendSoportFormS  ' - Envía formulario de soporte
   CreateDamageMap  ' - Crea daño en el mapa
   MovimentWeapon   ' - Movimiento del arma
   MovimentShield   ' - Movimiento del escudo
   ChangeHour       ' - Cambia hora del cliente (Noche, día, etc)
   SendTorneoF      ' - Envía form del panelTorneo.
   
   'Accounts in developed
   ReceivedAccount
   UpdateAccount
   
   ' - - - - - - - - - - - - - - - - -
   
   SendUserInventory  ' - Recibe los objetos del inventario
   UploadScreen
   SendRanking
   
   UpdateUsers        ' - Recibe usuarios onlines.
   UpdateTagTarget    ' - Updatea labels del S.Targets
End Enum

' Declares
Public PuedeScreen      As Boolean
Public CounterScreen    As Byte

Sub JAJ()

MsgBox ServerDaoPacketID.UpdateTagTarget
MsgBox ClientDaoPacketID.REQUESTSCREEN
End Sub

'ORDEN I N V I O L A B L E DE MANEJO DE RUTINAS Y FUNCIONES.
'HANDLE + PAQUETE
'WRITE + "DAO" + PAQUETE
'FUNCTION DUNKAN + PAQUETE

Public Sub HandleDAOProtocol()

' / Author: maTih
' / Note: Handling Dao Packet IDs sending by server.

'Remove byte of memory
Call incomingData.ReadByte

Dim DAOPacketID As Byte

'We read the fact that used yo remove the memory

DAOPacketID = incomingData.PeekByte

Select Case DAOPacketID

    'WHAT PACKAGE RECEIVED?
    
    Case ServerDaoPacketID.UpdateTargets
        Call HandleUpdateTargets
        
    Case ServerDaoPacketID.QuestListLVL
        Call HandleQuestListLVL
        
   Case ServerDaoPacketID.ParchamentMsg
        Call HandleParchamentMessage
    
   Case ServerDaoPacketID.CreateParticle
        Call HandleCreateParticle
    
    Case ServerDaoPacketID.CreateProyectile
        Call HandleCreateProyectile

  'Recibe la acción para mostrar el formulario de quest
   Case ServerDaoPacketID.SendQuestForm
        Call HandleSendQuestForm

   'Recibe la acción para mostrar el formulario de soportes
   Case ServerDaoPacketID.SendSoportFormS
        Call HandleSendSupportForm
        
   'Crea el daño sobre el mapa
   Case ServerDaoPacketID.CreateDamageMap
        Call HandleCreateDamageMap
        
   'Movimiento del arma sobre el char
   Case ServerDaoPacketID.MovimentWeapon
        Call HandleMovimentWeapon
        
   'Movimiento de escudo sobre el char
   Case ServerDaoPacketID.MovimentShield
        Call HandleMovimentShield
        
   'Recibe la hora del servidor
   Case ServerDaoPacketID.ChangeHour
        Call HandleChangeHour
        
   'Recibe la acción para mostrar un torneo
   Case ServerDaoPacketID.SendTorneoF
        Call HandleShowTorneoForm
        
   'Recibe la cuenta entera
   Case ServerDaoPacketID.ReceivedAccount
        Call HandleReceivedAccount
        
   'Update X Slot del user account
   Case ServerDaoPacketID.UpdateAccount
        Call HandleUpdateAccount
   
   'Recibe el inventario y lo agrega en una listbox (Subasta)
   Case ServerDaoPacketID.SendUserInventory
       Call HandleSendUserInventory
       
    'Envía una foto al servidor FTP
    Case ServerDaoPacketID.UploadScreen
       Call HandleScreenUpload
       
    'Recibe el ranking y sus personajes
    Case ServerDaoPacketID.SendRanking
       Call HandleSendRAnking
       
    Case ServerDaoPacketID.UpdateUsers
       Call handleUpdateUsers
       
     Case ServerDaoPacketID.UpdateTagTarget
       Call HandleUpdateTagTarget
    
   Case Else
   
        WriteDenounce UserName & "Rrecibió un paquete sin codear, packetID = " & DAOPacketID
   
End Select

End Sub

Private Sub HandleUpdateTargets()

'  /  Author    :  maTih.-
'  /  Note      :  Recieved targets to be click.

Dim Buffer      As New clsByteQueue
Dim loopX       As Long

    Buffer.CopyBuffer incomingData
    
    Buffer.ReadByte
    
    targetS = Split(Buffer.ReadASCIIString(), "|")
   
    For loopX = LBound(targetS()) To UBound(targetS())
        frmMain.targetS(loopX).Caption = targetS(loopX)
        frmMain.targetS(loopX).AutoSize = True
        frmMain.targetS(loopX).Visible = True
        frmMain.targetS(loopX).ForeColor = vbBlue
    Next loopX
   
    incomingData.CopyBuffer Buffer
End Sub

Private Sub HandleQuestListLVL()

'  /  Author    :  maTih.-
'  /  Note      :  Recieved quest list of level the user.

Dim Buffer      As New clsByteQueue
Dim loopX       As Long

    Buffer.CopyBuffer incomingData
    
    Buffer.ReadByte
    
    QuestList = Split(Buffer.ReadASCIIString, "|")
    QuestInfo = Split(Buffer.ReadASCIIString, "|")
    
    frmPergamino.Message.Caption = QuestList(LBound(QuestList())) & vbNewLine & QuestInfo(LBound(QuestInfo()))
    
    incomingData.CopyBuffer Buffer
End Sub

Private Sub HandleParchamentMessage()

'  /  Author    :  maTih.-
'  /  Note      :  Show parchament form and prepare message.

Dim Buffer      As New clsByteQueue
Dim Message     As String

With Buffer
    .CopyBuffer incomingData
    .ReadByte
    
    Message = .ReadASCIIString()
    
    frmPergamino.Message.Caption = Message
    
    frmPergamino.Show , frmMain
End With

    incomingData.CopyBuffer Buffer

End Sub

Private Sub HandleCreateParticle()

' / Author: Dunkan
' / Note  : Sending particles from one char to another

    Call incomingData.ReadByte
    
    Dim SendChar        As Integer  'UserIndex
    Dim ReceivedChar    As Integer  'Víctima
    Dim ParticleID      As Byte     'Particle ID
    
    SendChar = incomingData.ReadInteger()
    ReceivedChar = incomingData.ReadInteger()
    ParticleID = incomingData.ReadByte()
    
    Engine_UTOV_Particle SendChar, ReceivedChar, ParticleID

End Sub

Private Sub HandleCreateProyectile()

' / Author: Dunkan

With incomingData

    Call .ReadByte
    
    Dim CharSending      As Integer
    Dim CharRecieved     As Integer
    Dim GrhIndex         As Integer
    
    CharSending = .ReadInteger()
    CharRecieved = .ReadInteger()
    GrhIndex = .ReadInteger()
    
    Engine_Projectile_Create CharSending, CharRecieved, GrhIndex, 0

End With

End Sub

Private Sub HandleSendQuestForm()

' / Author: maTih
' / Note: Received by server data : quest recompense & Num of quests

Call incomingData.ReadByte

Dim NumQ            As Byte
Dim Recompense()    As String
Dim LoopC           As Long

NumQ = incomingData.ReadByte

ReDim Recompense(1 To NumQ)

    For LoopC = 1 To NumQ
        Recompense(LoopC) = incomingData.ReadASCIIString
    Next LoopC

End Sub

Private Sub HandleSendSupportForm()

' / Author: maTih
' / Note: Received by server data : support numbers and names

Call incomingData.ReadByte

Dim idName          As Byte
Dim SupportsSends() As String
Dim LoopC           As Long

idName = incomingData.ReadByte()

ReDim SupportsSends(1 To idName)

    For LoopC = 1 To idName
        SupportsSends(LoopC) = incomingData.ReadASCIIString()
    Next LoopC

End Sub

Private Sub HandleCreateDamageMap()

' / Author: maTih
' / Note: Received by server data : X , Y & Damage.

Call incomingData.ReadByte

Dim x       As Byte
Dim y       As Byte
Dim Damage  As Integer
Dim Valor   As Byte

x = incomingData.ReadByte
y = incomingData.ReadByte

Damage = incomingData.ReadInteger

Engine_Damage_Create x, y - 1, Damage, 255, 0, 0

End Sub

Private Sub HandleMovimentWeapon()

' / Author: maTih


Call incomingData.ReadByte

Dim charIndex As Integer          'Char Index

charIndex = incomingData.ReadInteger

With charlist(charIndex)
    .InMoviment = True
    .Arma.WeaponWalk(.Heading).Started = 1
    .Escudo.ShieldWalk(.Heading).Started = 1
End With

End Sub

Private Sub HandleMovimentShield()

Call incomingData.ReadByte

Dim charIndex As Integer          'CHAR INDEX

charIndex = incomingData.ReadInteger

With charlist(charIndex)
    .InMoviment = True
    .Escudo.ShieldWalk(.Heading).Started = 1
End With

End Sub

Private Sub HandleShowTorneoForm()

' / Author: maTih
' / Note: Received package it form show !

'Remove byte of memory
Call incomingData.ReadByte

'This loop users will go
Dim i               As Long

'Number of users to add to list
Dim NumUsers        As Byte
Dim ArrayUsers()    As String

'We read the amount and generate the loop
NumUsers = incomingData.ReadByte

'Resize the matrix and loop start!
ReDim ArrayUsers(1 To NumUsers)

    'Now let's walk to one and adding to the list
    For i = 1 To NumUsers
        ArrayUsers(i) = incomingData.ReadASCIIString
    Next i

End Sub

Private Sub HandleChangeHour()

' / Author: maTih
' / Note: Time received sent by server to handle the state of the day

'Otra ves sopa, kill of memory 1 byte .
Call incomingData.ReadByte

'**** | Received Info for handling state of day | ****'
Dim HourReceived        As Byte
Dim MinutesReceiveds    As Byte

HourReceived = incomingData.ReadByte
MinutesReceiveds = incomingData.ReadByte

Debug.Print "Hora: " & HourReceived
Debug.Print "Minutos: " & MinutesReceiveds

'Ema maneja esto ah como lo tenías hecho vos, ahora
'La hora la controla el servidor ^^
End Sub

Private Sub HandleUpdateAccount()

' / Author: maTih
' / Note: Update slot of account.

Call incomingData.ReadByte

'VARIABLES OF HANDLING DATA
Dim slotR   As Byte    'received slot .
Dim Name    As String
Dim Clase   As String
Dim Nivel   As Byte
Dim Cuerpo  As Integer
Dim Cabeza  As Integer
Dim Arma    As Byte
Dim Escudo  As Byte
Dim Casco   As Byte

slotR = incomingData.ReadByte
Name = incomingData.ReadASCIIString
Clase = incomingData.ReadASCIIString
Nivel = incomingData.ReadByte
Cuerpo = incomingData.ReadInteger
Cabeza = incomingData.ReadInteger
Arma = incomingData.ReadByte
Escudo = incomingData.ReadByte
Casco = incomingData.ReadByte

    If Cuentas.CantidadPersonajes = 0 Then
        ReDim Cuentas.charInfo(1 To 1) As CharData
            Cuentas.CantidadPersonajes = 1
            Cuentas.charInfo(1).Head = Cabeza
            Cuentas.charInfo(1).Body = Cuerpo
            Cuentas.charInfo(1).Weapon = Arma
            Cuentas.charInfo(1).Shield = Escudo
            Cuentas.charInfo(1).Helmet = Casco
            Cuentas.charInfo(1).Name = Name
            Cuentas.charInfo(1).Nivel = Nivel
    Else
        If Cuentas.charInfo(slotR).Name <> "NothingPJ" Then
            Cuentas.CantidadPersonajes = Cuentas.CantidadPersonajes + 1
            ReDim Cuentas.charInfo(1 To Cuentas.CantidadPersonajes) As CharData
                Cuentas.charInfo(slotR).Head = Cabeza
                Cuentas.charInfo(slotR).Body = Cuerpo
                Cuentas.charInfo(slotR).Weapon = Arma
                Cuentas.charInfo(slotR).Shield = Escudo
                Cuentas.charInfo(slotR).Helmet = Casco
                Cuentas.charInfo(slotR).Name = Name
                Cuentas.charInfo(slotR).Nivel = Nivel
        Else
            Cuentas.charInfo(slotR).Head = Cabeza
            Cuentas.charInfo(slotR).Body = Cuerpo
            Cuentas.charInfo(slotR).Weapon = Arma
            Cuentas.charInfo(slotR).Shield = Escudo
            Cuentas.charInfo(slotR).Helmet = Casco
            Cuentas.charInfo(slotR).Name = Name
            Cuentas.charInfo(slotR).Nivel = Nivel
        End If
        
    End If

End Sub

Private Sub HandleReceivedAccount()

' / Author: maTih
' / Note: Receive data for charfile and the characters of the account

Dim Buffer As New clsByteQueue

Call Buffer.CopyBuffer(incomingData)

Call Buffer.ReadByte

'VARIABLES OF HANDLING DATA
Dim CantPj      As Byte
Dim slotR()     As Byte    'received slot .
Dim Names()     As String
Dim Clase()     As String
Dim Nivel()     As Byte
Dim Cuerpo()    As Integer
Dim Cabeza()    As Integer
Dim Arma()      As Byte
Dim Escudo()    As Byte
Dim Casco()     As Byte
Dim i           As Long

CantPj = Buffer.ReadByte

    If CantPj > 0 Then
        For i = 1 To CantPj
            slotR(i) = Buffer.ReadByte
            Names(i) = Buffer.ReadASCIIString
            Clase(i) = Buffer.ReadASCIIString
            Nivel(i) = Buffer.ReadByte
            Cuerpo(i) = Buffer.ReadInteger
            Cabeza(i) = Buffer.ReadInteger
            Arma(i) = Buffer.ReadByte
            Escudo(i) = Buffer.ReadByte
            Casco(i) = Buffer.ReadByte
        Next i
    Else
        slotR(0) = Buffer.ReadByte
        Names(0) = Buffer.ReadASCIIString
        Clase(0) = Buffer.ReadASCIIString
        Nivel(0) = Buffer.ReadByte
        Cuerpo(0) = Buffer.ReadInteger
        Cabeza(0) = Buffer.ReadInteger
        Arma(0) = Buffer.ReadByte
        Escudo(0) = Buffer.ReadByte
        Casco(0) = Buffer.ReadByte
    End If

    Call incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleSendUserInventory()
' / Author : maTih
' / Note: Received obj names of userinventory.

'Is not an interface, so it is NEW
'Use the Auxiliary buffer
Dim Buffer As New clsByteQueue

Call Buffer.CopyBuffer(incomingData)

Call Buffer.ReadByte   'remove packetID

'THE NUMBER OF ITEMS THAT HAVE IN YOUR INVENTORY
Dim amountItems As Byte
Dim tmpS()      As String
Dim i           As Long

amountItems = Buffer.ReadByte

ReDim tmpS(1 To amountItems) As String

    'declares one bucle, This will bucle through the objects.
    For i = 1 To amountItems
        tmpS(i) = Buffer.ReadASCIIString
    'Pending of modification
    'FALTA HACERLO, LO GUARDO EN UN STRING PARA NO ROMPER TODO
    Next i

    'cerramos y destruimos el buffer auxiliar, ya tenemos nuestros datos
    Call incomingData.CopyBuffer(Buffer)

End Sub

Private Sub HandleScreenUpload()

' / Author: maTih

With incomingData

    Call .ReadByte
    
    Dim lastscren As Byte
    
    lastscren = .ReadByte
    
    'funcion para subir la foto al servidor ftp..
    
    'uploadfoto lastscren + 1

End With

End Sub

Private Sub HandleUpdateTagTarget()

With incomingData

Call .ReadByte

' Save buffer data in variables

Dim TName         As String     ' Name
Dim aOptions      As Byte       ' Amount Options
Dim tOptions()    As String     ' Options strings
Dim LoopC         As Long

    ' Save data
    
    TName = .ReadASCIIString()
    aOptions = .ReadByte()
    
    ReDim tOptions(1 To aOptions) As String
    
    For LoopC = 1 To aOptions
        tOptions(LoopC) = .ReadASCIIString()
    Next LoopC

End With

End Sub

Private Sub handleUpdateUsers()

' / Author : maTih

'not used auxilliarbuffer , used only package containg strings

Call incomingData.ReadByte

  'Stored in this variable the amount received

 UsuariosOnline = incomingData.ReadInteger()
 MsgBox UsuariosOnline
 
End Sub

Private Sub HandleSendRAnking()

' / Author: maTih

Dim buff As New clsByteQueue

Call buff.CopyBuffer(incomingData)

With buff

'remove packetID for memory (1 byte)
Call .ReadByte

'here loop for 1 to 10.
Dim i As Long
For i = 1 To 10

tRAnk.UserNames(i) = .ReadASCIIString
tRAnk.UserLevels(i) = .ReadByte
tRAnk.UserFrags(i) = .ReadInteger
tRAnk.UserClases(i) = .ReadASCIIString

Next i
End With

Call incomingData.CopyBuffer(buff)
End Sub

Public Sub WriteDAOAcceptQuest(ByVal QuestName As String)

' / Author: maTih
' / Note: Send package of accept quest by name
' / Parameters sends: QuestName of the quest.

    With outgoingData
         .WriteByte 1
         .WriteByte ClientDaoPacketID.AcceptQuest
         .WriteASCIIString QuestName
    End With

End Sub

Public Sub WriteDAOSendReto(ByVal RetoMode As Byte, ByVal Opponent As String, ByVal OpponentTwo As String, ByVal Couple As String, ByVal GLD As Long, ByVal itemDrop As Byte)

' / Author: maTih
' / Note: Send Package of Retos to server .
' / Parameters sends: RetoMode, Opponent, OpponentTwo, GLD & itemDrop(boolean)

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.SendReto)
    Call .WriteByte(RetoMode)
    Call .WriteASCIIString(Opponent)
    Call .WriteASCIIString(OpponentTwo)
    Call .WriteASCIIString(Couple)
    Call .WriteLong(GLD)
   ' Call .WriteBoolean(CBool(itemDrop))
End With

End Sub

Public Sub WriteDAOAcceptReto(ByVal targetName As String)

' / Author: maTih
' / Note: Send package of accept Reto to server.
' / Parameters sends: TargetName.

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.AcceptReto)
    Call .WriteASCIIString(targetName)
End With

End Sub

Public Sub WriteDAOPartyStart()

' / Author: maTih
' / Note: Send Package Of start party
' Null arguments to server, handle of targetUser in server.

Call outgoingData.WriteByte(1)
Call outgoingData.WriteByte(ClientDaoPacketID.partystart)
End Sub

Public Sub WriteDAOChangeExpParty(ByVal tSlot As Byte, ByVal tSLot2, ByVal NewExp As Byte, ByVal NewExp2 As Byte)

' / Author: maTih
' / Note: send package of change exp of actuality party.
' / Parameters sends: tslot 1 and 2(used by handling array) , newexp & newexp2, is new experencie.

With outgoingData
    Call .WriteByte(1)  'HARDCODED :D
    Call .WriteByte(ClientDaoPacketID.ChangeExpParty)
    Call .WriteByte(tSlot)
    Call .WriteByte(tSLot2)
    Call .WriteByte(NewExp)
    Call .WriteByte(NewExp2)
End With

End Sub

Public Sub WriteDAOSendSupport(ByVal SupportMessage As String)

' / Author  -  maTih
' / Note:  Send package of send support.
' / Parameters sends: SupportMessage .

With outgoingData
    Call .WriteByte(1) 'HARDCODED :D
    Call .WriteByte(ClientDaoPacketID.SendSoport)
    Call .WriteASCIIString(SupportMessage)
End With

End Sub

Public Sub WriteDAOSupportForm()

' / Author: maTih
' / Note: Send package is admin received form and show.
'Null arguments send of server.

Call outgoingData.WriteByte(1) 'Hardcoded ! :D
Call outgoingData.WriteByte(ClientDaoPacketID.SendSoportForm)
End Sub

Public Sub WriteDAOSupportRead(ByVal SlotS As Byte)

' / Author: maTih
' / Note: Send package of read support by slot (is privilegies of user , NO LE PASAMOS KBIDA XD)
' / Parameters sends: Slot , byte.

With outgoingData
    Call .WriteByte(1)   'Como el enum ahora empieza desde 1, todo HARDCODEADO Y FEO :D
    Call .WriteByte(ClientDaoPacketID.ReadSoport)
    Call .WriteByte(SlotS)
End With

End Sub

Public Sub WriteDAOResponseSupport(ByVal slot As Byte, ByVal ResponseMessage As String)

' / Author: maTih
' / Parameters send: Slot , responseMessage.

With outgoingData

    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.RespondSoport)
    Call .WriteByte(slot)
    Call .WriteASCIIString(ResponseMessage)
    
End With

End Sub

Public Sub WriteDAOAcceptQuestByIndex(ByVal slot As Byte)

' / Author: maTih
' / Note: Send package of acept quest by QuestIndex(slot)

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.QuestAcept)
    Call .WriteByte(slot)
End With

End Sub

'System of accounts in developed
Public Sub WriteDAOCreateNewAccount(ByVal AccountName As String, ByVal AccountPassword As String, ByVal AccountEmail As String, ByVal AccountPIN As Byte)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.CreateNewAccount)
    Call .WriteASCIIString(AccountName)
    Call .WriteASCIIString(AccountPassword)
    Call .WriteASCIIString(AccountEmail)
    Call .WriteByte(AccountPIN)
End With

End Sub

Public Sub WriteDAOLoginExistingAccount(ByVal AccountName As String, ByVal AccountPass As String)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.LoginExistingAccount)
    Call .WriteASCIIString(AccountName)
    Call .WriteASCIIString(AccountPass)
End With
    
End Sub

Public Sub WriteDAOLoginCharAccount(ByVal charSlot As Byte)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.LoginCharAccount)
    Call .WriteByte(charSlot)
End With

End Sub

Public Sub WriteDAOCreateSubast(ByVal slot As Byte, ByVal Amount As Integer, ByVal MInime As Long)

' / Author: maTih

With outgoingData
    Call .WriteByte(1) 'hardcoding feo.
    Call .WriteByte(ClientDaoPacketID.CreateSubast)
    Call .WriteByte(slot)
    Call .WriteInteger(Amount)
    Call .WriteLong(MInime)
End With

End Sub

Public Sub WriteDAOOffertSubast(ByVal SendOff As Long)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.OffertSubast)
    Call .WriteLong(SendOff)
End With

End Sub

Public Sub WriteDAOCreateTorneo(ByVal Cupos As Byte, ByVal TIPO As Byte, ByVal PrecioInscripcion As Long)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.CreateTorneo)
    Call .WriteByte(Cupos)
    Call .WriteByte(TIPO)
    Call .WriteLong(PrecioInscripcion)
End With

End Sub

Public Sub WriteDAODragObj(ByVal tSlot As Byte, ByVal tSLot2 As Byte)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.DragObj)
    Call .WriteByte(tSlot)
    Call .WriteByte(tSLot2)
End With

End Sub

Public Sub WriteDAODragObjTarget(ByVal x As Byte, ByVal y As Byte, ByVal slot As Byte, ByVal Cant As Integer)


' / Author: maTih
' / Note  : Drag InventObj for TargetX - TargetY , and Amount

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.DragDropObj)
    Call .WriteByte(slot)
    Call .WriteByte(x)
    Call .WriteByte(y)
    Call .WriteInteger(Cant)
End With

End Sub

Public Sub WriteScreenDenounce()

' / Author: maTih

    If PuedeScreen = False Then ShowConsoleMsg "Debes esperar 2 minutos para enviar cada foto denuncia.": Exit Sub
    
    PuedeScreen = False
    CounterScreen = 2
    
    With outgoingData
        Call .WriteByte(1)
        Call .WriteByte(ClientDaoPacketID.ScreenDenounce)
    End With
    
End Sub

Public Sub WriteDAOScreenForClient(ByVal UserName As String)

' / Author: maTih

With outgoingData
    Call .WriteByte(1)
    Call .WriteByte(ClientDaoPacketID.REQUESTSCREEN)
    Call .WriteASCIIString(UserName)
End With

End Sub
