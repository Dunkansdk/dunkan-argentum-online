Attribute VB_Name = "Dunkan_Protocol"
'***************************************************************

'Author: maTih
'Note : Handle dunkan AO packets

'DOCUMENTATION :

'THE CREATOR OF THIS MODULE WILL NOT BE PUBLIC
'IS STRICTLY FORBIDDEN THE PUBLICATION OF THE SAME WITHOUT
'AUTHORIZATION OF CREATOR, THE MODULE FOR maTih WAS SCHEDULED
'FOR DUNKANAO, ALL RIGHTS RESERVED TO THE SAME.
'DOCUMENTATION IS TOTALLY OWN CODE OF CREATOR AND MAY NOT BE THE SAME UNAUTHORIZED PUBLIC

'FOR USE    :                 -

'NEW PACKAGES SENT USING THIS MODULE OCCUPIES A ONE BYTE IN THE MEMORY
'PASSING PEEKBYTE WRITING BLOCK A
'FOR USE AS DECLARE AUXILIARBUFFER DUNKANBUFFER
'AS SENDING AN EXTRA BYTE ID.
'THE ENUM OF THE PACKAGES WAS CHANGED FROM ONE TO START
'SO SEND PACKAGES TO SEND ONE MUST GO, ALL HARDCODING :D

'***************************************************************


Option Explicit

Private DunkanBuffer As New clsByteQueue

'HANDLES THE PACKETS OF CLIENT
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
    
    OffertSubast    'Usuario OFERTA SUBASTA ACTUAL.
    CreateSubast    'USUARIO INICIA UNA SUBASTA
    RequestRanking  'Usuario pide form de ranking.
   
End Enum

'HANDLES THE PACKETS OFF SERVER
Private Enum ServerDaoPacketID
       UpdateVentaSlot = 1
       UpdateVenta
       ShowVentaForm
       UpdateTargets
       SendQuestLvl
       ParchamentMsg
       createparticle   ' - Crea particula sobre el char y victimchar
       CreateProjectile ' - Crea proyectiles (flechas)
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
    
       SendUserInventory  'Envia la lista de los objetos del usuario(usado para la subasta)
       SendRanking        'Envia la lista de ranking
       
       UpdateUsers        'Envia cantidad de onlines
       UpdateTagTarget    'Updatea labels del S.Targets
       
End Enum

Sub XD()

MsgBox ServerDaoPacketID.UpdateTagTarget
MsgBox ClientDaoPacketID.RequestRanking

End Sub

'Modificated 26 / 11 / 2011
'Declares auxliarbuffer for Send Data
'Agrego encryptINI$ y decryptINI$ (author : nose.)

Public Sub HandleDAOProtocol(ByVal UserIndex As Integer)

' / Author : maTih
' / Note   : Handle the incomingdata of DAO Protocol
' / Date   : 25 / 11 / 11

Dim DAOPacketID As Byte

With UserList(UserIndex)

    Call .incomingData.ReadByte

    DAOPacketID = .incomingData.PeekByte
   
    Select Case DAOPacketID
         Case ClientDaoPacketID.AcceptQuest
            Call HandleAcceptQuest(UserIndex)
            
         Case ClientDaoPacketID.SendReto
             Call HandleSendReto(UserIndex)
          
         Case ClientDaoPacketID.AcceptReto
             Call HandleAcceptReto(UserIndex)
          
         Case ClientDaoPacketID.partystart
             Call HandlePartyStart(UserIndex)
        
         Case ClientDaoPacketID.ChangeExpParty
             Call HandleChangeExpParty(UserIndex)
        
         Case ClientDaoPacketID.SendSoport
             Call HandleSendSupport(UserIndex)
          
         Case ClientDaoPacketID.SendSoportForm
             Call HandleSendSupportForm(UserIndex)
          
         Case ClientDaoPacketID.ReadSoport
             Call HandleReadSupport(UserIndex)
          
         Case ClientDaoPacketID.RespondSoport
             Call HandleResponseSupport(UserIndex)
          
         Case ClientDaoPacketID.QuestAcept
             Call HandleQuestAccept(UserIndex)
          
         Case ClientDaoPacketID.QuestRequestInf
             Call HandleQuestRequestInfo(UserIndex)
          
         Case ClientDaoPacketID.BankGoldTrans
             Call HandleBankGLDTransferencie(UserIndex)
          
         Case ClientDaoPacketID.EventIngress
        
         Case ClientDaoPacketID.DragObj
             Call HandleDragObj(UserIndex)
          
         Case ClientDaoPacketID.DragDropObj
             Call HandleDragDropObj(UserIndex)
          
         Case ClientDaoPacketID.RecoverAccount
        
         Case ClientDaoPacketID.RequestTorneoF
             Call HandleShowTOrneoForm(UserIndex)
          
         Case ClientDaoPacketID.ActionTorneo
             Call HandleActionTorneo(UserIndex)
          
         Case ClientDaoPacketID.IngressTorneo
         '   Call handleingresstorneo(UserIndex)
          
         Case ClientDaoPacketID.CreateNewAccount
             Call HandleCreateNewAccount(UserIndex)
          
         Case ClientDaoPacketID.LoginExistingAccount
             Call HandleLoginExistingAccount(UserIndex)
          
         Case ClientDaoPacketID.LoginCharAccount
             Call HandleLoginCharAccount(UserIndex)
          
         Case ClientDaoPacketID.OffertSubast
             Call HandleOffertSubast(UserIndex)
          
         Case ClientDaoPacketID.CreateSubast
             Call HandleCreateSubast(UserIndex)
         
         Case ClientDaoPacketID.RequestRanking
             Call HandleRequestRanking(UserIndex)
        
    End Select

End With

End Sub

Private Sub HandleAcceptQuest(ByVal UserIndex As Integer)

' / Author - maTih
' / Note : AcceptQuest by name.
' / Parameters received by client : QuestName.

'Use the Auxiliary Buffer for received data

Dim buffer      As New clsByteQueue
Dim questName   As String
    
    buffer.CopyBuffer UserList(UserIndex).incomingData
    
    buffer.ReadByte
    
    questName = buffer.ReadASCIIString
    
    Quest_Aceptar UserIndex, questName
    
    UserList(UserIndex).incomingData.CopyBuffer buffer
    
End Sub

Private Sub HandleSendReto(ByVal UserIndex As Integer)

'Author - maTih
'Note : SendReto for users name.
'Parameters received by client : RetoMode , Opponent , Opponent Two , Couple , Gold & Drop
'Use the Auxiliary Buffer for received data

With UserList(UserIndex)

    'If you find a bug we headed to the error handler
    On Error GoTo ErrorHandler
    
    Dim buffer As New clsByteQueue
    
    'Copy Data for Auxiliary Buffer
    
    Call buffer.CopyBuffer(.incomingData)
    
    'Remove 1 bytes of memory
    
    Call buffer.ReadByte
    
    Dim RetoMode        As Byte     '1 for 1vs1 mode, 2 for 2vs2 mode
    Dim Opponent        As String   'Opponent name
    Dim OpponentTwo     As String   'Opponent Two name
    Dim Couple          As String   'Couple name
    Dim Gold            As Long     'Gold wagered
    Dim Drop            As Boolean  'Drop is boolean , true is drop items, false not.
    Dim ReffE           As String   'ByRef string error
    RetoMode = buffer.ReadByte()
    Opponent = buffer.ReadASCIIString()
    OpponentTwo = buffer.ReadASCIIString()
    Couple = buffer.ReadASCIIString()
    Gold = buffer.ReadLong()
   ' Drop = buffeR.ReadBoolean()
    'Disclaimer: outstanding handling data
    
    If Retos_PuedeIniciar(RetoMode, UserIndex, Opponent, OpponentTwo, Couple, Gold, 1, ReffE) Then
            Retos_Arranca2v2 UserIndex, NameIndex(Couple), NameIndex(Opponent), NameIndex(OpponentTwo), Gold, 1
    Else
            WriteConsoleMsg UserIndex, ReffE, FontTypeNames.FONTTYPE_GUILD
    End If
    
    'If we get here and complete the data arrived, we destroy the buffer auxliar
    Call .incomingData.CopyBuffer(buffer)
    
End With
    
ErrorHandler:
    Debug.Print "There has been a driving mistake Send Challenge package"

End Sub

Private Sub HandleAcceptReto(ByVal UserIndex As Integer)

' / Author: maTih
' / Note : SendReto for users name.
' / Parameters received by client : Name of the opponent to handle the challenge
' - Use the Auxiliary Buffer for received data

'If you find a bug we headed to the error handler
On Error GoTo ErrorHandler

    With UserList(UserIndex)
    
        Dim buffer As New clsByteQueue
        
        'Copy Data for Auxiliary Buffer
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove 1 bytes of memory
        Call buffer.ReadByte
        
        Dim ReceivedData As String          'Data Recieved for client
        ReceivedData = buffer.ReadASCIIString()
        
        'Disclaimer: outstanding handling data
        
        'If we get here and complete the data arrived, we destroy the buffer auxliar
        Call .incomingData.CopyBuffer(buffer)
        
    End With
    
ErrorHandler:
    Debug.Print "There has been a driving mistake Send Challenge package"
    
End Sub

Private Sub HandlePartyStart(ByVal UserIndex As Integer)

'Author - maTih
'Note : Launch party with the TargetUser

With UserList(UserIndex)

    'Do not use the auxiliary buffer, only to receive incoming data with data
    
    'Remove for memory 1 bytes
    Call .incomingData.ReadByte
    
    'Declare targetUserIndex to handle the routine
    Dim targetUserIndex As Integer
    
    targetUserIndex = .flags.targetUser
    'Aclaration: outstanding handling data

End With

End Sub

Private Sub HandleChangeExpParty(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Set new experience for users in party
' / Parameters received by the client: New user experience for me, New experience for another user.. pointer array my user party, party pointer array of other user

With UserList(UserIndex)

    'Verification to verify that the data arrived complete
    '1 bytes for targetChangeExp and 1 bytes for targetUserChangeExp
    'An extra byte for the packet
    
    'If .incomingData.length < 3 Then Exit Sub | Cancelled by testing
    
    'Remove for memory 1 byte
    Call .incomingData.ReadByte
    
    Dim MeNewExp    As Byte     'New Experience for Me
    Dim YouNewExp   As Byte     'New experience for Other User Index
    Dim SlotMe      As Byte     'Pointer for partyArray
    Dim SlotYou     As Byte     'Pointer for partyArray
    
    Dim PointerUser As Integer  'Pointer for UserList array
    
    'Remove data for incomingData
    MeNewExp = .incomingData.ReadByte
    YouNewExp = .incomingData.ReadByte
    
    SlotMe = .incomingData.ReadByte
    SlotYou = .incomingData.ReadByte
    
    'Now we handle the received data
    
    PointerUser = NameIndex(.PartyArray(SlotYou))
    
    If PointerUser <= 0 Then Exit Sub 'If the user is not connected to send data to avoid empty slots

End With

End Sub

Private Sub HandleSendSupport(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Send a support and added to the array
' / Parameters received by client : SupportMessage

' - Use the Auxiliary Buffer for received data

With UserList(UserIndex)

    Dim buffer          As New clsByteQueue
    Dim SupportMessage  As String
    
    'Copybuffer for use the handle incomingData
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte
    
    SupportMessage = buffer.ReadASCIIString()
    
    'Disclaimer: outstanding handling data
    
    'If we get here and complete the data arrived, we destroy the buffer auxliar
    Call .incomingData.CopyBuffer(buffer)

End With

End Sub

Private Sub HandleSendSupportForm(ByVal UserIndex As Integer)

' / Author: maTih
'Note : Administrator requires the form with the supports messages

With UserList(UserIndex)

    Call .incomingData.ReadByte

    If Not .flags.Privilegios = PlayerType.User Then
        'Here send the packet with the supports in the array
    End If

End With
' / Author: maTih
End Sub

Private Sub HandleReadSupport(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: If the person sending the package is a Game Master in the message to the console will pass the message of the pointer,
' if we send a user admin response
' / Parameters received by client : SlotSupport , if not gamemaster is show response

With UserList(UserIndex)

    'Remove for memory buffer 1 byte
    Call .incomingData.ReadByte
    
    Dim slotSupport As Byte
    
    'Remove for incomingData 1 byte
    slotSupport = .incomingData.ReadByte
    
    If .flags.Privilegios = PlayerType.User Then
        'If the person sending the package is a user's what you said admin
    Else
        'If not, search the array index
    End If

End With

End Sub

Private Sub HandleResponseSupport(ByVal UserIndex As Integer)
' / Author: maTih
' / Note : Game-Master responds to an index support
' / Parameters received by client : slotSupport used by search in array _
                                    Response message by goto user
                                    
With UserList(UserIndex)

    Dim buffer As New clsByteQueue
    
    'use auxilliar buffer for handling strings
    
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte 'kill 1 byte for memory
    
    Dim slotSupport     As Byte
    Dim ResponseMessage As String
    
    slotSupport = buffer.ReadByte()
    ResponseMessage = buffer.ReadASCIIString()
    
    'If .flags.Privilegios < PlayerType.Consejero Then Exit Sub
    'If the user privileges are, exit
    
    'If we get here and complete the data arrived, we destroy the buffer auxliar
    Call .incomingData.CopyBuffer(buffer)

End With

End Sub

Private Sub HandleQuestAccept(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Accept quest by index
' / Parameters received by client : questIndex , used by slot for array , the quest.

With UserList(UserIndex)

    Call .incomingData.ReadByte 'kill memory 1 byte
    
    Dim QuestIndex As Byte
        QuestIndex = .incomingData.ReadByte

End With
End Sub

Private Sub HandleQuestRequestInfo(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: User requires information about your current Quest

With UserList(UserIndex)

    Call .incomingData.ReadByte

End With

End Sub

Private Sub HandleBankGLDTransferencie(ByVal UserIndex As Integer)
' / Author: maTih
' / Note: Transfer gold from a vault to the other
' / Parameters received by client : targetName used by array of userlist, amount used for gold.

Dim buffer As New clsByteQueue

With UserList(UserIndex)

    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte
    
    Dim targetName  As String
    Dim Amount      As Long
    
    targetName = buffer.ReadASCIIString()
    Amount = buffer.ReadLong()
    
    Dim tmpPointer  As Integer  'Pointer used by array UserList
    Dim Actuality   As Byte     'Actuality gold for char
    
    tmpPointer = NameIndex(targetName)
    
    If Amount > .Stats.GLD Then          'check gold user
        If tmpPointer <= 0 Then          'if pointer not exist, go charfile
            If FileExist(CharPath & targetName & ".chr") Then
            
                Actuality = val(GetVar(CharPath & targetName & ".chr", "STATS", "GLD"))
                
                'if char exist, write!
                WriteVar CharPath & targetName & ".chr", "STATS", "GLD", Actuality + Amount
                WriteUpdateGold UserIndex
                
            Else
            
                WriteConsoleMsg UserIndex, "El personaje " & targetName & " no existe.", FontTypeNames.FONTTYPE_GUILD

            End If
            
        Else ' if pointer is > 0 , user online.
        
            .Stats.GLD = .Stats.GLD - Amount
            UserList(tmpPointer).Stats.GLD = UserList(tmpPointer).Stats.GLD + Amount
            WriteUpdateGold tmpPointer
            WriteUpdateGold UserIndex
            
        End If
        
    Else  'I try to transfer gold had not
    
        WriteConsoleMsg UserIndex, .name & " No tienes tanto oro!!", FontTypeNames.FONTTYPE_GUILD
    
    End If
    
    'If we get here and complete the data arrived, we destroy the buffer auxliar
    Call .incomingData.CopyBuffer(buffer)

End With

End Sub

Private Sub HandleDragObj(ByVal UserIndex As Integer)

' / Author - Unknown
' / Note : Drag objs of user Inventory
' / Parameters received by client : ObjSlot1 & ObjSlot2

Dim ObjSlot1    As Byte
Dim ObjSlot2    As Byte
Dim tmpUserObj  As UserOBJ
 

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

Private Sub HandleDragDropObj(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Drag & Drop inventory items to the target
' / Parameters recieved by client : Slot , X , Y , Amount


With UserList(UserIndex)
  
  Call .incomingData.ReadByte
  
  Dim tSlot        As Byte        'Slot destinado a dragear
  Dim tX           As Byte        'Posicion X target
  Dim tY           As Byte        'Posición Y target
  Dim Amount       As Integer     'Cantidad target
  
  tSlot = .incomingData.ReadByte
  tX = .incomingData.ReadByte
  tY = .incomingData.ReadByte
  Amount = .incomingData.ReadInteger

   'si la cantidad que ingresó el usuario es mayor a la que tiene
   'tomamos como amount todo.
   
   If Amount > .Invent.Object(tSlot).Amount Then
    Amount = .Invent.Object(tSlot).Amount
   End If

  'Declaraciones de uso, targets NPCS/USUARIOS.
  
  Dim tNpc         As Integer
  Dim tUser        As Integer
  
  tUser = MapData(.Pos.Map, tX, tY).UserIndex
  tNpc = MapData(.Pos.Map, tX, tY).NpcIndex
  
 
  'Hacemos condicionales, dando prioridades, primero usuarios & NPc's
    
  If tUser > 0 Then
   DragDrop_DragUsuario UserIndex, tUser, .Invent.Object(tSlot).ObjIndex, Amount, tSlot
   Exit Sub
  ElseIf tNpc > 0 Then
   DragDrop_DragNpc UserIndex, tNpc, .Invent.Object(tSlot).ObjIndex, Amount, tSlot
   Exit Sub
  End If

  'Bueno si estamos acá, no hay ni usuarios ni Npcs, drageamos al piso
    
    DragDrop_AlPiso UserIndex, .Invent.Object(tSlot).ObjIndex, Amount, CInt(tX), CInt(tY), tSlot
   
End With
  
  
End Sub

Private Sub HandleShowTOrneoForm(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Send package of show frmTorneo.

With UserList(UserIndex)

Call .incomingData.ReadByte

    If .flags.Privilegios > PlayerType.Consejero Then
        WriteDaoSendTorneoForm UserIndex
    End If

End With

End Sub

Public Sub HandleActionTorneo(ByVal UserIndex As Integer)

' / Author: maTih
' / Note:  Handling actions of torneo list users

With UserList(UserIndex)

    Call .incomingData.ReadByte
    
    Dim actioN  As Byte     'parameters of whats action?
    Dim Slot    As Byte     'slot 1
    Dim slot2   As Byte     'slot 2.
    
    actioN = .incomingData.ReadByte
    Slot = .incomingData.ReadByte
    slot2 = .incomingData.ReadByte
    
    'DECLARES FOR USE.
    Dim IndexA As Integer           'Obtenemos indices
    Dim IndexB As Integer           'Obtenemos Indices
    
    Select Case actioN
        Case 1                       'send to a duel, slot against slot2
        
            IndexA = NameIndex(ArrayTorneoList(Slot))
            IndexB = NameIndex(ArrayTorneoList(slot2))
            
            'we indexA to corner a
            WarpUserChar IndexA, 272, 50, 66, False
            
            'we indexB to corner b
            WarpUserChar IndexB, 272, 66, 50, False
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo > " & UserList(IndexA).name & " vs " & UserList(IndexB).name, FontTypeNames.FONTTYPE_GUILD)
      
        Case 2                       'disqualifies a user of tournament
        
            'We will use the slot for warp.
             IndexA = NameIndex(ArrayTorneoList(Slot))
             
             WarpUserChar IndexA, 1, 50, 50, False
             SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo > " & UserList(IndexA).name & " descalificado.", FontTypeNames.FONTTYPE_GUILD)
     
        Case 3                         'sets the winner by slot.
        
            IndexA = NameIndex(ArrayTorneoList(Slot))
            WarpUserChar IndexA, 1, 50, 50, False
            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo > " & UserList(IndexA).name & " Ganó el torneo.", FontTypeNames.FONTTYPE_GUILD)
    
    End Select

End With

End Sub

Private Sub HandleCreateNewAccount(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Create New Account, handling is here.
' / Parameters received by client : Name , Pass, Email , PIN.

With UserList(UserIndex)

    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte
    
    Dim name    As String
    Dim Pass    As String
    Dim Email   As String
    Dim Pin     As Byte
    
    'This variable will be spent as a ref for the error by message
    Dim ErrorS  As String
    
    'Here will store the received data to be handled
    name = buffer.ReadASCIIString()
    Pass = buffer.ReadASCIIString()
    Email = buffer.ReadASCIIString()
    Pin = buffer.ReadByte()
    
    If Account_CanCreate(name, Pass, Email, Pin, ErrorS) Then
        Accounts_Create name, Pass, Email, Pin
        WriteErrorMsg UserIndex, " La cuenta " & name & " ha sido creada!!"
    Else
        WriteErrorMsg UserIndex, ErrorS
    End If
    
    'If we get here and complete the data arrived, we destroy the buffer auxliar
    Call .incomingData.CopyBuffer(buffer)

End With

End Sub

Private Sub HandleLoginExistingAccount(ByVal UserIndex As Integer)

With UserList(UserIndex)

    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte
    
    Dim NameA As String
    Dim PassA As String
    
    NameA = buffer.ReadASCIIString
    PassA = buffer.ReadASCIIString
    
    If Account_Exist(NameA) Then
    
        If Account_CheckPass(NameA, PassA) Then
            'ACCOUNT_SENDACCOUNT NAMEA
        Else
            WriteErrorMsg UserIndex, "Contraseña incorrecta"
        End If
        
    Else
        WriteErrorMsg UserIndex, " La cuenta no existe."
    End If
    
    Call .incomingData.CopyBuffer(buffer)

End With

End Sub

Private Sub HandleLoginCharAccount(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Login existing char of account by index.

With UserList(UserIndex)

    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte
    
    Dim receivedaccName As String
    Dim targetSlot      As Byte      'apuntamos al slot recibido
    
    receivedaccName = buffer.ReadASCIIString
    targetSlot = buffer.ReadByte
    
    If Account_PjSlot(targetSlot, receivedaccName) Then
        Account_LoginChar UserIndex, targetSlot, receivedaccName
    Else
        WriteErrorMsg UserIndex, "Slot inválido , o bien No hay pj"
    End If
    
    Call .incomingData.CopyBuffer(buffer)

End With

End Sub

Private Sub HandleOffertSubast(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: offert subast actuality
' / Parameters received by client : newOffert (type data is Long)

With UserList(UserIndex)

    Call .incomingData.ReadByte
    
    Dim newOffert   As Long
    Dim refError    As String
    
    'GET DATA.
    newOffert = .incomingData.ReadLong
    
    If Not Subasta_UsuarioOferta(UserIndex, newOffert, refError) Then
        WriteConsoleMsg UserIndex, refError, FontTypeNames.FONTTYPE_CENTINELA
    End If

End With

End Sub

Private Sub HandleCreateTorneo(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Create Tournament by parameters receiveds.

With UserList(UserIndex)

    Call .incomingData.ReadByte
    
    'Metemos data en estas variables y limpiamos su socket
    Dim Cupos   As Byte
    Dim Tipo    As Byte
    Dim Precio  As Long
    Dim Tmp     As String
    
    Cupos = .incomingData.ReadByte
    Tipo = .incomingData.ReadByte
    Precio = .incomingData.ReadLong
    
    If Torneo_PuedeCrear(Cupos, Tipo, Precio, Tmp) Then
        Torneo_Crear UserIndex, Cupos, Tipo, Precio
    Else
        WriteConsoleMsg UserIndex, Tmp, FontTypeNames.FONTTYPE_GUILD
    End If

End With

End Sub

Private Sub HandleCreateSubast(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Start new subast.
' / Parameters received by client : slot, amount, minime.

With UserList(UserIndex)

    Call .incomingData.ReadByte
    
    'get data and save in variables
    Dim Slot    As Byte
    Dim Amount  As Integer
    Dim minime  As Long
    
    Slot = .incomingData.ReadByte
    Amount = .incomingData.ReadInteger
    minime = .incomingData.ReadLong
    
    'VARIABLES OF USE ERROR MESSAGE-
    Dim refError As String
    
    If Subasta_UserPuede(UserIndex, Slot, Amount, minime, refError) Then
        Subasta_Inicia UserIndex, Slot, Amount, minime
    Else
        WriteConsoleMsg UserIndex, refError, FontTypeNames.FONTTYPE_GUILD
    End If

End With

End Sub

Private Sub HandleRequestRanking(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)

    Call .incomingData.ReadByte
    
    If Not .flags.Muerto = 0 Then Exit Sub
    WriteDaoSendRanking UserIndex

End With

End Sub

Public Function DunkanSend_CreateParticle(ByVal CharIndex As Integer, ByVal TargetCharIndex As Integer, ByVal ParticleIndex As Byte) As String

' / Author: maTih

With DunkanBuffer

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.createparticle)
    Call .WriteInteger(CharIndex)
    Call .WriteInteger(TargetCharIndex)
    Call .WriteByte(ParticleIndex)
    
    DunkanSend_CreateParticle = .ReadASCIIStringFixed(.length)
    
End With

End Function

Public Function DunkanSend_CreateProjectile(ByVal CharSending As Integer, ByVal CharRecieved As Integer, GrhIndex As Integer) As String

' / Author : maTih

With DunkanBuffer

Call .WriteByte(1)

Call .WriteByte(ServerDaoPacketID.CreateProjectile)

Call .WriteInteger(CharSending)
Call .WriteInteger(CharRecieved)
Call .WriteInteger(GrhIndex)

DunkanSend_CreateProjectile = .ReadASCIIStringFixed(.length)

End With

End Function

Public Function DunkanSend_CreateDamageMap(ByVal X As Integer, ByVal Y As Integer, ByVal Damage As Integer) As String

' / Author: maTih

With DunkanBuffer

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.CreateDamageMap)
    Call .WriteByte(X)
    Call .WriteByte(Y)
    Call .WriteInteger(Damage)
    
    DunkanSend_CreateDamageMap = .ReadASCIIStringFixed(.length)
    
End With

End Function

Public Function DunkanSend_CreateMovimentWeapon(ByVal CharIndex As Integer) As String

' / Author: maTih

With DunkanBuffer

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.MovimentWeapon)
    Call .WriteInteger(CharIndex)
    
    DunkanSend_CreateMovimentWeapon = .ReadASCIIStringFixed(.length)
    
End With

End Function

Public Function DunkanSend_CreateMovimentShield(ByVal CharIndex As Integer) As String

' / Author: maTih

With DunkanBuffer

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.MovimentShield)
    Call .WriteInteger(CharIndex)
    
    DunkanSend_CreateMovimentShield = .ReadASCIIStringFixed(.length)
    
End With

End Function

Public Sub WriteDaoUpdateVentaSlot(ByVal UserIndex As Integer, ByVal vIndex As Integer, ByVal Slot As Byte)

' / Author  : maTih.-

Dim tmpObjData As ObjData

tmpObjData = ObjData(UserList(vIndex).VentaInv.Objs(Slot).ObjIndex)

With UserList(UserIndex).outgoingData
    .WriteByte 1
    .WriteByte ServerDaoPacketID.UpdateVentaSlot
    .WriteByte Slot
    
    'Enviamos GrhIndex , ObjIndex y Amount
    
    .WriteInteger tmpObjData.GrhIndex
    .WriteInteger UserList(vIndex).VentaInv.Objs(Slot).ObjIndex
    .WriteInteger UserList(vIndex).VentaInv.Objs(Slot).Amount
    
    'Enviamos el precio y el nombre del objeto.
    
    .WriteASCIIString tmpObjData.name
    .WriteLong UserList(vIndex).VentaInv.Objs(Slot).Precio
    
    'Detalles del objeto.
    
    .WriteByte tmpObjData.Minhit
    .WriteByte tmpObjData.MaxHIT
    
    .WriteByte tmpObjData.MinDef
    .WriteByte tmpObjData.MaxDef
End With

End Sub

Public Sub WriteDaoShowVentaForm(ByVal UserIndex As Integer)

' / Author  : maTih.-

With UserList(UserIndex).outgoingData
    .WriteByte 1
    .WriteByte ServerDaoPacketID.ShowVentaForm
End With

End Sub

Public Sub WriteDaoParchamentMessage(ByVal UserIndex As Integer, ByVal Message As String)

' / Author  : maTih.-

With UserList(UserIndex).outgoingData
    .WriteByte 1
    .WriteByte ServerDaoPacketID.ParchamentMsg
    .WriteASCIIString Message
End With

End Sub

Public Sub WriteDaoUpdateTargets(ByVal UserIndex As Integer, ByVal targetString As String)

' / Author  : maTih.-

With UserList(UserIndex).outgoingData
    .WriteByte 1
    .WriteByte ServerDaoPacketID.UpdateTargets
    .WriteASCIIString targetString
End With

End Sub

Public Sub WriteDaoSendQuestLVL(ByVal UserIndex As Integer, ByVal NameList As String, ByVal infoList As String)

' / Author  : maTih.-

With UserList(UserIndex).outgoingData
    .WriteByte 1
    .WriteByte ServerDaoPacketID.SendQuestLvl
    .WriteASCIIString NameList
    .WriteASCIIString infoList
End With

End Sub

Public Sub WriteDaoSendQuestForm(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex).outgoingData

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.SendQuestForm)
    Call .WriteByte(CantQ)
    
    Dim i As Long
    For i = 1 To CantQ
        Call .WriteASCIIString(tQuest(i).Recompense)
    Next i
    
End With

End Sub

Public Sub WriteDaoSendUserInventory(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)

    'El 1 hardcodeado en TODOS los subs, es el byteArray que usaremos como ID.
    Call .outgoingData.WriteByte(1)
    Call .outgoingData.WriteByte(ServerDaoPacketID.SendUserInventory)
    'primero le enviamos la cantida de objetos.
    Call .outgoingData.WriteByte(.Invent.NroItems)
    'Aca recorremos los objetos y le enviamos los nombres.
    Dim i As Long
    For i = 1 To .Invent.NroItems
        Call .outgoingData.WriteASCIIString(ObjData(.Invent.Object(i).ObjIndex).name)
    Next i
    
End With

'ENJOY!
End Sub

Public Sub WriteDaoSendSupportForm(ByVal UserIndex As Integer)

' / Author: maTih

'No tengo ganas.
Dim tmpM As Byte
With UserList(UserIndex).outgoingData

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.SendSoportFormS)
    Call .WriteByte(LastMensaje)
    
    Dim i As Long
    For i = 1 To LastMensaje
        If Users(i) <> vbNullString Then
            Call .WriteASCIIString(Users(i))
        End If
    Next i

End With

End Sub

Public Sub WriteDaoSendTorneoForm(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex).outgoingData

    Call .WriteByte(1)
    Call .WriteByte(ServerDaoPacketID.SendTorneoF)

End With

End Sub



Public Sub WriteDaoSendRanking(ByVal UserIndex As Integer)

' / Author: maTih
' / Note: Send Ranking list to Index.

With UserList(UserIndex)

'El uno hardCodeado es el ByteArray de el ServerPacketID.
'Como es privado, se usa solo en el modulo protocol, por eso está
'Hardcodeado :D

    Call .outgoingData.WriteByte(1)
    Call .outgoingData.WriteByte(ServerDaoPacketID.SendRanking)
    
    Dim i As Long
    
    For i = 1 To 10
        Call .outgoingData.WriteASCIIString(tRanking.UserNames(i) & "< " & tRanking.UserTag(i) & ">")
        Call .outgoingData.WriteByte(tRanking.UserLevels(i))
        Call .outgoingData.WriteInteger(tRanking.UserFrags(i))
        Call .outgoingData.WriteASCIIString(tRanking.UserClases(i))
    Next i

End With

End Sub
#If SeguridadDunkan Then
Public Sub WriteDaoUpdateTagTargets(ByVal UserIndex As Integer, ByVal TargetTipe As Byte, ByVal TargetIndex As Integer)

Dim name As String
Dim Stat(1 To 1) As String
Dim Cantidad As Byte

If TargetTipe = 2 Then
name = Npclist(TargetIndex).name
Else
name = UserList(TargetIndex).name
End If

If TargetTipe = 2 Then
'stats dde vida
Stat(1) = Npclist(TargetIndex).Stats.MinHp & "/" & Npclist(TargetIndex).Stats.MaxHp
Cantidad = 1
Else
Cantidad = 1

If UserList(UserIndex).PartyIndex > 0 Then
If Parties(UserList(UserIndex).PartyIndex).EsPartyLeader(UserIndex) Then

Stat(1) = "Invitar a party"

ElseIf UserList(TargetIndex).PartyIndex > 0 Then
If Parties(UserList(TargetIndex).PartyIndex).EsPartyLeader(TargetIndex) Then

Stat(2) = "Solicitar party"

End If
End If
With UserList(UserIndex)

Call .outgoingData.WriteByte(1)
Call .outgoingData.WriteByte(ServerDaoPacketID.UpdateTagTarget)
Call .outgoingData.WriteASCIIString(name)

End With


End Sub
#End If
Function EncryptINI$(Strg$, Password$)

'Author - Unknown
'Note : Encrypt Strg$ loop for Password$

   Dim b$, S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, p$
   j = 1
   For i = 1 To Len(Password$)
     p$ = p$ & Asc(mid$(Password$, i, 1))
   Next
   
   For i = 1 To Len(Strg$)
     A1 = Asc(mid$(p$, j, 1))
     j = j + 1: If j > Len(p$) Then j = 1
     A2 = Asc(mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     S$ = S$ + b$
   Next
   
   EncryptINI$ = S$
   
End Function
 
Function DecryptINI$(Strg$, Password$)

'Author - Unknown
'Note : Decrypt Strg$ loop for Password$

   Dim b$, S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, p$
   j = 1
   For i = 1 To Len(Password$)
     p$ = p$ & Asc(mid$(Password$, i, 1))
   Next
   
   For i = 1 To Len(Strg$) Step 2
     A1 = Asc(mid$(p$, j, 1))
     j = j + 1: If j > Len(p$) Then j = 1
     b$ = mid$(Strg$, i, 2)
     A3 = val("&H" + b$)
     A2 = A1 Xor A3
     S$ = S$ + Chr$(A2)
   Next
   
   DecryptINI$ = S$
   
End Function

