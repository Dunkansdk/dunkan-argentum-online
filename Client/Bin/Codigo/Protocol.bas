Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Public Enum ServerPacketID
    SendCuenta
    CuentaSV
    Compra
    ' Dunkan AO Packets. maTih.-
    CreateSpell             ' CREA HECHIZOS/PARTICULAS.
    CreateMeditation        ' CREA MEDITACION.
    CreateArrow             ' CREA FLECHAS.
    CreateDamage            ' CREA DAÑO.
    ' Dunkan AO Packets. maTih.-
    logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    Disconnect              ' FINOK
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    ErrorMsg                ' ERR
    Dumb                    ' DUMB
    SetcharInvisible        ' NOVER
    MeditateToggle          ' MEDOK
    DumbNoMore              ' NESTUP
    ParalizeOK              ' PARADOK
    Pong
    UpdateTagAndStatus
    
    AddSlots
    MultiMessage
End Enum

Public Enum ClientPacketID
    LogCuenta
    LogPJ
    Compra
    Potear
    Bando
    LoginExistingChar       'OLOGIN
    'Dunkan ao Packets
    DragInventory
    DragToPos
    SpawnBot
    'Dunkan ao Packets
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    WorkLeftClick           'WLC
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    MoveSpell               'DESPHE
    Online                  '/ONLINE
    Quit                    '/SALIR
    Meditate                '/MEDITAR
    Ping                    '/PING
    GMCommands
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
End Enum

Public FontTypes(20) As tFont


''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 255
        .Green = 255
        .Blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 32
        .Green = 51
        .Blue = 223
        .bold = 1
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 65
        .Green = 190
        .Blue = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 65
        .Green = 190
        .Blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 130
        .Green = 130
        .Blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 255
        .Green = 180
        .Blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).Green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 255
        .Green = 255
        .Blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).Green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .Red = 130
        .Green = 130
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .Red = 255
        .Green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .Green = 200
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .Red = 255
        .Green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .Red = 30
        .Green = 255
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .Blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .Red = 30
        .Green = 150
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
        .bold = 1
    End With
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
Dim daoID As Byte
daoID = incomingData.PeekByte

    Select Case daoID
          ' Dunkan AO Packets. maTih.-
            
          Case ServerPacketID.Compra
               Call HandleCompra
            
          Case ServerPacketID.CreateSpell
            Call HandleCreateSpell
            
          Case ServerPacketID.CreateMeditation
            Call HandleCreateMeditation
            
          Case ServerPacketID.CreateArrow
            Call HandleCreateArrow
            
          Case ServerPacketID.CreateDamage
            Call HandleCreateDamage
        
          ' Dunkan AO Packets. maTih.-
          Case ServerPacketID.logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
            
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
            
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave

        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              ' LLU
            Call HandleRainToggle
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes

        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb

        Case ServerPacketID.SetcharInvisible            ' NOVER
            Call HandleSetInvisible

        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK

        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
            
        
           
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
            
        Case ServerPacketID.SendCuenta
            Call Mod_DunkanProtocol.HandleSendCuenta
            
#If SeguridadAlkon Then
        Case Else
            Call HandleIncomingDataEx
#Else
        Case Else
            'ERROR : Abort!
            Exit Sub
#End If
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If
End Sub

Public Sub HandleMultiMessage()

    Dim BodyPart As Byte
    Dim Daño As Integer
    
With incomingData
    Call .ReadByte
    
    Select Case .ReadByte
        Case eMessages.DontSeeAnything
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_NO_VES_NADA_INTERESANTE, 65, 190, 156, False, False, True)
        
        Case eMessages.NPCSwing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.NPCKillUser
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldUser
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldOther
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserSwing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.SafeModeOn
            Call frmMain.ControlSM(eSMType.sSafemode, True)
        
        Case eMessages.SafeModeOff
            Call frmMain.ControlSM(eSMType.sSafemode, False)
        
        Case eMessages.ResuscitationSafeOff
            Call frmMain.ControlSM(eSMType.sResucitation, False)
         
        Case eMessages.ResuscitationSafeOn
            Call frmMain.ControlSM(eSMType.sResucitation, True)
        
        Case eMessages.NobilityLost
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
        
        Case eMessages.CantUseWhileMeditating
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
        
        Case eMessages.NPCHitUser
            Select Case incomingData.ReadByte()
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.UserHitNPC
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
        
        Case eMessages.UserAttackedSwing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserHittedByUser
            Dim AttackerName As String
            
            AttackerName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
            BodyPart = incomingData.ReadByte()
            Daño = incomingData.ReadInteger()
            
            Select Case BodyPart
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.UserHittedUser

            Dim VictimName As String
            
            VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
            BodyPart = incomingData.ReadByte()
            Daño = incomingData.ReadInteger()
            
            Select Case BodyPart
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.WorkRequestTarget
            UsingSkill = incomingData.ReadByte()
            
            frmMain.MousePointer = 2
            
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select

        Case eMessages.HaveKilledUser
            Dim level As Long
            Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & charlist(.ReadInteger).Nombre & MENSAJE_22, 255, 0, 0, True, False)
            level = .ReadLong
            Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & level & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
            If ClientSetup.bKill And ClientSetup.bActive Then
                If level / 2 > ClientSetup.byMurderedLevel Then
                    isCapturePending = True
                End If
            End If
        Case eMessages.UserKill
            Call ShowConsoleMsg(charlist(.ReadInteger).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)
            If ClientSetup.bDie And ClientSetup.bActive Then _
                isCapturePending = True
        Case eMessages.EarnExp
            Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadLong & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        Case eMessages.GoHome
            Dim Distance As Byte
            Dim Hogar As String
            Dim tiempo As Integer
            Distance = .ReadByte
            tiempo = .ReadInteger
            Hogar = .ReadASCIIString
            Call ShowConsoleMsg("Estás a " & Distance & " mapas de distancia de " & Hogar & ", este viaje durará " & tiempo & " segundos.", 255, 0, 0, True)
            Traveling = True
        Case eMessages.FinishHome
            Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
            Traveling = False
        Case eMessages.CancelGoHome
            Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
            Traveling = False
    End Select
End With
End Sub



''
' Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    
    ' Variable initialization
    EngineRun = True
    Nombres = True
    
    'Set connected state
    Call SetConnected
        
    If bShowTutorial Then frmTutorial.Show
    
  '  If frmMain.PicInv.Visible = True Then Inventario.DrawInventory
    
    'Show tip
    If tipf = "1" And PrimeraVez Then
        Call CargarTip
        frmtip.Visible = True
        PrimeraVez = False
    End If
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
#If UsarWrench = 1 Then
    frmMain.Socket1.Disconnect
#Else
    If frmMain.Winsock1.State <> sckClosed Then _
        frmMain.Winsock1.Close
#End If
    
    'Hide main form
    frmMain.Visible = False
    
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    

    
    'Show connection form

    'Reset global vars
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    bRain = False
    bFogata = False
    SkillPoints = 0
    Comerciando = False
    'new
    Traveling = False
    'Delete all kind of dialogs
    Call CleanDialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        charlist(i).Invisible = False
    Next i

    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> frmMain.Name Then
            Unload frm
        End If
    Next
    
            frmConnect.Show
            
    For i = 1 To MAX_INVENTORY_SLOTS
        Call Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", False)
    Next i
    
#If SeguridadAlkon Then
    Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If

    Call Audio.PlayMIDI("2.mid")
End Sub
''
' Handles the CommerceEnd message.


''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, True)
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, False)
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, False)
End Sub

''
' Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, True)
End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    frmMain.shpEnergia.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 98)
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    
    If UserMinMAN <> 0 And UserMaxMAN <> 0 Then
        frmMain.shpMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 98)
    End If
    
    'frmMain.shpMana.Left = 550 + (frmMain.shpMana.Width - 98)

End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    
    If (UserMaxHP < 0) Then UserMaxHP = 0
    If (UserMinHP < 0) Then UserMinHP = 0
    
        frmMain.lblVida = UserMinHP & "/" & UserMaxHP
        frmMain.shpVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 98)
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If
    
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tmpMap As Long
    
    tmpMap = incomingData.ReadLong()
        
    #If SeguridadDunkan Then
    UserMap = KeyMapUsate(tmpMap)
    #End If
    
    UserMap = CInt(tmpMap)
'TODO: Once on-the-fly editor is implemented check for map version before loading....
'For now we just drop it
    Call incomingData.ReadInteger
        
#If SeguridadAlkon Then
    Call InitMI
#End If
    
    If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
        Call SwitchMap(UserMap)
        If bLluvia(UserMap) = 0 Then
            If bRain Then
                Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
            End If
        End If
    Else
        'no encontramos el mapa en el hd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call CloseClient
    End If
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    
    If UserPos.X <> 0 Then
    'Remove char from old position
        If MapData(UserPos.X, UserPos.Y).charIndex = UserCharIndex Then
            MapData(UserPos.X, UserPos.Y).charIndex = 0
        End If
    End If
    
    'Set new pos
    UserPos.X = incomingData.ReadByte()
    UserPos.Y = incomingData.ReadByte()
    
    'Set char
    MapData(UserPos.X, UserPos.Y).charIndex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                
    'Update pos label
    frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Select Case incomingData.ReadByte()
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
    End Select
End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String
    
    attacker = charlist(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Victim As String
    
    Victim = charlist(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat        As String
    Dim charIndex   As Integer
    Dim R           As Byte
    Dim G           As Byte
    Dim B           As Byte
    Dim Spell       As Byte
    
    chat = Buffer.ReadASCIIString()
    charIndex = Buffer.ReadInteger()
    
    R = Buffer.ReadByte()
    G = Buffer.ReadByte()
    B = Buffer.ReadByte()
    
    Spell = Buffer.ReadBoolean()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(charIndex).Active Then _
        Call Dialogos.CreateDialog(Trim$(chat), charIndex, RGB(R, G, B), CByte(Spell))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                R = 255
            Else
                R = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                G = 255
            Else
                G = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                B = 255
            Else
                B = Val(str)
            End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), R, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
        

    End If
'    Call checkText(chat)
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/07/08 (NicoNZ)
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim str As String
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim tmp As Integer
    Dim Cont As Integer
    
    
    chat = Buffer.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                R = 255
            Else
                R = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                G = 255
            Else
                G = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                B = 255
            Else
                B = Val(str)
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), R, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos = charlist(UserCharIndex).Pos
    
    If UserPos.X = 0 Then Exit Sub
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim charIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim Weapon As Integer
    Dim Shield As Integer
    Dim Helmet As Integer
    Dim privs As Integer
    Dim NickColor As Byte
    
    charIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    Weapon = Buffer.ReadInteger()
    Shield = Buffer.ReadInteger()
    Helmet = Buffer.ReadInteger()
    

    
    With charlist(charIndex)

        Call SetCharacterFx(charIndex, Buffer.ReadInteger(), Buffer.ReadInteger())
        
        .Nombre = Buffer.ReadASCIIString()
        NickColor = Buffer.ReadByte()

        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        privs = Buffer.ReadByte()
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .Priv = Log(privs) / Log(2)
        Else
            .Priv = 0
        End If
    End With
    
    Call MakeChar(charIndex, Body, Head, Heading, X, Y, Weapon, Shield, Helmet)
    
    Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim charIndex As Integer
    charIndex = incomingData.ReadInteger
    charlist(charIndex).Nombre = incomingData.ReadASCIIString
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charIndex As Integer
    
    charIndex = incomingData.ReadInteger()
    
    Call EraseChar(charIndex)
    Call RefreshAllChars
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    charIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    With charlist(charIndex)
    
        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        If .Priv <> 1 And .Priv <> 2 And .Priv <> 3 And .Priv <> 5 And .Priv <> 25 Then
            Call DoPasosFx(charIndex)
        End If
        
    End With
    
    Call MoveCharbyPos(charIndex, X, Y)
    
    Call RefreshAllChars
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'***************************************************
    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charIndex As Integer
    Dim tempint As Integer
    Dim HeadIndex As Integer
    
    charIndex = incomingData.ReadInteger()
    
    With charlist(charIndex)
        tempint = incomingData.ReadInteger()
        
        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(tempint)
            .iBody = tempint
        End If
        
        
        HeadIndex = incomingData.ReadInteger()
        
        If HeadIndex < LBound(HeadData()) Or HeadIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(HeadIndex)
            .iHead = HeadIndex
        End If
        
        .Muerto = (HeadIndex = CASPER_HEAD)

        .Heading = incomingData.ReadByte()
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Casco = CascoAnimData(tempint)
        
        Call SetCharacterFx(charIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    End With
    
    Call RefreshAllChars
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    MapData(X, Y).OBJGrh.GrhIndex = incomingData.ReadInteger()
    
    Call InitGrh(MapData(X, Y).OBJGrh, MapData(X, Y).OBJGrh.GrhIndex)
    
    'Guardo el EffectIndex del array de particulas()
    'Si es un teleport.
    If Not MapData(X, Y).OBJGrh.GrhIndex <> 669 And Not MapData(X, Y).Particle <> 0 Then MapData(X, Y).Particle = Mod_DX8_Particle.Effect_Teleport_Begin(Engine_TPtoSPX(X), Engine_TPtoSPY(Y) - 30, 1, 300, 42)
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    MapData(X, Y).OBJGrh.GrhIndex = 0
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If incomingData.ReadBoolean() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadByte()
    
    If currentMidi Then
        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", incomingData.ReadInteger())
    Else
        'Remove the bytes to prevent errors
        Call incomingData.ReadInteger
    End If
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadByte()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
        
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

'

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa
End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4)
    If bRain Then
        If bLluvia(UserMap) Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            If bTecho Then
                Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
            End If
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    
    bRain = Not bRain
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    charIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(charIndex, fX, Loops)
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()

    'Stats
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    
    Dim BWidth As Byte
    
    '***************************
    
    ' - Mana
    If UserMaxMAN > 0 Then _
    frmMain.shpMana.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 98)

    ' - Vida
    frmMain.shpVida.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 98)
    
    '***************************
    
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If

End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill
        Case Magia
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
        Case Pesca
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
        Case Talar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
        Case Mineria
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
    End Select
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    
    Dim objIndex As Integer
    Dim Name As String
    Dim amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim value As Single
    Dim UsaSlot As Boolean
    
    slot = Buffer.ReadByte()
    objIndex = Buffer.ReadInteger()
    Name = Buffer.ReadASCIIString()
    amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadInteger()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    MaxDef = Buffer.ReadInteger()
    MinDef = Buffer.ReadInteger


    Call Inventario.SetItem(slot, objIndex, amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, Name, UsaSlot)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    MaxInventorySlots = incomingData.ReadByte
End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡Has terminado de trabajar!", .Red, .Green, .Blue, .bold, .italic)
    End With
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub


''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .objIndex = Buffer.ReadInteger()
        .Name = Buffer.ReadASCIIString()
        .amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
        .Valor = Buffer.ReadLong()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(slot, .objIndex, .amount, _
                .Equipped, .GrhIndex, .OBJType, .MaxHit, _
                .MinHit, .MaxDef, .MinDef, .Valor, .Name, True)
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    UserHechizos(slot) = Buffer.ReadInteger()
    
    If slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(slot - 1) = Buffer.ReadASCIIString()
    Else
        Call frmMain.hlst.AddItem(Buffer.ReadASCIIString())
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                For i = 1 To NUMATRIBUTES
                    .lblAtributos(i).Caption = UserAtributos(i)
                Next i
                
                .UpdateStats
            End If
        End With
    Else
        LlegaronAtrib = True
    End If
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmasHerrero(i)
            .Name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .objIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    With frmHerrero
        ' Inicializo los inventarios
        Call InvLingosHerreria(1).Initialize(DirectD3D8, .picLingotes0, 3, , , , , , False)
        Call InvLingosHerreria(2).Initialize(DirectD3D8, .picLingotes1, 3, , , , , , False)
        Call InvLingosHerreria(3).Initialize(DirectD3D8, .picLingotes2, 3, , , , , , False)
        Call InvLingosHerreria(4).Initialize(DirectD3D8, .picLingotes3, 3, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1, True)
    End With
    
    For i = 1 To Count
        With ArmasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmasHerrero(k).objIndex Then
                        j = j + 1
                
                        ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                        HerreroMejorar(j).Name = .Name
                        HerreroMejorar(j).GrhIndex = .GrhIndex
                        HerreroMejorar(j).objIndex = .objIndex
                        HerreroMejorar(j).UpgradeName = ArmasHerrero(k).Name
                        HerreroMejorar(j).UpgradeGrhIndex = ArmasHerrero(k).GrhIndex
                        HerreroMejorar(j).LinH = ArmasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(j).LinP = ArmasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(j).LinO = ArmasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            .Name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .objIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    j = UBound(HerreroMejorar)
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmadurasHerrero(k).objIndex Then
                        j = j + 1
                
                        ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                        HerreroMejorar(j).Name = .Name
                        HerreroMejorar(j).GrhIndex = .GrhIndex
                        HerreroMejorar(j).objIndex = .objIndex
                        HerreroMejorar(j).UpgradeName = ArmadurasHerrero(k).Name
                        HerreroMejorar(j).UpgradeGrhIndex = ArmadurasHerrero(k).GrhIndex
                        HerreroMejorar(j).LinH = ArmadurasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(j).LinP = ArmadurasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(j).LinO = ArmadurasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ObjCarpintero(i)
            .Name = Buffer.ReadASCIIString()        'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .Madera = Buffer.ReadInteger()          'The wood needed
            .MaderaElfica = Buffer.ReadInteger()    'The elfic wood needed
            .objIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    With frmCarp
        ' Inicializo los inventarios
        Call InvMaderasCarpinteria(1).Initialize(DirectD3D8, .picMaderas0, 2, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(DirectD3D8, .picMaderas1, 2, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(DirectD3D8, .picMaderas2, 2, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(DirectD3D8, .picMaderas3, 2, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With
    
    For i = 1 To Count
        With ObjCarpintero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ObjCarpintero(k).objIndex Then
                        j = j + 1
                
                        ReDim Preserve CarpinteroMejorar(j) As tItemsConstruibles
                        
                        CarpinteroMejorar(j).Name = .Name
                        CarpinteroMejorar(j).GrhIndex = .GrhIndex
                        CarpinteroMejorar(j).objIndex = .objIndex
                        CarpinteroMejorar(j).UpgradeName = ObjCarpintero(k).Name
                        CarpinteroMejorar(j).UpgradeGrhIndex = ObjCarpintero(k).GrhIndex
                        CarpinteroMejorar(j).Madera = ObjCarpintero(k).Madera - .Madera * 0.85
                        CarpinteroMejorar(j).MaderaElfica = ObjCarpintero(k).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call MsgBox(Buffer.ReadASCIIString())
    
    If frmConnect.Visible Then
#If UsarWrench = 1 Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    
    Call InitCartel(tmp, Buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With NPCInventory(slot)
        .Name = Buffer.ReadASCIIString()
        .amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadInteger()
        .objIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
        .ItemUsableByUser = Buffer.ReadBoolean()
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    frmMain.lblHambre = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblSed = UserMinAGU & "/" & UserMaxAGU

    Dim BWidth As Byte
    
    BWidth = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 75)
    
    frmMain.shpHambre.Width = 75 - BWidth
    frmMain.shpHambre.Left = 584 + (75 - frmMain.shpHambre.Width)
    
    frmMain.shpHambre.Visible = (BWidth <> 75)
    '*********************************
    
    BWidth = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 75)
    
    frmMain.shpSed.Width = 75 - BWidth
    frmMain.shpSed.Left = 584 + (75 - frmMain.shpSed.Width)
    
    frmMain.shpSed.Visible = (BWidth <> 75)
    
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .Promedio = incomingData.ReadLong()
    End With
    
    LlegoFama = True
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    
    Call frmMain.LightSkillStar(True)
End Sub


''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charIndex As Integer
    
    charIndex = incomingData.ReadInteger()
    charlist(charIndex).Invisible = incomingData.ReadBoolean()
    charlist(charIndex).timeInvi = incomingData.ReadInteger()
    
    
#If SeguridadAlkon Then
    If charlist(charIndex).Invisible Then
        Call MI(CualMI).SetInvisible(charIndex)
    Else
        Call MI(CualMI).ResetInvisible(charIndex)
    End If
#End If

End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
    UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
    UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()
    
    With frmCrearPersonaje
        .lblAtributos(eAtributos.Fuerza) = UserAtributos(eAtributos.Fuerza)
        .lblAtributos(eAtributos.Agilidad) = UserAtributos(eAtributos.Agilidad)
        .lblAtributos(eAtributos.Inteligencia) = UserAtributos(eAtributos.Inteligencia)
        .lblAtributos(eAtributos.Carisma) = UserAtributos(eAtributos.Carisma)
        .lblAtributos(eAtributos.Constitucion) = UserAtributos(eAtributos.Constitucion)
        
        .UpdateStats
    End With
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Now the server send the percentage of progress of the skills.
'***************************************************
    If incomingData.length < 2 + NUMSKILLS * 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserClase = incomingData.ReadByte
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        PorcentajeSkills(i) = incomingData.ReadByte()
    Next i
    LlegaronSkills = True
End Sub

''
' Handles the TrainerCreatureList message.


''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the TradeOK message.


''
' Handles the SendNight message.

Private Sub HandleSendNight()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tBool As Boolean 'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?
    tBool = incomingData.ReadBoolean()
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub



''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
End Sub



''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'*************************************Su**************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGMPanelForm message.



''


''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, True)
    
    pingTime = 0
End Sub



''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim charIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    
    charIndex = Buffer.ReadInteger()
    NickColor = Buffer.ReadByte()
    UserTag = Buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With charlist(charIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        .Nombre = UserTag
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginExistingChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
#If SeguridadAlkon Then
        Call .WriteASCIIStringFixed(MD5HushYo)
#End If
    End With
End Sub

''
' Writes the "ThrowDices" message to the outgoing data buffer.


''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginNewChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
        
'#If SeguridadAlkon Then
'        Call .WriteASCIIStringFixed(UserPassword)
'#Else
'        Call .WriteASCIIString(UserPassword)
'#End If
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
#If SeguridadAlkon Then
        Call .WriteASCIIStringFixed(MD5HushYo)
#End If
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(UserHead)
        
        'Call .WriteASCIIString(UserEmail)
        
        Call .WriteByte(UserHogar)
        Call .WriteASCIIString(Cuenta.Name)
    End With
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Yell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal charIndex As Integer, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Whisper" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteInteger(charIndex)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseSpellMacro" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)

    End With
End Sub


''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub


''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpellInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
    End With
End Sub


' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub


''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub
''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub



''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Silence)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Working" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Hiding" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(Time)
    End With
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharStats" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharStats)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharGold)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInventory" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharSkills" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Forgive" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Forgive)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.banChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld(ByVal StartMAP As Integer, ByVal EndMAP As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanWorld" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanWorld)
    Call outgoingData.WriteInteger(StartMAP)
    Call outgoingData.WriteInteger(EndMAP)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.nickToIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
            Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TeleportCreate)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Radio)
    End With
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetCharDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CitizenMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CitizenMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CriminalMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CriminalMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumb)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.dumpIPTables)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CouncilKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            For i = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)
    End With
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIAll)
        
        Call .WriteByte(midiID)
    End With
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemovePunishment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetMOTD" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetMOTD)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ImperialArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnOffServer" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnCriminal" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResetFactions)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharMail)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterPassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
    End With
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
    End With
End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
    End With
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowServerForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowServerForm)
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Night" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.night)
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Restart)
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetAutoUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatColor" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChatColor)
        
        Call .WriteByte(R)
        Call .WriteByte(G)
        Call .WriteByte(B)
    End With
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount
End Sub

''
' Writes the "ShareNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/2009
'Writes the "SetIniVar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)
    End With
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.IsWritable Then
        'Put data back in the bytequeue
        Call outgoingData.WriteASCIIStringFixed(sdData)
        
        Exit Sub
    End If
    
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

#If SeguridadAlkon Then
    Dim data() As Byte
    
    data = StrConv(sdData, vbFromUnicode)
    
    Call DataSent(data)
    
    sdData = StrConv(data, vbUnicode)
#End If
    'Send data!
#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub

' *****************************************************
' ********************* DUNKAN AO *********************
' *****************************************************


