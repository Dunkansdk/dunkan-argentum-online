Attribute VB_Name = "mod_QuestByLVL"
Option Explicit

Type questReward
     Objetos()          As Obj              'Objetos :P
     Oro                As Long             'Oro para usuario.
     ExpReward          As Long             'Experiencia.
End Type

Type questList
     Nivel              As Byte             'Nivel requerido para la quest.
     NPCs               As Byte             'Cuantos npcs.
     NPCNum             As Integer          'Que npc.
     Info               As String           'Informacion para el usuario.
     name               As String           'Nombre de la quest.
     Recompensa         As questReward      'Recompensas de las quest.
End Type

Public questArray()     As questList

Sub Quest_Load()

' @author   :  maTih.-
' @date     :  20110125
' @note     :  load array of quests

Dim loopX       As Long
Dim QuestFile   As String

QuestFile = App.Path & "\Dat\Quest.txt"

ReDim questArray(1 To val(GetVar(QuestFile, "INIT", "Cantidad"))) As questList

For loopX = 1 To val(GetVar(QuestFile, "INIT", "Cantidad"))

    With questArray(loopX)
    
         .Nivel = val(GetVar(QuestFile, "Q" & loopX, "Nivel"))
         .NPCs = val(GetVar(QuestFile, "Q" & loopX, "CantidadNPCs"))
         .NPCNum = val(GetVar(QuestFile, "Q" & loopX, "NpcNumero"))
         .Info = GetVar(QuestFile, "Q" & loopX, "Informacion")
         .name = GetVar(QuestFile, "Q" & loopX, "Nombre")
         
         Quest_LoadPremiers loopX
    End With
    
Next loopX

End Sub

Sub Quest_LoadPremiers(ByVal QuestIndex As Byte)

' @author   :  maTih.-
' @date     :  20110125
' @note     :  Load  the premiers of questIndex

Dim loopX       As Long
Dim QuestPath   As String

    QuestPath = App.Path & "\Dat\Quest.txt"

    With questArray(QuestIndex).Recompensa
    
        ReDim .Objetos(1 To val(GetVar(QuestPath, "Q" & QuestIndex, "NumeroObjs")))
    
        For loopX = 1 To val(GetVar(QuestPath, "Q" & QuestIndex, "NumeroObjs"))
            .Objetos(loopX).ObjIndex = val(ReadField(1, GetVar(QuestPath, "Q" & QuestIndex, "Premio" & loopX), Asc("-")))
            .Objetos(loopX).Amount = val(ReadField(2, GetVar(QuestPath, "Q" & QuestIndex, "Premio" & loopX), Asc("-")))
        Next loopX
        
            .ExpReward = val(GetVar(QuestPath, "Q" & loopX, "ExpRecompensa"))
            .Oro = val(GetVar(QuestPath, "Q" & loopX, "OroRecompensa"))
    End With

End Sub

Sub Quest_Enviar(ByVal Usuario As Integer, ByRef refName As String, ByRef refInfo As String)

' @author   :  maTih.-
' @date     :  20110125
' @note     :  Send posibility quest of the userlevel

    Dim loopX           As Long
    Dim myELV           As Byte
    Dim NameString      As String
    Dim InfoString      As String
    
    myELV = UserList(Usuario).Stats.ELV
    
    For loopX = 1 To UBound(questArray())
        If questArray(loopX).Nivel = myELV Then
           NameString = NameString & questArray(loopX).name & "|"
           InfoString = InfoString & questArray(loopX).Info & "|"
        End If
    Next loopX
    
    refName = NameString
    refInfo = InfoString
    
    WriteDaoSendQuestLVL Usuario, NameString, InfoString
    
End Sub

Sub Quest_Aceptar(ByVal UserIndex As Integer, ByVal questName As String)

' @author   :  maTih.-
' @date     :  20110125
' @note     :  User accept quest by questName

Dim QuestIndex      As Byte

Quest_GetQuestIndexByName QuestIndex, questName

With UserList(UserIndex).qUser
     .NpcKilleds = 0
     .NPCNum = questArray(QuestIndex).NPCNum
     .QuestIndex = QuestIndex
End With

    WriteConsoleMsg UserIndex, "Aceptaste la misión! " & questArray(QuestIndex).Info, FontTypeNames.FONTTYPE_CITIZEN

End Sub

Sub Quest_GetQuestIndexByName(ByRef QuestIndex As Byte, ByVal qName As String)

' @author   :  maTih.-
' @date     :  20110125
' @note     :  Convert questName to index

Dim loopX       As Long

For loopX = 1 To UBound(questArray())
    If questArray(loopX).name = qName Then
        QuestIndex = loopX
        Exit Sub
    End If
Next loopX

End Sub

Sub Quest_UserCheckNPCs(ByVal UserIndex As Integer, ByVal NPCKill)

    With UserList(UserIndex).qUser
         
         If .NPCNum = NPCKill Then
            .NpcKilleds = .NpcKilleds + 1
            
            If .NpcKilleds >= questArray(.QuestIndex).NPCs Then
                
                Quest_PremierUser UserIndex
                Exit Sub
            End If
            
            WriteConsoleMsg UserIndex, "Te restan asesinar : " & (questArray(.QuestIndex).NPCs - .NpcKilleds) & " - " & Quest_GetNPCName(.NPCNum) & IIf((questArray(.QuestIndex).NPCs - .NpcKilleds) > 1, "s", ""), FontTypeNames.FONTTYPE_CITIZEN
            
        End If
         
    End With

End Sub

Sub Quest_PremierUser(ByVal UserIndex As Integer)

' @author   :  maTih.-
' @date     :  20110125
' @note     :  Premier userIndex to premierList of quest, and clear type.

    With UserList(UserIndex)
    
    Dim loopX           As Long
    Dim fMessage        As String
        
        'Quest start string.
        fMessage = "Has terminado la misión! Ganaste : "
        
        'Exp of quest not null?
        If questArray(.qUser.QuestIndex).Recompensa.ExpReward > 0 Then
        
            .Stats.Exp = .Stats.Exp + questArray(.qUser.QuestIndex).Recompensa.ExpReward
            
            'Update of Client.
            WriteUpdateExp UserIndex
            
            'Check level of the user.
            CheckUserLevel UserIndex
            
            fMessage = fMessage & questArray(.qUser.QuestIndex).Recompensa.ExpReward & " puntos de experiencia."
            
        End If
        
        'Gold of Quest not null?
        If questArray(.qUser.QuestIndex).Recompensa.Oro > 0 Then
            
            .Stats.GLD = .Stats.GLD + questArray(.qUser.QuestIndex).Recompensa.Oro
            
            WriteUpdateGold UserIndex
            
            fMessage = fMessage & questArray(.qUser.QuestIndex).Recompensa.Oro & " Monedas de oro!!"
            
        End If
        
        'Message to user.
        WriteConsoleMsg UserIndex, fMessage, FontTypeNames.FONTTYPE_CITIZEN
        
        If UBound(questArray(.qUser.QuestIndex).Recompensa.Objetos()) <= 0 Then Exit Sub
        
        'Fill the premierObjs to userIndex.
        For loopX = 1 To UBound(questArray(.qUser.QuestIndex).Recompensa.Objetos())
                
            If questArray(.qUser.QuestIndex).Recompensa.Objetos(loopX).ObjIndex <> 0 Then
                
                MeterItemEnInventario UserIndex, questArray(.qUser.QuestIndex).Recompensa.Objetos(loopX)
            
                WriteConsoleMsg UserIndex, "Has ganado : " & questArray(.qUser.QuestIndex).Recompensa.Objetos(loopX).Amount & " - " & ObjData(questArray(.qUser.QuestIndex).Recompensa.Objetos(loopX).ObjIndex).name, FontTypeNames.FONTTYPE_CITIZEN
            
            End If
        
        Next loopX

    End With

End Sub

Function Quest_GetNPCName(ByVal NPCNum As String) As String

' @author   :  maTih.-
' @date     :  20110125
' @note     :  Return NPCName by npcNumber.

    Quest_GetNPCName = GetVar(App.Path & "\Dat\NPCs.dat", "NPC" & NPCNum, "Name")
    
End Function

