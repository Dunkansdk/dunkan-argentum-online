Attribute VB_Name = "moduleQuest"
' / - Module Quest's - - - - - - - - - - - - -

' / Author: maTih
' / Coded for: DunkanAO

' - Tabulado y organizado por Dunkan
 
'Módulo para controlar quest de usuarios.
 
Option Explicit
 
Public Type tQuestU
   nNPC     As Integer  'Cuantos NPCS
   rNPC     As Byte     'Nro del NPC
   rUser    As Byte     'Cantidad de usuarios
End Type
 

Public Type TypQuest
   RecompenseGold   As Long      'Oro de premio
   RecompenseExp    As Long      'Eexperiencia de premio
   RequieredNPC     As Byte      'NPC Cuantos
   'RemaingNPC      As Byte      'NPC Faltan
   RequieredNPCI    As Integer   'NPC Num Requerido
   RequieredUser    As Byte      'Users Cuantos
   'RemaingUser     As Byte      'Users Faltan
   Recompense       As String    'Recompensa text.
End Type
 
Public tQuest() As TypQuest
Public CantQ As Byte

Public Sub Quest_Cargar()

' / Author: maTih

On Error GoTo ErrorHandler
 
CantQ = val(GetVar(App.Path & "\Dat\Quest.maTih", "INIT", "NumQuest"))

If CantQ = 0 Then Exit Sub

ReDim tQuest(1 To CantQ) ' Lento !! Optimizar. - Dunkan
 
Dim loopC As Long

For loopC = 1 To CantQ
    With tQuest(loopC)
        .RequieredNPC = val(GetVar(App.Path & "\Dat\Quest.maTih", "QUEST" & loopC, "CantidadNPC"))
        .RequieredNPCI = val(GetVar(App.Path & "\Dat\Quest.maTih", "QUEST" & loopC, "NpcNumero"))
        .RequieredUser = val(GetVar(App.Path & "\Dat\Quest.maTih", "QUEST" & loopC, "CantidadUsuarios"))
        .RecompenseGold = val(GetVar(App.Path & "\Dat\Quest.maTih", "QUEST" & loopC, "RecompensaORO"))
        .RecompenseExp = val(GetVar(App.Path & "\Dat\Quest.maTih", "QUEST" & loopC, "RecompensaEXP"))
        .Recompense = GetVar(App.Path & "\Dat\Quest.maTih", "QUEST" & loopC, "RecompensaTexto")
    End With
Next loopC

ErrorHandler:
    LogCriticEvent "Error cargando quest"
End Sub
 
Public Sub Quest_uRequiereInfo(ByVal User As Integer)

' / Author: maTih

With UserList(User)

    If .Stats.QuestIndex <= 0 Then Exit Sub
    
End With

End Sub
 
Public Sub Quest_uInicia(ByVal UserIndex As Integer, ByVal QuestSelected As Byte)

' / Author: maTih

On Error GoTo errHandler

With UserList(UserIndex)
 
    If QuestSelected <= 0 Then LogCriticEvent "CATASTROFE EN QUESTS, SELECTED DIO NEGATIVO ó CERO": Exit Sub
     
    If .flags.Muerto <> 0 Then
        WriteConsoleMsg UserIndex, "Estás muerto, solo puedes iniciar una quest estando vivo!", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If
     
    If Not .Stats.QuestIndex = 0 Then
        WriteConsoleMsg UserIndex, "Ya estás en una quest!", FontTypeNames.FONTTYPE_GUILD
        Exit Sub
    End If
 
1   .Stats.QuestIndex = QuestSelected

    Debug.Print .Stats.QuestIndex
    
    If tQuest(QuestSelected).RequieredNPC > 0 Then
2       .Quest.rNPC = tQuest(QuestSelected).RequieredNPC
3       .Quest.nNPC = tQuest(QuestSelected).RequieredNPCI
4   End If

5   If tQuest(QuestSelected).RequieredUser > 0 Then
6       .Quest.rUser = tQuest(QuestSelected).RequieredUser
7   End If

8   WriteConsoleMsg UserIndex, "Has aceptado la quest!", FontTypeNames.FONTTYPE_GUILD
9   WriteConsoleMsg UserIndex, "La recompensa de la quest es : " & tQuest(QuestSelected).Recompense & "!", FontTypeNames.FONTTYPE_GUILD
 
End With

errHandler:
    Debug.Print "Error en la línea nº " & Erl() & " del módulo 'Quest_uInicia'"
    
End Sub
 
Public Function Quest_Finaliza(ByVal UserIndex As Integer)

' / Author: maTih

With UserList(UserIndex)
 
    .Stats.GLD = .Stats.GLD + tQuest(.Stats.QuestIndex).RecompenseGold
        WriteUpdateGold UserIndex
    .Stats.ELU = .Stats.ELU + tQuest(.Stats.QuestIndex).RecompenseExp
        CheckUserLevel UserIndex
    .Stats.QuestIndex = 0
        WriteConsoleMsg UserIndex, "La quest ha finalizado!", FontTypeNames.FONTTYPE_GUILD

End With

End Function
 
Public Function ENpcNAME(ByVal Numero As Integer) As String

' / Author: maTih

    ENpcNAME = GetVar(App.Path & "\Dat\NPCs.dat", "NPC" & Numero, "Name")

End Function
