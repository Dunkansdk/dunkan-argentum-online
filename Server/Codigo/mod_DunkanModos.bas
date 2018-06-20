Attribute VB_Name = "mod_DunkanModos"
'
' @ Creación y logeo de personajes [INVITADOS]

Option Explicit


Function raza_By_Clase(ByVal classIndex As Byte) As eRaza

'
' @ Devuelve una raza para una clase.

Select Case classIndex
       
       Case eClass.Assasin                               '< Asesino : DROW.
            raza_By_Clase = eRaza.Drow
       
       Case eClass.Bard, eClass.Druid                    '< Bardo/Druida : elfo
            raza_By_Clase = eRaza.Elfo
            
       Case eClass.Cleric                                '< Clerigo : DROW.
            raza_By_Clase = eRaza.Drow
        
       Case eClass.Paladin, eClass.Hunter, eClass.Mage   '< Paladín/Cazador/Mago : humano.
            raza_By_Clase = eRaza.Humano
       
       Case eClass.Warrior                               '< Guerrero : enano
            raza_By_Clase = eRaza.Enano
       
End Select

End Function

Function Body_by_Raza(ByVal Raza As String) As Integer

'
' @ Cuerpo para la raza.

Dim end_Body    As Integer

  Select Case Raza
         Case eRaza.Humano
            end_Body = 21
        Case eRaza.Drow
            end_Body = 32
        Case eRaza.Elfo
            end_Body = 210
        Case eRaza.Gnomo
            end_Body = 222
        Case eRaza.Enano
            end_Body = 53
  End Select
  
  Body_by_Raza = end_Body

End Function

Function Hit_By_Clase(ByVal Class As eClass) As Integer

'
' @ Golpe por personaje.

            Select Case Class
                Case eClass.Warrior
                    Hit_By_Clase = RandomNumber(2, 3)
                
                Case eClass.Hunter
                    Hit_By_Clase = RandomNumber(2, 3)
                
                Case eClass.Paladin
                    Hit_By_Clase = RandomNumber(2, 3)
                
                Case eClass.Cleric
                    Hit_By_Clase = 2
                
                Case eClass.Druid
                    Hit_By_Clase = 2
                
                Case eClass.Assasin
                    Hit_By_Clase = RandomNumber(2, 3)
                
                Case eClass.Bard
                    Hit_By_Clase = 2
            End Select
            
            Hit_By_Clase = (Hit_By_Clase * 40)

End Function

Function Head_By_Raza(ByVal Raza As String) As Integer
 
'
' @ Cabeza random para el personaje.
 
Dim TmpHead As Integer
 
  Select Case Raza
 
   Case "HUMANO"
   
   TmpHead = RandomNumber(1, 40)
   
   Case "ELFO"
   
   TmpHead = RandomNumber(102, 112)
   
   Case "DROW"
   
   TmpHead = RandomNumber(200, 210)
   
   Case "GNOMO"
   
   TmpHead = RandomNumber(402, 407)
   
   Case "ENANO"
   
   TmpHead = RandomNumber(303, 307)
 
  End Select
 
Head_By_Raza = TmpHead
 
End Function

Function GetMAN(ByVal Raza As String, ByVal Clase As String) As Integer
 
'
' @ Mana de los pjs-
 
Dim tmpMana As Integer
 
Select Case Clase
 
  Case "DRUIDA", "BARDO", "CLERIGO"
 
  Select Case Raza
 
         Case "HUMANO"
        
         tmpMana = 1460
        
         Case "DROW", "ELFO"
        
         tmpMana = 1760
        
         Case "GNOMO"
        
         tmpMana = 1860
        
         Case "ENANO"
        
         tmpMana = 1350
   
  End Select
 
   Case "MAGO"
   
  Select Case Raza
 
         Case "HUMANO"
        
         tmpMana = 2000
        
         Case "DROW", "ELFO"
        
         tmpMana = 2250
        
         Case "GNOMO"
        
         tmpMana = 2600
        
         Case "ENANO"
        
         tmpMana = 1894
   
  End Select
 
   Case "PALADIN", "ASESINO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      tmpMana = 711
     
      Case "DROW", "ELFO"
     
      tmpMana = 862
     
      Case "GNOMO"
     
      tmpMana = 952
     
      Case "ENANO"
     
      tmpMana = 655
   
  End Select
 
   Case "CAZADOR", "GUERRERO"
   
   tmpMana = 0
   
End Select
 
GetMAN = tmpMana
 
End Function
 
Function GetHP(ByVal Raza As String, ByVal Clase As String) As Integer
 
' @ Vida de los pjs.
 
Dim tmpHP(tVar.tinteger) As Integer
 
Select Case Clase
 
  Case "DRUIDA", "BARDO", "CLERIGO"
 
  Select Case Raza
 
      Case "HUMANO"
     
      Call Memory.SetInteger(tmpHP, 385 Xor 128)
     
      Case "DROW", "ELFO"
     
      Call Memory.SetInteger(tmpHP, 362 Xor 128)
     
      Case "GNOMO"
     
      Call Memory.SetInteger(tmpHP, 355 Xor 128)
     
      Case "ENANO"
     
      Call Memory.SetInteger(tmpHP, 400 Xor 128)
   
  End Select
 
   Case "MAGO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      Call Memory.SetInteger(tmpHP, 315 Xor 128)
     
      Case "DROW", "ELFO"
     
      Call Memory.SetInteger(tmpHP, 290 Xor 128)
     
      Case "GNOMO"
     
      Call Memory.SetInteger(tmpHP, 275 Xor 128)
     
      Case "ENANO"
     
      Call Memory.SetInteger(tmpHP, 321 Xor 128)
   
  End Select
 
   Case "PALADIN", "ASESINO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      Call Memory.SetInteger(tmpHP, 372 Xor 128)
     
      Case "DROW", "ELFO"
     
      Call Memory.SetInteger(tmpHP, 325 Xor 128)
     
      Case "GNOMO"
     
      Call Memory.SetInteger(tmpHP, 350 Xor 128)
     
      Case "ENANO"
     
      Call Memory.SetInteger(tmpHP, 400 Xor 128)
   
  End Select
 
   Case "CAZADOR", "GUERRERO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      Call Memory.SetInteger(tmpHP, 415 Xor 128)
     
      Case "DROW", "ELFO"
     
      Call Memory.SetInteger(tmpHP, 400 Xor 128)
     
      Case "GNOMO"
     
      Call Memory.SetInteger(tmpHP, 386 Xor 128)
     
      Case "ENANO"
     
      Call Memory.SetInteger(tmpHP, 429 Xor 128)
   
  End Select
   
End Select
 
GetHP = Memory.GetInteger(tmpHP)
 
End Function

Public Function Checkear_Memoria(ByVal UserIndex As Integer) As Boolean

'
' @ Checkea que no alla edición en la memoria de lavida del usuario.

With UserList(UserIndex).Stats
     
     Checkear_Memoria = (.MaxHp = (.Original_MaxHP Xor 128))
     
     'Edición de memoria !
     If (Checkear_Memoria = True) Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Se detectó una modificación en la memoria de : " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_CITIZEN))
        Call CloseSocket(UserIndex)
     End If
End With

End Function

Public Sub Crear_Personaje(ByVal UserIndex As Integer, ByVal claseIndex As Byte, ByRef UserName As String)

'
' @ Crea el personaje en la memoria.

Dim loopC   As Long

With UserList(UserIndex)
     
     'Setea clase & raza.
     .Clase = claseIndex
     .Raza = raza_By_Clase(claseIndex)

    'Siempre es hombre por que me daba mucha paja , sorry.
    .Genero = eGenero.Hombre
     .Name = UserName
     Call Cuerpo_Muerto(UserIndex)
     
     'Solo tres intentos para buscar nombres.
     Dim tmp_Name   As String
     
     tmp_Name = UserName
     
     If CheckForSameName(tmp_Name) Then
        For loopC = 1 To 3
            'Si no encuentra usuario con este nombre cierra el bucle
            tmp_Name = .Name & "(" & CStr(loopC) & ")"
            If Not CheckForSameName(tmp_Name) Then Exit For
        Next loopC
     
     'No encontró nick.
     If loopC > 3 Then
        WriteErrorMsg UserIndex, "Ya hay otro usuario con tu nombre."
        Exit Sub
     Else
        .Name = tmp_Name
     End If
     
    End If
     'Busca mana, maná, vida y energía.
           
     With .Stats

          .MaxHp = GetHP(UCase$(ListaRazas(UserList(UserIndex).Raza)), UCase$(ListaClases(UserList(UserIndex).Clase)))
          .Original_MaxHP = .MaxHp
          .MaxHp = (.MaxHp Xor 128)
          .MinHp = 0
          .MaxMAN = GetMAN(UCase$(ListaRazas(UserList(UserIndex).Raza)), UCase$(ListaClases(UserList(UserIndex).Clase)))
          .MinMAN = .MaxMAN
          'TODO : Esto ai que rajarlo xd
          .MinAGU = 1000
          .MaxAGU = 1000
          .MinHam = 1000
          .MaxHam = 1000
          .MinSta = 1000
          'Dar hechizos.
          .UserHechizos(35) = 10        'Remover paralisis.
          .UserHechizos(34) = 24        'Inmovilizar.
          '.UserAtributos(2) = 25       'Apocalipsis.
          '.UserAtributos(32) = 23       'Descarga.
          'Obtiene hit.
          .MaxHIT = Hit_By_Clase(UserList(UserIndex).Clase)
          .Minhit = .MaxHIT
          .ELV = 40
          
          'Skills & atributos.
          For loopC = 1 To NUMSKILLS
              .UserSkills(loopC) = 100
          Next loopC
          
          For loopC = 1 To NUMATRIBUTOS
               .UserAtributos(loopC) = 18
          Next loopC
          
          'Obtiene la fuerza y agilidad mínima.
         .UserAtributos(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza) + ModRaza(UserList(UserIndex).Raza).Fuerza
         .UserAtributos(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad) + ModRaza(UserList(UserIndex).Raza).Agilidad
          
          .UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, (.UserAtributos(eAtributos.Fuerza) * 2))
          .UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, (.UserAtributos(eAtributos.Agilidad) * 2))
     End With
     
     .flags.Muerto = 1
    
     'Update hechizos & inventario.
     Call InvUsuario.UpdateUserInv(True, UserIndex, 0)
     Call modHechizos.UpdateUserHechizos(True, UserIndex, 0)
     
     With .OrigChar
          .Head = Head_By_Raza(UCase$(ListaRazas(UserList(UserIndex).Raza)))
          .body = Body_by_Raza(UserList(UserIndex).Raza)
     End With
     
     WriteUpdateHP UserIndex
     WriteUpdateMana UserIndex
     
     'Update pos
     .Pos = Server_Info.Mapa

     .Pos.X = 50
     .Pos.Y = 50
     
     Dim NullObj    As UserOBJ
     
     For loopC = 1 To UBound(.Invent.Object())
         .Invent.Object(loopC) = NullObj
     Next loopC
     
     .CurrentInventorySlots = 20
     'Connect : p
     Call Connect(UserIndex)
     
End With

End Sub

Sub Cuerpo_Muerto(ByVal UserIndex As Integer)

'
' @ Setea el char muerto.

With UserList(UserIndex)

     'Setea el char
     With .Char
          .body = iCuerpoMuerto
          .Head = iCabezaMuerto
          .CascoAnim = NingunCasco
          .WeaponAnim = NingunArma
          .ShieldAnim = NingunEscudo
          .heading = eHeading.SOUTH
     End With

     .flags.Desnudo = 1

End With

End Sub

Sub Connect(ByVal UserIndex As Integer)

'
' @ Conecta usuario invitado.

 
With UserList(UserIndex)
   
    .showName = True 'Por default los nombres son visibles
    
    .Pos.X = 50
    .Pos.Y = 50
    
    'Info
    Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index

    Call WriteChangeMap(UserIndex, .Pos.map, MapInfo(.Pos.map).MapVersion) 'Carga el mapa
   
    .flags.ChatColor = vbWhite
       
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
    
   
    'Crea  el personaje deel usuario
    Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
   
    MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = UserIndex
   
    Call WriteUserCharIndexInServer(UserIndex)
    
    Call WritePosUpdate(UserIndex)
    
    ''[/el oso]
   
    'Call WriteUpdateUserStats(userIndex)
   Call WriteUpdateMana(UserIndex)
   Call WriteUpdateHP(UserIndex)
   
    NumUsers = NumUsers + 1
    .flags.UserLogged = True
   
    'MapInfo(.Pos.map).NumUsers = MapInfo(.Pos.map).NumUsers + 1

    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    

    Call WriteLoggedMessage(UserIndex)

End With

End Sub
