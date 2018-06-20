Attribute VB_Name = "mod_DunkanInvitados"
 
Option Explicit


Sub Connect(ByVal UserIndex As Integer, ByVal NickName As String, ByVal classIndex As Byte)

'
' @ Conecta usuario invitado.

 
With UserList(UserIndex)

    'Siempre es hombre por que me daba mucha paja , sorry.
    .Genero = eGenero.Hombre
   
    Call DarCuerpoDesnudo(UserIndex)
   
    .Stats.MaxMAN = DarManaByClaseAndRaza(UCase$(ListaRazas(.Raza)), UCase$(ListaClases(.Clase)))
    .Stats.MinMAN = .Stats.MaxMAN
    .Stats.MaxHp = DarHPByClaseAndRaza(UCase$(ListaRazas(.Raza)), UCase$(ListaClases(.Clase)))
    .Stats.MinHp = .Stats.MaxHp
   
    .Stats.MinAGU = 100
    .Stats.MaxAGU = 100
    .Stats.MinHam = 100
    .Stats.MaxHam = 100
    .Stats.MaxSta = 800
    .Stats.MinSta = 770
   
    LlenarPosition UserIndex
   
    LlenarFlags UserIndex
   
    LlenarUserChar UserIndex
   
    LlenarHechizos UserIndex
   
    LlenarSkillsYAtributos UserIndex
   
    LlenarInventario UserIndex
   
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)
   
    'Nombre de sistema
    .Name = NickName
   
    .showName = True 'Por default los nombres son visibles
 
    'Info
    Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
    Call WriteChangeMap(UserIndex, .Pos.map, MapInfo(.Pos.map).MapVersion) 'Carga el mapa
    Call WritePlayMidi(UserIndex, Val(ReadField(1, MapInfo(.Pos.map).Music, 45)))
   
    .flags.ChatColor = vbWhite
       
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
   
    With .Pos
         .map = Server_Info.Mapa.map
         .X = 50
         .Y = 50
    End With
   
    'Crea  el personaje del usuario
    Call MakeUserChar(False, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
   
    Call WriteUserCharIndexInServer(UserIndex)
    ''[/el oso]
   
    Call WriteUpdateHP(UserIndex)
    Call WriteUpdateMana(UserIndex)
   
    NumUsers = NumUsers + 1
    .flags.UserLogged = True

    MapInfo(.Pos.map).NumUsers = MapInfo(.Pos.map).NumUsers + 1
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))

    Call WriteLoggedMessage(UserIndex)
    
    Call FlushBuffer(UserIndex)
    
End With

End Sub
 
Function DarRazaByString(ByVal Raza As String) As eRaza
 
'
' @ Raza by nombre
 
Dim tmpN As eRaza
 
Select Case Raza
 
  Case "HUMANO"
  tmpN = eRaza.Humano
 
  Case "ELFO"
  tmpN = eRaza.Elfo
 
  Case "ELFO OSCURO"
  tmpN = eRaza.Drow
 
  Case "GNOMO"
  tmpN = eRaza.Gnomo
 
  Case "ENANO"
  tmpN = eRaza.Enano
 
End Select
 
DarRazaByString = tmpN
 
End Function
 
Function DarClaseByString(ByVal ClaseS As String) As Byte
 
'
' @ Clase by Nombre.
 
Dim claseN As Byte
 
Select Case ClaseS
 
  Case "CLERIGO"
 
  claseN = eClass.Cleric
 
  Case "MAGO"
 
  claseN = eClass.Mage
 
  Case "BARDO"
 
  claseN = eClass.Bard
 
  Case "PALADIN"
 
  claseN = eClass.Paladin
 
  Case "ASESINO"
 
  claseN = eClass.Assasin
 
  Case "GUERRERO"
 
  claseN = eClass.Warrior
 
  Case "CAZADOR"
 
  claseN = eClass.Hunter
 
  Case "DRUIDA"
 
  claseN = eClass.Druid
 
End Select
 
DarClaseByString = claseN
 
End Function
 
Function DarManaByClaseAndRaza(ByVal Raza As String, ByVal Clase As String) As Integer
 
'
' @ Mana de los pjs-
 
Dim tmpMana As Integer
 
Select Case Clase
 
  Case "DRUIDA", "BARDO", "CLERIGO"
 
  Select Case Raza
 
      Case "HUMANO"
     
      tmpMana = 1460
     
      Case "ELFO OSCURO", "ELFO"
     
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
     
      Case "ELFO OSCURO", "ELFO"
     
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
     
      Case "ELFO OSCURO", "ELFO"
     
      tmpMana = 862
     
      Case "GNOMO"
     
      tmpMana = 952
     
      Case "ENANO"
     
      tmpMana = 655
   
  End Select
 
   Case "CAZADOR", "GUERRERO"
   
   tmpMana = 0
   
End Select
 
DarManaByClaseAndRaza = tmpMana
 
End Function
 
Function DarHPByClaseAndRaza(ByVal Raza As String, ByVal Clase As String) As Integer
 
' @ Vida de los pjs.
 
Dim tmpHP As Integer
 
Select Case Clase
 
  Case "DRUIDA", "BARDO", "CLERIGO"
 
  Select Case Raza
 
      Case "HUMANO"
     
      tmpHP = 385
     
      Case "ELFO OSCURO", "ELFO"
     
      tmpHP = 362
     
      Case "GNOMO"
     
      tmpHP = 355
     
      Case "ENANO"
     
      tmpHP = 400
   
  End Select
 
   Case "MAGO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      tmpHP = 315
     
      Case "ELFO OSCURO", "ELFO"
     
      tmpHP = 290
     
      Case "GNOMO"
     
      tmpHP = 275
     
      Case "ENANO"
     
      tmpHP = 321
   
  End Select
 
   Case "PALADIN", "ASESINO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      tmpHP = 372
     
      Case "ELFO OSCURO", "ELFO"
     
      tmpHP = 325
     
      Case "GNOMO"
     
      tmpHP = 350
     
      Case "ENANO"
     
      tmpHP = 400
   
  End Select
 
   Case "CAZADOR", "GUERRERO"
   
  Select Case Raza
 
      Case "HUMANO"
     
      tmpHP = 415
     
      Case "ELFO OSCURO", "ELFO"
     
      tmpHP = 400
     
      Case "GNOMO"
     
      tmpHP = 386
     
      Case "ENANO"
     
      tmpHP = 429
   
  End Select
   
End Select
 
DarHPByClaseAndRaza = tmpHP
 
End Function
 
Function DarHeadByRaza(ByVal Raza As String) As Integer
 
'
' @ Cabeza random para el personaje.
 
Dim TmpHead As Integer
 
  Select Case Raza
 
   Case "HUMANO"
   
   TmpHead = RandomNumber(1, 40)
   
   Case "ELFO"
   
   TmpHead = RandomNumber(102, 112)
   
   Case "ELFO OSCURO"
   
   TmpHead = RandomNumber(200, 210)
   
   Case "GNOMO"
   
   TmpHead = RandomNumber(402, 407)
   
   Case "ENANO"
   
   TmpHead = RandomNumber(303, 307)
 
  End Select
 
DarHeadByRaza = TmpHead
 
End Function
 
Sub LlenarInventario(ByVal UserIndex As Integer)
 
'
' @ Darle items !
 
With UserList(UserIndex)
 
.CurrentInventorySlots = 1
.Invent.Object(1).objIndex = 986
.Invent.Object(1).Amount = 100
 
End With
 
End Sub
 
Sub LlenarSkillsYAtributos(ByVal UserIndex As Integer)
 
'
' @ Agrega skills & Atributos.
 
Dim i As Long
 
For i = 1 To NUMSKILLS
 UserList(UserIndex).Stats.UserSkills(i) = 100
Next i
 
For i = 1 To NUMATRIBUTOS
 UserList(UserIndex).Stats.UserAtributos(i) = 18
Next i
 
End Sub
 
Sub LlenarHechizos(ByVal UserIndex As Integer)
 
'
' @ Llena hechizos - MODIFICAR!
 
With UserList(UserIndex).Stats
 
.UserHechizos(35) = 10
.UserHechizos(34) = 24
.UserHechizos(33) = 25
.UserHechizos(32) = 23
 
UpdateUserHechizos True, UserIndex, 0
 
End With
 
End Sub
 
Sub LlenarUserChar(ByVal UserIndex As Integer)
 
'
' @ Llena el char.

With UserList(UserIndex).Char
    .CascoAnim = NingunCasco
    .ShieldAnim = NingunEscudo
    .WeaponAnim = NingunArma
    .heading = eHeading.SOUTH
    .FX = 0
End With
 
End Sub
 
Sub LlenarFlags(ByVal UserIndex As Integer)
 
'
' @ Resetea flags generales.
 
With UserList(UserIndex).flags
 
.Escondido = 0
.targetNPC = 0
.TargetNpcTipo = eNPCType.Comun
.TargetOBJ = 0
.targetUser = 0
.Privilegios = PlayerType.User
 
End With

End Sub
 
Sub LlenarPosition(ByVal User As Integer)
 
'
' @ Llena la posición con los datos del servidor.
 
With UserList(User)
.Pos = Server_Info.Mapa
End With
 
End Sub
