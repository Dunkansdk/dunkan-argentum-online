VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   6210
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   4890
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   4920
      TabIndex        =   45
      Top             =   240
      Width           =   495
   End
   Begin InetCtlsObjects.Inet FTP 
      Left            =   5160
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ComboBox cmd_MaxUser 
      Height          =   330
      Left            =   3000
      TabIndex        =   42
      Top             =   1200
      Width           =   1815
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "#"
      TabIndex        =   39
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame f_Map 
      Caption         =   "Mapa"
      Height          =   1335
      Left            =   2880
      TabIndex        =   36
      Top             =   1680
      Width           =   1935
      Begin VB.CommandButton cmd_Cambiar 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox c_Map 
         Height          =   330
         Left            =   120
         TabIndex        =   37
         Text            =   "Elige un mapa"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame f_Clases 
      Caption         =   "Clases permitidas"
      Height          =   1815
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   2535
      Begin VB.CommandButton cmd_Clases 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chk_Class 
         Caption         =   "CLASE"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_Cerrar 
      Caption         =   "CERRAR SERVER"
      Height          =   495
      Left            =   2640
      TabIndex        =   25
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Frame f_Bots 
      Caption         =   "Bots"
      Height          =   1935
      Left            =   2760
      TabIndex        =   20
      Top             =   3360
      Width           =   1935
      Begin VB.ComboBox cmd_Bando 
         Height          =   330
         Left            =   240
         TabIndex        =   41
         Text            =   "Criminal"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Invocar 
         Caption         =   "Invocar"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton op_Class 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton op_Class 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton op_Class 
         Caption         =   "Options"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame f_Reglas 
      Caption         =   "Reglas"
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   2175
      Begin VB.CheckBox chk_Death 
         Caption         =   "DeathMatch"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chk_Respawn 
         Caption         =   "Auto Resu"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chk_Resu 
         Caption         =   "Vale resucitar"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chk_Invi 
         Caption         =   "Vale invisibilidad"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Text            =   "SV AO"
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "ABRIR SERVER"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wskData 
      Left            =   5160
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timerBarcos 
      Interval        =   40
      Left            =   8760
      Top             =   780
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   0
      Width           =   2055
   End
   Begin VB.Timer timerIA 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   8760
      Top             =   300
   End
   Begin VB.TextBox txtChat 
      Height          =   2775
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   3600
      Top             =   5880
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   5880
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   2640
      Top             =   5880
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   6000
      TabIndex        =   2
      Top             =   1440
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios jugando :"
      Height          =   255
      Left            =   360
      TabIndex        =   44
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad maxima de usuarios"
      Height          =   375
      Left            =   3000
      TabIndex        =   43
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña del servidor :"
      Height          =   255
      Left            =   2520
      TabIndex        =   40
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del servidor :"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Clase_Bot   As Integer
Public ESPERANDO_MAPS  As Boolean

Public ESCUCHADAS As Long

Private Downloading_Obj   As Boolean  '< Se está descargando obj.Dat?
Private Action_Proccessed As Byte     '< Accion actual del winsock.
Private Server_Name       As String   '< Nombre del servidor.
Private Server_IP         As String   '< IP Del servidor (host)
Private Server_MaxUsers   As Byte     '< Cantidad máxima de usuarios
Private Action_Close      As Boolean  '< Cerrar el socket
Private Class_Per(1 To 7) As Boolean
Private Server_List_Index As Integer  '< Indice del server
Private Action_ChangeRulz As Byte     '< Regla que se cambia
Public Internet_Protocol As String

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
        With UserList(iUserIndex)
            'Conexion activa? y es un usuario loggeado?
            If .ConnID <> -1 And .flags.UserLogged Then
                'Actualiza el contador de inactividad
                If .flags.Traveling = 0 Then
                    .Counters.IdleCount = .Counters.IdleCount + 1
                End If
                
                If .Counters.IdleCount >= IdleLimit Then
                    Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")
                    'mato los comercios seguros
                    Call Cerrar_Usuario(iUserIndex)
                End If
            End If
        End With
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
Resume Next

End Sub



Private Sub chk_Class_Click(Index As Integer)

Server_Info.Clase(Index + 1) = Not Server_Info.Clase(Index + 1)

End Sub

Private Sub chk_Death_Click()

Server_Info.DeathMathc = (chk_Death.Value = 1)

Dim msj As String

msj = IIf(Server_Info.DeathMathc = True, "Servidor> El modo de juego ahora es DeathMatch.", "Servidor> El modo de juego ahora es Ciudas vs Crimis.")

SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(msj, FontTypeNames.FONTTYPE_CENTINELA)

End Sub

Private Sub chk_Invi_Click()

Server_Info.Invisibilidad = (chk_Invi.Value = 1)

Dim msj As String

msj = IIf(Server_Info.Resucitar = True, "Servidor> Invisibilidad ahora está activado.", "Servidor> Invisibilidad ahora está desactivado.")

SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(msj, FontTypeNames.FONTTYPE_CENTINELA)

Exit Sub

'Cambia el estado de la invisibilidad
If wskData.State = sckConnected Then
    wskData.SendData CStr(3) & "@" & CStr(Server_List_Index) & "@" & CStr(chk_Invi.Value)
Else
    If wskData.State <> sckClosed Then wskData.Close
    Action_ChangeRulz = 1
    wskData.Connect "localhost", 555
    '201.212.2.132
End If
End Sub

Private Sub chk_Respawn_Click()

Dim msj         As String
Dim resultBox   As Integer

Server_Info.AutoRespawn = (chk_Respawn.Value = 1)

'Actualiza el tiempo de respawn
If Server_Info.AutoRespawn Then
   resultBox = Val(InputBox$("Tiempo [en segundos] del auto-resucitar (mínimo : 1 - máximo : 60)"))
   'No ingreso un número o ingresó 0
   If (Not resultBox <> 0) Then
      MsgBox "Tiempo incorrecto, auto Resucitar quedará sin efecto."
      Server_Info.AutoRespawn = False
      Exit Sub
   ElseIf (resultBox < 1) Or (resultBox > 60) Then
   'Ingresó un número negativo o mayor a 60
      MsgBox "Tiempo incorrecto, auto Resucitar quedará sin efecto."
      Server_Info.AutoRespawn = False
      Exit Sub
   End If
   Server_Info.TiempoRespawn = (resultBox)
End If
    
msj = IIf(Server_Info.AutoRespawn = True, "Servidor> Auto-Resucitar ahora está activado.", "Servidor> Auto-Resucitar ahora está desactivado.")

SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(msj, FontTypeNames.FONTTYPE_CENTINELA)

End Sub

Private Sub chk_Resu_Click()

' @ Setea el resucitar activado/desactivado.

Dim msj As String

Server_Info.Resucitar = (chk_Resu.Value = 1)

msj = IIf(Server_Info.Resucitar = True, "Servidor> Resucitar ahora está activado.", "Servidor> Resucitar ahora está desactivado.")

SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(msj, FontTypeNames.FONTTYPE_CENTINELA)

End Sub

Private Sub cmd_Cambiar_Click()

'
' @ Cambia el mapa

Dim i   As Long

'Mapa válido?
If (c_Map.ListIndex <> -1) Then
   If (c_Map.list(c_Map.ListIndex) <> vbNullString) Then
      'Setea el nuevo map.
      Server_Info.Mapa.map = (c_Map.ListIndex + 1)
      Server_Info.Mapa.X = 50
      Server_Info.Mapa.Y = 50
      For i = 1 To LastUser
          'Si está logeado.
          If UserList(i).ConnID <> -1 Then
             'Si está muerto lo revive
             If UserList(i).flags.Muerto <> 0 Then
                Call RevivirUsuario(i)
             End If
             'Llena la vida
             UserList(i).Stats.MinHp = UserList(i).Stats.MaxHp
             'Actualiza
             WriteUpdateHP CInt(i)
             'Telep
             Call WarpUserChar(CInt(i), (c_Map.ListIndex + 1), RandomNumber(50, 55), RandomNumber(50, (50 + CInt(i))), True)
          End If
      Next i
      
      'Avisa
      Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Cambió al mapa " & MapInfo(c_Map.ListIndex + 1).Name & ".", FontTypeNames.FONTTYPE_CENTINELA))
      
  End If
Else
    'limpia la lista
    c_Map.Clear
    'actualiza
    For i = 1 To NumMaps
        c_Map.AddItem MapInfo(i).Name
    Next i
End If
End Sub

Private Sub cmd_Cerrar_Click()

' Va form por form

    Dim loopF   As Form
    
    For Each loopF In Forms
        Unload loopF
    Next loopF

End Sub

Private Sub cmd_Clases_Click()

Dim i    As Long
Dim iTxt As String

iTxt = "Servidor> Actualizacion en las clases permitidas"

For i = 0 To 7

    If (chk_Class(i).Value = 1) Then
       iTxt = iTxt & vbNewLine & chk_Class(i).Caption & ": Activada"
    Else
       iTxt = iTxt & vbNewLine & chk_Class(i).Caption & ": Prohibida"
    End If
    
Next i

SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(iTxt, FontTypeNames.FONTTYPE_CENTINELA)

End Sub

Private Sub cmd_Invocar_Click()

'Invoca un bot.

If Clase_Bot <> 0 Then
    Call mod_IA.ia_Spawn(Clase_Bot, Ullathorpe, "Bot <" & op_Class(Clase_Bot - 1).Caption & ">", False, cmd_Bando.ListIndex = 0)
End If

End Sub

Private Sub cmdAbrir_Click()

If (cmd_MaxUser.ListIndex = -1) Then
   MsgBox "Cantidad máxima de usuarios inválida."
   Exit Sub
End If

If (c_Map.ListIndex = -1) Then
   MsgBox "Elije un mapa!"
   Exit Sub
End If

'Call mod_DunkanGeneral.Load_MapData(c_Map.ListIndex + 1)


Server_MaxUsers = cmd_MaxUser.ListIndex + 2

Server_IP = "localhost"

Action_Proccessed = 1

Server_Name = txtName.Text
If wskData.State = 8 Then wskData.Close

If wskData.State <> sckConnected Then
    wskData.Connect "localhost", 555
Else
    wskData.SendData "2@" & Server_Name & "@" & Server_IP & "@" & Server_MaxUsers
End If

End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub cmdSa_Click()
Form1.Show , frmMain
End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub



Private Sub Command3_Click()

Dim i   As Long
Dim S   As Byte

For i = 1 To UBound(ObjData())
    If bla(i) = False Then
       With ObjData(i)
            S = S + 1
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & vbNewLine
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").MaxHit = " & .MaxHIT
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").MinHit = " & .Minhit
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").MinDef = " & .MinDef
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").MaxDef = " & .MaxDef
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").ObjType = " & .OBJType
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").DefensaMagicaMin = " & .DefensaMagicaMin
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").DefensaMagicaMax = " & .DefensaMagicaMax
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").DuracionEfecto = " & .DuracionEfecto
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").TipoPocion = " & .TipoPocion
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").WeaponAnim = " & .WeaponAnim
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").ShieldAnim = " & .ShieldAnim
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").CascoAnim = " & .CascoAnim
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").WeaponRazaEnanaAnim = " & .WeaponRazaEnanaAnim
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").Ropaje = " & .Ropaje
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").Apuñala = " & .Apuñala
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").GrhIndex = " & .GrhIndex
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").GrhIndexSecundario = " & .GrhSecundario
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").MinModificador = " & .MinModificador
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").MaxModificador = " & .MaxModificador
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").Municion = " & .Municion
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").Name = " & .Name
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").StaffDamageBonus = " & .StaffDamageBonus
            Form1.Text1.Text = Form1.Text1.Text & vbNewLine & "Objdata(" & CStr(S) & ").StaffPower = " & .StaffPower
       End With
    End If
Next i

Form1.Show

End Sub

Function bla(ByVal S As Integer) As Boolean

Dim i As Long
Dim j As Long

For i = 1 To UBound(Buy())
    For j = 1 To UBound(Buy(i).Buys())
        If Buy(i).Buys(j).Objeto.objIndex = S Then bla = False: Exit Function
    Next j
Next i

bla = True

End Function

Private Sub Command4_Click()
On Error Resume Next


End Sub

Private Sub Form_Load()

'Carga los nombres de las clases de los bots.

op_Class(0).Caption = "Clerigo"
op_Class(1).Caption = "Mago"
op_Class(2).Caption = "Cazador"

cmd_Bando.AddItem "Criminal", 0
cmd_Bando.AddItem "Ciudadano", 1

Dim i    As Long

For i = 2 To 20
    cmd_MaxUser.AddItem "Usuarios:" & CStr(i)
Next i

'Clases
For i = 1 To 8
    chk_Class(i - 1).Caption = ListaClases(i)
    chk_Class(i - 1).Value = 1
    Server_Info.Clase(i) = True
Next i

'Setea el nuevo map.
Server_Info.Mapa.map = 1
Server_Info.Mapa.X = 50
Server_Info.Mapa.Y = 50

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                'PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim loopC As Integer

For loopC = 1 To MaxUsers
    If UserList(loopC).ConnID <> -1 Then Call CloseSocket(loopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub



Private Sub FTP_StateChanged(ByVal State As Integer)

If State > 10 Then Downloading_Obj = False

End Sub

Private Sub GameTimer_Timer()
'********************************************************
'Author: Unknown
'Last Modify Date: -
'********************************************************
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror
    
    Dim jj As Long
    Dim kk As Long
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To MaxUsers 'LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '¿User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    'If there is anything to be sent, we send it
                    Call FlushBuffer(iUserIndex)
                    
                    Call DoTileEvents(iUserIndex, .Pos.map, .Pos.X, .Pos.Y)

                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                    
                    If .flags.Muerto = 0 Then
                        
                       
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If

                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If 'UserLogged
                

            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuCerrar_Click()


If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
End If

End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Public Function DescargarObjDat() As Boolean

'
' @ Descarga el obj.dat desde el servidor FTP.

Dim txt_Usuario As String
Dim txt_Pass    As String
Dim txt_Host    As String
Dim txt_Remoto  As String
Dim txt_Local   As String

txt_Remoto = "/Dat/obj.dat"
txt_Local = "C:\Obj.dat"
txt_Host = "ftp://dunkanao.zxq.net"
txt_Usuario = "dunkanao_zxq"
txt_Pass = "asdasd22"

    With Inet1
        
        .URL = txt_Host
        .UserName = txt_Usuario
        .Password = txt_Pass
        .RemotePort = 21
        Call .Execute(, "Get " & txt_Remoto & " " & txt_Local)
        DoEvents
         
         DescargarObjDat = dir$(txt_Local) <> vbNullString
         
     End With

End Function


Private Sub op_Class_Click(Index As Integer)
Clase_Bot = Index + 1
End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo Errhandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

Errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.Description)
    Resume Next
End Sub




Private Sub t_Update_Maps_Timer()

End Sub


Private Sub timerBarcos_Timer()

Dim loopI   As Long

#If Barcos <> 0 Then
For loopI = 1 To NUMEMBARCACIONES
    If Embarcaciones(loopI).Zarpo Then
        Call mod_Embarcaciones.Mover(CByte(loopI))
    End If
Next loopI

#End If

End Sub

Private Sub timerIA_Timer()
#If ConBots Then

'Acción de los bots * maTih.-
Dim loopX   As Long

For loopX = 1 To MAX_BOTS

    If ia_Bot(loopX).Invocado Then ia_Action loopX

Next loopX

#End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
    
    Debug.Print Datos
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(MiDato)
    Else
        Exit Sub
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub

Private Sub tLluvia_Timer()

If c_Map.ListIndex = -1 Then
    Dim i   As Long
    
    For i = 1 To NumMaps
        c_Map.AddItem MapInfo(i).Name
    Next i
    tLluvia.Enabled = False
End If

End Sub




#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub wskData_Connect()

' @ Avisa que cambia datos.

If Action_ChangeRulz <> 0 Then
   If Action_ChangeRulz = 1 Then    'Toggle INVI.
      wskData.SendData "3@" & Server_List_Index & "@" & CStr(1)
      Action_Proccessed = 100
   End If
End If

If Action_Proccessed <> 0 Then
   If Action_Proccessed = 1 Then
      wskData.SendData "2@" & Server_Name & "@" & Server_IP & "@" & Server_MaxUsers
      Action_Proccessed = 100
    End If
    
    If Action_Proccessed = 2 Then
      wskData.SendData "4@" & Server_Name & "@" & CStr(NumUsers)
      Action_Proccessed = 100
    End If
    
End If

End Sub

Public Sub EnviarNumUsers()

'
' @ Envia la cantidad de usuarios online.

If wskData.State <> sckConnected Then
'   wskData.Connect "localhost", 555
   Action_Proccessed = 2
Else
   wskData.SendData "4@" & Server_Name & "@" & CStr(NumUsers)
End If

End Sub

Private Sub wskData_DataArrival(ByVal bytesTotal As Long)

' @ Handle the incoming Data

Dim incoming_Data   As String

wskData.GetData incoming_Data, vbString

Select Case Left$(incoming_Data, 2)
       Case "12"
            Server_List_Index = Val(ReadElement(incoming_Data, 1))
End Select

End Sub

Public Function ReadElement(ByVal recData As String, ByVal sComparePos As Byte) As String

' @ maTih.-   /    Devuelve un elemento de una cadena.

Dim arrData()   As String   '< Array para la collection de elementos.

'Crea el array.
arrData = Split(recData, "@")

'Compara elementos.
If ((sComparePos - 1) <= UBound(arrData())) Then
    'Devuelve.
    ReadElement = arrData(sComparePos - 1)
End If

End Function


Private Sub wskData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskData.Close
End Sub

Private Sub wskData_SendComplete()

    'Hay que cerrar el winsock?
    If Action_Close Then
       wskData.Close
    'Feisimo ! XD
    ElseIf Action_Proccessed = 100 Then
       wskData.SendData "@CLOSE"
       Action_Proccessed = 0
       Action_Close = True
    End If

End Sub

