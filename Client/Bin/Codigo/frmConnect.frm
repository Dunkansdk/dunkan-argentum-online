VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   597
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   797
      TabIndex        =   4
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton Command1 
         Caption         =   "Listar"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.ListBox lst_Svs 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Frame Frame3 
         Caption         =   "Info del servidor"
         Height          =   2415
         Left            =   6480
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
         Begin VB.Label lblIP 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblNumUsers 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblName 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   7320
         Width           =   3015
         Begin VB.TextBox txtPasswd 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   600
            Width           =   2460
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   225
            TabIndex        =   6
            Top             =   240
            Width           =   2460
         End
         Begin VB.Image imgConectarse 
            Height          =   375
            Left            =   120
            Top             =   990
            Width           =   1335
         End
      End
      Begin MSWinsockLib.Winsock wskData 
         Left            =   0
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Image imgTeclas 
         Height          =   375
         Left            =   6840
         Top             =   9000
         Width           =   1335
      End
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4890
      TabIndex        =   0
      Text            =   "7666"
      Top             =   2760
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5760
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox WebAuxiliar 
      Height          =   360
      Left            =   960
      ScaleHeight     =   300
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgVerForo 
      Height          =   465
      Left            =   450
      Top             =   6120
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   9960
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgBorrarPj 
      Height          =   375
      Left            =   8400
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgCodigoFuente 
      Height          =   375
      Left            =   6840
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgReglamento 
      Height          =   375
      Left            =   5280
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgManual 
      Height          =   375
      Left            =   3720
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgRecuperar 
      Height          =   375
      Left            =   2160
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgCrearCuenta 
      Height          =   375
      Left            =   600
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit

Private pendientAction  As Byte
Private Server_Add_Name As String       '< Nombre del servidor a agregar.
Private Server_Add_IP   As String       '< Ip del servidor a agregar.
Public Now_Server_Ip    As String
Private Connecting_ToSV As Boolean      '< Conectandose a unserver?
Private cBotonCrearPj As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton
Private cBotonTeclas As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Command1_Click()

On Error Resume Next

With Winsock1
     'Si no está conectado.
     If .State = 8 Then .Close
     
     If .State <> sckConnected Then
        .Connect "localhost", 555
        pendientAction = 1
        Exit Sub
     End If
     
     .SendData "0"
     
End With

End Sub

Private Sub Form_Activate()

Engine_LoadMap_Connect

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    'Port
    PortTxt.Visible = True
    'Label4.Visible = True
    
    'Server IP
    PortTxt.Text = "7666"
    IPTxt.Text = "192.168.0.2"
    IPTxt.Visible = True
    'Label5.Visible = True
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
    
    
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If

'EstadoLogin = reon

'frmMain.Winsock1.Connect CurServerIp, CurServerPort
    
    PortTxt.Text = Config_Inicio.Puerto
 
     '[CODE]:MatuX
    '
    '  El código para mostrar la versión se genera acá para
    ' evitar que por X razones luego desaparezca, como suele
    ' pasar a veces :)
       version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    '[END]'
    
    Me.Picture = LoadPicture(App.path & "\..\Resources\Graphics\VentanaConectar.jpg")
    
    Call LoadButtons

    Call CheckLicenseAgreement
        
End Sub

Private Sub CheckLicenseAgreement()
    'Recordatorio para cumplir la licencia, por si borrás el Boton sin leer el code...
    Dim i As Long
    
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls(i).Name = "imgCodigoFuente" Then
            Exit For
        End If
    Next i
    
    If i = Me.Controls.Count Then
        MsgBox "No debe eliminarse la posibilidad de bajar el código de sus servidor. Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor, incurriendo de esta forma en un delito punible por ley." & vbCrLf & vbCrLf & vbCrLf & _
                "Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo, no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.", vbCritical Or vbApplicationModal
    End If

End Sub

Private Sub LoadButtons()
    
    txtPasswd.Text = "lkasdoi"
    txtNombre.Text = "dunkan"
    
    Dim GrhPath As String
    
    GrhPath = DirGraficos
    
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonReglamento = New clsGraphicalButton
    Set cBotonCodigoFuente = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cBotonForo = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    Set cBotonTeclas = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

                                    
    Call cBotonRecuperarPass.Initialize(imgRecuperar, GrhPath & "BotonRecuperarPass.jpg", _
                                    GrhPath & "BotonRecuperarPassRollover.jpg", _
                                    GrhPath & "BotonRecuperarPassClick.jpg", Me)
                                    
    Call cBotonManual.Initialize(imgManual, GrhPath & "BotonManual.jpg", _
                                    GrhPath & "BotonManualRollover.jpg", _
                                    GrhPath & "BotonManualClick.jpg", Me)
                                    
    Call cBotonReglamento.Initialize(imgReglamento, GrhPath & "BotonReglamento.jpg", _
                                    GrhPath & "BotonReglamentoRollover.jpg", _
                                    GrhPath & "BotonReglamentoClick.jpg", Me)
                                    
    Call cBotonCodigoFuente.Initialize(imgCodigoFuente, GrhPath & "BotonCodigoFuente.jpg", _
                                    GrhPath & "BotonCodigoFuenteRollover.jpg", _
                                    GrhPath & "BotonCodigoFuenteClick.jpg", Me)
                                    
    Call cBotonBorrarPj.Initialize(imgBorrarPj, GrhPath & "BotonBorrarPersonaje.jpg", _
                                    GrhPath & "BotonBorrarPersonajeRollover.jpg", _
                                    GrhPath & "BotonBorrarPersonajeClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirConnect.jpg", _
                                    GrhPath & "BotonBotonSalirRolloverConnect.jpg", _
                                    GrhPath & "BotonSalirClickConnect.jpg", Me)
                                    
    Call cBotonForo.Initialize(imgVerForo, GrhPath & "BotonVerForo.jpg", _
                                    GrhPath & "BotonVerForoRollover.jpg", _
                                    GrhPath & "BotonVerForoClick.jpg", Me)
                                    
    Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonConectarse.jpg", _
                                    GrhPath & "BotonConectarseRollover.jpg", _
                                    GrhPath & "BotonConectarseClick.jpg", Me)
                                    
    Call cBotonTeclas.Initialize(imgTeclas, GrhPath & "BotonTeclas.jpg", _
                                    GrhPath & "BotonTeclasRollover.jpg", _
                                    GrhPath & "BotonTeclasClick.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgBorrarPj_Click()

On Error GoTo errH
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)

    Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgCodigoFuente_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el código de sus servidor de esta forma.
'Caso contrario estarían violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es libre, es de todos. Mantengamoslo así. Si tanto te gusta el juego y querés los
'cambios que hacemos nosotros, compartí los tuyos. Es un cambio justo. Si no estás de acuerdo,
'no uses nuestro código, pues nadie te obliga o bien utiliza una versión anterior a la 0.12.0.
'***********************************
    Call ShellExecute(0, "Open", "https://sourceforge.net/project/downloading.php?group_id=67718&filename=AOServerSrc0.12.2.zip&a=42868900", "", App.path, SW_SHOWNORMAL)

End Sub

Private Sub imgConectarse_Click()
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    
    If wskData.State <> sckClosed Then wskData.Close
   
    Cuenta.Name = txtNombre.Text
    Cuenta.Pass = txtPasswd.Text
 
    frmNewCuenta.Personaje_Index = 1
 
    Acc_Data.acc_Name = txtNombre.Text
 
    'If CheckAccData(False, False) = True Then
    EstadoLogin = E_MODO.Normal

    frmMain.Winsock1.Connect CurServerIp, CurServerPort
    'End If
End Sub

Private Sub imgLeerMas_Click()
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgCrearCuenta_Click()
    Form1.Show
End Sub

Private Sub imgManual_Click()
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/manual/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgRecuperar_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgReglamento_Click()
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/reglamento.html", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgSalir_Click()
    prgRun = False
End Sub

Private Sub imgServArgentina_Click()
    Call Audio.PlayWave(SND_CLICK)
    IPTxt.Text = IPdelServidor
    PortTxt.Text = PuertoDelServidor
End Sub

Private Sub imgTeclas_Click()
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    txtPasswd.SetFocus
End Sub

Private Sub imgVerForo_Click()
    Call ShellExecute(0, "Open", "http://www.alkon.com.ar/foro/argentum-online.53/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub lst_Svs_Click()

' @ Requiere ip y nombre del servidor clickeado.

If Not Winsock1.State <> sckConnected Then
        If (lst_Svs.Text <> vbNullString) Then
            Winsock1.SendData "1@" & CStr(lst_Svs.ListIndex + 1)
        End If
Else
    Winsock1.Close
    Winsock1.Connect "localhost", 555
    pendientAction = 2
End If

End Sub

Private Sub MainViewPic_Click()

Now_Server_Ip = "localhost"

End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectarse_Click
End Sub

Private Sub Winsock1_Connect()

' @ Acciones del conectar al servidorList

If pendientAction = 1 Then
    Winsock1.SendData "0"
ElseIf pendientAction = 2 Then
    If (lst_Svs.Text <> vbNullString) Then
        Winsock1.SendData "1@" & CStr(lst_Svs.ListIndex + 1)
    End If
ElseIf pendientAction = 3 Then
    Winsock1.SendData "2@" & Server_Add_Name & "@" & Server_Add_IP
End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next

' @ Maneja la data entrante.

Dim gData   As String
Dim LoopC   As Long
Dim tmp()   As String

Winsock1.GetData gData, vbString

If Left$(gData, 1) = "0" Then
    If InStr(1, gData, "0@") <> 0 Then
        gData = Replace(gData, "0@", "")
    End If

    tmp = Split(gData, ",")

    lst_Svs.Clear

    For LoopC = LBound(tmp()) To UBound(tmp())
        lst_Svs.AddItem tmp(LoopC)
    Next LoopC

    Winsock1.Close
ElseIf Left$(gData, 1) = "1" Then
    tmp = Split(gData, "@")
    lblName.Caption = "Nombre:" & tmp(2)
    If InStr(1, tmp(1), "CLOSE") <> 0 Then
        lblIP.Caption = "IP:" & Left$(tmp(1), InStr(1, tmp(1), "CLOSE") - 1)
    Else
        lblIP.Caption = "IP:" & tmp(1)
    End If
    
    lblNumUsers.Caption = "Users:" & tmp(3)
    Now_Server_Ip = mid$(lblIP.Caption, 4)
    
End If

Winsock1.Close

End Sub

