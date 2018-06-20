VERSION 5.00
Begin VB.Form FrmCuenta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4410
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PlayerView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   900
      Left            =   2520
      ScaleHeight     =   3.5
      ScaleMode       =   0  'User
      ScaleWidth      =   916.667
      TabIndex        =   6
      Top             =   240
      Width           =   885
   End
   Begin VB.CommandButton boton 
      Caption         =   "SALIR"
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   960
   End
   Begin VB.CommandButton boton 
      Caption         =   "BORRAR"
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   960
   End
   Begin VB.CommandButton boton 
      Caption         =   "CREAR"
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   960
   End
   Begin VB.CommandButton boton 
      Caption         =   "LOGUEAR"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   960
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "FrmAcconunt.frx":0000
      Left            =   45
      List            =   "FrmAcconunt.frx":0002
      TabIndex        =   0
      Top             =   45
      Width           =   2280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "FrmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PJSeleccionado As Byte

Private Sub Form_Load()
Me.Caption = "  Cuenta de " & Cuenta.Name
End Sub

Private Sub boton_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 1 'CREAR PJ
            If Cuenta.CantPj = 8 Then
                MsgBox "No tienes más espacio para continuar creando personajes."
                Exit Sub
            End If
        'List1.Clear
        EstadoLogin = Dados
        
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
        'Unload Me
        Exit Sub
        
    Case 0  'CONECTAR PJ
        If List1.Text = "" Then MsgBox "No seleccionaste ningún personajes.": Exit Sub
        
        UserName = List1.Text
        List1.Clear
        EstadoLogin = Normal
        
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
        Exit Sub
    
    Case 2 'BORRAR PJ
            If List1.Text = "" Then MsgBox "No seleccionaste ningún personajes.": Exit Sub
            
            IndexSelectedUSer = List1.ListIndex + 1
        If MsgBox("Al borrar un personaje de su cuenta perderá todo lo que hay en él." & vbCrLf & "¿Está totalmente seguro que decea eliminar el mismo?", vbInformation + vbYesNo, "Eliminar Personaje de la cuenta.") = vbYes Then
            List1.Clear
            EstadoLogin = BorrarPJ
            
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
            
         '   List1.Text = ""
            'Cuenta.pjs(IndexSelectedUSer).NamePJ = ""
            Exit Sub
        Else
            Exit Sub
        End If
    Case 3
        frmConnect.Show
        Unload Me
End Select
End Sub

Private Sub List1_DblClick()
Call Audio.PlayWave(SND_CLICK)

Call boton_Click(0)


End Sub

