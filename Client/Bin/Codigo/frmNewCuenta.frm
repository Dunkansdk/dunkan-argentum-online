VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmNewCuenta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " -  Cuenta -"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCuenta 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   2415
      Left            =   240
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   2
      Top             =   240
      Width           =   3975
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   7320
         Width           =   3015
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
            TabIndex        =   12
            Top             =   240
            Width           =   2460
         End
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
            TabIndex        =   11
            Top             =   600
            Width           =   2460
         End
         Begin VB.Image imgConectarse 
            Height          =   375
            Left            =   120
            Top             =   990
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Lista de servidores"
         Height          =   2655
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   2895
         Begin VB.ListBox lst_Svs 
            Height          =   1815
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Listar"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   2160
            Width           =   2175
         End
         Begin MSWinsockLib.Winsock wskData 
            Left            =   240
            Top             =   960
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSWinsockLib.Winsock Winsock1 
            Left            =   240
            Top             =   480
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Info del servidor"
         Height          =   2415
         Left            =   3480
         TabIndex        =   3
         Top             =   2880
         Width           =   2055
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
            TabIndex        =   6
            Top             =   360
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
            TabIndex        =   5
            Top             =   1200
            Width           =   1815
         End
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
            TabIndex        =   4
            Top             =   1680
            Width           =   1815
         End
      End
      Begin VB.Image imgTeclas 
         Height          =   375
         Left            =   6840
         Top             =   9000
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmd_Log 
      Caption         =   "Conectarse"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ListBox lst_Pjs 
      Height          =   2205
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmNewCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Personaje_Index  As Byte
Public Eleccion_Index   As Byte

Private Sub cmd_Log_Click()

'Cierra el winsock.
If frmMain.Winsock1.State <> sckClosed Then
    frmMain.Winsock1.Close
    DoEvents
End If

'Setea el estadoLogin
EstadoLogin = E_MODO.Normal

'Se conecta al servidor.
frmMain.Winsock1.Connect CurServerIp, CurServerPort

End Sub

Private Sub lst_Pjs_Click()
    Personaje_Index = (lst_Pjs.ListIndex + 1)
End Sub

Private Sub picCuenta_DblClick()

If (Personaje_Index <> 0) Then
    cmd_Log_Click
End If

End Sub

Private Sub picCuenta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (X > 21 And X < 40) And (Y > 5 And Y < 60) Then
   Personaje_Index = 1
End If

If (X > 79 And X < 120) And (Y > 5 And Y < 60) Then
   Personaje_Index = 2
End If

If (X > 150 And X < 220) And (Y > 5 And Y < 60) Then
   Personaje_Index = 3
End If

If (X > 190 And X < 240) And (Y > 5 And Y < 60) Then
   Personaje_Index = 4
End If

If (X > 21 And X < 40) And (Y > 80 And Y < 150) Then
   Personaje_Index = 5
End If

If (X > 79 And X < 120) And (Y > 80 And Y < 150) Then
   Personaje_Index = 6
End If

If (X > 150 And X < 185) And (Y > 80 And Y < 150) Then
   Personaje_Index = 7
End If

If (X > 190 And X < 240) And (Y > 80 And Y < 150) Then
   Personaje_Index = 8
End If

End Sub
