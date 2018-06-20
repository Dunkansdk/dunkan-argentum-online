VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear Cuenta"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   135
      TabIndex        =   7
      Top             =   1275
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREAR CUENTA"
      Height          =   330
      Left            =   720
      TabIndex        =   6
      Top             =   1320
      Width           =   2670
   End
   Begin VB.TextBox TMail 
      Height          =   285
      Left            =   1155
      TabIndex        =   4
      Top             =   945
      Width           =   2760
   End
   Begin VB.TextBox TPass 
      Height          =   285
      Left            =   1155
      TabIndex        =   2
      Top             =   585
      Width           =   2760
   End
   Begin VB.TextBox TName 
      Height          =   285
      Left            =   1155
      TabIndex        =   0
      Top             =   225
      Width           =   2760
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   5
      Top             =   990
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   630
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   270
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
       Cuenta.Name = UCase(LTrim(TName.Text))
        Cuenta.Pass = TPass.Text
        Cuenta.Email = TMail.Text
 
        If Cuenta.Name = "" Then
        MsgBox ("Ingrese un nombre.")
       Exit Sub
         End If
   
       If Cuenta.Pass = "" Then
        MsgBox ("Ingrese un password.")
        Exit Sub
    End If
   
    If Not CheckMailString(Cuenta.Email) Then
        MsgBox "Direccion de mail invalida."
        Exit Sub
    End If
 
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
   
        'If CheckAccData(True, True) = True Then
            EstadoLogin = CrearCuenta
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
        'End If
       
        Exit Sub
End Sub

