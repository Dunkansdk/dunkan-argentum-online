VERSION 5.00
Begin VB.Form frmRanking 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ranking"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbClases 
      Height          =   315
      Left            =   3000
      TabIndex        =   13
      Text            =   "Eliga clase"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblMasVida 
      Caption         =   "USUARIOS CON MAS VIDA"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblMasNivel 
      Caption         =   "USUARIOS CON MAS NIVEL"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblMasOro 
      Caption         =   "USUARIOS CON MAS ORO"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   9
      Top             =   4680
      Width           =   4935
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   3960
      Width           =   4935
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   5175
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label lblUser 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
   End
End
Attribute VB_Name = "frmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim loopX   As Long
    
    For loopX = 1 To NUMCLASES
        cmbClases.AddItem ListaClases(loopX)
    Next loopX
    
    'Call Mod_DunkanProtocol.WriteRequestRanking(1, 0)
    
End Sub

Public Sub LimpiarUsers()

Dim loopX   As Long

For loopX = 0 To (10 - 1)
    lblUser(loopX).Caption = "Ninguno"
Next loopX

End Sub



Private Sub LblMasVida_Click()
    If cmbClases.ListIndex <> -1 Then
        'Call Mod_DunkanProtocol.WriteRequestRanking(3, cmbClases.ListIndex + 1)
    Else
        MsgBox "Elige una clase!"
    End If
End Sub

Private Sub LblMasNivel_Click()
    'Call Mod_DunkanProtocol.WriteRequestRanking(2, 0)
End Sub

Private Sub lblMasOro_Click()
    'Call Mod_DunkanProtocol.WriteRequestRanking(1, 0)
End Sub

