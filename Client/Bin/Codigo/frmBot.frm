VERSION 5.00
Begin VB.Form frmBot 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "INVOKATE UN BOT KPO"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SPAWN"
      Height          =   1215
      Left            =   1920
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame fPos 
      Caption         =   "Posición de spawn."
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtMap 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "posY:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "posX:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox cClase 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Text            =   "Clase"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "BotName <CLAN>"
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cRuta 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "Viajar hacia"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CheckBox chkViaja 
      Caption         =   "Viajante"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkViaja_Click()
    cRuta.Enabled = (chkViaja.value = 1)
End Sub

Private Sub Command1_Click()
    If cClase.ListIndex <> -1 Then
        If Val(txtMap.Text) <> 0 Then
            If Val(txtX.Text) <> 0 Then
               If Val(txtY.Text) <> 0 Then
                  Call mod_DunkanProtocol.WriteSpawnBot(txtTag.Text, cClase.ListIndex + 1, Val(txtMap.Text), Val(txtX.Text), Val(txtY.Text), (chkViaja.value = 1))
               End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
'DEFAULT POS.
    txtMap.Text = CStr(UserMap)
    txtX.Text = CStr(UserPos.X + 5)
    txtY.Text = CStr(UserPos.Y)
    
'DEFAULT CLASS.
    cClase.AddItem "Clerigo"
    cClase.AddItem "Mago"
    cClase.AddItem "Cazador"

'DEFAULT (NOT MODIFY) PATH
    cRuta.AddItem "Medio de ulla hasta Muelle de nix"
    
End Sub
