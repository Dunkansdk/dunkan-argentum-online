VERSION 5.00
Begin VB.Form frmBando 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "PK O CIUDA CHE??"
   ClientHeight    =   1245
   ClientLeft      =   6210
   ClientTop       =   6435
   ClientWidth     =   2055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_Bando 
      Caption         =   "CIUDADANO"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmd_Bando 
      Caption         =   "CRIMINAL"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmBando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Bando_Click(Index As Integer)

WriteBando (Index = 0)

End Sub
