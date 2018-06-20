VERSION 5.00
Begin VB.Form frmNewVenta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Crea nueva venta"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCant 
      Height          =   285
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   3720
      ScaleHeight     =   162
      ScaleMode       =   0  'User
      ScaleWidth      =   162
      TabIndex        =   1
      Top             =   480
      Width           =   2430
   End
   Begin VB.PictureBox picUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   120
      ScaleHeight     =   162
      ScaleMode       =   0  'User
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   480
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblMenos 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblMas 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmNewVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
