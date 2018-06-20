VERSION 5.00
Begin VB.Form f_Mountain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mountain Form - Dunkansdk"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   960
      TabIndex        =   12
      Top             =   1680
      Width           =   990
   End
   Begin VB.Frame FraInformation 
      Caption         =   "Information"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdGenerateMountain 
         Caption         =   "Generate Mountain"
         Height          =   240
         Left            =   2280
         TabIndex        =   11
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtPosY 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "46"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPosX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Text            =   "58"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtAltura 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "120"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtRadioY 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "4"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtRadioX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "10"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRadioY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Radio Y:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblPosY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pos Y:"
         Height          =   195
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblPosX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pos X:"
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblAltura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Altura:"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblRadio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Radio X:"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "f_Mountain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerateMountain_Click()
    Generate_Mountain txtAltura, txtRadioX, txtRadioY, txtPosX, txtPosY
End Sub

