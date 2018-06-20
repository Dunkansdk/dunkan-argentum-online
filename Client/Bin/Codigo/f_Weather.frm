VERSION 5.00
Begin VB.Form f_Weather 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather - Dunkansdk"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
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
   ScaleHeight     =   3030
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraIntensidad 
      Caption         =   "Intensidad"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
      Begin VB.HScrollBar sNiebla 
         Height          =   255
         Left            =   1320
         Max             =   90
         Min             =   40
         TabIndex        =   11
         Top             =   1080
         Value           =   50
         Width           =   1695
      End
      Begin VB.HScrollBar sNieve 
         Height          =   255
         Left            =   1320
         Max             =   250
         Min             =   50
         TabIndex        =   9
         Top             =   720
         Value           =   150
         Width           =   1695
      End
      Begin VB.HScrollBar sLluvia 
         Height          =   255
         Left            =   1320
         Max             =   250
         Min             =   50
         TabIndex        =   7
         Top             =   360
         Value           =   150
         Width           =   1695
      End
      Begin VB.Label aaaa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De la niebla:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblDeLaNieve 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De la nieve:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblDeLa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De la lluvia:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame FraWeaterGame 
      Caption         =   "Weather Game"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton OptQueTenga 
         Caption         =   "Que tenga niebla"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton OptQueNinguno 
         Caption         =   "Que ninguno"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton OptQueNieve 
         Caption         =   "Que nieve"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton OptQueLlueva 
         Caption         =   "Que llueva"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "f_Weather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OptQueLlueva_Click()
LastWeather = 1
End Sub

Private Sub OptQueNieve_Click()
LastWeather = 2
End Sub

Private Sub OptQueNinguno_Click()
LastWeather = 0
End Sub

Private Sub OptQueTenga_Click()
LastWeather = 3
End Sub

Private Sub sLluvia_Change()
Settings.RainIntensity = sLluvia.value
End Sub

Private Sub sNiebla_Change()
Settings.FogIntensity = sNiebla.value
End Sub

Private Sub sNieve_Change()
Settings.SnowIntensity = sNieve.value
End Sub
