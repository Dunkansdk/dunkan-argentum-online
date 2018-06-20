VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Settings - AO"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Aceptar y guardar"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Video"
      Height          =   3615
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   960
         ItemData        =   "frmOptions.frx":0000
         Left            =   120
         List            =   "frmOptions.frx":0010
         TabIndex        =   22
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CheckBox chkVsync 
         Caption         =   "&Sincronización vertical"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmOptions.frx":0033
         Left            =   120
         List            =   "frmOptions.frx":003A
         TabIndex        =   19
         Text            =   "Software"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "&Uso máximo de memoria:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "&Aceleración:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Los cambios que se realicen en relación con el video se verán al reiniciar el cliente."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Graphics"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   120
         Max             =   160
         Min             =   40
         TabIndex        =   16
         Top             =   4320
         Value           =   80
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   120
         Max             =   160
         Min             =   50
         TabIndex        =   14
         Top             =   3720
         Value           =   100
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   180
         Min             =   60
         TabIndex        =   12
         Top             =   3120
         Value           =   90
         Width           =   2415
      End
      Begin VB.CheckBox chkDamage 
         Caption         =   "&Daño en el mapa"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkWeater 
         Caption         =   "&Lluvia, nieve y niebla"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox chkAmbiente 
         Caption         =   "Día, noche, tarde, mañana"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   2640
         Max             =   11
         Min             =   6
         TabIndex        =   7
         Top             =   480
         Value           =   10
         Width           =   2535
      End
      Begin VB.CheckBox chkRadius 
         Caption         =   "&Radio de luz"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chkMinimap 
         Caption         =   "&Minimapa (TAB)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkProyectil 
         Caption         =   "&Proyectiles"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkWater 
         Caption         =   "&Movimiento del agua"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "&Intensidad de la nieve:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "&Intensidad de la lluvia:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "&Intensidad de la niebla:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "&Tamaño del buffer:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

