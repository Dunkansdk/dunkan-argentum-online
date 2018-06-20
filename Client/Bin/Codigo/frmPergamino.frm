VERSION 5.00
Begin VB.Form frmPergamino 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4965
   ClientLeft      =   5565
   ClientTop       =   2805
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aceptar mision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Message 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1965
   End
End
Attribute VB_Name = "frmPergamino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim pergaminoPATH As String
    
    pergaminoPATH = App.path & "\..\Resources\Graphics\Pergamino.gif"
    
    Me.Picture = LoadPicture(pergaminoPATH)
    
    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
   
    MakeFormTransparent Me, vbWhite
    Message.Caption = "HOLA ASKO FRAKA VIENBENIDO A UN JUEGO DE ROLL XD"
End Sub

Private Sub Message_Click()
    Unload Me
End Sub
