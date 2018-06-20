VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMapa 
      Caption         =   "cmdMAPA"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   10560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   10455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMapa_Click()

Dim i   As Long

For i = 1 To UBound(ObjData())
    If Objeto(i) = False Then
    
    End If
Next i

End Sub

Function Objeto(ByVal o As Integer) As Boolean

Dim i   As Long
Dim j   As Long
Objeto = False

For i = 1 To UBound(mod_DunkanCs.Buy())
    For j = 1 To UBound(mod_DunkanCs.Buy(i).Buys())
        With Buy(i).Buys(j)
             If .Objeto.objIndex = o Then
                Exit Function
             End If
        End With
    Next j
Next i

Objeto = True

End Function

