VERSION 5.00
Begin VB.Form frmVilla 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   3090
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdComprar 
      Caption         =   "comprar"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox picInv 
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
      Height          =   1680
      Left            =   240
      ScaleHeight     =   113.4
      ScaleMode       =   0  'User
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   240
      Width           =   2430
   End
   Begin VB.Label lblAmount 
      Caption         =   "Cantidad:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblPrecio 
      Caption         =   "Precio:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmVilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComprar_Click()

If Not Val(txtAmount.Text) <> 0 Then Exit Sub

If Venta_SelectSlot <> 0 Then
   Call Mod_DunkanProtocol.WriteBuyVenta(Venta_SelectSlot, Val(txtAmount.Text))
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim i           As Long
Dim eraseInv    As ObjVenta

For i = 1 To 20
    Venta_Inventory(i) = eraseInv
Next i

End Sub

Private Sub PicInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim SelectSlot  As Byte

SelectSlot = mod_DX8_Invent.CliCkearItem(CInt(X), CInt(Y))

If SelectSlot <> 0 Then
   Venta_SelectSlot = SelectSlot
   
   lblPrecio.Caption = "Precio:" & mod_DX8_Invent.GetPrecio()
   
End If


End Sub

