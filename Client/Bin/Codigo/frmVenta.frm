VERSION 5.00
Begin VB.Form frmVenta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "BÑABÑA"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textAmount 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Text            =   "1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdComprar 
      Caption         =   "COMPRAR"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox picVenta 
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
      Left            =   240
      ScaleHeight     =   162
      ScaleMode       =   0  'User
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   240
      Width           =   2430
   End
End
Attribute VB_Name = "frmVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComprar_Click()
Dim SelectSlot As Byte

SelectSlot = InvVentaUser(1).SelectedItem

'no select.
If Not SelectSlot <> 0 Then Exit Sub

'no amount.
If Not Val(textAmount.Text) <> 0 Then Exit Sub

Call Mod_DunkanProtocol.WriteBuyVenta(SelectSlot, Val(textAmount.Text))

End Sub
