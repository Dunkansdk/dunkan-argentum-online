VERSION 5.00
Begin VB.Form frmCrearVenta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Crear venta wACho"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   3000
      ScaleHeight     =   162
      ScaleMode       =   0  'User
      ScaleWidth      =   162
      TabIndex        =   1
      Top             =   360
      Width           =   2430
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
      Height          =   2400
      Left            =   360
      ScaleHeight     =   162
      ScaleMode       =   0  'User
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   360
      Width           =   2430
   End
End
Attribute VB_Name = "frmCrearVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private NewInv As New clsGrapchicalInventory

