VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frm_Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Servidor HOST"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "cmdHide"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmd_Quitar 
      Caption         =   "Quitar los offlie"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock wskTest 
      Left            =   4680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timerCheckConnections 
      Interval        =   60000
      Left            =   3000
      Top             =   120
   End
   Begin VB.ListBox lst_Svs 
      Height          =   2205
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmd_Armar 
      Caption         =   "Armar Lista de prueba [TEST]"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock wskData 
      Left            =   4200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl_lst 
      Caption         =   "Lista de servers (TOTAL:)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Pendient_Send_Server_Index  As Integer
Private Puede_Conectarse            As Boolean
Private Server_Index_Analized       As Integer

Private Sub cmd_Armar_Click()

Dim i   As Long

ReDim Server_List(1 To 5)

For i = 1 To 5
    Server_List(i).Nombre = "Server de juan XD"
    Server_List(i).Online = True
    Server_List(i).NumUsers = 5 + i
    Server_List(i).Internet_Protocol = "localhost"
Next i

Call mod_Declares.Actualizar_Lista

End Sub

Private Sub cmd_Quitar_Click()
timerCheckConnections.Enabled = Not timerCheckConnections.Enabled
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()

' @ Prepara el winsock.

Puede_Conectarse = True

Server_List_Port = 555

With wskData
     'Inicializa el puerto.
     .LocalPort = Server_List_Port
     
     'Deja el winsock a la escucha.
     .Listen
End With

Call mod_Declares.Inicializar(Server_List())

End Sub

Private Sub lst_Svs_Click()

If lst_Svs.ListIndex <> -1 Then
   MsgBox Server_List(lst_Svs.ListIndex + 1).MaxUsers
End If

End Sub

Private Sub timerCheckConnections_Timer()

' @ Checkea las conexiones actuales y borra las que no están online.

Dim i   As Long

Static Count_Checks As Byte

Count_Checks = Count_Checks + 1

If Count_Checks >= 60 Then

    'Si no hay svs.
    If UBound(Server_List()) <> 0 Then
       If Not Server_List(1).Nombre <> vbNullString Then Exit Sub
    End If
    
    For i = 1 To UBound(Server_List())
    'Mientras no pueda conectarse mantiene la interfaz
        Do While (Puede_Conectarse = False)
                 DoEvents
        Loop
    
        'Cierra el socket
        If wskTest.State <> sckClosed Then
           wskTest.Close
           DoEvents
        End If
        
        'Conecta.
        wskTest.Connect Server_List(i).Internet_Protocol, 7666
        Server_Index_Analized = i
        Puede_Conectarse = False
    Next i

    Call mod_Declares.Actualizar_Lista

    Count_Checks = 0
End If
End Sub

Private Sub wskData_ConnectionRequest(ByVal requestID As Long)

' @ Acepta una petición para conectarse.

With wskData
     'Cierra el socket.
     If .State <> sckClosed Then .Close
     
     .LocalPort = 0
     
     'Acepta.
     .Accept requestID
     
End With

End Sub

Private Sub wskData_DataArrival(ByVal bytesTotal As Long)

' @ Recibe la data del cliente.

Dim recData As String

With wskData
    .GetData recData, vbString
End With

'Largo de la cadena
If Len(recData) <> 0 Then
    'Si no es "close".
    If recData <> "@CLOSE" Then
        Call mod_General.Handle_Incoming_Data(recData)
    Else
        wskData.Close
        wskData.LocalPort = 555
        wskData.Listen
    End If
End If

End Sub

Private Sub wskData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskData.Close
End Sub

Private Sub wskData_SendComplete()

If Waiting_Close Then
   wskData.Close
   wskData.LocalPort = 555
   wskData.Listen
   Waiting_Close = False
End If

End Sub

Private Sub wskData_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

wskData.Close
wskData.LocalPort = 555
wskData.Listen

End Sub

Private Sub wskTest_Connect()

'
' @ Se conecta

Puede_Conectarse = True

'maTih.-
Server_List(Server_Index_Analized).Online = True

End Sub

Private Sub wskTest_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

'
' @ Cierra la conexión.

wskTest.Close

Server_List(Server_Index_Analized).Online = False

Puede_Conectarse = True

End Sub

