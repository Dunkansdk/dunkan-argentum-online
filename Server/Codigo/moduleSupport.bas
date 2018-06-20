Attribute VB_Name = "moduleSupport"
' / Module Support - - - - - - - - - - - - -

' / Author: maTih
' / Coded for: Dunkan AO

' - Tabulado y organizado por Dunkan
' Dunkan: Optimize varios bucles, SIEMPRE USAR LONG!!!

Option Explicit

Public Users()      As String
Public Mensajes()   As String
Public LastMensaje  As Byte

Public Sub Soporte_Agrega(ByVal name As String, ByVal strMsg As String)

' / Author: maTih

ReDim Preserve Users(1 To Soporte_SlotLibre)
ReDim Preserve Mensajes(1 To Soporte_SlotLibre)

Users(LastMensaje) = name

Mensajes(LastMensaje) = strMsg

If NameIndex(name) > 0 Then
    WriteConsoleMsg NameIndex(name), "El soporte ha sido enviado, cuando un administrador te responda, el ícono de la carta se iluminará, recuerda que si tu soporte es inválido , o contiene insultos podrás ser penado.", FontTypeNames.FONTTYPE_GUILD
End If

End Sub
Public Function Soporte_SlotLibre() As Byte

' / Author: maTih

Dim loopC   As Long

If LastMensaje = 0 Then
    LastMensaje = LastMensaje + 1
    Soporte_SlotLibre = LastMensaje
    Exit Function
End If

For loopC = 1 To LastMensaje
    If Mensajes(loopC) <> vbNullString Then
        Soporte_SlotLibre = LastMensaje + 1
        LastMensaje = LastMensaje + 1
    Else
        Soporte_SlotLibre = loopC
    End If
Next loopC

End Function
Public Sub Soporte_ModuloInicia()

' / Author: maTih

LastMensaje = 0

End Sub

Public Sub Soporte_LeerUser(ByVal RequestUser As Integer)

' / Author: maTih

Dim loopC   As Long
Dim tempMensaje As String

For loopC = 1 To LastMensaje
    If UCase$(Users(loopC)) = UCase$(UserList(RequestUser).name) Then
        tempMensaje = Mensajes(loopC)
    End If
Next loopC

WriteConsoleMsg RequestUser, "El administrador respondió : " & tempMensaje, FontTypeNames.FONTTYPE_GUILD

End Sub


Public Sub Soporte_LeerGM(ByVal sId As Byte, ByVal RequestID As Integer)

' / Author: maTih

If Mensajes(sId) <> vbNullString Then

    WriteConsoleMsg RequestID, Users(sId) & ">" & Mensajes(sId), FontTypeNames.FONTTYPE_GUILD
    Soporte_Leido sId

End If

End Sub

Public Sub Soporte_Leido(ByVal sId As Byte)

' / Author: maTih

If LenB(Users(sId)) > 0 Then
    Users(sId) = vbNullString
    Mensajes(sId) = vbNullString
End If

End Sub

Public Sub Soporte_Responde(ByVal name As String, ByVal Responded As String)

' / Author: maTih

Dim tU      As Integer

If NameIndex(name) > 0 Then

    tU = NameIndex(name)
    With UserList(tU)
        .Mensaje = Responded
        WriteConsoleMsg tU, "Un administrador te ha respondido el soporte que enviaste, podrás leerlo clikeando el ícono de la carta.", FontTypeNames.FONTTYPE_GUILD
    End With
    
Else

    If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
        WriteVar CharPath & UCase$(name) & ".chr", "SOPORTE", "Mensaje", Responded
    End If
    SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El soporte ha " & name & " se ha grabado en el charfile", FontTypeNames.FONTTYPE_GUILD)

End If

End Sub

Public Function Soporte_Envio(ByVal name As String) As Boolean

' / Author: maTih

Dim loopC As Long

For loopC = 1 To LastMensaje

    If UCase$(Users(loopC)) = UCase$(name) Then
        Soporte_Envio = True
    Else
        Soporte_Envio = False
    End If

Next loopC

End Function

