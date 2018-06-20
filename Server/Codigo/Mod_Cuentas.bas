Attribute VB_Name = "Mod_Cuentas"
Option Explicit

Public Type pjs
    namePj As String
    LvlPJ As Byte
    ClasePJ As eClass
End Type

Public Type Acc
    Name As String
    Pass As String
    CantPjs As Byte
    PJ(1 To 8) As pjs
End Type

Public Cuenta As Acc

Public Sub CrearCuenta(ByVal userIndex As Integer, ByVal Name As String, ByVal Pass As String, ByVal Email As String)
Dim ciclo As Byte
'¿Posee caracteres invalidos?
If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
    Call WriteErrorMsg(userIndex, "Nombre invalido.")
    Exit Sub
End If

'Si ya existe la cuenta
If FileExist(App.Path & "\Cuentas\" & Name & ".bao", vbNormal) Then
    Call WriteErrorMsg(userIndex, "El nombre de la cuenta ya existe, por favor ingresa otro.")
    Exit Sub
End If

Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "CUENTA", "NOMBRE", Name)
Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "CUENTA", "PASSWORD", Pass)
Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "CUENTA", "MAIL", Email)
Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "CUENTA", "FECHA_CREACION", Now)
Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "CUENTA", "BAN", "0")

'************************RELLENO LOS PJs************************'
Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "PERSONAJES", "CANTIDAD_PJS", "0")
For ciclo = 1 To 8
    Call WriteVar(App.Path & "\Cuentas\" & Name & ".bao", "PERSONAJES", "PJ" & ciclo, "")
Next ciclo
'************************************************************'

Call EnviarCuenta(userIndex, "", "", "", "", "", "", "", "", "0", "1")
End Sub

Public Sub ConectarCuenta(ByVal userIndex As Integer, ByVal Name As String, ByVal Pass As String)

'
' @ Modificado por maTih.-

Dim loopX   As Long

With Cuenta

        .CantPjs = 1
        
        For loopX = 1 To 8
            'Llena los nicks.
            .PJ(loopX).namePj = Name
        Next loopX
        
        Call EnviarCuenta(userIndex, .PJ(1).namePj, .PJ(2).namePj, .PJ(3).namePj, .PJ(4).namePj, _
        .PJ(5).namePj, .PJ(6).namePj, .PJ(7).namePj, .PJ(8).namePj, .CantPjs, "1")

End With
End Sub

Public Sub AgregarPersonaje(ByVal userIndex As Integer, ByVal CuentaName As String, ByVal UserName As String)
Dim CantidadPJs As Byte
CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "CANTIDAD_PJS")

WriteVar App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs + 1
WriteVar App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ" & (CantidadPJs + 1), UserName

WriteVar App.Path & "\Charfile\" & UserName & ".CHR", "INIT", "CUENTA", UCase(CuentaName)
With Cuenta
'Actualizamos la cuenta.
        .CantPjs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "CANTIDAD_PJS")
        
        .PJ(1).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ1")
        .PJ(2).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ2")
        .PJ(3).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ3")
        .PJ(4).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ4")
        .PJ(5).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ5")
        .PJ(6).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ6")
        .PJ(7).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ7")
        .PJ(8).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ8")
        
        Call EnviarCuenta(userIndex, .PJ(1).namePj, .PJ(2).namePj, .PJ(3).namePj, .PJ(4).namePj, _
        .PJ(5).namePj, .PJ(6).namePj, .PJ(7).namePj, .PJ(8).namePj, .CantPjs, "1")
        
        Call WriteVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)

End With
End Sub

Public Sub BorrarPersonaje(ByVal userIndex As Integer, ByVal CuentaName As String, ByVal IndiceUser As String)
Dim CantidadPJs As Byte
Dim namePj As String
Dim c As String
Dim d As String
Dim f As String
Dim g As String
Dim H As Byte
Dim i As String
Dim j As String
    
'Consulto el nombre del PJ a eliminar
namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ" & IndiceUser)

CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "CANTIDAD_PJS")

WriteVar App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs - 1
WriteVar App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ" & IndiceUser, ""

        Call WriteErrorMsg(userIndex, "Personaje eliminado con éxito.")
    Call Kill(App.Path & "\Charfile\" & namePj & ".CHR")

With Cuenta
'Actualizamos la cuenta.
        .CantPjs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "CANTIDAD_PJS")
        
        .PJ(1).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ1")
        .PJ(2).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ2")
        .PJ(3).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ3")
        .PJ(4).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ4")
        .PJ(5).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ5")
        .PJ(6).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ6")
        .PJ(7).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ7")
        .PJ(8).namePj = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "PERSONAJES", "PJ8")
        
        Call EnviarCuenta(userIndex, .PJ(1).namePj, .PJ(2).namePj, .PJ(3).namePj, .PJ(4).namePj, _
        .PJ(5).namePj, .PJ(6).namePj, .PJ(7).namePj, .PJ(8).namePj, .CantPjs, "1")
        
        Call WriteVar(App.Path & "\Cuentas\" & CuentaName & ".bao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)

End With

End Sub

