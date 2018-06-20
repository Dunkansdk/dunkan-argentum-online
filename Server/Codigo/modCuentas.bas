Attribute VB_Name = "modCuentas"
Option Explicit

' / Author: maTih
' / Note: Here is to manage thenew system of accounts

' - Tabulado y organizado por Dunkan

Public AccountPath As String    'Path of accounts

Public Function Account_CanCreate(ByVal Nombre As String, ByVal Password As String, ByVal Email As String, ByVal Pin As Byte, ByRef msjError As String) As Boolean

' / Author: maTih

If Account_Exist(Nombre) Then
    msjError = "Ya existe una cuenta con ese nombre."
    Account_CanCreate = False
    Exit Function
End If

If Not AsciiValidos(Password) Then
    msjError = "La contraseña contiene carácteres inválidos."
    Account_CanCreate = False
    Exit Function
End If

If Not Pin > 99 Then ' Feo, optimizar.
    msjError = "El Pin debe contener 3 carácteres númericos."
    Account_CanCreate = False
    Exit Function
End If

Account_CanCreate = True

End Function

Public Function Account_Exist(ByVal NameAcc As String) As Boolean

' / Author: maTih

Account_Exist = FileExist(App.Path & "\Accounts\" & NameAcc & ".Acc")

End Function

Public Sub Accounts_Create(ByVal Nombre As String, ByVal Password As String, ByVal Email As String, ByVal Pin As Byte)

' / Author: maTih

Dim loopC   As Long

    WriteVar AccountPath & Nombre & ".Acc", "INICIO", "Contraseña", Password
    WriteVar AccountPath & Nombre & ".Acc", "INICIO", "Email", Email
    WriteVar AccountPath & Nombre & ".Acc", "INICIO", "Banned", 0
    WriteVar AccountPath & Nombre & ".Acc", "INICIO", "Pin", CStr(Pin)

    For loopC = 1 To 10
        WriteVar AccountPath & Nombre & ".Acc", "PERSONAJES", "Personaje" & loopC, "NothingPJ"
        WriteVar AccountPath & Nombre & ".Acc", "PERSONAJES", "Cantidad", 0
    Next loopC

    Accounts_SendAccount Nombre

End Sub

Public Function Account_CheckPass(ByVal Name As String, ByVal Pass As String) As Boolean

' / Author: maTih

Dim tmpPass As String

    tmpPass = GetVar(AccountPath & Name & ".chr", "INICIO", "Contraseña")
    Account_CheckPass = (tmpPass = Pass)
    
End Function

Public Function Account_PjSlot(ByVal Slot As Byte, ByVal NameAcc As String) As Boolean

' / Author: maTih

Dim pjSlot      As String

    pjSlot = GetVar(AccountPath & NameAcc & ".chr", "PERSONAJES", "PERSONAJE" & Slot)
    
    If pjSlot = "NothingPJ" Then
        Account_PjSlot = False
        Exit Function
    End If
    
    Account_PjSlot = True

End Function
Public Sub Accounts_SendAccount(ByVal NameACCOUNT As String)

' / Author: maTih


End Sub

Public Sub Account_LoginChar(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal NameAcc As String)

' / Author: maTih

Dim namePj  As String

    namePj = GetVar(AccountPath & NameAcc & ".Acc", "PERSONAJES", "PERSONAJE" & Slot)

    ConnectUser UserIndex, namePj

End Sub
