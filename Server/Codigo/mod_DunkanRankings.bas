Attribute VB_Name = "mod_DunkanRankings"
' Programado y diseñado por maTih.-

Type subRankVida
     Nick                   As String       'Nombre del pj.
     Vida                   As Integer      'Vida del pj.
     Nivel                  As Byte         'NIvel del pj.
End Type

Type rankVidas
     Vidas(1 To 10)         As subRankVida
End Type

Type rankOro
     Nick                   As String       'Nombre del pj.
     Oro                    As Long         'Oro del pj.
End Type

Type rankNivel
     Nick                   As String       'Nombre del pj.
     Nivel                  As Byte         'Nivel del pj.
End Type

Type ALLRanks
     Oro(1 To 10)           As rankOro
     Nivel(1 To 10)         As rankNivel
     Vida(1 To NUMCLASES)   As rankVidas
     ArchivoOro             As String       'Archivo del ranking.
     ArchivoNivel           As String       'Archivo del ranking.
     ArchivoVida            As String       'Archivo del ranking.
End Type

Public Ranking      As ALLRanks

Sub CargarTodos()

' @ Carga los tres tipos de rankings.

'Initialize file.
Ranking.ArchivoNivel = App.Path & "\Ranking\Nivel.txt"
Ranking.ArchivoOro = App.Path & "\Ranking\Oro.txt"
Ranking.ArchivoVida = App.Path & "\Ranking\Vida.txt"

CargarVidas
CargarNiveles
CargarOros

End Sub

Sub GrabarNiveles(Optional ByVal Posicion As Integer = -1)

' @ Guarda el ranking de vidas.

With Ranking

Dim LoopC   As Long

'Una posicion?
If Posicion <> -1 Then
   'Graba.
   WriteVar Ranking.ArchivoNivel, "RANKING" & Posicion, "Nick", .Nivel(Posicion).Nick
   WriteVar Ranking.ArchivoNivel, "RANKING" & Posicion, "Nivel", CStr(.Nivel(Posicion).Nivel)
   Exit Sub
End If

'Guarda todas.
For LoopC = 1 To 10
     
    'Graba.
    WriteVar Ranking.ArchivoNivel, "RANKING" & LoopC, "Nick", .Nivel(LoopC).Nick
    WriteVar Ranking.ArchivoNivel, "RANKING" & LoopC, "Nivel", CStr(.Nivel(LoopC).Nivel)
     
Next LoopC
    
End With

End Sub

Sub GrabarOros(Optional ByVal Posicion As Integer = -1)

' @ Guarda el ranking de vidas.

With Ranking

Dim LoopC   As Long

'Una posicion?
If Posicion <> -1 Then
   'Graba.
   WriteVar Ranking.ArchivoOro, "RANKING" & Posicion, "Nick", .Oro(Posicion).Nick
   WriteVar Ranking.ArchivoOro, "RANKING" & Posicion, "Oro", .Oro(Posicion).Oro
   Exit Sub
End If

'Guarda todas.
For LoopC = 1 To 10
     
    'Graba.
    WriteVar Ranking.ArchivoOro, "RANKING" & LoopC, "Nick", .Oro(LoopC).Nick
    WriteVar Ranking.ArchivoOro, "RANKING" & LoopC, "Oro", CStr(.Oro(LoopC).Oro)
     
Next LoopC
    
End With

End Sub


Sub GrabarVidas(ByVal Clase As eClass, Optional ByVal Posicion As Integer = -1)

' @ Guarda el ranking de vidas.

With Ranking.Vida(Clase)

Dim LoopC   As Long

'Una posicion?
If Posicion <> -1 Then
   'Graba.
   WriteVar Ranking.ArchivoVida, "CLASE" & Clase, "RANKING" & Posicion & "Nick", .Vidas(Posicion).Nick
   WriteVar Ranking.ArchivoVida, "CLASE" & Clase, "RANKING" & Posicion & "Nivel", CStr(.Vidas(Posicion).Nivel)
   WriteVar Ranking.ArchivoVida, "CLASE" & Clase, "RANKING" & Posicion & "Vida", CStr(.Vidas(Posicion).Vida)
   Exit Sub
End If

'Guarda todas.
For LoopC = 1 To 10
     
    'Graba.
    WriteVar Ranking.ArchivoVida, "CLASE" & Clase, "RANKING" & LoopC & "Nick", .Vidas(LoopC).Nick
    WriteVar Ranking.ArchivoVida, "CLASE" & Clase, "RANKING" & LoopC & "Nivel", CStr(.Vidas(LoopC).Nivel)
    WriteVar Ranking.ArchivoVida, "CLASE" & Clase, "RANKING" & LoopC & "Vida", CStr(.Vidas(LoopC).Vida)
     
Next LoopC
    
End With

End Sub

Sub CargarVidas()

' @ Carga el rankin de vidas

Dim LoopX   As Long
Dim LoopC   As Long     'Bucle C < Clases!

For LoopC = 1 To NUMCLASES
    For LoopX = 1 To 10
    
        With Ranking.Vida(LoopC).Vidas(LoopX)
             .Nick = GetVar(Ranking.ArchivoVida, "CLASE" & LoopC, "RANKING" & LoopX & "Nick")
             .Nivel = val(GetVar(Ranking.ArchivoVida, "CLASE" & LoopC, "RANKING" & LoopX & "Nivel"))
             .Vida = val(GetVar(Ranking.ArchivoVida, "CLASE" & LoopC, "RANKING" & LoopX & "Vida"))
        End With
    
    Next LoopX
Next LoopC

End Sub

Sub CargarNiveles()

' @ Carga el rankin de niveles

Dim LoopX   As Long

For LoopX = 1 To 10
    
    With Ranking.Nivel(LoopX)
         .Nick = GetVar(Ranking.ArchivoNivel, "RANKING" & LoopX, "Nick")
         .Nivel = val(GetVar(Ranking.ArchivoNivel, "RANKING" & LoopX, "Nivel"))
    End With
    
Next LoopX

End Sub

Sub CargarOros()

' @ Carga el rankin de oro

Dim LoopX   As Long

For LoopX = 1 To 10
    
    With Ranking.Oro(LoopX)
         .Nick = GetVar(Ranking.ArchivoOro, "RANKING" & LoopX, "Nick")
         .Oro = val(GetVar(Ranking.ArchivoOro, "RANKING" & LoopX, "Oro"))
    End With
    
Next LoopX

End Sub

Sub ActualizarOros(ByVal UserIndex As Integer, ByVal Posicion As Byte)

' @ Agrega un usuario al ranking de oro.

With Ranking
    
     Dim tempTipo(1 To 10)  As rankOro
     Dim BucleX             As Long
    
     'Si entra en la última pos
     If Not Posicion <> 10 Then
        .Oro(10).Nick = UserList(UserIndex).name
        .Oro(10).Oro = UserList(UserIndex).Stats.GLD
        Call Protocol.WriteConsoleMsg(UserIndex, "Ingresaste al ranking de oro! en la posición 10!", FontTypeNames.FONTTYPE_DIOS)
        Call GrabarOros(10)
     End If
     
     'Si no entra en la última pos.
     'Usamos un buffer.
     For BucleX = 1 To 10
         tempTipo(BucleX) = .Oro(BucleX)
     Next BucleX
     
     'Desde la posición que entra hacia 1+
     For BucleX = Posicion To 9
         .Oro(BucleX + 1) = tempTipo(BucleX)
     Next BucleX
     
     .Oro(Posicion).Nick = UserList(UserIndex).name
     .Oro(Posicion).Oro = UserList(UserIndex).Stats.GLD
     
     'Avisa al user.
     Call Protocol.WriteConsoleMsg(UserIndex, "Tu posición en el ranking ha sido actualizada! Estás en la posición #" & Posicion & " del ranking de oro!", FontTypeNames.FONTTYPE_DIOS)
     
     Call GrabarOros
     
End With

End Sub

Sub ActualizarVidas(ByVal UserIndex As Integer, ByVal Posicion As Byte)

' @ Agrega un usuario al ranking de vida.

With Ranking

Dim punteroClass    As eClass
Dim tData(1 To 10)  As subRankVida
Dim BucleX          As Long

punteroClass = UserList(UserIndex).Clase

    With .Vida(punteroClass)
         
         'Si entra en la pos 10
         If Not Posicion <> 10 Then
            .Vidas(10).Nick = UserList(UserIndex).name
            .Vidas(10).Nivel = UserList(UserIndex).Stats.ELV
            .Vidas(10).Vida = UserList(UserIndex).Stats.MaxHp
            Call GrabarNiveles(10)
         End If
         
         'Entra en otra pos
         'Copiamos los datos
         For BucleX = 1 To 10
             tData(BucleX) = .Vidas(BucleX)
         Next BucleX
         
         'Corermos la pos
         For BucleX = Posicion To 9
             .Vidas(BucleX + 1) = tData(BucleX)
         Next BucleX
         
         'guarda los datos
         .Vidas(Posicion).Vida = UserList(UserIndex).Stats.MaxHp
         .Vidas(Posicion).Nivel = UserList(UserIndex).Stats.ELV
         .Vidas(Posicion).Nick = UserList(UserIndex).name
         
         Call Protocol.WriteConsoleMsg(UserIndex, "Tu posición en el ranking ha sido actualizada! estás en la posición " & Posicion & " del ranking de vidas!", FontTypeNames.FONTTYPE_DIOS)
         
         Call GrabarVidas(punteroClass)
         
    End With

End With

End Sub

Sub ActualizarNiveles(ByVal UserIndex As Integer, ByVal Posicion As Byte)

' @ Agrega un usuario al ranking de nivel.

With Ranking

     Dim tempTipo(1 To 10)  As rankNivel
     Dim BucleX             As Long
    
     'Si entra en la última pos
     If Not Posicion <> 10 Then
        .Nivel(10).Nick = UserList(UserIndex).name
        .Nivel(10).Nivel = UserList(UserIndex).Stats.ELV
        Call Protocol.WriteConsoleMsg(UserIndex, "Ingresaste al ranking de niveles! en la posición 10!", FontTypeNames.FONTTYPE_DIOS)
        Call GrabarNiveles(10)
     End If
     
     'Si no entra en la última pos.
     'Usamos un buffer.
     For BucleX = 1 To 10
         tempTipo(BucleX) = .Nivel(BucleX)
     Next BucleX
     
     'Desde la posición que entra hacia 1+
     For BucleX = Posicion To 9
         .Nivel(BucleX + 1) = tempTipo(BucleX)
     Next BucleX
     
     'Guardo sus datos
     .Nivel(Posicion).Nick = UserList(UserIndex).name
     .Nivel(Posicion).Nivel = UserList(UserIndex).Stats.ELV
     
     'Avisa al user.
     Call Protocol.WriteConsoleMsg(UserIndex, "Ingresaste al ranking de niveles! en la posición " & Posicion & "!", FontTypeNames.FONTTYPE_DIOS)
     
     Call GrabarNiveles
     
End With

End Sub

Function IngresaOro(ByVal UserIndex As Integer) As Byte

' @ Comprueba si puede entrar al ranking de oro y devuelve la posición.

With Ranking
    
     Dim LoopX      As Long
     Dim antesPos   As Byte
     
     antesPos = PosicionEnOro(UserIndex)
     
     For LoopX = 1 To 10
         'El usuario tiene más oro?
         If .Oro(LoopX).Oro < UserList(UserIndex).Stats.GLD Then
            'Siempre que no sea el mismo.
            If UCase(.Oro(LoopX).Nick) <> UCase$(UserList(UserIndex).name) Then
               'Escaló pos?
               If antesPos <> 0 Then
                  If antesPos > LoopX Then
                    'Guarda.
                    IngresaOro = CByte(LoopX)
                    Exit Function
                  End If
               Else 'No estaba en el ranking.
                  IngresaOro = CByte(LoopX)
                  Exit Function
               End If
            End If
        End If
     Next LoopX

End With

End Function

Function IngresaVidas(ByVal UserIndex As Integer) As Byte

' @ Comprueba si puede entrar al ranking de vidas y devuelve la posición.

With Ranking

     Dim LoopX      As Long
     Dim antesPos   As Byte
          
     antesPos = PosicionEnVidas(UserIndex)
     
     For LoopX = 1 To 10
         'El usuario tiene más vida?
         If .Vida(UserList(UserIndex).Clase).Vidas(LoopX).Vida < UserList(UserIndex).Stats.MaxHp Then
            'Siempre que no sea el mismo.
            If UCase(.Vida(UserList(UserIndex).Clase).Vidas(LoopX).Nick) <> UCase$(UserList(UserIndex).name) Then
               'Escaló pos?
               If antesPos <> 0 Then
                  If antesPos > LoopX Then
                    'Guarda.
                    IngresaVidas = CByte(LoopX)
                    Exit Function
                  End If
               Else 'No estaba en el ranking.
                  IngresaVidas = CByte(LoopX)
                  Exit Function
               End If
            End If
        End If
     Next LoopX

End With

End Function

Function IngresaNivel(ByVal UserIndex As Integer) As Byte

' @ Comprueba si puede entrar al ranking de niveles y devuelve la posición.

With Ranking

     Dim LoopX      As Long
     Dim antesPos   As Byte
     
     antesPos = PosicionEnNivel(UserIndex)
     
     For LoopX = 1 To 10
         'El usuario tiene más nivel?
         If .Nivel(LoopX).Nivel < UserList(UserIndex).Stats.ELV Then
            'Siempre que no sea el mismo.
            If UCase(.Nivel(LoopX).Nick) <> UCase$(UserList(UserIndex).name) Then
               'Escaló pos?
               If antesPos <> 0 Then
                  If antesPos > LoopX Then
                    'Guarda.
                    IngresaNivel = CByte(LoopX)
                    Exit Function
                  End If
               Else 'No estaba en el ranking.
                  IngresaNivel = CByte(LoopX)
                  Exit Function
               End If
            End If
        End If
     Next LoopX

End With

End Function

Function PosicionEnOro(ByVal UserIndex As Integer) As Byte

' @ Se fija si un usuario está en el ranking y devuelve la posición

Dim LoopX   As Long

For LoopX = 1 To 10
    
    With Ranking.Oro(LoopX)
         If UCase$(.Nick) = UCase$(UserList(UserIndex).name) Then
            PosicionEnOro = CByte(LoopX)
            Exit Function
         End If
    End With
    
Next LoopX

PosicionEnOro = 0

End Function

Function PosicionEnNivel(ByVal UserIndex As Integer) As Byte

' @ Se fija si un usuario está en el ranking y devuelve la posición

Dim LoopX   As Long

For LoopX = 1 To 10
    
    With Ranking.Nivel(LoopX)
         If UCase$(.Nick) = UCase$(UserList(UserIndex).name) Then
            PosicionEnNivel = CByte(LoopX)
            Exit Function
         End If
    End With
    
Next LoopX

PosicionEnNivel = 0

End Function

Function PosicionEnVidas(ByVal UserIndex As Integer) As Byte

' @ Se fija si un usuario está en el ranking y devuelve la posición

Dim LoopX   As Long

For LoopX = 1 To 10
    
    With Ranking.Vida(UserList(UserIndex).Clase).Vidas(LoopX)
         If UCase$(.Nick) = UCase$(UserList(UserIndex).name) Then
            PosicionEnVidas = CByte(LoopX)
            Exit Function
         End If
    End With
    
Next LoopX

PosicionEnVidas = 0

End Function

Function ListaOro() As String

' @ Lista los ranking de oro.

Dim LoopX   As Long

For LoopX = 1 To 10
    If Ranking.Oro(LoopX).Nick <> vbNullString Then
        ListaOro = ListaOro & Ranking.Oro(LoopX).Nick & " Oro: " & Ranking.Oro(LoopX).Oro & ","
    Else
        ListaOro = ListaOro & "Ninguno,"
    End If
Next LoopX

ListaOro = Left$(ListaOro, Len(ListaOro) - 1)

End Function

Function ListaNivel() As String

' @ Lista los ranking de oro.

Dim LoopX   As Long

For LoopX = 1 To 10
    If Ranking.Nivel(LoopX).Nick <> vbNullString Then
        ListaNivel = ListaNivel & Ranking.Nivel(LoopX).Nick & " Nivel: " & Ranking.Nivel(LoopX).Nivel & ","
    Else
        ListaNivel = ListaNivel & "Ninguno,"
    End If
Next LoopX

ListaNivel = Left$(ListaNivel, Len(ListaNivel) - 1)

End Function

Function ListaVidas(ByVal Clase As eClass) As String

' @ Lista los ranking de oro.

Dim LoopX   As Long
Dim RazaStr As String
Dim RazaPj  As Byte

For LoopX = 1 To 10
    RazaPj = val(GetVar(CharPath & Ranking.Vida(Clase).Vidas(LoopX).Nick & ".chr", "INIT", "Raza"))
    
    If RazaPj <> 0 Then
       RazaStr = " Raza: " & ListaRazas(RazaPj)
    Else
       RazaStr = vbNullString
    End If
    
    If Ranking.Vida(Clase).Vidas(LoopX).Nick <> vbNullString Then
        ListaVidas = ListaVidas & Ranking.Vida(Clase).Vidas(LoopX).Nick & " Vida:" & Ranking.Vida(Clase).Vidas(LoopX).Vida & RazaStr & ","
    Else
        ListaVidas = ListaVidas & "Ninguno,"
    End If
Next LoopX

ListaVidas = Left$(ListaVidas, Len(ListaVidas) - 1)

End Function
