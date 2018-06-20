Attribute VB_Name = "Mod_DX8_SelectClase"
Option Explicit

Private Type structSelectMenuChar
    Head            As Integer
    Body            As Integer
    arrCabezas()    As Integer
End Type

Private SelectMenuChar As structSelectMenuChar

Public Sub SelectCharHead(ByRef Heads() As Integer)

' '
' Inicializa las cabezas de los personajes

ReDim Heads(1 To 8) As Integer

Dim i As Long

    For i = 1 To 8
    
        With SelectMenuChar
    
                If (i = eClass.Mage) Or (i = eClass.Paladin) Or (i = eClass.Hunter) Then
                    .Head = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)
                ElseIf (i = eClass.Cleric) Or (i = eClass.Assasin) Then
                    .Head = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)
                ElseIf (i = eClass.Bard) Or (i = eClass.Druid) Then
                    .Head = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)
                Else
                    .Head = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)
                End If
            
            Heads(i) = .Head
        
        End With
    
    Next i

End Sub

Public Sub Engine_Render_Cuenta()

' '
' Renderiza el menu para seleccionar las clases
    
Dim i               As Long
Dim X               As Integer
Dim Y               As Integer
Dim notY            As Integer

With SelectMenuChar

    'Inicializa las cabezas.
    If Not .Head <> 0 Then
       Call SelectCharHead(.arrCabezas)
       .Head = 1
    End If
    
    If (frmNewCuenta.Personaje_Index > 4) Then
    
        X = ((frmNewCuenta.Personaje_Index * 100) - (4 * 100))
        Y = 180
        Engine_Draw_Box X - 30, Y - 40, 93, 93, D3DColorARGB(50, 1, 200, 1)
        
    Else
    
        X = (frmNewCuenta.Personaje_Index * 100)
        Y = 80
        Engine_Draw_Box X - 30, Y - 40, 93, 93, D3DColorARGB(50, 1, 200, 0)
        
    End If
    
    For i = 1 To 8
            
        If (i > 4) Then
            X = ((i * 100) - (4 * 100))
            Y = 180
        Else
            X = (i * 100)
            Y = 80
        End If

        If (i = eClass.Mage) Or (i = eClass.Paladin) Or (i = eClass.Hunter) Then
           .Body = 21
        ElseIf (i = eClass.Cleric) Or (i = eClass.Assasin) Then
           .Body = 32
        ElseIf (i = eClass.Bard) Or (i = eClass.Druid) Then
           .Body = 210
        Else
           .Body = 53
        End If
        
        'Offset de la cabeza / enanos.
        If (i <> eClass.Warrior) Then
           notY = 5
        Else
           notY = -5
        End If
        
        'Si tiene cuerpo dibuja
        If (.Body <> 0) Then
        
            Engine_Draw_Box X - 30, Y - 40, 93, 93, D3DColorARGB(100, 100, 100, 100)

            DDrawTransGrhtoSurface BodyData(.Body).Walk(3), X, Y, 1, 0, DefaultColor(), 150

            If (.arrCabezas(i) <> 0) Then
                DDrawTransGrhtoSurface HeadData(.arrCabezas(i)).Head(3), X, Y - notY, 1, 0, DefaultColor()
            End If

            Engine_RenderText X - 1, Y + 33, ListaClases(i), D3DColorARGB(100, 255, 255, 255)
            
        End If

    Next i
    
    Engine_Draw_Box 70, 240, 194, 60, D3DColorARGB(100, 155, 0, 0)
    
    Engine_Draw_Box 270, 240, 194, 60, D3DColorARGB(100, 0, 0, 155)

End With
    
End Sub

Public Sub Click_SelectClase(ByVal X As Integer, ByVal Y As Integer)

    

End Sub

