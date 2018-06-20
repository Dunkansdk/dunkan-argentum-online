Attribute VB_Name = "ModDibujadoPj"
Option Explicit
 
Sub DrawGrh(Grh As Grh, ByVal X As Byte, ByVal Y As Byte)
 
Dim r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer
 
    If Grh.GrhIndex <= 0 Then Exit Sub
   
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
       
    With r2
        .Left = GrhData(iGrhIndex).sX
        .Top = GrhData(iGrhIndex).sY
        .Right = .Left + GrhData(iGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
    End With
   
    With auxr
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
   
    BackBufferSurface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Call BackBufferSurface.BltToDC(FrmCuenta.PlayerView.hdc, auxr, auxr)
 
End Sub
 
Sub DibujarPJ(Index As Byte)
 
Dim Grh As Grh
Dim Pos As Integer
Dim r2 As RECT
 
    With r2
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
With Cuenta
'Call BackBufferSurfaceBltColorFill r2, vbBlack
Call BackBufferSurface.BltColorFill(r2, vbBlack)
  '  If .pjs(Index).muerto = 0 Then
        'Dibuja Cuerpo
      '  Grh = BodyData(.pjs(Index).Body).Walk(3)
       ' Call DrawGrh(Grh, 12, 15)
        'Dibuja Cabeza
       ' Pos = BodyData(.pjs(Index).Body).HeadOffset.Y + GrhData(GrhData(Grh.GrhIndex).Frames(1)).pixelHeight
      '  Grh = HeadData(.pjs(Index).Head).Head(3)
      '  Call DrawGrh(Grh, 17, Pos)
        'Dibuja Casco
       ' If .pjs(Index).Casco <> 2 Then
      '      Grh = CascoAnimData(.pjs(Index).Casco).Head(3)
      '      Call DrawGrh(Grh, 17, Pos)
      '  End If
        'Dibuja Arma
      '  If .pjs(Index).Arma <> 2 Then
      '      Grh = WeaponAnimData(.pjs(Index).Arma).WeaponWalk(3)
      '      Call DrawGrh(Grh, 12, 15)
      '  End If
        'Dibuja Escudo
       ' If .pjs(Index).Escu <> 2 Then
      '      Grh = ShieldAnimData(.pjs(Index).Escu).ShieldWalk(3)
      '      Call DrawGrh(Grh, 12, 17)
     '   End If
   ' Else
        'Dibuja Cuerpo muerto
  '      Grh = BodyData(8).Walk(3)
   '     Call DrawGrh(Grh, 12, 15)
   '     'Dibuja Cabeza muerta
   '     Pos = BodyData(8).HeadOffset.Y + GrhData(GrhData(Grh.GrhIndex).Frames(1)).pixelHeight
   '     Grh = HeadData(500).Head(3)
   '     Call DrawGrh(Grh, 18, 3)
 '   End If
    FrmCuenta.PlayerView.Refresh
   End With
End Sub
