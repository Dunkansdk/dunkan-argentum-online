Attribute VB_Name = "Direct3D"
Option Explicit

' Dunkan: Estoy hay que sacarlo a la mierda.

Public idx As Long 'current material index for 2d mapping and material editor


Public usematcoloronly As Long
Public usematcolorandtexture As Long
Public usetextureonly As Long
Public usetextureandtransparency As Long

Public ScreenShotSurface As Direct3DSurface8
Public SrcPalette As PALETTEENTRY
Public SrcRect As RECT

Public PDirect3DX As New D3DX8
Public PDirectX As New DirectX8
Public PDirect3D As Direct3D8
Public PDevice As Direct3DDevice8
Public DisplaySettings As D3DDISPLAYMODE
Public PStructParameter As D3DPRESENT_PARAMETERS
'Const PD3DFVF = (D3DFVF_NORMAL Or D3DFVF_TEX1 Or D3DFVF_XYZ)


Public Const PD3DFVF = D3DFVF_XYZ Or D3DFVF_TEX1

Public MatProj As D3DMATRIX
Public RotMatrix As D3DMATRIX
Public TranMatrix As D3DMATRIX
Public ScaleMatrix As D3DMATRIX
Public matWorld As D3DMATRIX
Public MatView As D3DMATRIX
Public PLight As D3DLIGHT8

Public tempvertex(2) As TLVERTEX 'D3DVERTEX   'triangle temporaire

'len(tempvertex)= 32 : 4 * 8 single in memory

Public RenduD3d As Boolean
'Public ptexture() As Direct3DTexture8

Global Pcouleur As D3DMATERIAL8

Public vFrom As D3DVECTOR
Public vTo As D3DVECTOR
Public vUp As D3DVECTOR
Public LastX! ' ! = As Single
Public LastY!
Public RotX!
Public RotY!
Public RotZ!

Public ProjMatrix As D3DMATRIX
Public PasRotation As Single
Public PasMouvement As Single
Public PasZoom As Single

Public PivotObjetX As Single
Public PivotObjetY As Single
Public PivotObjetZ As Single
Public offsetz As Single
Public Sub INITRenduDirect3d()


    Set PDirect3D = PDirectX.Direct3DCreate()
    PDirect3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DisplaySettings
    
'type de rendu de textures
'    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTFN_POINT '0
'    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTFN_LINEAR '1
'    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTFN_ANISOTROPIC '2
 
'    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_POINT
'    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_LINEAR
'    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_FLATCUBIC
'    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_GAUSSIANCUBIC
'    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTFG_ANISOTROPIC

        'model render type
'    D3DDevice.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_POINT
'    D3DDevice.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_WIREFRAME
'    D3DDevice.SetRenderState D3DRENDERSTATE_FILLMODE, D3DFILL_SOLID


    With PStructParameter
        .BackBufferCount = 1
        .AutoDepthStencilFormat = CheckZBuffer(DisplaySettings)
        .EnableAutoDepthStencil = 1
        .BackBufferFormat = D3DFMT_X8R8G8B8 'DisplaySettings.Format
        .hDeviceWindow = frmMain.MainViewPic.hWnd
        .SwapEffect = D3DSWAPEFFECT_FLIP
        .Windowed = 1

    End With

    Set PDevice = PDirect3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.MainViewPic.hWnd, CheckHardwareTL(), PStructParameter)
   
    Set ScreenShotSurface = PDevice.CreateImageSurface(DisplaySettings.Width, DisplaySettings.Height, D3DFMT_A8R8G8B8)

    PDevice.SetVertexShader PD3DFVF
   
    PDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
    ' Turn on the zbuffer
    PDevice.SetRenderState D3DRS_ZENABLE, 1
    ' Turn off lighting
    PDevice.SetRenderState D3DRS_LIGHTING, 1
'   MsgBox 3.141592654 / 3

          


D3DXMatrixPerspectiveFovLH ProjMatrix, 3.141592654 / 3#, 1#, 1#, 1000000#
PDevice.SetTransform D3DTS_PROJECTION, ProjMatrix
    vFrom.X = 0
    vFrom.Y = 0
    'auto zoom our scene to fit best size
    offsetz = (Zmax + Zmin)
    If offsetz = 0 Then
    offsetz = (-Zmax / 2)    'zmax = -zmin -> the model is symetrical on Z
    Else
    offsetz = -(Zmax - Zmin) 'zmax <> -zmin -> the model is not symetrical on Z
    End If
   
    
    vFrom.Z = offsetz


    vTo.X = 0
    vTo.Y = 0
    vTo.Z = 0

    vUp.X = 16000
    vUp.Y = 16000
    vUp.Z = 0

    RotX = 0
    RotY = 0
    RotZ = 0
    
Pcouleur.Ambient.a = 1
Pcouleur.Ambient.R = 1 '255
Pcouleur.Ambient.G = 1 '200
Pcouleur.Ambient.b = 1 '255
    
'    PLight.type = D3DLIGHT_DIRECTIONAL
'
'    PLight.Ambient.R = 1#
'    PLight.Ambient.G = 1#
'    PLight.Ambient.b = 1#
'
'
'    PLight.diffuse.R = 1#
'    PLight.diffuse.G = 1#
'    PLight.diffuse.b = 1#

'    PLight.position.x = 0
'    PLight.position.y = 0
'    PLight.position.Z = -1000
'
'    PLight.Direction.x = 1000
'    PLight.Direction.y = 0
'    PLight.Direction.Z = 0

'    PLight.Phi = 3.1415

'    PLight.Range = 1000#
'    PLight.Attenuation0 = 1#
'
'    PDevice.SetLight 0, PLight
'    PDevice.LightEnable 0, 1

    D3DXMatrixIdentity matWorld
    D3DXMatrixIdentity RotMatrix
    D3DXMatrixIdentity ScaleMatrix
    D3DXMatrixIdentity TranMatrix

    RenduD3d = True

End Sub





Public Sub rendu()
    On Error Resume Next

Dim i As Long, j As Long, k As Long
Dim TempMatrix As D3DMATRIX
Dim point1 As Long, point2 As Long, point3 As Long
Dim pcount As Long

If PDevice Is Nothing Then Exit Sub


'material color

    Pcouleur.power = 0.5
    Pcouleur.diffuse.R = 1#
    Pcouleur.diffuse.G = 1#
    Pcouleur.diffuse.b = 1#
    Pcouleur.Ambient.R = 1#
    Pcouleur.Ambient.G = 1#
    Pcouleur.Ambient.b = 1#
    Pcouleur.Specular.R = 1#
    Pcouleur.Specular.G = 1#
    Pcouleur.Specular.b = 1#



    'xyz axis fix for first view (center object at Origin(0,0,0) )
     vFrom.X = vFrom.X + (Xmax - (-Xmin)) '/ 2    'offset x
     vFrom.Y = vFrom.Y + (Ymax - (-Ymin)) '/ 2   'offset y
     vFrom.Z = vFrom.Z - (Zmax - Zmin) * 2       'offset z

    Do While RenduD3d = True
'    DoEvents
        vTo.X = vFrom.X
        vTo.Y = vFrom.Y '+ 10
        vTo.Z = vFrom.Z + 1000
        
'        Form1.Label2.Caption = vFrom.z
        
'    PLight.position.X = vFrom.X
'    PLight.position.Y = vFrom.Y
'    PLight.position.Z = vFrom.Z
'
'    PLight.Direction.X = vTo.X
'    PLight.Direction.Y = vTo.Y
'    PLight.Direction.Z = vTo.Z
        
    'here is our camera view , we can rotate and translate around our model by moving and rotating this camera


'translation
'D3DXMatrixTranslation TempMatrix, 0.001, 0, 0
'D3DXMatrixMultiply TranMatrix, TranMatrix, TempMatrix

'scaling
'D3DXMatrixScaling TempMatrix, 0.5, 0.5, 0.5
'D3DXMatrixMultiply ScaleMatrix, ScaleMatrix, TempMatrix
    PDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HBBCCFF, 1000000#, 1
    
    PDevice.BeginScene



    D3DXMatrixLookAtLH MatView, vFrom, vTo, vUp
    
    'rotation for fixing the default view (get front /perspective view) we setup the world angles here :
    D3DXMatrixRotationX TempMatrix, 0
    D3DXMatrixRotationY TempMatrix, 0.8
    D3DXMatrixRotationZ TempMatrix, 0.8
    
    D3DXMatrixMultiply MatView, MatView, TempMatrix

    PDevice.SetTransform D3DTS_VIEW, MatView
 
    PDevice.SetRenderState D3DRS_ZENABLE, 1
    PDevice.SetRenderState D3DRS_CULLMODE, 1
    
'    'FUCKING LIGHTS DONT WORKS :o/
'     PDevice.SetLight 0, PLight
'     PDevice.LightEnable 0, 1




    PDevice.SetRenderState D3DRS_AMBIENT, &H50FFFFFF
'    PDevice.SetRenderState D3DRS_LIGHTING, 1


        'It does, so turn alpha blending on
 
     
     
         For i = 0 To UBound(ObjetVertex)
         DoEvents
                        For j = 0 To ObjetVertex(i).nbrFace - 1
                                point1 = ObjetVertex(i).VertexFace(j).a
                                point2 = ObjetVertex(i).VertexFace(j).b
                                point3 = ObjetVertex(i).VertexFace(j).C
                                
                                
                                'create a temporary triangle and offset it to object Origine axis (not world origine axis)
                                tempvertex(0).X = ObjetVertex(i).tVertex(point1).X '+ ObjetVertex(i).locale.Oaxis.x '+ ObjetVertex(i).locale.Xaxis.X
                                tempvertex(0).Y = ObjetVertex(i).tVertex(point1).Y '+ ObjetVertex(i).locale.Oaxis.y '+ ObjetVertex(i).locale.Xaxis.Y
                                tempvertex(0).Z = ObjetVertex(i).tVertex(point1).Z  '+ ObjetVertex(i).locale.Oaxis.z '+ ObjetVertex(i).locale.Xaxis.Z

                                
    
                                tempvertex(1).X = ObjetVertex(i).tVertex(point2).X '+ ObjetVertex(i).locale.Oaxis.x '+ ObjetVertex(i).locale.Xaxis.X
                                tempvertex(1).Y = ObjetVertex(i).tVertex(point2).Y '+ ObjetVertex(i).locale.Oaxis.y '+ ObjetVertex(i).locale.Xaxis.Y
                                tempvertex(1).Z = ObjetVertex(i).tVertex(point2).Z  '+ ObjetVertex(i).locale.Oaxis.z '+ ObjetVertex(i).locale.Xaxis.Z


                                tempvertex(2).X = ObjetVertex(i).tVertex(point3).X '+ ObjetVertex(i).locale.Oaxis.x '+ ObjetVertex(i).locale.Xaxis.X
                                tempvertex(2).Y = ObjetVertex(i).tVertex(point3).Y '+ ObjetVertex(i).locale.Oaxis.y '+ ObjetVertex(i).locale.Xaxis.Y
                                tempvertex(2).Z = ObjetVertex(i).tVertex(point3).Z  '+ ObjetVertex(i).locale.Oaxis.z '+ ObjetVertex(i).locale.Xaxis.Z

                                
                                If ObjetVertex(i).bmapped = True Then
                                    tempvertex(0).TU = ObjetVertex(i).mapping(point1).TU
                                    tempvertex(0).TV = 1 - ObjetVertex(i).mapping(point1).TV
                                    tempvertex(1).TU = ObjetVertex(i).mapping(point2).TU
                                    tempvertex(1).TV = 1 - ObjetVertex(i).mapping(point2).TV
                                    tempvertex(2).TU = ObjetVertex(i).mapping(point3).TU
                                    tempvertex(2).TV = 1 - ObjetVertex(i).mapping(point3).TV
                                End If
                                


                                
                         'material color only
                         If usematcoloronly = 1 Then
                            PDevice.SetRenderState D3DRS_LIGHTING, 1
                            PDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
                            PDevice.SetMaterial Materials3dsarray(ObjetVertex(i).VertexFaceMaterialindex(j)).materialD3D
                            PDevice.SetTexture 0, Nothing
                        End If
                         
                         
                         'material color + texture
                         If usematcolorandtexture = 1 Then
                            PDevice.SetRenderState D3DRS_LIGHTING, 1
                            PDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
                            PDevice.SetMaterial Materials3dsarray(ObjetVertex(i).VertexFaceMaterialindex(j)).materialD3D
                            PDevice.SetTexture 0, ptexture(ObjetVertex(i).VertexFaceMaterialindex(j))
                        End If
                         
                         
                         'texture only
                         If usetextureonly = 1 Then
                            PDevice.SetRenderState D3DRS_LIGHTING, 1
                            PDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
                            PDevice.SetMaterial Pcouleur
                            PDevice.SetTexture 0, ptexture(ObjetVertex(i).VertexFaceMaterialindex(j))
                         End If
                         
                         
                         'texture + transparent color
                        If usetextureandtransparency = 1 Then
                            PDevice.SetRenderState D3DRS_LIGHTING, 0
                            PDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
                            PDevice.SetMaterial Pcouleur
                            PDevice.SetTexture 0, ptexture(ObjetVertex(i).VertexFaceMaterialindex(j))
                            'State for just drawing transparent color-keyed without alpha-blending
                            Call PDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
                            Call PDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)

                         End If
                            
                        PDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, tempvertex(0), Len(tempvertex(0))
                       Next
                     
        Next
        ' Add Anti-Aliasing for textures
        PDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, 2
        PDevice.SetTextureStageState 0, D3DTSS_MINFILTER, 2

        
        
        'on remet la scene correctement
        Rotate RotX, RotY + 1.6, RotZ

        D3DXMatrixMultiply matWorld, RotMatrix, TranMatrix
'        D3DXMatrixMultiply matWorld, ScaleMatrix, matWorld
        PDevice.SetTransform D3DTS_WORLD, matWorld
          


        PDevice.EndScene
        PDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    Loop

End Sub
Public Sub AdjustScale(sx As Single, sy As Single, sz As Single)

    With ScaleMatrix
        .m11 = sx
        .m22 = sy
        .m33 = sz
    End With

End Sub
 
'Public Sub SetPosition(tranX As Single, tranY As Single, tranZ As Single)
''    TranslateMatrix TranMatrix, MakeVector(tranX, tranY, tranZ)
'End Sub

'Private Sub prv_SetWorldMatrix()
'    D3DXMatrixIdentity mWorldMatrix
'    D3DXMatrixMultiply mWorldMatrix, mRotMatrix, mTranMatrix
'    D3DXMatrixMultiply mWorldMatrix, mScaleMatrix, mWorldMatrix
'End Sub

Sub SetSize(ByVal sx As Single, ByVal sy As Single, ByVal sz As Single)
'Dim TempMatrix As D3DMATRIX

    ScaleMatrix.m11 = sx
    ScaleMatrix.m22 = sy
    ScaleMatrix.m33 = sz
    '-------------------------------------------
'    prv_SetWorldMatrix
    '-------------------------------------------
End Sub

Sub SetPosition(ByVal sx As Single, ByVal sy As Single, ByVal sz As Single)
    TranMatrix.m41 = sx
    TranMatrix.m42 = sy
    TranMatrix.m43 = sz
    '-------------------------------------------
'    prv_SetWorldMatrix
    '-------------------------------------------
End Sub



Public Sub Rotate(ByVal RotX1!, ByVal RotY1!, ByVal RotZ1!)
Dim TempMatrix As D3DMATRIX
    D3DXMatrixIdentity RotMatrix
    D3DXMatrixRotationX TempMatrix, RotX1
    D3DXMatrixMultiply RotMatrix, TempMatrix, RotMatrix
    D3DXMatrixRotationY TempMatrix, RotY1
    D3DXMatrixMultiply RotMatrix, TempMatrix, RotMatrix
    D3DXMatrixRotationZ TempMatrix, RotZ1
    D3DXMatrixMultiply RotMatrix, TempMatrix, RotMatrix
End Sub



'    Dim light As D3DLIGHT8
'    With light
'    .type = D3DLIGHT_POINT
'    .Ambient = CreateD3DColorVal(1, 1, 1, 1)
'    .diffuse = CreateD3DColorVal(1, 1, 1, 1)
'    .specular = CreateD3DColorVal(1, 1, 1, 1)
'    .position = MakeVector(0, MD3Model.BoxMax.y + 5, 0)
'    .Attenuation0 = 0
'    .Attenuation1 = 0.1
'    .Attenuation2 = 0
'    .Range = 64
'    End With
'    D3Ddevice.SetLight 0, light
'    D3Ddevice.LightEnable 0, True
    

'Private Function creationpoint(X As Single, Y As Single, z As Single, tu As Single, tv As Single) As PVertex
'creationpoint.X = X
'creationpoint.Y = Y
'creationpoint.z = z
'creationpoint.tu = tu
'creationpoint.tv = tv
'End Function
'
'Public Function Quaternion(Axe As D3DVECTOR, Angle As Single) As D3DQUATERNION
'  With Quaternion
'    .X = Sin(Angle / 2) * Axe.X
'    .Y = Sin(Angle / 2) * Axe.Y
'    .z = Sin(Angle / 2) * Axe.z
'    .w = Cos(Angle / 2)
'  End With
'End Function

Private Function CheckZBuffer(Mode As D3DDISPLAYMODE) As Long
    If PDirect3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
    CheckZBuffer = D3DFMT_D16
    End If
    
    If PDirect3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16_LOCKABLE) = D3D_OK Then
    CheckZBuffer = D3DFMT_D16_LOCKABLE
    End If
    
    If PDirect3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24S8) = D3D_OK Then
    CheckZBuffer = D3DFMT_D24S8
    End If
    
    If PDirect3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24X4S4) = D3D_OK Then
    CheckZBuffer = D3DFMT_D24X4S4
    End If
    
    If PDirect3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
    CheckZBuffer = D3DFMT_D24X8
    End If
    
    If PDirect3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Mode.format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
    CheckZBuffer = D3DFMT_D32
    End If
End Function

Private Function CheckHardwareTL() As Long
Dim DevCaps As D3DCAPS8
    PDirect3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DevCaps

    If (DevCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
        CheckHardwareTL = D3DCREATE_HARDWARE_VERTEXPROCESSING
    Else
        CheckHardwareTL = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    End If
End Function


Public Sub LoadMatFile()
'load materials from a binary file directly
Dim matcount As Long, i As Long
idx = 0
If Dir(file3dsname & filetitle & ".mat") <> "" Then
          Open file3dsname & filetitle & ".mat" For Binary Access Read As #2
            Get #2, , matcount
           
            If matcount <= 0 Then
            Close #2
            Exit Sub
            End If
            nbmaterials = matcount
            ReDim Materials3dsarray(matcount - 1)
            
            Get #2, , Materials3dsarray
          Close #2
          
        If matcount > 0 Then
            RenduD3d = False
                ReDim ptexture(matcount - 1)
        
                    For i = 0 To matcount - 1

                            If Materials3dsarray(i).Diffuse_Texture_FileName <> "" Then
                            
                                If Dir(Materials3dsarray(i).Diffuse_Texture_FileName) <> "" Then
                                    Set ptexture(i) = Nothing
'                                    Set ptexture(i) = PDirect3DX.CreateTextureFromFile(PDevice, Materials3dsarray(i).Diffuse_Texture_FileName)
'                                   Set ptexture(i) = PDirect3DX.CreateTextureFromFileEx(PDevice, Materials3dsarray(i).Diffuse_Texture_FileName, ByVal 0, ByVal 0, D3DX_DEFAULT, _
'             0, DisplaySettings.Format, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_DITHER, &HFFFF00FF, ByVal 0, ByVal 0)
                                    Set ptexture(i) = PDirect3DX.CreateTextureFromFileEx(PDevice, Materials3dsarray(i).Diffuse_Texture_FileName, D3DX_DEFAULT, D3DX_DEFAULT, _
                                    1&, 0&, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFF00FF, ByVal 0, ByVal 0)
                            
                                End If
                            End If
                    Next

                    idx = 0
               
                    RenduD3d = True
          End If
End If

End Sub


