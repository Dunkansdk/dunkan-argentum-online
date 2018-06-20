Attribute VB_Name = "Mod_DX8_Test3D"
Option Explicit

Dim lnm As Long

'3 - How to read chunks ?
'------------------------
'This is the representation of a chunk :
'
'Offset   Length    Name
'
'0        2         Chunk-ID
'2        4         Chunk-length = 6+n+m
'6        n Data
'6+n      m         Sub-chunks


Public Type RGB 'rgb color for materails (ambiant,diffuse and specular
        R As Byte
        G As Byte
        b As Byte
End Type

Public Type Materialpercentinfos  'struct infos for reading percent values (see UDT Materail3ds)
            nextchunk As Integer
            matcolorlen As Long
            Percent As Byte
            SubChunkID As Integer
End Type


Public Type Materialcolorinfos     'struct infos for reading rgbcolor and texture filenames (see UDT Materail3ds)
            datatype As Integer     'rbg bytes or rgbsingles ?
            matcolorlen As Long     'len of chunk
            MatRGB As RGB           'rgbcolor byte format
            SubChunkID As Integer
End Type

Public ptexture() As Direct3DTexture8     'Direct3DBaseTexture8

Public Type Material3ds
    Material_Name As String
    Ambiant_Color As Materialcolorinfos  'default face color in .x file aka Material.Power
    Ambiant_Texture_FileName As String
    
    Diffuse_Color As Materialcolorinfos  'diffuse color in .x file
    Diffuse_Texture_FileName As String
    
    Specular_Color As Materialcolorinfos 'specular color in .x file
    Specular_Texture_FileName As String
    
    Shininess_Percent As Materialpercentinfos
    Shininess_Strength_Percent As Materialpercentinfos
    Opacity_Percent As Materialpercentinfos
    Opacity_FallOff_Percent As Materialpercentinfos
    Reflexion_Blur As Materialpercentinfos
    TwoSided As Boolean
    Self_Illumination_Percent As Materialpercentinfos    'equal to .x material power or emmisive ?
    Wired  As Boolean
    FaceMap As Boolean
    materialD3D As D3DMATERIAL8
End Type

Public Type TypeBloc
    Entete As Integer
    Longeur As Long
End Type


Public Type tVertex '= invert = D3DVECTOR
    X As Single
    Y As Single 'z
    Z As Single 'y
End Type

Public Type mapcoor
    TU As Single
    TV As Single
End Type

Public Type Face
        a  As Integer
        b  As Integer
        C  As Integer
        IV4 As Integer
End Type

Public Type Localinfos
        Xaxis As D3DVECTOR
        Yaxis As D3DVECTOR
        Zaxis As D3DVECTOR
        Oaxis As D3DVECTOR
End Type

Public Type tObjet
    nbrVertex As Long
    tVertex() As tVertex 'D3DVECTOR
    VertexFace() As Face
    VertexFaceMaterialindex() As Integer  'material index 0 ou 1 ou ... ou nb material
    nbrFace As Long
    mapping() As mapcoor
    bmapped As Boolean
    nbmappingpoint As Long
    locale As Localinfos
End Type

'scene materials array
Public Materials3dsarray() As Material3ds
Public nbmaterials As Long


' Premier bloc
Public Const MAIN3DS                  As Long = &H4D4D  'M3DMAGIC '*3DS file*'
Public Const NULL_CHUNK               As Long = &H0
  Public Const SMAGIC                 As Long = &H2D2D
  Public Const LMAGIC                 As Long = &H2D3D
  Public Const MLIBMAGIC              As Long = &H3DAA   '*MLI file*'
  Public Const MATMAGIC               As Long = &H3DFF
  Public Const CMAGIC                 As Long = &HC23D   '*PRJ file*'
  Public Const M3D_VERSION            As Long = &H2
  Public Const M3D_KFVERSION          As Long = &H5
  Public Const MESH_VERSION           As Long = &H3D3E

' Bloc principal
Public Const EDIT3DS                  As Long = &H3D3D  'MDATA  : Début de la configuration de l'element 3d et de l'editeur
Public Const EDIT_CONFIG1             As Long = &H100   'MASTER_SCALE : world scale
'Public Const EDIT_CONFIG2           As Long = &H3E3D   '????

'*************************************************
Public Const BACKGROUND_BITMAP      As Long = &H1100 'BIT_MAP
Public Const USE_BIT_MAP            As Long = &H1101

Public Const BACKGROUND_COLOR       As Long = &H1200 'SOLID_BGND
Public Const USE_BACKG_COLOR        As Long = &H1201 'USE_SOLID_BGND

Public Const V_GRADIENT             As Long = &H1300
Public Const USE_V_GRADIENT         As Long = &H1301

Public Const LO_SHADOW_BIAS         As Long = &H1400
Public Const HI_SHADOW_BIAS         As Long = &H1410
Public Const SHADOW_MAP_SIZE        As Long = &H1420
Public Const SHADOW_SAMPLES         As Long = &H1430
Public Const SHADOW_RANGE           As Long = &H1440
Public Const SHADOW_FILTER          As Long = &H1450
Public Const RAY_BIAS               As Long = &H1460

Public Const O_CONSTS               As Long = &H1500
  

Public Const EDIT_AMBIENT           As Long = &H2100 'AMBIENT_LIGHT

'***************************************************
Public Const EDIT_UNKNW09           As Long = &H2201
Public Const EDIT_UNKNW13           As Long = &H2000

Public Const FOG                    As Long = &H2200
Public Const FOG_BGND               As Long = &H2210
Public Const LAYER_FOG              As Long = &H2302
Public Const DISTANCE_CUE           As Long = &H2300
Public Const DCUE_BGND              As Long = &H2310
Public Const USE_FOG                As Long = &H2201
Public Const USE_LAYER_FOG          As Long = &H2303
Public Const USE_DISTANCE_CUE       As Long = &H2301


Public Const Material_Name          As Long = &HA000 'MAT_NAME
Public Const EDIT_MATERIAL          As Long = &HAFFF 'MAT_ENTRY
Public Const MAT_AMBIENT            As Long = &HA010
Public Const MAT_DIFFUSE            As Long = &HA020
Public Const MAT_SPECULAR           As Long = &HA030
Public Const MAT_SHININESS          As Long = &HA040
Public Const MAT_SHIN2PCT           As Long = &HA041
Public Const MAT_TRANSPARENCY       As Long = &HA050
Public Const MAT_XPFALL             As Long = &HA052
Public Const MAT_USE_XPFALL         As Long = &HA240 'boolean
Public Const MAT_REFBLUR            As Long = &HA053 'boolean
Public Const MAT_SHADING            As Long = &HA100 'boolean
Public Const MAT_USE_REFBLUR        As Long = &HA250 'boolean
Public Const MAT_SELF_ILLUM         As Long = &HA080 'boolean
Public Const MAT_TWO_SIDE           As Long = &HA081 'boolean
Public Const MAT_DECAL              As Long = &HA082 'boolean
Public Const MAT_ADDITIVE           As Long = &HA083
Public Const MAT_WIRE               As Long = &HA085 'boolean
Public Const MAT_FACEMAP            As Long = &HA088 'boolean
Public Const MAT_PHONGSOFT          As Long = &HA08C 'boolean
Public Const MAT_WIREABS            As Long = &HA08E 'boolean
Public Const MAT_WIRE_SIZE          As Long = &HA087 'single
Public Const MAT_TEXMAP             As Long = &HA200

Public Const MAT_SXP_TEXT_DATA      As Long = &HA320
Public Const MAT_TEXMASK            As Long = &HA33E
Public Const MAT_SXP_TEXTMASK_DATA  As Long = &HA32A
Public Const MAT_TEX2MAP            As Long = &HA33A
Public Const MAT_SXP_TEXT2_DATA     As Long = &HA321
Public Const MAT_TEX2MASK           As Long = &HA340
Public Const MAT_SXP_TEXT2MASK_DATA As Long = &HA32C
Public Const MAT_OPACMAP            As Long = &HA210
Public Const MAT_SXP_OPAC_DATA      As Long = &HA322
Public Const MAT_OPACMASK           As Long = &HA342
Public Const MAT_SXP_OPACMASK_DATA  As Long = &HA32E
Public Const MAT_BUMPMAP            As Long = &HA230
Public Const MAT_SXP_BUMP_DATA      As Long = &HA324
Public Const MAT_BUMPMASK           As Long = &HA344
Public Const MAT_SXP_BUMPMASK_DATA  As Long = &HA330
Public Const MAT_SPECMAP            As Long = &HA204
Public Const MAT_SXP_SPEC_DATA      As Long = &HA325
Public Const MAT_SPECMASK           As Long = &HA348
Public Const MAT_SXP_SPECMASK_DATA  As Long = &HA332
Public Const MAT_SHINMAP            As Long = &HA33C
Public Const MAT_SXP_SHIN_DATA      As Long = &HA326
Public Const MAT_SHINMASK           As Long = &HA346
Public Const MAT_SXP_SHINMASK_DATA  As Long = &HA334
Public Const MAT_SELFIMAP           As Long = &HA33D
Public Const MAT_SXP_SELFI_DATA     As Long = &HA328
Public Const MAT_SELFIMASK          As Long = &HA34A
Public Const MAT_SXP_SELFIMASK_DATA As Long = &HA336
Public Const MAT_REFLMAP            As Long = &HA220
Public Const MAT_REFLMASK           As Long = &HA34C
Public Const MAT_SXP_REFLMASK_DATA  As Long = &HA338
Public Const MAT_ACUBIC             As Long = &HA310
Public Const MAT_MAPNAME            As Long = &HA300  'file for texture (BMP or other)
Public Const MAT_MAP_TILING         As Long = &HA351
Public Const MAT_MAP_TEXBLUR        As Long = &HA353
Public Const MAT_MAP_USCALE         As Long = &HA354
Public Const MAT_MAP_VSCALE         As Long = &HA356
Public Const MAT_MAP_UOFFSET        As Long = &HA358
Public Const MAT_MAP_VOFFSET        As Long = &HA35A
Public Const MAT_MAP_ANG            As Long = &HA35C
Public Const MAT_MAP_COL1           As Long = &HA360
Public Const MAT_MAP_COL2           As Long = &HA362
Public Const MAT_MAP_RCOL           As Long = &HA364
Public Const MAT_MAP_GCOL           As Long = &HA366
Public Const MAT_MAP_BCOL           As Long = &HA368



Public Const EDIT_OBJECT             As Long = &H4000 'NAMED_OBJECT

Public Const OBJ_TRIMESH             As Long = &H4100
Public Const OBJECT_HIDDEN           As Long = &H4010
Public Const OBJ_VIS_LOFTER          As Long = &H4011
Public Const OBJECT_CAST             As Long = &H4012 'OBJ_DOESNT_CAST
Public Const OBJECT_MATT             As Long = &H4013 'OBJ_MATTE
Public Const OBJ_FAST                As Long = &H4014
Public Const OBJECT_EXTP             As Long = &H4015 'OBJ_PROCEDURAL
Public Const OBJ_FROZEN              As Long = &H4016
Public Const OBJECT_SHAD             As Long = &H4017 'OBJ_DONT_RECVSHADOW

'OBJ_TRIMESH
Public Const TRI_VERTEXL            As Long = &H4110 'POINT_ARRAY
Public Const TRI_VERTEXOPTIONS      As Long = &H4111 'POINT_FLAG_ARRAY
Public Const TRI_FACEL1             As Long = &H4120 'FACE_ARRAY
Public Const TRI_MATERIAL           As Long = &H4130 'MSH_MAT_GROUP
Public Const TRI_MAPPINGCOORS       As Long = &H4140 'TEX_VERTS
Public Const TRI_SMOOTH             As Long = &H4150 ' SMOOTH_GROUP
Public Const TRI_LOCAL              As Long = &H4160 'MESH_MATRIX
Public Const TRI_VISIBLE            As Long = &H4165 'MESH_COLOR
Public Const TRI_MAPPINGSTANDARD    As Long = &H4170 'MESH_TEXTURE_INFO
Public Const MSH_BOXMAP             As Long = &H4190 'BoundingBox


'définition de OBJ_CAMERA
Public Const OBJ_CAMERA              As Long = &H4700
Public Const CAM_SEE_CONE            As Long = &H4710
Public Const CAM_RANGES              As Long = &H4720

'définition de OBJ_LIGHT
Public Const OBJ_LIGHT               As Long = &H4600 'N_DIRECT_LIGHT
Public Const LIT_SPOTLIGHT           As Long = &H4610
Public Const LIT_OFF                 As Long = &H4620
Public Const LIT_ATTENUATE           As Long = &H4625
Public Const LIT_RAYSHAD             As Long = &H4627
Public Const LIT_SHADOWED            As Long = &H4630
Public Const LIT_LOCAL_SHADOW2       As Long = &H4641
Public Const LIT_SEE_CONE            As Long = &H4650
Public Const LIT_SPOT_RECTANGULAR    As Long = &H4651
Public Const LIT_SPOT_OVERSHOOT      As Long = &H4652
Public Const LIT_SPOT_PROJECTOR      As Long = &H4653
Public Const LIT_EXCLUDE             As Long = &H4654
Public Const LIT_SPOT_ROLL           As Long = &H4656
Public Const LIT_SPOT_ASPECT         As Long = &H4657
Public Const LIT_RAY_BIAS            As Long = &H4658
Public Const LIT_INNER_RANGE         As Long = &H4659
Public Const LIT_OUTER_RANGE         As Long = &H465A
Public Const LIT_MULTIPLIER          As Long = &H465B
 

'définition de KEYF3DS
Public Const KEYF3DS                As Long = &HB000  ' KFDATA :Debut de la configuration des frames clé
Public Const AMBIENT_NODE_TAG       As Long = &HB001
Public Const KEYF_OBJDES            As Long = &HB002
Public Const CAMERA_NODE_TAG        As Long = &HB003
Public Const TARGET_NODE_TAG        As Long = &HB004
Public Const LIGHT_NODE_TAG         As Long = &HB005
Public Const L_TARGET_NODE_TAG      As Long = &HB006
Public Const SPOTLIGHT_NODE_TAG     As Long = &HB007
Public Const KEYF_FRAMES            As Long = &HB008
Public Const KEYF_UNKNWN01          As Long = &HB009
Public Const KFHDR                  As Long = &HB00A
Public Const NODE_HDR               As Long = &HB010
Public Const INSTANCE_NAME          As Long = &HB011
Public Const PIVOT                  As Long = &HB013
Public Const BOUNDBOX               As Long = &HB014
Public Const MORPH_SMOOTH           As Long = &HB015
Public Const POS_TRACK_TAG          As Long = &HB020
Public Const ROT_TRACK_TAG          As Long = &HB021
Public Const SCL_TRACK_TAG          As Long = &HB022
Public Const FOV_TRACK_TAG          As Long = &HB023
Public Const ROLL_TRACK_TAG         As Long = &HB024
Public Const COL_TRACK_TAG          As Long = &HB025
Public Const MORPH_TRACK_TAG        As Long = &HB026
Public Const HOT_TRACK_TAG          As Long = &HB027
Public Const FALL_TRACK_TAG         As Long = &HB028
Public Const HIDE_TRACK_TAG         As Long = &HB029
Public Const NODE_ID                As Long = &HB030



' Ceci défini les différent bloc de type couleur.
Public Const COL_RGB                As Long = &H10 'COLOR_F
Public Const COL_TRU                As Long = &H11 'COLOR_24
Public Const LIN_COLOR_24           As Long = &H12
Public Const LIN_COLOR_F            As Long = &H13

Public Const INT_PERCENTAGE         As Long = &H30
Public Const FLOAT_PERCENTAGE       As Long = &H31

'Définition des block pour l'affichage des fenetre
Public Const WINDOW_TOP                As Long = &H1
Public Const WINDOW_BOTTOM             As Long = &H2
Public Const WINDOW_LEFT               As Long = &H3
Public Const WINDOW_RIGHT              As Long = &H4
Public Const WINDOW_FRONT              As Long = &H5
Public Const WINDOW_BACK               As Long = &H6
Public Const WINDOW_USER               As Long = &H7
Public Const WINDOW_CAMERA             As Long = &HFFFF
Public Const WINDOW_LIGHT              As Long = &H9
Public Const WINDOW_DISABLED           As Long = &H10
Public Const WINDOW_BOGUS              As Long = &H11


Public Const Version_3DS            As Long = &H2

Public Const DEFAULT_VIEW           As Long = &H3000
Public Const VIEW_TOP               As Long = &H3010
Public Const VIEW_BOTTOM            As Long = &H3020
Public Const VIEW_LEFT              As Long = &H3030
Public Const VIEW_RIGHT             As Long = &H3040
Public Const VIEW_FRONT             As Long = &H3050
Public Const VIEW_BACK              As Long = &H3060
Public Const VIEW_USER              As Long = &H3070
Public Const VIEW_CAMERA            As Long = &H3080
Public Const VIEW_WINDOW            As Long = &H3090


Public Const POLY_2D                As Long = &H5000
Public Const SHAPE_OK               As Long = &H5010
Public Const SHAPE_NOT_OK           As Long = &H5011
Public Const SHAPE_HOOK             As Long = &H5020
Public Const PATH_3D                As Long = &H6000
Public Const PATH_MATRIX            As Long = &H6005
Public Const SHAPE_2D               As Long = &H6010
Public Const M_SCALE                As Long = &H6020
Public Const M_TWIST                As Long = &H6030
Public Const M_TEETER               As Long = &H6040
Public Const M_FIT                  As Long = &H6050
Public Const M_BEVEL                As Long = &H6060
Public Const XZ_CURVE               As Long = &H6070
Public Const YZ_CURVE               As Long = &H6080
Public Const INTERPCT               As Long = &H6090
Public Const DEFORM_LIMIT           As Long = &H60A0

Public Const USE_CONTOUR            As Long = &H6100
Public Const USE_TWEEN              As Long = &H6110
Public Const USE_SCALE              As Long = &H6120
Public Const USE_TWIST              As Long = &H6130
Public Const USE_TEETER             As Long = &H6140
Public Const USE_FIT                As Long = &H6150
Public Const USE_BEVEL              As Long = &H6160



Public Const VIEWPORT_LAYOUT_OLD    As Long = &H7000
Public Const EDIT_VIEW1             As Long = &H7001 'VIEWPORT_LAYOUT
Public Const VIEWPORT_DATA_OLD      As Long = &H7010
Public Const EDIT_VIEW_P1           As Long = &H7012 'VIEWPORT_DATA_3
Public Const EDIT_VIEW_P2           As Long = &H7011 'VIEWPORT_DATA
Public Const EDIT_VIEW_P3           As Long = &H7020 'VIEWPORT_SIZE
Public Const NETWORK_VIEW           As Long = &H7030

Dim matindex() As Integer 'get index of face in materails
Dim Index As Long
'*******************************
Public FilePos                      As Long ' Position dans la lecture du fichier
Dim NumeroFichier                   As Integer
Public nbrSolide                    As Integer
Public nbrVertex                    As Long
Public ObjetVertex()                As tObjet

Dim iV1 As Integer, IV4 As Integer, a As D3DVECTOR, b As D3DVECTOR, Normale As D3DVECTOR
Public KeyFrameStart As Integer, KeyFrameEnd As Integer

Public Function LireFichier()
Dim l As Long, i As Long
On Error Resume Next ' handle texture coordinates crash on some 3ds files 'fixed ;op
nbmaterials = 0

ReDim Materials3dsarray(nbmaterials)
ReDim ptexture(nbmaterials)
 
NumeroFichier = FreeFile
FilePos = 1
nbrSolide = 0
ReDim Polygone(0)
ReDim ObjetVertex(0)

Open PathFile For Binary As #NumeroFichier
    
    While Not EOF(NumeroFichier)
        DoEvents
        LireBloc FilePos
    Wend
Close #NumeroFichier


Dim TempMinX As Single
Dim TempMaxX As Single
Dim TempMinY As Single
Dim TempMaxY As Single
Dim TempMinZ As Single
Dim TempMaxZ As Single

TempMinX = ObjetVertex(1).tVertex(0).X
TempMaxX = ObjetVertex(1).tVertex(0).X
TempMinY = ObjetVertex(1).tVertex(0).Y
TempMaxY = ObjetVertex(1).tVertex(0).Y
TempMinZ = ObjetVertex(1).tVertex(0).Z
TempMaxZ = ObjetVertex(1).tVertex(0).Z

For l = 1 To UBound(ObjetVertex)
    For i = 0 To UBound(ObjetVertex(l).tVertex)
        If ObjetVertex(l).tVertex(i).X > TempMaxX Then TempMaxX = ObjetVertex(l).tVertex(i).X
        If ObjetVertex(l).tVertex(i).X < TempMinX Then TempMinX = ObjetVertex(l).tVertex(i).X
        If ObjetVertex(l).tVertex(i).Y > TempMaxY Then TempMaxY = ObjetVertex(l).tVertex(i).Y
        If ObjetVertex(l).tVertex(i).Y < TempMinY Then TempMinY = ObjetVertex(l).tVertex(i).Y
        If ObjetVertex(l).tVertex(i).Z > TempMaxZ Then TempMaxZ = ObjetVertex(l).tVertex(i).Z
        If ObjetVertex(l).tVertex(i).Z < TempMinZ Then TempMinZ = ObjetVertex(l).tVertex(i).Z
    Next
Next

'ModeleCenter.x = (TempMinX + TempMaxX) / 2 =offsetx
'ModeleCenter.y = (TempMinY + TempMaxY) / 2=offsety
'ModeleCenter.Z = (TempMinZ + TempMaxZ) / 2= offsetz


Ymax = TempMaxY
Ymin = TempMinY

Xmax = TempMaxX
Xmin = TempMinX

Zmax = TempMaxZ
Zmin = TempMinZ


'create array of textures in memory
'For i = 0 To nbmaterials
'
' If Dir(Materials3dsarray(nbmaterials - 1).Diffuse_Texture_FileName) <> "" Then
'create blank texture
    


' End If
'
'Next

End Function

Public Function LireBloc(ByRef fpos As Long)
    Dim i As Long, h As String
    Dim tBloc As TypeBloc
    Dim TempLong As Long
    Dim TempString As String
    Dim TempInteger As Integer
    Dim TempSingle As Single
    Dim TempChar
    Dim TempByte As Byte
    
    Get #NumeroFichier, fpos, tBloc
    
Select Case tBloc.Entete ' = Hex(tBloc.Entete) give the chunk ID
    
    Case MAIN3DS 'starting reference chunk to start reading skip it
'        Form1.List1.AddItem "Bloc de départ trouvé."
        tBloc.Longeur = 6
    
    Case Version_3DS
        Get #NumeroFichier, , TempInteger
        TempString = CStr(TempInteger)

       
    Case EDIT3DS
'        Form1.List1.AddItem "Donné sur l'editeur trouvé."
        tBloc.Longeur = 6
        
    Case KEYF3DS 'animation
'        Form1.List1.AddItem "Donnée sur l'animation trouvée."
        TempLong = fpos + 6
        While TempLong <> (fpos + tBloc.Longeur)
            LireBloc TempLong
        Wend
        
    Case TRI_MAPPINGCOORS 'coordonnées uv de mapping
            'get the texture coordinates count
            Get #NumeroFichier, , TempInteger
            nbpoint = TempInteger
            ReDim ObjetVertex(nbrSolide).mapping(nbpoint)
             Get #NumeroFichier, , ObjetVertex(nbrSolide).mapping()
            ObjetVertex(nbrSolide).bmapped = True
           
    
    Case TRI_LOCAL 'local coordinates system (le point zero est interessant ici
        Get #NumeroFichier, , ObjetVertex(nbrSolide).locale
          
    
    Case TRI_VISIBLE 'l'objet est visible ?
'        Get #NumeroFichier, , TempByte
         

'    Case &H4710, &H4711
'          Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'          Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'            Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'                      Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'                      Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'                      Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'                      Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte
'                      Get #NumeroFichier, , TempByte
'            Form1.List1.AddItem " " & TempByte

'********************* données sur les materials start**************************************
    Case Material_Name '&HA000  'fixé le 02/02/2005
            nbmaterials = nbmaterials + 1 'add a new material
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        While TempChar <> Chr(0)
            TempString = TempString & TempChar
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        Wend
            Materials3dsarray(nbmaterials - 1).Material_Name = TempString
          
            ReDim Preserve Materials3dsarray(nbmaterials)
            ReDim Preserve ptexture(nbmaterials)
              
              
        TempLong = fpos + 7 + Len(TempString)
        While TempLong <> (fpos + tBloc.Longeur)
            LireBloc TempLong
        Wend
              
              
 Case MAT_MAPNAME 'MAPPING_FILENAME '&HA300 texture file filename (eg: image.JPG ou truc.bmp)
    
'            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
'        While TempChar <> Chr(0)
'            TempString = TempString & TempChar
'            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
'        Wend
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        While TempChar <> Chr(0)
            TempString = TempString & TempChar
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        Wend
        

'           TempString = String(tBloc.Longeur - 6, Chr(0))
'           Get #NumeroFichier, , TempString
           
           Materials3dsarray(nbmaterials - 1).Diffuse_Texture_FileName = TempString
        
              
'        TempLong = fpos + 7 + Len(TempString)
'        While TempLong <> (fpos + tBloc.Longeur)
'            LireBloc TempLong
'        Wend
        
    Case MAT_AMBIENT '&HA010 'Get material ambiant color
        Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Ambiant_Color

            Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.R = Materials3dsarray(nbmaterials - 1).Ambiant_Color.MatRGB.R / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.G = Materials3dsarray(nbmaterials - 1).Ambiant_Color.MatRGB.G / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.b = Materials3dsarray(nbmaterials - 1).Ambiant_Color.MatRGB.b / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.a = 1#
         


    Case MAT_DIFFUSE '&HA020 'Get material diffuse color
        Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Diffuse_Color
        
            Materials3dsarray(nbmaterials - 1).materialD3D.diffuse.R = Materials3dsarray(nbmaterials - 1).Diffuse_Color.MatRGB.R / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.diffuse.G = Materials3dsarray(nbmaterials - 1).Diffuse_Color.MatRGB.G / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.diffuse.b = Materials3dsarray(nbmaterials - 1).Diffuse_Color.MatRGB.b / 255
                     
           'swap diffuse color as ambiant color
           'this will be fix as when i kown how to set material in dx8 to diffuse as default :/
           Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.R = Materials3dsarray(nbmaterials - 1).materialD3D.diffuse.R
           Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.G = Materials3dsarray(nbmaterials - 1).materialD3D.diffuse.G
           Materials3dsarray(nbmaterials - 1).materialD3D.Ambient.b = Materials3dsarray(nbmaterials - 1).materialD3D.diffuse.b
                         
                      
                      
                      
        
        Form1.List1.AddItem "         |------> Diffuse Color RGB(" & _
        Materials3dsarray(nbmaterials - 1).Diffuse_Color.MatRGB.R & ", " & _
        Materials3dsarray(nbmaterials - 1).Diffuse_Color.MatRGB.G & ", " & _
        Materials3dsarray(nbmaterials - 1).Diffuse_Color.MatRGB.b & " )"

    Case MAT_SPECULAR '&HA030 'Get material Specular color
        Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Specular_Color
        
            Materials3dsarray(nbmaterials - 1).materialD3D.Specular.R = Materials3dsarray(nbmaterials - 1).Specular_Color.MatRGB.R / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.Specular.G = Materials3dsarray(nbmaterials - 1).Specular_Color.MatRGB.G / 255
            Materials3dsarray(nbmaterials - 1).materialD3D.Specular.b = Materials3dsarray(nbmaterials - 1).Specular_Color.MatRGB.b / 255

        Form1.List1.AddItem "         |------> Specular Color RGB(" & _
        Materials3dsarray(nbmaterials - 1).Specular_Color.MatRGB.R & ", " & _
        Materials3dsarray(nbmaterials - 1).Specular_Color.MatRGB.G & ", " & _
        Materials3dsarray(nbmaterials - 1).Specular_Color.MatRGB.b & " )"

    Case MAT_SHININESS '&HA040  ' Material shininess percent  'fixé le 11/02/2005
         Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Shininess_Percent

         Form1.List1.AddItem "         |---> Shininess Percent : " & Materials3dsarray(nbmaterials - 1).Shininess_Percent.Percent & " %"

    Case MAT_SHIN2PCT '&HA041 'Material shininess strength percent  'fixé le 11/02/2005
          Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Shininess_Strength_Percent

          Form1.List1.AddItem "         |---> Shininess Strength Percent : " & Materials3dsarray(nbmaterials - 1).Shininess_Strength_Percent.Percent & " %"

    Case MAT_TRANSPARENCY '&HA050 'Material Opacity/Transparency Percent  'fixé le 11/02/2005
          Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Opacity_Percent

          Form1.List1.AddItem "         |---> Opacity Percent : " & 100 - Materials3dsarray(nbmaterials - 1).Opacity_Percent.Percent & " %"
'
    Case MAT_XPFALL '&HA052 'Material Opacity/Transparency falloff percent  'fixé le 11/02/2005
          Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Opacity_FallOff_Percent

          Form1.List1.AddItem "         |---> Opacity Falloff Percent : " & Materials3dsarray(nbmaterials - 1).Opacity_FallOff_Percent.Percent & " %"

    Case MAT_REFBLUR '&HA053 'Material Reflexion/Blur Percent  'fixé le 11/02/2005
          Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Reflexion_Blur

          Form1.List1.AddItem "         |---> Reflexion/Blur Percent : " & (100 + Materials3dsarray(nbmaterials - 1).Reflexion_Blur.Percent) / 100 & " %"

    Case MAT_USE_XPFALL '&HA240 'boolean
        Get #NumeroFichier, , TempByte

            Form1.List1.AddItem "         |---> Use transparency Fall off : " & CBool(TempByte)


    Case MAT_TWO_SIDE '&HA081   'ajout le 11/02/2005 'if exist then 2sided=true
          Materials3dsarray(nbmaterials - 1).TwoSided = True

          Form1.List1.AddItem "         |---> 2 sided : " & Materials3dsarray(nbmaterials - 1).TwoSided


    Case &HA084 'Material Self Illumination Percent  'fixé le 11/02/2005
          Get #NumeroFichier, , Materials3dsarray(nbmaterials - 1).Self_Illumination_Percent

          Form1.List1.AddItem "         |---> Self Illumination Percent : " & Materials3dsarray(nbmaterials - 1).Self_Illumination_Percent.Percent & " %"


    Case MAT_WIRE '&HA085 'Material wire 'ajout le 11/02/2005 'if exist then wire=true
          Materials3dsarray(nbmaterials - 1).Wired = True

          Form1.List1.AddItem "         |---> Wire : " & Materials3dsarray(nbmaterials - 1).Wired


    Case MAT_SELF_ILLUM '&HA080
    'if exist =true
     Form1.List1.AddItem "         |---> Material self illumination : true" '& Materials3dsarray(nbmaterials - 1).TwoSided


    Case MAT_DECAL      '&HA082 'if exist=true

    Case MAT_ADDITIVE   ' &HA083 'boolean true if exist

    Case MAT_PHONGSOFT  '&HA08C 'boolean true if exist

    Case MAT_WIREABS    '&HA08E 'boolean true if exist



   Case MAT_WIRE_SIZE '&HA087 'Material wire thickness  'ajout le 11/02/2005
        Get #NumeroFichier, , TempSingle

        Form1.List1.AddItem "         |---> Material wire thickness : " & TempSingle 'Materials3dsarray(nbmaterials - 1).FaceMap


    Case MAT_FACEMAP '&HA088 'Material Face map 'boolean (on/off) 'if exist then Face Map=true
'          Get #NumeroFichier, , TempInteger 'donne le sub chunk a lire
'          Get #NumeroFichier, , TempByte 'sub chunk len in bytes = 9
          Materials3dsarray(nbmaterials - 1).FaceMap = True

          Form1.List1.AddItem "         |---> Face Map : " & Materials3dsarray(nbmaterials - 1).FaceMap

 
    Case EDIT_MATERIAL '&HAFFF 'add a new material
    tBloc.Longeur = 6


    Case MAT_MAP_TILING '&HA351 'Material MAP TILING
                  Form1.List1.AddItem "         |---> Material MAP TILING : True"
                  ' Get #NumeroFichier, , Tempbyte 'ou Templong

    Case MAT_MAP_TEXBLUR '&HA353
        Get #NumeroFichier, , TempSingle
           Form1.List1.AddItem "         |---> Map texture Blur : " & TempSingle


    Case MAT_MAP_USCALE '&HA354 'Material Map u scale (texture scale sur x)
        Get #NumeroFichier, , TempSingle
           Form1.List1.AddItem "         |---> Map v scale : " & TempSingle


    Case MAT_MAP_VSCALE '&HA356 'Material Map v scale (texture scale sur y)
        Get #NumeroFichier, , TempSingle
           Form1.List1.AddItem "         |---> Map u scale : " & TempSingle


    Case MAT_MAP_UOFFSET '&HA358 'Material Map u offset (texture placement sur x)
        Get #NumeroFichier, , TempSingle
           Form1.List1.AddItem "         |---> Map v offset : " & TempSingle


    Case MAT_MAP_VOFFSET '&HA35A 'Material Map v offset (texture placement sur y)
        Get #NumeroFichier, , TempSingle
           Form1.List1.AddItem "         |---> Map u offset : " & TempSingle


    Case MAT_MAP_ANG '&HA35C 'Material Map Rotation Angle (texture placement Angle)
        Get #NumeroFichier, , TempSingle
           Form1.List1.AddItem "         |---> Map Rotation Angle : " & TempSingle


    Case MAT_MAP_COL1  ' &HA360
        Get #NumeroFichier, , TempByte 'red color
        Get #NumeroFichier, , TempByte 'green color
        Get #NumeroFichier, , TempByte 'blue color


    Case MAT_MAP_COL2  ' &HA362
        Get #NumeroFichier, , TempByte 'red color
        Get #NumeroFichier, , TempByte 'green color
        Get #NumeroFichier, , TempByte 'blue color


    Case MAT_MAP_RCOL  ' &HA364
        Get #NumeroFichier, , TempByte 'red 1 color
        Get #NumeroFichier, , TempByte 'red 2 color
        Get #NumeroFichier, , TempByte 'red 3 color


    Case MAT_MAP_GCOL  ' &HA366
        Get #NumeroFichier, , TempByte 'green 1 color
        Get #NumeroFichier, , TempByte 'green 2 color
        Get #NumeroFichier, , TempByte 'green 3 color


    Case MAT_MAP_BCOL  ' &HA368
        Get #NumeroFichier, , TempByte 'blue 1 color
        Get #NumeroFichier, , TempByte 'blue 2 color
        Get #NumeroFichier, , TempByte 'blue 3 color


    Case EDIT_CONFIG1
        Get #NumeroFichier, , TempSingle
'        Echelle = TempSingle

        Form1.List1.AddItem "echelle" & TempSingle
        
        
'    Case MESH_VERSION
'        Get #NumeroFichier, , TempInteger
'        Form1.List1.AddItem "mesh version :" & TempInteger
''        MeshVersion = TempInteger
        
    '*****************************************************************************************
    Case EDIT_OBJECT 'object name '&H4000
        'name of the object "6Abbey.000" "Box01" etc...
        
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        While TempChar <> Chr(0)
            TempString = TempString & TempChar
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        Wend

       Form1.List1.AddItem "Nom du Mesh trouvé : " & TempString & " "
       
        TempLong = fpos + Len(TempString) + 7 ' + 6 + Len(Chr(0))
        While TempLong < (fpos + tBloc.Longeur)
            LireBloc TempLong
        Wend
        

    Case OBJECT_HIDDEN 'object visibility on/off not really usefull
        Get #NumeroFichier, , TempByte

        If TempByte = 0 Then
            Form1.List1.AddItem "         |---> Object Visibility : True"
        Else
            Form1.List1.AddItem "         |---> Object Visibility : False"
        End If


    Case OBJECT_CAST 'Object doesn't cast on/off not really usefull
        Get #NumeroFichier, , TempByte

        If TempByte = 0 Then
            Form1.List1.AddItem "         |---> Object Cast : True"
        Else
            Form1.List1.AddItem "         |---> Object Cast : False"
        End If


    Case OBJECT_MATT 'Matte object on/off not really usefull
        Get #NumeroFichier, , TempByte

        If TempByte = 0 Then
            Form1.List1.AddItem "         |---> Object Matte : True"
        Else
            Form1.List1.AddItem "         |---> Object Matte : False"
        End If


    Case OBJECT_EXTP 'External process on on/off not really usefull
        Get #NumeroFichier, , TempByte

        If TempByte = 0 Then
            Form1.List1.AddItem "         |---> Object External process on : True"
        Else
            Form1.List1.AddItem "         |---> Object External process on : False"
        End If


    Case OBJECT_SHAD 'Object doesn't receive shadows on/off not really usefull
        Get #NumeroFichier, , TempByte

        If TempByte = 0 Then
            Form1.List1.AddItem "         |---> Object Receive Shadows : True"
        Else
            Form1.List1.AddItem "         |---> Object Receive Shadows : False"
        End If
        
  '***********************************************************************************************

        
    'Sous bloc d'EDIT_OBJET doit etre lu par l'edit_objet coordonnées xyz de chaque point
    Case TRI_VERTEXL
        nbrSolide = nbrSolide + 1
        ReDim Preserve ObjetVertex(nbrSolide)
        Get #NumeroFichier, , TempInteger
        ReDim ObjetVertex(nbrSolide).tVertex(TempInteger - 1)
        Get #NumeroFichier, , ObjetVertex(nbrSolide).tVertex

         
    'index des points des triangles + maillage du triangle AB,BC,AC   'Face de l'objet 3d.
    Case TRI_FACEL1 '&H4120
        Get #NumeroFichier, , TempInteger  'cube= 12 : 6 faces * 2 = 12 triangles

        'bug here
        If TempInteger > 0 Then
        nbrVertex = CLng(TempInteger) * 3 'nombre de triangles * 3 = nombre de points
        ObjetVertex(nbrSolide).nbmappingpoint = nbpoint
        ObjetVertex(nbrSolide).nbrVertex = nbrVertex
        Form1.List1.AddItem "         |---> Nombre de points trouvé : " & nbrVertex
        ObjetVertex(nbrSolide).nbrFace = TempInteger ' nombre de faces
        Form1.List1.AddItem "         |---> Nombre de triangles trouvé : " & TempInteger
'        ReDim ObjetVertex(nbrSolide).VertexFace(0)
        ReDim ObjetVertex(nbrSolide).VertexFaceMaterialindex(CLng(TempInteger) - 1) 'face material index for texturing aka texture(i)
        ReDim Preserve ObjetVertex(nbrSolide).VertexFace(CLng(TempInteger) - 1)
        
        Get #NumeroFichier, , ObjetVertex(nbrSolide).VertexFace
        TempLong = fpos + 8 + 8 * CLng(TempInteger)
        
        
        While TempLong <> (fpos + tBloc.Longeur)
            LireBloc TempLong
        Wend
        End If
        
      Case TRI_MATERIAL '&H4130 'faceindex , material index
      'Get the material face name
              TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        While TempChar <> Chr(0)
            TempString = TempString & TempChar
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        Wend
        
        
        'get the number of face that use this material
        Get #NumeroFichier, , TempInteger 'number of entries
         Form1.List1.AddItem "TRI_MATERIAL Faces mapping : " & TempString & " : " & TempInteger & " Faces"
       
    Index = -1
For i = 0 To nbmaterials - 1
    If Materials3dsarray(i).Material_Name = TempString Then
        Index = i
        Exit For
    End If
Next
       
        'get the face index of the face using this material
        'if they are in this part we know witch material/bitmap it use for mapping !
        ReDim matindex(TempInteger - 1)
        Get #NumeroFichier, , matindex
'fill the matindex for all the triangles using this material:
For i = 0 To TempInteger - 1
    If Index <> -1 Then
    ObjetVertex(nbrSolide).VertexFaceMaterialindex(matindex(i)) = Index '0,1,2 and so on...
    End If
Next

      Case MAT_SHADING
      
      Case TRI_MAPPINGSTANDARD 'standart mapping only 'MESH_TEXTURE_INFO
            Get #NumeroFichier, , TempLong

            Select Case TempLong
             Case 0
             Form1.List1.AddItem "planar"
             Case 1
             Form1.List1.AddItem "cylindrical"
             Case 2
             Form1.List1.AddItem "spherical"
            End Select

             Get #NumeroFichier, , TempSingle 'x_tiling
             Form1.List1.AddItem "         |---> Map X_tiling : " & TempSingle
             Get #NumeroFichier, , TempSingle 'Y_tiling
              Form1.List1.AddItem "         |---> Map Y_tiling : " & TempSingle
             Get #NumeroFichier, , TempSingle 'icon_x
              Form1.List1.AddItem "         |---> Map icon_x : " & TempSingle
             Get #NumeroFichier, , TempSingle 'icon_y
              Form1.List1.AddItem "         |---> Map icon_y : " & TempSingle
             Get #NumeroFichier, , TempSingle 'icon_z
              Form1.List1.AddItem "         |---> Map icon_z : " & TempSingle

'
    Case &HAFFF   'material editor chunk skip it
        tBloc.Longeur = 6

'
    Case &HB008 'Frame
        Get #NumeroFichier, , TempLong
        KeyFrameStart = TempLong
        Get #NumeroFichier, , TempLong
        KeyFrameEnd = TempLong

    Case &HB013  'Pivot de l'objet
        Get #NumeroFichier, , PivotObjetX
        Get #NumeroFichier, , PivotObjetY
        Get #NumeroFichier, , PivotObjetZ

    Case &HB010
        TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        While TempChar <> Chr(0)
            TempString = TempString & TempChar
            TempChar = StrConv(InputB(1, NumeroFichier), vbUnicode)
        Wend
        h = TempString
    MsgBox TempString
        Get #NumeroFichier, , TempSingle
'        Get #NumeroFichier, , TempInteger
        Get #NumeroFichier, , TempInteger
        
    Case &HB020 'track

    Case &HB021 'Animation par quaternion

        
'     Case &HA100, &HA200, &HA353, &H4130, &H4150, &HB001, &HB002, &HB008, &HB009, &HB00A, &HB010, &HB013, &HB020, &HB021
Case OBJ_TRIMESH, BACKGROUND_COLOR, EDIT_UNKNW13, &HA200, &HA204, &HA210, &HA220, &HA230, &H4100, &H1200, &H0  'OBJ_TRIMESH, OBJ_CAMERA, &HA200, BACKGROUND_COLOR, EDIT_UNKNW13, &HA204, &HA210, &HA220, &HA230, &H1200, &HA353, &HB001, &HB002, &HB009, &HB00A, &H0, &H4710, &H4720
     
        TempLong = fpos + 6
        While TempLong <> (fpos + tBloc.Longeur)
            LireBloc TempLong
        Wend


Case &H4600 'Lumière
Get #NumeroFichier, , TempSingle 'pos x
Get #NumeroFichier, , TempSingle 'pos y
Get #NumeroFichier, , TempSingle 'pos z


Case &H4610 ' Lumière Spot

Get #NumeroFichier, , TempSingle 'pos x
Get #NumeroFichier, , TempSingle 'pos y
Get #NumeroFichier, , TempSingle 'pos z
Get #NumeroFichier, , TempSingle 'hotspot
Get #NumeroFichier, , TempSingle 'falloff


Case &H4620 'allumée ou eteinte ?



Case OBJ_CAMERA '&H4700

Get #NumeroFichier, , TempSingle 'pos x
Get #NumeroFichier, , TempSingle 'pos y
Get #NumeroFichier, , TempSingle 'pos z
Get #NumeroFichier, , TempSingle ' Cible x
Get #NumeroFichier, , TempSingle ' Cible y
Get #NumeroFichier, , TempSingle ' Ciblez
Get #NumeroFichier, , TempSingle ' angle
Get #NumeroFichier, , TempSingle ' lens


'Case &H4620, &H4625, &H4627, &H4630, &H4641, &H4650, &H4651, &H4652, &H4653, &H4654, &H4656, &H4657, &H4658, &H4659, &H465A, &H465B
'
'                TempLong = fpos + 6
'        While TempLong <> (fpos + tBloc.Longeur)
'            LireBloc TempLong
'        Wend
        
'Case CAM_SEE_CONE, CAM_RANGES, &H4730, &H4740, &H4750, &H4760, &H4770, &H4780, &H4790
'
'                        TempLong = fpos + 6
'        While TempLong <> (fpos + tBloc.Longeur)
'            LireBloc TempLong
'        Wend
        
'Case &HA250, &H7001, &H1400, &H1420, &H1450, &H1500, &H2100, &H1100, &H1300, &H1301, &H2200, &H2300, &H3000
'
'        TempLong = fpos + 6
'        While TempLong <> (fpos + tBloc.Longeur)
'            LireBloc TempLong
'        Wend
        
        
'    Case Else

    
'OBJ_TRIMESH, OBJ_CAMERA, BACKGROUND_COLOR, EDIT_UNKNW13, &HA204, &HA210, &HA220, &HA230 '&H4100, &H1200, &H3000
'&HA100
'&HA200
'&HA351
'&HA353
'&H30
'&H4130
'&H4150
'&HB001
'&HB002
'&HB008
'&HB009
'&HB00A
'&HB010
'&HB013
'&HB020
'&HB021
'&H0

'MsgBox Hex(tBloc.Entete)
'
'            TempLong = fpos + 6
'        While TempLong <> (fpos + tBloc.Longeur)
'            LireBloc TempLong
'        Wend

        
End Select
    
    fpos = fpos + tBloc.Longeur
End Function


