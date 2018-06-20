Attribute VB_Name = "modInfoArchi"
Option Explicit

' Dunkan: Esto tendría que volar.

Public PathFile As String
Public bmapped As Boolean
Public nbpoint As Long
Public TextureFilename As String
Public TextureHeight As Long
Public TextureWidth As Long

'Public Echelle As String
'Public MeshVersion As Long

Public ObjectName As String
Public filelentitle As Long


Public filetitle As String      'path of our file
Public file3dsname As String    'filename without extension

'auto scene size vars
Public Ymax As Single
Public Ymin As Single
Public Xmax As Single
Public Xmin As Single
Public Zmax As Single
Public Zmin As Single

'screenshot tools
Public Type POINTAPI
X As Long
Y As Long
End Type

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long


