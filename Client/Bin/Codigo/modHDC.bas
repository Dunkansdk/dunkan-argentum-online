Attribute VB_Name = "Mod_DX8_HDC"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' Módulo creado para dibujar los personajes en el CP. Boskorcha AO - Dunkansdk 2011

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public MinEleccion As Integer, MaxEleccion As Integer
Public Actual As Integer

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Sub DrawGrhtoHdc(ByVal desthDC As Long, ByVal grh_index As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)

' / Author: Emanuel Matias 'Dunkan'
' / Note: Dibujar pictures del 'Crear Personaje'

'On Error Resume Next
    
    Dim file_path   As String
    Dim src_x       As Integer
    Dim src_y       As Integer
    Dim src_width   As Integer
    Dim src_height  As Integer
    Dim hdcsrc      As Long
    Dim MaskDC      As Long
    Dim PrevObj     As Long
    Dim PrevObj2    As Long
    Dim screen_x    As Integer
    Dim screen_y    As Integer
    
    screen_x = destRect.Left
    screen_y = destRect.Top
    
    If grh_index <= 0 Then Exit Sub

    If GrhData(grh_index).numFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If

        file_path = DirGraficos & CStr(GrhData(grh_index).FileNum) & ".bmp"
       
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
           
        hdcsrc = CreateCompatibleDC(desthDC)
         
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
       
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
 
        DeleteDC hdcsrc
 
End Sub
