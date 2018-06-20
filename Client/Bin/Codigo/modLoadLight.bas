Attribute VB_Name = "modLoadLight"
Option Explicit

Public Default_RGB(3) As Long
Public Alpha_RGB(3) As Long
Public Shadow_RGB(3) As Long

Public Function initializeColors()

' Author: Dunkan
    
    Default_RGB(0) = D3DColorXRGB(255, 255, 255)
    Default_RGB(1) = D3DColorXRGB(255, 255, 255)
    Default_RGB(2) = D3DColorXRGB(255, 255, 255)
    Default_RGB(3) = D3DColorXRGB(255, 255, 255)
    
    Alpha_RGB(0) = D3DColorARGB(100, 255, 255, 255)
    Alpha_RGB(1) = D3DColorARGB(100, 255, 255, 255)
    Alpha_RGB(2) = D3DColorARGB(100, 255, 255, 255)
    Alpha_RGB(3) = D3DColorARGB(100, 255, 255, 255)
    
    Shadow_RGB(0) = 1677721600
    Shadow_RGB(1) = 1677721600
    Shadow_RGB(2) = 1677721600
    Shadow_RGB(3) = 1677721600
    
End Function

