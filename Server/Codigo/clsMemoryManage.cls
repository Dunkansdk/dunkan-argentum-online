VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Este modulo manipula la memoria atraves de 2 variables de array grande. Cada una de estas variables cada determinado tiempo seran copiadas y luego seran borradas de la memoria cambiandolas de lugar y haciendo mas dificil su ubicación.
'Una ves que se cambie la variable que se manipulara entonces se comenzara a trabajar sobre esta y nuevamente sera cambiada.
'La clase solamente tendra 2 variables con array grande y un identificador para saber con cual de las 2 variables se esta trabajando
'Luego los datos que se quieran guardar seran escritos de forma azarosa dentro de este array dividiendo las variables.
'Para poder obtener el valor de la variable se tendran que declarar variables fuera de la clase y estas deberan ser usadas con la clase.
'Lo que haran las funciones para realizar el Write y Get en la memoria sera indicar a las variables que nos dan un valor encriptado de la posicion en bytes en la variable principal
'Para facilitar haremos un ejemplo:
'La variables principales se llamaran VMain().data() as byte
'VMain() tendra un array del 0 al 1.
'y el data sera de 10240bytes por defeult pero al iniciar la clase se podra especificar (VMain(X).data(10240) as byte)
'Cada determinado tiempo se cambiara del 0 al 1 y del 1 al 0 cambiando la ubicacion en la memoria.
'Luego en la variable data() se pondran los datos al azar.
'Por ejemplo si yo quiero guardar una variable INTEGER entonces se dividira en 2bytes y cada byte se ubicara en una posicion azarosa en el data()
'Luego con la variable indicada que debera ser una variable de 4 bytes por array (en caso de ser INTEGER) Los primeros 2bytes serviran para ubicar el primer byte en data() y los otros 2bytes seran para ubicar el segundo byte en data()
'De esta forma confundiremos a los que intentaran editar la memoria :D
'Para hacer mas complicado todo, es sugerible siempre utilizar variables superiores a 2bytes y usar valores superiores al 256.

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private LMain As Integer
Private NMain As Byte
Private VMain() As tData
Private HMain() As Byte

Private Type tData
    Data() As Byte
End Type

Public Enum tVar
    tByte = 0
    tinteger = 1
    tLong = 3
    tSingle = 3
    tDouble = 7
    tSmallString = 10
    tMediumString = 100
    tLongString = 1000
End Enum

Public Function GetData() As Byte()
    GetData = VMain(NMain).Data
End Function

Public Function GetAct() As Byte()
    GetAct = HMain
End Function

Public Sub SetData(ByRef Data() As Byte, ByRef Act() As Byte)
    
    Dim loopC As Long
    LMain = UBound(Data)
    
    For loopC = 0 To UBound(Data) - 1
        If Act(loopC) = 1 Then
            VMain(NMain).Data(loopC) = Data(loopC)
            HMain(loopC) = Act(loopC)
        End If
    Next loopC
    
End Sub

Public Function GetLength() As Long
     GetLength = LMain
End Function

Public Function UsingData() As Integer
    Dim loopC As Long
    Dim Count As Integer
    
    For loopC = 0 To LMain - 1
        If HMain(loopC) = 1 Then
            Count = Count + 1
        End If
    Next loopC
    
    UsingData = Count
    
End Function

Public Sub Initialize(Optional ByRef NMainStart = 0, Optional ByRef length As Long = 10240) '¿Con que VMain empiezo a manipular y que tamaño manipulo?
    ReDim VMain(1) As tData
    
    LMain = length
    NMain = NMainStart
    
    ReDim VMain(NMain).Data(LMain - 1) As Byte
    ReDim HMain(LMain - 1) As Byte
End Sub

Public Sub Refresh(Optional ByRef NMainStart = 0, Optional ByRef length As Long = 10240) 'Borra todo y lo resetea. Por si queremos borrar todas las variables ^^
    Call ZeroMemory(HMain(0), LMain + 1)
    Call ZeroMemory(VMain(NMain).Data(0), LMain + 1)
    
    LMain = length
    NMain = NMainStart
    
    ReDim VMain(1) As tData
    ReDim VMain(NMain).Data(LMain - 1) As Byte
    ReDim HMain(LMain - 1) As Byte
End Sub

Public Sub ChangeMain()
    Dim TempNMain As Byte
    TempNMain = IIf(NMain = 0, 1, 0)
    
    ReDim VMain(TempNMain).Data(LMain - 1) As Byte
    
    Call CopyMemory(VMain(TempNMain).Data(0), VMain(NMain).Data(0), LMain + 1)
    Call ZeroMemory(VMain(NMain).Data(0), LMain + 1)
    
    ReDim VMain(NMain).Data(0) As Byte
    
    NMain = TempNMain
End Sub

Private Function FindMemorySpace(ByRef DownValue As Integer, ByRef UpValue As Integer, Optional ByRef Bucle As Boolean = False) As Integer
Dim loopC As Long

    For loopC = DownValue To UpValue
        If HMain(loopC) = 0 Then
            FindMemorySpace = loopC
            Exit Function
        End If
    Next loopC

    If Not Bucle Then
        FindMemorySpace = FindMemorySpace(0, DownValue, True)
    End If

End Function

Private Function RandomFind() As Integer
    RandomFind = FindMemorySpace(Rnd * LMain, LMain - 1)
End Function

' '
' Integers

Public Function SetInteger(ByRef V_In() As Integer, ByRef Value As Integer)
    Dim Data(1) As Byte
    CopyMemory Data(0), Value, 2
    
    Dim loopC As Long
    For loopC = 0 To 1
        HMain(V_In(loopC)) = 0
    Next loopC
    
    For loopC = 0 To 1
        V_In(loopC) = RandomFind()
        HMain(V_In(loopC)) = 1
        VMain(NMain).Data(V_In(loopC)) = Data(loopC) Xor V_In(loopC) - (256 + (Fix(V_In(loopC) / 256) - 1) * 256)
    Next loopC
    
    ZeroMemory Data(0), 2
End Function

Public Function GetInteger(ByRef V_In() As Integer) As Integer
    Dim loopC As Long
    Dim Data(1) As Byte
    
    'EEEEEEEEEEEEEE DUNKAN HIJO DE PUTA !!!!!!
    For loopC = 0 To 1
        Data(loopC) = VMain(NMain).Data(V_In(loopC)) Xor V_In(loopC) - (256 + (Fix(V_In(loopC) / 256) - 1) * 256)
    Next loopC
    
    CopyMemory GetInteger, Data(0), 2
    ZeroMemory Data(0), 2
    
    'Call SetInteger(V_In(), GetInteger)
End Function

Public Sub DelInteger(ByRef V_In() As Integer)
    Dim loopC As Long
    For loopC = 0 To 1
        HMain(V_In(loopC)) = 0
    Next loopC
End Sub

' '
' Longs

Public Function SetLong(ByRef V_In() As Integer, ByRef Value As Long)
    Dim Data(3) As Byte
    CopyMemory Data(0), Value, 4
    
    Dim loopC As Long
    For loopC = 0 To 3
        HMain(V_In(loopC)) = 0
    Next loopC
    
    For loopC = 0 To 3
        V_In(loopC) = RandomFind()
        HMain(V_In(loopC)) = 1
        VMain(NMain).Data(V_In(loopC)) = Data(loopC) Xor V_In(loopC) - (256 + (Fix(V_In(loopC) / 256) - 1) * 256)
    Next loopC
    
    ZeroMemory Data(0), 4
End Function

Public Function GetLong(ByRef V_In() As Integer) As Long
    Dim loopC As Long
    Dim Data(3) As Byte
    
    For loopC = 0 To 3
        Data(loopC) = VMain(NMain).Data(V_In(loopC)) Xor V_In(loopC) - (256 + (Fix(V_In(loopC) / 256) - 1) * 256)
    Next loopC
    
    CopyMemory GetLong, Data(0), 4
    ZeroMemory Data(0), 4
    
    'Call SetLong(V_In(), GetLong)
End Function

Public Sub DelLong(ByRef V_In() As Integer)
    Dim loopC As Long
    For loopC = 0 To 3
        HMain(V_In(loopC)) = 0
    Next loopC
End Sub

'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||| Single |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Public Function SetSingle(ByRef V_In() As Integer, ByRef Value As Single)
    Dim Data(3) As Byte
    CopyMemory Data(0), Value, 4
    Dim i As Long
    For i = 0 To 3
        HMain(V_In(i)) = 0
    Next i
    For i = 0 To 3
        V_In(i) = RandomFind()
        HMain(V_In(i)) = 1
        VMain(NMain).Data(V_In(i)) = Data(i) Xor V_In(i) - (256 + (Fix(V_In(i) / 256) - 1) * 256)
    Next i
    ZeroMemory Data(0), 4
End Function

Public Function GetSingle(ByRef V_In() As Integer) As Single
    Dim i As Long
    Dim Data(3) As Byte
    For i = 0 To 3
        Data(i) = VMain(NMain).Data(V_In(i)) Xor V_In(i) - (256 + (Fix(V_In(i) / 256) - 1) * 256)
    Next i
    CopyMemory GetSingle, Data(0), 4
    ZeroMemory Data(0), 4
    'Call SetSingle(V_In(), GetSingle)
End Function

Public Sub DelSingle(ByRef V_In() As Integer)
    Dim i As Long
    For i = 0 To 3
        HMain(V_In(i)) = 0
    Next i
End Sub

'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||| Double |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Public Function SetDouble(ByRef V_In() As Integer, ByRef Value As Double)
    Dim Data(7) As Byte
    CopyMemory Data(0), Value, 8
    Dim i As Long
    For i = 0 To 7
        HMain(V_In(i)) = 0
    Next i
    For i = 0 To 7
        V_In(i) = RandomFind()
        HMain(V_In(i)) = 1
        VMain(NMain).Data(V_In(i)) = Data(i) Xor V_In(i) - (256 + (Fix(V_In(i) / 256) - 1) * 256)
    Next i
    ZeroMemory Data(0), 8
End Function

Public Function GetDouble(ByRef V_In() As Integer) As Double
    Dim i As Long
    Dim Data(7) As Byte
    For i = 0 To 7
        Data(i) = VMain(NMain).Data(V_In(i)) Xor V_In(i) - (256 + (Fix(V_In(i) / 256) - 1) * 256)
    Next i
    CopyMemory GetDouble, Data(0), 8
    ZeroMemory Data(0), 8
    'Call SetDouble(V_In(), GetDouble)
End Function

Public Sub DelDouble(ByRef V_In() As Integer)
    Dim i As Long
    For i = 0 To 7
        HMain(V_In(i)) = 0
    Next i
End Sub


