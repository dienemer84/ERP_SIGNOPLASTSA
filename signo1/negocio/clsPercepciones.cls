VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vPercepcion As String
Dim vId As Long
Dim vValido As Integer
Dim vPorcentaje As Double
Public codigo As String
Public Property Let Id(nId As Long)
    vId = nId
End Property
Public Property Get Id() As Long
    Id = vId
End Property

Public Property Let Percepcion(nPercepcion As String)
    vPercepcion = nPercepcion
End Property
Public Property Get Percepcion() As String
    Percepcion = vPercepcion
End Property

Public Property Let valido(nValido As Integer)
    vValido = nValido
End Property

Public Property Get valido() As Integer
    valido = vValido
End Property

Public Property Let Porcentaje(nPorcentaje As Double)
    vPorcentaje = nPorcentaje
End Property
Public Property Get Porcentaje() As Double
    Porcentaje = vPorcentaje
End Property


