VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPercepcionesAplicadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vPercepcion As clsPercepciones
Private vMonto As Double
Public Id As Long
Public Property Let Percepcion(nPercepcion As clsPercepciones)
    Set vPercepcion = nPercepcion
End Property
Public Property Get Percepcion() As clsPercepciones
    Set Percepcion = vPercepcion
End Property
Public Property Let Monto(nMonto As Double)
    vMonto = nMonto
End Property
Public Property Get Monto() As Double
    Monto = vMonto
End Property


