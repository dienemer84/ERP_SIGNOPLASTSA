VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIvaAplicado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vIva As clsTipoIVA
Private vMonto As Double

Public Property Let iva(niva As clsTipoIVA)
Set vIva = niva
End Property
Public Property Get iva() As clsTipoIVA
Set iva = vIva
End Property


Public Property Let monto(nMonto As Double)
vMonto = nMonto
End Property
Public Property Get monto() As Double
monto = vMonto
End Property


 


