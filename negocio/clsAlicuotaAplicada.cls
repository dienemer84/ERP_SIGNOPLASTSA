VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlicuotaAplicada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vMonto As Double
Private vAlicuota As clsAlicuotas
Public id As Long

Public Property Let Monto(nMonto As Double)
    vMonto = nMonto
End Property
Public Property Get Monto() As Double
    Monto = vMonto
End Property
Public Property Let Alicuota(nAlicuota As clsAlicuotas)
    Set vAlicuota = nAlicuota
End Property
Public Property Get Alicuota() As clsAlicuotas
    Set Alicuota = vAlicuota
End Property
