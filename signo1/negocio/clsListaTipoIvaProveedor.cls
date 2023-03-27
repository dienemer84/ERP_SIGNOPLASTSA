VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlicuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vId As Long
Private vAlicuota As Double



Public Property Let Id(nId As Long)
    vId = nId
End Property


Public Property Get Id() As Long
    Id = vId
End Property

Public Property Let alicuota(nAlicuota As Double)
    vAlicuota = nAlicuota

End Property


Public Property Get alicuota() As Double
    alicuota = vAlicuota
End Property
