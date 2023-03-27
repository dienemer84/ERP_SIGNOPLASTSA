VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private vDetalle As String
Private vAlicuota As Double
Private vValido As Boolean
Private vIdIVA As Long
'Private vTipoFactura As clsTipoFactura


Public Property Let idIVA(nvalue As Long)
    vIdIVA = nvalue
End Property
Public Property Get idIVA() As Long
    idIVA = vIdIVA
End Property

Public Property Let detalle(nDetalle As String)
    vDetalle = nDetalle
End Property
Public Property Get detalle() As String
    detalle = vDetalle
End Property
Public Property Let alicuota(nAlicuota As Double)
    vAlicuota = nAlicuota
End Property
Public Property Get alicuota() As Double
    alicuota = vAlicuota
End Property

Public Property Let valido(nValido As Boolean)
    vValido = nValido
End Property
'Public Property Set TipoFactura(nValue As clsTipoFactura)
'    Set vTipoFactura = nValue
'End Property
'Public Property Get TipoFactura() As clsTipoFactura
'    Set TipoFactura = vTipoFactura
'End Property
