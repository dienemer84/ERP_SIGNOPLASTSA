VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfigFacturaProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vId As Long
Private vDiscrimina As Boolean
Public TipoIvaProveedor As clsTipoIvaProveedor
Private vTipoFactura As String
Public FormateaNumero As Boolean
Private mAlicuotas As New Collection

Public Property Let alicuotas(nAlicuotas As Collection)
    Set mAlicuotas = nAlicuotas
End Property

Public Property Get alicuotas() As Collection
    Set alicuotas = mAlicuotas
End Property


Public Property Let TipoFactura(nTipoFactura As String)
    vTipoFactura = nTipoFactura
End Property

Public Property Get TipoFactura() As String
    TipoFactura = vTipoFactura
End Property


Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get Id() As Long
    Id = vId
End Property
Public Property Let Discrimina(nvalue As Boolean)
    vDiscrimina = nvalue
End Property
Public Property Get Discrimina() As Boolean
    Discrimina = vDiscrimina
End Property
