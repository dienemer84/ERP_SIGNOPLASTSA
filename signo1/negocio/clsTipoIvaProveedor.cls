VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoIvaProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private vDetalle As String
Private mConfigFacturas As Collection
Private vId As Long
Public Property Let Id(nId As String)
    vId = nId
End Property
Public Property Get Id() As String
    Id = vId
End Property
Public Property Let configFacturas(nConfigFacturas As Collection)
    Set mConfigFacturas = nConfigFacturas
End Property
Public Property Get configFacturas() As Collection
    If mConfigFacturas Is Nothing Then Set mConfigFacturas = New Collection
    Set configFacturas = mConfigFacturas
End Property
Public Property Let detalle(nDetalle As String)
    vDetalle = nDetalle
End Property
Public Property Get detalle() As String
    detalle = vDetalle
End Property


