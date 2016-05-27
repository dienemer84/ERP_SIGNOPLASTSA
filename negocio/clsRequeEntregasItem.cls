VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequeEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vCantidad As Double
Dim vTipo As tipoEntrega
Dim vId As Long
Dim vFecha As Date
Public Property Let Tipo(nvalue As tipoEntrega)
    vTipo = nvalue
End Property

Public Property Get Tipo() As tipoEntrega
    Tipo = vTipo
End Property
Public Property Let Cantidad(nvalue As Double)
    vCantidad = nvalue
End Property
Public Property Let FEcha(nvalue As Date)
    vFecha = nvalue
End Property
Public Property Get Cantidad() As Double
    Cantidad = vCantidad
End Property
Public Property Get FEcha() As Date
    FEcha = vFecha
End Property
Public Property Let id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get id() As Long
    id = vId
End Property



