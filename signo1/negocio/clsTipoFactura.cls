VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vId As Long
Private vNumeracion As Long
Private vTipo As String
Private vDiscrimina As Boolean
Public PuntoVenta As PuntoVenta
Public ExcentoIVA As Boolean



Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get Id() As Long
    Id = vId
End Property


'Public Property Let numeracion(nvalue As Long)
'    vNumeracion = nvalue
'End Property
'Public Property Get numeracion() As Long
'    numeracion = vNumeracion
'End Property


Public Property Let Tipo(nvalue As String)
    vTipo = nvalue
End Property
Public Property Get Tipo() As String
    Tipo = vTipo
End Property

Public Property Let Discrimina(nvalue As Boolean)
    vDiscrimina = nvalue
End Property
Public Property Get Discrimina() As Boolean
    Discrimina = vDiscrimina
End Property






