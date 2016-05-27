VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private vRubros As New Collection
Private vCuit As String
Private vId As Long
Private vRazonSocial As String
Private mTipoIva As clsTipoIvaProveedor
Private vDireccion As String
Private vCiudad As String
Private vCP As String
Private vTel As String
Private vFax As String
Private vEmail As String
Private vContacto As String
Public Moneda As clsMoneda

Private mContactos As New Collection
Private vForma_pago As String
Private vPago_contra_entrega As Boolean
Private vPago_dolares As Boolean
Private vEstado As EstadoProveedor
Private vBonificacion As Double
Private vCuentasContables As New Collection
Private vIibb As String
Private vRazonFantasia As String
Public Property Let contactos(nvalue As Collection)
    Set mContactos = nvalue
End Property
Public Property Get contactos() As Collection
    Set contactos = mContactos
End Property
Public Property Let rubros(nvalue As Collection)
    Set vRubros = nvalue
End Property
Public Property Get rubros() As Collection
    Set rubros = vRubros
End Property
Public Property Let IIBB(nvalue As String)
    vIibb = nvalue
End Property
Public Property Get IIBB() As String
    IIBB = vIibb
End Property
Public Property Let razonFantasia(nvalue As String)
    vRazonFantasia = UCase(nvalue)
End Property
Public Property Get razonFantasia() As String
    razonFantasia = UCase(vRazonFantasia)
End Property
Public Property Let cuentasContables(nvalue As Collection)
    Set vCuentasContables = nvalue
End Property
Public Property Get cuentasContables() As Collection
    Set cuentasContables = vCuentasContables
End Property
Public Property Let RazonSocial(nvalue As String)
    vRazonSocial = nvalue
End Property
Public Property Get RazonSocial() As String
    RazonSocial = vRazonSocial
End Property
Public Property Let Cuit(nvalue As String)
    vCuit = nvalue
End Property
Public Property Get Cuit() As String
    Cuit = vCuit
End Property
Public Property Let bonificacion(nvalue As Double)
    vBonificacion = nvalue
End Property
Public Property Get bonificacion() As Double
    bonifacion = vBonificacion
End Property
Public Property Let id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get id() As Long
    id = vId
End Property
Public Property Let direccion(nvalue As String)
    vDireccion = nvalue
End Property
Public Property Get direccion() As String
    direccion = vDireccion
End Property
Public Property Let Ciudad(nvalue As String)
    vCiudad = nvalue
End Property
Public Property Get Ciudad() As String
    Ciudad = vCiudad
End Property
Public Property Let cp(nvalue As String)
    vCP = nvalue
End Property
Public Property Get cp() As String
    cp = vCP
End Property
Public Property Let tel(nvalue As String)
    vTel = nvalue
End Property
Public Property Get tel() As String
    tel = vTel
End Property
Public Property Let Fax(nvalue As String)
    vFax = nvalue
End Property
Public Property Get Fax() As String
    Fax = vFax
End Property
Public Property Let email(nvalue As String)
    vEmail = nvalue
End Property
Public Property Get email() As String
    email = vEmail
End Property
Public Property Let contacto(nvalue As String)
    vContacto = nvalue
End Property
Public Property Get contacto() As String
    contacto = vContacto
End Property
Public Property Let FormaPago(nvalue As String)
    vForma_pago = nvalue
End Property
Public Property Get FormaPago() As String
    FormaPago = vForma_pago
End Property
Public Property Let pagocontraEntrega(nvalue As Boolean)
    vPago_contra_entrega = nvalue
End Property
Public Property Get pagocontraEntrega() As Boolean
    pagocontraEntrega = vPago_contra_entrega
End Property
Public Property Let pagoDolares(nvalue As Boolean)
    vPago_dolares = nvalue
End Property
Public Property Get pagoDolares() As Boolean
    pagoDolares = vPago_dolares
End Property
Public Property Let estado(nvalue As EstadoProveedor)
    vEstado = nvalue

End Property
Public Property Get estado() As EstadoProveedor
    estado = vEstado
End Property
Public Property Set TipoIVA(nTipoIva As clsTipoIvaProveedor)
    Set mTipoIva = nTipoIva
End Property
Public Property Get TipoIVA() As clsTipoIvaProveedor
    Set TipoIVA = mTipoIva
End Property










