VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vIdPersona As Long
Private vId As Long
Private vNombre As String
Private vDomicilio As String
Private vLocalidad As String
Private vCargo As String
Private vTelefono As String
Private vCelular As String
Private vEmail As String
Private vProvincia As String
Private vDetalle As String
Private vCiudad As String
Private vPais As String
Private vTipo As Long


Public Property Get idPersona() As Long
    idPersona = vIdPersona
End Property

Public Property Get id() As Long
    id = vId
End Property


Public Property Get nombre() As String
    nombre = vNombre
End Property


Public Property Get Domicilio() As String
    Domicilio = vDomicilio
End Property


Public Property Get localidad() As String
    localidad = vLocalidad
End Property


Public Property Get Cargo() As String
    Cargo = vCargo
End Property


Public Property Get telefono() As String
    telefono = vTelefono
End Property


Public Property Get celular() As String
    celular = vCelular
End Property


Public Property Get email() As String
    email = vEmail
End Property


Public Property Get detalle() As String
    detalle = vDetalle
End Property


Public Property Get provincia() As String
    provincia = vProvincia
End Property


Public Property Get pais() As String
    pais = vPais
End Property


Public Property Get Tipo() As Long
    Tipo = vTipo
End Property


Public Property Let id(nvalue As Long)
    vId = nvalue
End Property


Public Property Let idPersona(nvalue As Long)
    vIdPersona = nvalue
End Property


Public Property Let nombre(nvalue As String)
    vNombre = nvalue
End Property


Public Property Let Domicilio(nvalue As String)
    vDomicilio = nvalue
End Property


Public Property Let localidad(nvalue As String)
    vLocalidad = nvalue
End Property


Public Property Let Cargo(nvalue As String)
    vCargo = nvalue
End Property


Public Property Let telefono(nvalue As String)
    vTelefono = nvalue
End Property


Public Property Let celular(nvalue As String)
    vCelular = nvalue
End Property


Public Property Let email(nvalue As String)
    vEmail = nvalue
End Property


Public Property Let detalle(nvalue As String)
    vDetalle = nvalue
End Property


Public Property Let provincia(nvalue As String)
    vProvincia = nvalue
End Property


Public Property Let pais(nvalue As String)
    vPais = nvalue
End Property


Public Property Let Tipo(nvalue As Long)
    vTipo = nvalue
End Property

