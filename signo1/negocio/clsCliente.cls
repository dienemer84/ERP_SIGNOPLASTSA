VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TipoIvaID As Long

Private mContacto As Collection
Private vId As Long
Private vRazon As String
Private vDomicilio As String
Public exLocalidad As String

Public TipoDocumento As TipoDocumento

Private vTelefono As String
Private vFax As String
Private vEmail As String
Private vCuit As String
Private mTipoIva As clsTipoIVA
Private vEstado As EstadoCliente
Private vPasswordSistema As String
Private vFP As Long    'en dias!!
Public FormaPago As String
Public ValidoRemitoFactura As Boolean
Public idMonedaDefault As Long

Public localidad As localidad
Public provincia As provincia

Public Property Get razonFixed()
Dim q As String
q = Replace(razon, " ", "")
q = Replace(razon, "/", "")
q = Replace(razon, "\", "")
q = Replace(razon, "'", "")
q = Replace(razon, ".", "")

razonFixed = q

End Property


Public Property Get contacto() As Collection
    Set contacto = mContacto

End Property
Public Property Let contacto(nvalue As Collection)
    Set mContacto = nvalue
End Property
Public Property Get id() As Long
    id = vId
End Property
Public Property Get razon() As String
    razon = vRazon
End Property
Public Property Get Domicilio() As String
    Domicilio = vDomicilio
End Property
Public Property Get telefono() As String
    telefono = vTelefono
End Property
Public Property Get Fax() As String
    Fax = vFax
End Property
Public Property Get email() As String
    email = vEmail
End Property
Public Property Get Cuit() As String
    Cuit = vCuit
End Property
Public Property Get estado() As EstadoCliente
    estado = vEstado
End Property
Public Property Get PasswordSistema() As String
    PasswordSistema = vPasswordSistema
End Property
Public Property Get FP() As Long
    FP = vFP
End Property
Public Property Let id(nvalue As Long)
    vId = nvalue
End Property
Public Property Let razon(nvalue As String)
    vRazon = nvalue
End Property
Public Property Let Domicilio(nvalue As String)
    vDomicilio = nvalue
End Property
Public Property Let telefono(nvalue As String)
    vTelefono = nvalue
End Property
Public Property Let Fax(nvalue As String)
    vFax = nvalue
End Property
Public Property Let email(nvalue As String)
    vEmail = nvalue
End Property
Public Property Let Cuit(nvalue As String)
    vCuit = nvalue
End Property

Public Property Let estado(nvalue As EstadoCliente)
    vEstado = nvalue
End Property
Public Property Let PasswordSistema(nvalue As String)
    vPasswordSistema = nvalue
End Property
Public Property Let FP(nvalue As Long)
    vFP = nvalue
End Property


Public Function CUITValido() As Boolean

    CUITValido = (IsNumeric(Me.Cuit) And Val(Me.Cuit) > 0)

End Function


Public Property Get TipoIVA() As clsTipoIVA
    If Not IsSomething(mTipoIva) Then Set mTipoIva = DAOTipoIva.GetById(Me.TipoIvaID)
    Set TipoIVA = mTipoIva
End Property

Public Property Set TipoIVA(nvalue As clsTipoIVA)
    Set mTipoIva = nvalue
End Property
