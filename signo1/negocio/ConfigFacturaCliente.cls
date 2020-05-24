VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigFacturaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Iva As clsTipoIVA
Public TipoFactura As clsTipoFactura
Public Moneda As clsMoneda
Public Cliente As clsCliente

Public numero As Long
Public Descuento As Double
Public FechaEmision As Boolean
Public idUsuarioEmision As clsUsuario
Public FechaAprobacion As Boolean
Public origenFacturado As Integer
Public OrdenCompra As String
Public estado As Integer
Public AlicuotaAplicada As Double
Public Discriminada As Boolean
Public VecesImpresa As Long
Public Tipo As Integer
Public Saldada As Double
Public Observaciones As String
Public AlicuotaPercIB As Double
Public CambioPatron As Double
Public FormaPago As Long
Public PropuestaPago As Date
Public Cancelada As Long
Public MotivoNC As Integer
Public m_cliente As clsCliente
Public m_ConfigFactura As ConfigFacturaCliente
Public m_Moneda As clsMoneda
Public m_Usuario As clsUsuario
Public idUsuarioAprobacion As clsUsuario




