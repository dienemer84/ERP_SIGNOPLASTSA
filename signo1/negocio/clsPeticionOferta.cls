VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeticionOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private vModificado As Date
Private vEstado As EstadoPO
Private vUsuarioCreador As clsUsuario
Private vIdReque As Long
Private vNumero As Long
Private vFechaEmision As Date
Public moneda As clsMoneda
Private vFechaSolicitada As Date
Private vProveedor As clsProveedor
Private vDetalle As New Collection    'of peticion oferta detalle
Public FormaDePago As String
Public CantidadDiasPago As Long
Public PorcentajeDescuento As Double
Public EntregaRetiramos As Boolean

Public Property Let Modificado(nvalue As Date)
    vModificado = nvalue
End Property
Public Property Get Modificado() As Date
    Modificado = vModificado
End Property

Public Property Let usuarioCreador(nvalue As clsUsuario)
    Set vUsuarioCreador = nvalue
End Property
Public Property Get usuarioCreador() As clsUsuario
    Set usuarioCreador = vUsuarioCreador
End Property
Public Property Let detalle(nvalue As Collection)
    Set vDetalle = nvalue
End Property
Public Property Let Proveedor(nvalue As clsProveedor)
    Set vProveedor = nvalue
End Property
Public Property Let FechaSolicitada(nvalue As Date)
    vFechaSolicitada = nvalue
End Property
Public Property Let FechaEmision(nvalue As Date)
    vFechaEmision = nvalue
End Property
Public Property Let numero(nvalue As Long)
    vNumero = nvalue
End Property
Public Property Let idReque(nvalue As Long)
    vIdReque = nvalue
End Property
Public Property Get detalle() As Collection
    Set detalle = vDetalle
End Property
Public Property Get Proveedor() As clsProveedor
    Set Proveedor = vProveedor
End Property
Public Property Get FechaSolicitada() As Date
    FechaSolicitada = vFechaSolicitada
End Property
Public Property Get FechaEmision() As Date
    FechaEmision = vFechaEmision
End Property
Public Property Get numero() As Long
    numero = vNumero
End Property
Public Property Get idReque() As Long
    idReque = vIdReque
End Property

Public Property Let estado(nvalue As EstadoPO)
    vEstado = nvalue
End Property
Public Property Get estado() As EstadoPO
    estado = vEstado
End Property


Public Function IsValid() As Boolean
    Dim det As clsPeticionOfertaDetalle
    Dim ret As Boolean

    If Me.detalle Is Nothing Then
        Me.detalle = DAOPeticionOfertaDetalle.FindAll(Me.numero)
    End If

    For Each det In Me.detalle
        ret = (det.Cantidad = 0 Or det.Valor = 0)
        If ret Then
            IsValid = False
            Exit Function
        End If
    Next det

    IsValid = True
End Function

