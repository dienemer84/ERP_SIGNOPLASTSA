VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeticionOfertaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vDetalleReque As clsRequeMateriales
Private vValor As Double
Private vFechaValor As Date
Private vTerminado As Boolean
Public POid As Long
Dim vId As Long
Public Cantidad As Long
Public Entregas As New Collection    '<EntregaPetOfDetalle>
Public EnOrdenCompra As Boolean    ' lelenarrrrrrrs

Public ProveedorId As Long    'shortcut al proveeedor de la PO
Public moneda As clsMoneda    'shortcut a Moneda de la PO

Public estado As EstadoPeticionOfertaDetalle

Public Property Get Total() As Double

    If vDetalleReque.Material.UnidadCompra = kg_ Then
        Total = vValor * vDetalleReque.Kg
    ElseIf vDetalleReque.Material.UnidadCompra = m2_ Then
        Total = vValor * vDetalleReque.m2
    ElseIf vDetalleReque.Material.UnidadCompra = Ml_ Then
        Total = vValor * vDetalleReque.ML
    Else
        Total = vValor * Cantidad
    End If
End Property


Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get Id() As Long
    Id = vId
End Property
Public Property Let Terminado(nvalue As Boolean)
    vTerminado = nvalue
End Property
Public Property Get Terminado() As Boolean
    Terminado = vTerminado
End Property
Public Property Let DetalleReque(nvalue As clsRequeMateriales)
    Set vDetalleReque = nvalue
End Property
Public Property Let Valor(nvalue As Double)
    vValor = nvalue
End Property
Public Property Let FechaValor(nvalue As Date)
    vFechaValor = nvalue
End Property
Public Property Get DetalleReque() As clsRequeMateriales
    Set DetalleReque = vDetalleReque
End Property
Public Property Get Valor() As Double
    Valor = vValor
End Property
Public Property Get FechaValor() As Date
    FechaValor = vFechaValor
End Property

Public Function IsValidCantidadEnEntregas() As Boolean
    Dim cont As Long
    Dim ent As EntregaPetOfDetalle
    For Each ent In Me.Entregas
        cont = cont + ent.Cantidad
    Next ent
    IsValidCantidadEnEntregas = (cont = Me.Cantidad)
End Function
