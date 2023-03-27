VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private deta As clsDetalleOrdenTrabajoCantidades
Public Id As Long
Public item As String
Public OrdenTrabajo As OrdenTrabajo
Public Pieza As Pieza
Public CantidadPedida As Double
Public NombrePiezaHistorico As String
Public Nota As String
Public ReservaStock As Long
Public idpedido As Long
Public IdMoneda As Long
Public Retirado As Boolean
Public Precio As Double
Public FechaEntrega As Date
Public EtiquetasImpresas As Double


Public CantidadEntregada As Double
Public CantidadFacturada As Double
Public CantidadFabricados As Double
Public CantidadEnviadasAStock As Double

Public CantidadEntregadaStatic As Double
Public CantidadFacturadaStatic As Double
Public CantidadFabricadosStatic As Double

Public CantidadTareas As Long
Public CantidadTareasFinalizadas As Long


Public MontoFacturado As Double

Public NotaProduccion As String
Public CantidadImpresionesDeRuta As Long
Public PrecioModificado As Boolean
Public EstadoProceso As EstadoProcesoDetalleOrdenTrabajo


'27-10-09
Public colCantidadesFabricadas As Collection
Public colCantidadesFacturadas As Collection
Public colCantidadesEntregadas As Collection



Public idDetalleOtPadre As Long
Private m_DetallesHijasMarcoPadre As Collection

'28-01-2010
Public idPresupuestoOrigen As Long

'25-2-2010
Public Descuento As Double


Public Property Get DescuentoFormateado() As Double
    DescuentoFormateado = 1 - (Me.Descuento / 100)
End Property
Public Property Get SubTotalConDescuento() As Double
    SubTotalConDescuento = Me.Precio * Me.DescuentoFormateado
End Property


Public Property Get TotalConDescuento() As Double


    Dim saldo_piezas As Double
    Dim tot As Double
    tot = 0
    saldo_piezas = Me.CantidadPedida - Me.Cantidad_Facturada

    If saldo_piezas > 0 Then
        tot = (saldo_piezas * (1 - (Me.Descuento / 100)) * (MonedaConverter.Convertir(Precio, Me.IdMoneda, MonedaConverter.Patron.Id)))
    End If

    TotalConDescuento = funciones.RedondearDecimales(Me.Monto_Facturado + tot)

End Property
Public Property Get TotalAFacturar() As Double
    TotalAFacturar = Me.TotalConDescuento - Me.Monto_Facturado
End Property
Public Property Get TotalSinDescuento() As Double
'    TotalSinDescuento = Me.Precio * Me.CantidadPedida
    Dim saldo_piezas As Double
    Dim tot As Double
    tot = 0
    saldo_piezas = Me.CantidadPedida - Me.Cantidad_Facturada
    If saldo_piezas > 0 Then
        tot = (saldo_piezas * Me.Precio)
    End If

    TotalSinDescuento = funciones.RedondearDecimales(Me.Monto_Facturado + tot)
End Property

Public Property Get Cantidad_Fabricada() As Double
    Dim c As Double
    c = 0

    'If Me.idDetalleOtPadre = -1 Then
    Dim deta2 As DetalleOrdenTrabajo
    If Me.DetallesHijasMarcoPadre.count > 0 Then

        For Each deta2 In Me.DetallesHijasMarcoPadre
            c = c + deta2.Cantidad_Fabricada
        Next deta2
    Else
        'For Each deta In colCantidadesFabricadas
        '     c = c + CDbl(deta.cantidad)
        '  Next
        'End If
        c = c + Me.CantidadFabricados
    End If
    ' Else

    'For Each deta In colCantidadesFabricadas
    '     c = c + CDbl(deta.cantidad)
    '  Next
    '   c = c + Me.CantidadFabricados
    ' End If
    Cantidad_Fabricada = c
End Property
Public Property Get Cantidad_Entregada() As Double
    Dim c As Double
    c = 0
    Dim deta2 As DetalleOrdenTrabajo
    '  If Me.idDetalleOtPadre = -1 Then
    If IsSomething(Me.DetallesHijasMarcoPadre) Then
        If Me.DetallesHijasMarcoPadre.count > 0 Then
            For Each deta2 In Me.DetallesHijasMarcoPadre
                c = c + deta2.Cantidad_Entregada
            Next deta2
        Else
            'For Each deta In colCantidadesEntregadas
            '     c = c + CDbl(deta.cantidad)
            '  Next
            c = c + Me.CantidadEntregada
        End If
    Else
        For Each deta In colCantidadesEntregadas
            c = c + CDbl(deta.Cantidad)
        Next
        c = c + Me.CantidadEntregada
    End If

    Cantidad_Entregada = c
End Property
Public Property Get Monto_Facturado() As Double
    Dim c As Double
    c = 0
    Dim deta2 As DetalleOrdenTrabajo



    '   If Me.idDetalleOtPadre = -1 Then

    If IsSomething(Me.DetallesHijasMarcoPadre) Then
        If Me.DetallesHijasMarcoPadre.count > 0 Then
            For Each deta2 In Me.DetallesHijasMarcoPadre
                c = c + deta2.Monto_Facturado
            Next deta2

        Else
            c = c + funciones.RedondearDecimales(Me.MontoFacturado * (1 - (Me.Descuento / 100)))
        End If
    Else
        c = c + funciones.RedondearDecimales(Me.MontoFacturado * (1 - (Me.Descuento / 100)))

    End If
    Monto_Facturado = c
End Property
Public Property Get Cantidad_Facturada() As Double
    Dim c As Double
    c = 0
    Dim deta2 As DetalleOrdenTrabajo


    'If Me.idDetalleOtPadre = -1 Then
    If IsSomething(Me.DetallesHijasMarcoPadre) Then
        If Me.DetallesHijasMarcoPadre.count > 0 Then
            For Each deta2 In Me.DetallesHijasMarcoPadre
                c = c + deta2.Cantidad_Facturada
            Next deta2
        Else
            c = c + Me.CantidadFacturada
        End If
    Else

        c = c + Me.CantidadFacturada
    End If
    Cantidad_Facturada = c



End Property


Public Property Get DetallesHijasMarcoPadre() As Collection
    If m_DetallesHijasMarcoPadre Is Nothing And Me.Id > 0 Then

        Set m_DetallesHijasMarcoPadre = DAODetalleOrdenTrabajo.FindAll("dp.IdDetalleOtPadre = " & Me.Id, True, True, True)

    End If
    Set DetallesHijasMarcoPadre = m_DetallesHijasMarcoPadre
End Property



Public Property Get CantidadConsumida() As Double
    Dim strsql As String
    Dim rs As Recordset
    Dim deta2 As DetalleOrdenTrabajo
    Dim T As Double
    For Each deta2 In Me.DetallesHijasMarcoPadre

        'If LenB(deta2.OrdenTrabajo.Descripcion) = 0 Then

        strsql = "Select estado,descripcion from pedidos where id=" & deta2.OrdenTrabajo.Id
        Set rs = conectar.RSFactory(strsql)

        deta2.OrdenTrabajo.descripcion = rs!descripcion
        deta2.OrdenTrabajo.estado = rs!estado
        '  End If

        If deta2.OrdenTrabajo.estado = EstadoOT_EnProceso Or deta2.OrdenTrabajo.estado = EstadoOT_Finalizado Then
            T = T + deta2.CantidadPedida
        End If
    Next
    CantidadConsumida = T
End Property



Public Function MarcoCantidadDisponibles() As Double

    MarcoCantidadDisponibles = Me.CantidadPedida - CantidadConsumida

End Function
