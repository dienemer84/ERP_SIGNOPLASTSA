VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public descripcion As String

Public Cliente As clsCliente  'centro de costos....
Public ClienteFacturar As clsCliente
Public FechaEntrega As Date
Public FechaCreado As Date
Public NroPresupuesto As Long
Public estado As EstadoOrdenTrabajo
Public Activa As Boolean
Public Entregada As Boolean
Public FechaCerrado As Date
Public usuario As clsUsuario
Public Descuento As Double
Public Moneda As clsMoneda
Public IdMoneda As Long
Public FechaAprobado As Date
Public UsuarioAprobado As clsUsuario
Public FechaModificado As Date
Public UsuarioModificado As clsUsuario
Public UsuarioFinalizado As clsUsuario
Public StockDescontado As Boolean
Public Anticipo As Double
Public AnticipoFacturado As Boolean
Public AnticipoFacturadoIdFactura As Long
Public FormaDePagoAnticipo As String
Public MismaFechaEntregaParaDetalles As Boolean
Public CantDiasAnticipo As Long
Public CantDiasSaldo As Long
Public ContaduriaImpreso As Integer
Public FormaDePagoSaldo As String
Public Detalles As New Collection
Private estadoAnterior As EstadoOrdenTrabajo
Private UsuarioModificadoAnterior As clsUsuario
Public OTMarcoIdPadre As Long    'si es 0 no es marco ni hija, -1 = padre, otro = hija
Public FechaInicioMarco As Date
Public FechaFinMarco As Date
Private m_FechasPreciosMarco As Collection
Public UltimaFechaActualizacionPrecios As Boolean
Public MontoTopeMarco As Double
Public Property Set FechasPreciosMarco(value As Collection)
    Set m_FechasPreciosMarco = value
End Property
Public Property Get FechasPreciosMarco() As Collection
    If Me.EsMarco Then
        If m_FechasPreciosMarco Is Nothing Then
            Set m_FechasPreciosMarco = New Collection
            If Me.id > 0 Then
                Dim r As New Recordset
                Set r = conectar.RSFactory("SELECT fecha FROM pedidos_fechas_precios WHERE id_ot_marco = " & Me.id)
                While Not r.EOF
                    m_FechasPreciosMarco.Add CDate(r!FEcha)
                    r.MoveNext
                Wend
                r.Close
            End If
        End If
    ElseIf m_FechasPreciosMarco Is Nothing Then
        Set m_FechasPreciosMarco = New Collection
    End If
    Set FechasPreciosMarco = m_FechasPreciosMarco
End Property
Public Property Get EsMarco() As Boolean
    EsMarco = (OTMarcoIdPadre = -1)
End Property
Public Property Get EsHija() As Boolean
    EsHija = (OTMarcoIdPadre >= 1)
End Property
Public Property Get NoEsMarcoNiHija() As Boolean
    NoEsMarcoNiHija = (OTMarcoIdPadre = 0)
End Property
Public Property Get TotalFacturado() As Double
    Dim deta As DetalleOrdenTrabajo
    Dim c As Double
    c = 0
    For Each deta In Detalles
        c = c + (deta.Monto_Facturado)
    Next
    TotalFacturado = c
End Property

Private Function CantidadFacturada() As Double
    Dim deta As DetalleOrdenTrabajo
    Dim c As Double
    c = 0
    For Each deta In Detalles
        c = c + (deta.Cantidad_Facturada)
    Next

    CantidadFacturada = c
End Function


'Public Function PorcentajeEntregas() As Double
'        Dim deta As DetalleOrdenTrabajo
'    Dim c As Double, d As Double
'    c = 0
'    d = 0
'    For Each deta In detalles
'        c = c + deta.CantidadEntregada
'        d = d + deta.CantidadPedida
'
'    Next
'
'    If d <> 0 Then
'         PorcentajeEntregas = (c * 100) / d
'    Else
'        PorcentajeEntregas = 0
'    End If
'
'End Function

Public Property Get PorcentajeFacturado() As Double

    Dim valor1 As Double
    Dim valor2 As Double

    'creo qu eestoy haciendo al pedo la conversion a moneda
    valor1 = MonedaConverter.Convertir(Me.Total, Me.Moneda.id, MonedaConverter.Patron.id)
    valor2 = MonedaConverter.Convertir(Me.TotalFacturado, Me.Moneda.id, MonedaConverter.Patron.id)
    If valor1 > 0 Then
        PorcentajeFacturado = funciones.RedondearDecimales((valor2 * 100) / valor1)
    Else
        PorcentajeFacturado = 0
    End If

End Property


Private Sub Class_Initialize()
    Me.FechaCreado = Now
    Me.Activa = True
    Me.NroPresupuesto = -1
    Me.estado = EstadoOrdenTrabajo.EstadoOT_Pendiente
    Set Me.usuario = funciones.GetUserObj
End Sub
Public Property Get IdFormateado() As String
    IdFormateado = Format(Me.id, "0000")
End Property
Public Property Get SubTotal() As Double

    Dim tot As Double: tot = 0
    Dim deta As DetalleOrdenTrabajo

    For Each deta In Me.Detalles
        'tot = tot + (deta.CantidadPedida * deta.Precio)
        tot = tot + deta.TotalSinDescuento
    Next deta

    If Me.EsMarco And tot = 0 Then
        SubTotal = Me.MontoTopeMarco
    Else
        SubTotal = tot
    End If
End Property

Public Property Get Total() As Double
    Dim col As Collection
    Dim tot As Double: tot = 0
    Dim deta As DetalleOrdenTrabajo
    Dim saldo_piezas As Double: saldo_piezas = 0
    Dim saldo_deta As Double
    Dim ca As clsDetalleOrdenTrabajoCantidades


    For Each deta In Detalles

        tot = tot + deta.TotalConDescuento


    Next deta



    If Me.EsMarco And tot = 0 Then    '
        Total = Me.MontoTopeMarco
    Else
        Total = tot
    End If
End Property

Public Property Get CondicionesComerciales() As String
    CondicionesComerciales = "Anticipo: " & Me.CondicionesComercialesAnticipo & ", Saldo: " & Me.CondicionesComercialesSaldo
End Property


Public Property Get CondicionesComercialesAnticipo() As String
    If Anticipo = 0 Then
        CondicionesComercialesAnticipo = "Sin Anticipo"
    Else
        Dim fact As String

        If Me.EsHija Then
            fact = "(Según contrato abierto)"
        Else
            If Not Me.AnticipoFacturado Then
                fact = " (No facturado)"
            Else
                fact = " (Facturado)"
            End If

        End If
        If LenB(Trim(FormaDePagoAnticipo)) = 0 Then
            If CantDiasAnticipo > 0 Then
                CondicionesComercialesAnticipo = Anticipo & "%, a " & CantDiasAnticipo & " días  " & fact
            Else
                CondicionesComercialesAnticipo = Anticipo & "%, Contado" & fact
            End If
        Else
            If CantDiasAnticipo > 0 Then
                CondicionesComercialesAnticipo = Anticipo & "%, a " & CantDiasAnticipo & " días, F.Pago: " & FormaDePagoAnticipo & fact
            Else
                CondicionesComercialesAnticipo = Anticipo & "%, F.Pago: " & FormaDePagoAnticipo & fact
            End If
        End If
    End If
End Property
Public Property Get CondicionesComercialesSaldo() As String
    If FormaDePagoSaldo <> Empty Then
        If CantDiasSaldo > 0 Then
            CondicionesComercialesSaldo = CantDiasSaldo & " días,  F.Pago: " & FormaDePagoSaldo
        Else
            CondicionesComercialesSaldo = "Contado,  F.Pago: " & FormaDePagoSaldo
        End If
    Else
        If CantDiasSaldo > 0 Then
            CondicionesComercialesSaldo = CantDiasSaldo & " días"
        Else
            CondicionesComercialesSaldo = "Contado"
        End If
    End If
End Property
Public Function CantidadPiezasReservadas() As Long
    Dim SUM As Long
    Dim deta As DetalleOrdenTrabajo
    SUM = 0
    For Each deta In Detalles
        SUM = SUM + deta.ReservaStock
    Next
    CantidadPiezasReservadas = SUM
End Function
Public Function ValidarProcesos() As Boolean
    Dim det As DetalleOrdenTrabajo
    Dim ok As Boolean
    ok = True
    For Each det In Detalles
        If det.EstadoProceso = EstProcDetOT_AunNoDefinido Then
            ok = False
            Exit For
        End If
    Next
    ValidarProcesos = ok
End Function
Public Function ValidarReservas() As Boolean
    Dim det As DetalleOrdenTrabajo
    Dim ok As Boolean
    ok = True
    For Each det In Detalles
        If det.Pieza.CantidadStock < det.ReservaStock Then
            ok = False
            Exit For
        End If
    Next
    ValidarReservas = ok
End Function


Public Property Get ProximaFechaActualizacionPrecios() As Date
    Dim fec As Variant
    Dim fechaRet As Date
    For Each fec In FechasPreciosMarco
        If DateDiff("d", Now, fec) >= 0 Then
            If CDbl(fechaRet) = 0 Then fechaRet = fec
            If fec < fechaRet Then
                fechaRet = fec
            End If
        End If
    Next fec
    ProximaFechaActualizacionPrecios = fechaRet
End Property


Public Property Get DetallesValidos() As Boolean
    If Me.NoEsMarcoNiHija Then
        DetallesValidos = True
    Else
        If Me.EsMarco Then

        Else

        End If
    End If
End Property


Public Property Get TodoEntregado() As Boolean
    Dim det As DetalleOrdenTrabajo
    Dim ok As Boolean
    ok = True
    For Each det In Detalles
        If det.CantidadEntregada < det.CantidadPedida Then
            ok = False
            Exit For
        End If
    Next
    TodoEntregado = ok
End Property

Public Property Get TodoFabricado() As Boolean
    Dim det As DetalleOrdenTrabajo
    TodoFabricado = True
    For Each det In Detalles
        TodoFabricado = ((det.CantidadFabricados + det.ReservaStock) >= det.CantidadPedida)
        If Not TodoFabricado Then Exit For
    Next
End Property


Public Property Get PuedeCerrarse() As Boolean
    Dim det As DetalleOrdenTrabajo
    PuedeCerrarse = True
    For Each det In Detalles
        PuedeCerrarse = (det.CantidadPedida <= (det.CantidadFabricados + det.ReservaStock) And (Me.estado = EstadoOT_Finalizado Or Me.estado = EstadoOT_EnProceso))
        If Not PuedeCerrarse Then Exit For
    Next
End Property


Public Property Get PorcentajeAvance() As Double
    Dim det As DetalleOrdenTrabajo
    Dim porc As Double
    Dim porc2 As Double
    Dim prom As Double
    For Each det In Me.Detalles
        DAODetalleOrdenTrabajo.CalcularPorcentajeAvanceYPromedioFabricado det.id, porc, prom
        porc2 = porc2 + porc
    Next

    If Me.Detalles.count > 0 Then
        PorcentajeAvance = funciones.RedondearDecimales(porc2 / Me.Detalles.count)
    Else
        PorcentajeAvance = 0
    End If
End Property

Public Property Get PorcentajeEntregas() As Double
    Dim det As DetalleOrdenTrabajo
    Dim sumaEntregas As Double: sumaEntregas = 0

    For Each det In Me.Detalles
        If det.CantidadPedida <> 0 Then
            sumaEntregas = sumaEntregas + ((det.CantidadEntregada * 100) / det.CantidadPedida)
        End If
    Next

    If Me.Detalles.count > 0 Then
        PorcentajeEntregas = funciones.RedondearDecimales(sumaEntregas / Me.Detalles.count)
    Else
        PorcentajeEntregas = 0
    End If
End Property


Public Property Get SpecificEventDescription() As String
    SpecificEventDescription = Cliente.razon
End Property

Public Property Get PorcentajeFabricacion() As Double
    Dim det As DetalleOrdenTrabajo
    Dim sumaFabricadas As Double: sumaFabricadas = 0

    For Each det In Me.Detalles
        If det.CantidadPedida <> 0 Then
            sumaFabricadas = sumaFabricadas + ((det.CantidadFabricados * 100) / det.CantidadPedida)
        End If
    Next

    If Me.Detalles.count > 0 Then
        PorcentajeFabricacion = funciones.RedondearDecimales(sumaFabricadas / Me.Detalles.count)
    Else
        PorcentajeFabricacion = 0
    End If
End Property


