VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public esCredito As Boolean
Public Id As Long
Public Cancelada As Long
Public FechaAprobacion As Date
Public MotivoNC As Integer
Public Historial As New Collection
Public origenFacturado As origenFacturado
Public FechaPropuestaPago As Date    'propuesta
Public numero As String
Public cliente As clsCliente
Public moneda As clsMoneda
Public CambioAPatron As Double
Public AprobadaAFIP As Boolean
Public FechaEntrega As Date
Public CAE As String
Public CAEVto As Date
Public CAEFechaProceso
Public Tipo As clsTipoFacturaDiscriminado
Public TipoIVA As clsTipoIVA
Public AnulacionAFIP As String
Public MotivosAnulacionAFIP As String
Public CaeSolicitarResponse As CAESolicitar
'agregado 31052020
Public ConceptoIncluir As ConceptoIncluir
'agregado 25-6-12
Public IdMonedaAjuste As Long
Public TipoCambioAjuste As Double

Public FechaEmision As Date
Public observaciones As String
'#197
Public observaciones_cancela As String
Public EstaImpresa As Boolean
Public EstaDiscriminada As Boolean
'Public TipoDocumento As tipoDocumentoContable
Public Saldado As TipoSaldadoFactura
Public AlicuotaAplicada As Double    'iva
Public AlicuotaPercepcionesIIBB As Double
Public estado As EstadoFacturaCliente
Public usuarioCreador As clsUsuario
Public CantDiasPago As Long    'FormaPago
Public OrdenCompra As String
Public UsuarioAprobacion As clsUsuario
Public Opcional27 As Integer
Public TextoAdicional As String


Public FechaVencimientoSQL As Date    'Fecha de vencimiento calculada por la base
Public TotalEstatico As New FacturaTotalEstatico

Private m_ots_facturadas_anticipo As Collection
Private m_detalles As Collection

Public fechaPago As Date    'para fce
Public CBU As String

'fce_nemer_28052020
Public FechaVtoDesde As Date
Public FechaVtoHasta As Date

'fce_nemer_02062020_#113
Public FechaServDesde As Date
Public FechaServHasta As Date

Public RecibosAplicadosId As String
Public MontoCobrado As Double

Public TasaAjusteMensual As Double  'de aca se empieza

Public Function Modificable() As Boolean

    Modificable = Not (Tipo.PuntoVenta.EsElectronico And Me.AprobadaAFIP)

End Function

Public Sub SetToNothingOTsFacturadasAnticipo()
    Set m_ots_facturadas_anticipo = Nothing
End Sub

Public Property Get DescripcionCreditoAdicional() As String
    Dim dsc As String


    If TipoDocumento = tipoDocumentoContable.Factura Then
        dsc = "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"
    End If
    If TipoDocumento = tipoDocumentoContable.notaCredito Then
        dsc = "NOTA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)"
    End If
    If TipoDocumento = tipoDocumentoContable.notaDebito Then
        dsc = "NOTA DE DÉBITO ELECTRÓNICA MiPyMEs  (FCE)"
    End If
    DescripcionCreditoAdicional = dsc
End Property

Public Property Get CodigoBarrasAfip() As String
    Dim numero As String
    numero = "30657604972" & Format(Me.GetCodigoDocumentoAfip, "00") & Format(Me.Tipo.PuntoVenta.PuntoVenta, "0000") & Me.CAE & Format(Me.CAEVto, "yyyymmdd")


    CodigoBarrasAfip = numero & calcularDVFactura(numero)
End Property


Public Property Get DetalleFacturaAnticipoOt(porcentajeAnticipo As Double) As FacturaDetalle
    Dim deta As FacturaDetalle

    Set DetalleFacturaAnticipoOt = Nothing




    For Each deta In Me.Detalles
        'cabeceada #1 para saber si es del mismo porcentaje de anticipo
        If deta.DescuentoAnticipo And InStr(1, deta.detalle, porcentajeAnticipo & "%") > 0 Then
            Set DetalleFacturaAnticipoOt = deta
            Exit Property
        End If
    Next deta
End Property

Public Property Get OTsFacturadasAnticipo() As Collection
    If Not IsSomething(m_ots_facturadas_anticipo) Then
        If Me.Id <> 0 Then
            Set m_ots_facturadas_anticipo = DAOOrdenTrabajo.FindAll("p.id_anticipo_factura = " & Me.Id)
        Else
            Set m_ots_facturadas_anticipo = New Collection
        End If
    End If

    Set OTsFacturadasAnticipo = m_ots_facturadas_anticipo
End Property

'Public Property Set OTsFacturadasAnticipo(value As Collection)
'    Set m_ots_facturadas_anticipo = value
'End Property



Public Function GetCodigoDocumentoAfip() As Integer
    If Tipo.TipoDoc = tipoDocumentoContable.notaDebito Then GetCodigoDocumentoAfip = 2
    If Tipo.TipoDoc = tipoDocumentoContable.notaCredito Then GetCodigoDocumentoAfip = 3
    If Tipo.TipoDoc = tipoDocumentoContable.Factura Then GetCodigoDocumentoAfip = 1


End Function
Public Function GetDescripciopnDocumentoAfip() As String
    If Tipo.TipoDoc = tipoDocumentoContable.notaDebito Then GetDescripciopnDocumentoAfip = "NOTA DE DEBITO"
    If Tipo.TipoDoc = tipoDocumentoContable.notaCredito Then GetDescripciopnDocumentoAfip = "NOTA DE CREDITO"
    If Tipo.TipoDoc = tipoDocumentoContable.Factura Then GetDescripciopnDocumentoAfip = "FACTURA"


End Function

Public Function getDescripcionCondicion() As String

    Dim A As String
    If Me.CantDiasPago = 1 Then
        A = Me.CantDiasPago & " dia FF"
    Else
        A = Me.CantDiasPago & " dias FF"
    End If


    getDescripcionCondicion = A & Me.observaciones

End Function


Public Function DiferenciaDiasEntrega() As Integer
    On Error GoTo err1
    If (IsDate(Me.FechaEntrega)) Then
        DiferenciaDiasEntrega = DateDiff("d", Me.FechaEntrega, Date)
    Else
        DiferenciaDiasEntrega = -1
    End If
    Exit Function
err1:
    DiferenciaDiasEntrega = -1
End Function



Public Property Get EsAnticipo() As Boolean
    EsAnticipo = Me.OTsFacturadasAnticipo.count > 0
End Property

Public Property Get Detalles() As Collection
'lazy load simulation
    If m_detalles Is Nothing And Me.Id <> 0 Then
        Set m_detalles = DAOFacturaDetalles.FindByFactura(Me.Id)
    End If

    Set Detalles = m_detalles
End Property


Public Property Let Detalles(nvalue As Collection)
    Set m_detalles = nvalue
End Property
Public Property Get EstaAtrasada() As Boolean
    EstaAtrasada = (Me.Saldado = SaldadoParcial Or Me.Saldado = NoSaldada) And ((Now - Vencimiento) > 0)
End Property
Public Property Get StringDiasAtraso() As String
    Dim atraso As Long


    atraso = DateDiff("d", Vencimiento, Date)


    'atraso = Date - Vencimiento

    If Me.estado = EstadoFacturaCliente.EnProceso Then
        StringDiasAtraso = "En Edición"
    ElseIf Me.estado = EstadoFacturaCliente.Anulada Then
        StringDiasAtraso = "Anulada"
    Else
        If Me.Saldado = SaldadoParcial Or Me.Saldado = NoSaldada Then
            StringDiasAtraso = atraso & " Días"
        Else
            If Me.Saldado = notaCreditoParcial Then
                StringDiasAtraso = "Saldada NC parcial"
            Else

                StringDiasAtraso = "Saldada"
            End If
        End If

        If atraso < 0 And Me.estado <> EstadoFacturaCliente.Aprobada Then
            StringDiasAtraso = "Faltan " & (atraso * -1) & " días"
        End If

    End If


End Property
Public Property Get NumeroFormateado() As String
    If (IsSomething(Tipo.PuntoVenta)) Then
        '   If Tipo.PuntoVenta.EsElectronico And Me.estado <> EstadoFacturaCliente.Aprobada Then
        '     NumeroFormateado = Me.Tipo.PuntoVenta.PuntoVenta & "- 0000"
        ' Else
        NumeroFormateado = Me.Tipo.PuntoVenta.PuntoVenta & "-" & Format(Me.numero, "0000")

        ' End If
    Else
        NumeroFormateado = "Error"
    End If
End Property


Public Property Get Vencimiento() As Date
    If IsDate(Me.FechaEntrega) Then
        Vencimiento = DateAdd("d", Me.CantDiasPago, Me.FechaEntrega)
    Else
        Vencimiento = DateAdd("d", Me.CantDiasPago, Me.FechaEmision)
    End If
End Property

Public Property Get EstaVencida() As Boolean
    EstaVencida = (Date > Vencimiento)
End Property

Public Property Get MostrarConcepto() As String
    If Me.ConceptoIncluir = 1 Then MostrarConcepto = "Productos"
    If Me.ConceptoIncluir = 2 Then MostrarConcepto = "Servicios"
    If Me.ConceptoIncluir = 3 Then MostrarConcepto = "Productos y Servicios"

End Property
Public Property Get total() As Double
    Dim totIva As Double
    Dim totPercepciones As Double
    Dim totSub As Double
    'Debug.Assert Me.numero <> 9126
    totIva = Me.TotalIVA
    totPercepciones = Me.totalPercepciones
    totSub = Me.TotalSubTotal
    total = funciones.RedondearDecimales(totIva + totPercepciones + totSub)
End Property

Public Property Get TipoDocumentoDescription() As String

    Dim retorno As String
    Select Case Me.TipoDocumento
    Case tipoDocumentoContable.Factura
        retorno = "FACTURA"
    Case tipoDocumentoContable.notaCredito
        retorno = "NOTA DE CREDITO"
    Case tipoDocumentoContable.notaDebito
        retorno = "NOTA DE DEBITO"
    End Select

    TipoDocumentoDescription = retorno
End Property

Public Property Get TipoDocumento() As tipoDocumentoContable

    TipoDocumento = Tipo.TipoDoc



End Property

Public Property Get GetShortDescription(justTipo As Boolean, fullShortDescription As Boolean) As String

    Dim retorno As String
    Select Case Me.TipoDocumento
    Case tipoDocumentoContable.Factura
        retorno = "FC"
    Case tipoDocumentoContable.notaCredito
        retorno = "NC"
    Case tipoDocumentoContable.notaDebito
        retorno = "ND"
    End Select

    If Not justTipo Then
        If IsSomething(Tipo) And IsSomething(Tipo.PuntoVenta) Then retorno = retorno & "-" & Me.Tipo.TipoFactura.Tipo & " | " & Me.Tipo.PuntoVenta.PuntoVenta & "-" & Format(Me.numero, "0000")
    End If

    If fullShortDescription Then
        If (estado = Anulada) Then
            retorno = retorno & "(A)"
        End If

        If (estado = CanceladaNC) Then
            retorno = retorno & "(CNC)"
        End If

    End If


    GetShortDescription = retorno
End Property

'suma de items con descuento, iva, etc

Public Property Get TotalNetoGravado()
    Dim det As FacturaDetalle
    Dim subT As Double

    'Debug.Assert Me.Numero <> 9124

    For Each det In Me.Detalles
        If det.IvaAplicado And Me.AlicuotaAplicada > 0 And Not Tipo.TipoFactura.ExcentoIVA Then subT = subT + funciones.RedondearDecimales(det.NetoGravado * det.Cantidad)
    Next det
    TotalNetoGravado = subT

End Property



Public Property Get TotalSubTotal() As Double

    Dim det As FacturaDetalle
    Dim subT As Double



    For Each det In Me.Detalles
        subT = subT + det.total
        'Debug.Print det.Total
    Next det

    TotalSubTotal = subT
End Property



Public Property Get TotalIVA() As Double
    TotalIVA = CalcularIva
End Property

Public Property Get TotalExento() As Double
'    TotalExento = CalcularIva(False)
    Dim det As FacturaDetalle
    Dim subT As Double
    For Each det In Me.Detalles
        If Not det.IvaAplicado Or Me.AlicuotaAplicada = 0 Or Tipo.TipoFactura.ExcentoIVA Then subT = subT + funciones.RedondearDecimales(det.NetoGravado * det.Cantidad)
    Next det
    TotalExento = subT
End Property

'muestra el iva, no importa si la fact esta discriminada o no, util para subdiario
Public Property Get TotalIVADiscrimandoONo() As Double

'   Debug.Assert Me.Numero <> "9144"

    Dim det As FacturaDetalle
    Dim totIva As Double
    For Each det In Me.Detalles
        If det.IvaAplicado Then
            totIva = totIva + funciones.RedondearDecimales(det.Cantidad * funciones.RedondearDecimales(det.NetoGravado) * (Me.AlicuotaAplicada / 100))
        End If
    Next det
    TotalIVADiscrimandoONo = totIva
End Property


Private Property Get CalcularIva(Optional conIva As Boolean = True) As Double
    Dim det As FacturaDetalle
    Dim totIva As Double
    'Debug.Assert Me.Numero <> "9144"

    If Me.EstaDiscriminada Then
        For Each det In Me.Detalles
            If conIva And det.IvaAplicado Then
                totIva = totIva + funciones.RedondearDecimales(det.total * (Me.AlicuotaAplicada / 100))
            End If
        Next det
    Else
        totIva = 0
    End If

    CalcularIva = totIva
End Property

Public Property Get totalPercepciones() As Double
    Dim det As FacturaDetalle
    Dim totPercep As Double

    If Me.estado = EstadoFacturaCliente.Aprobada Then
        totPercep = Me.TotalEstatico.TotalPercepcionesIB
    Else

        For Each det In Me.Detalles
            If det.IBAplicado Then
                'Debug.Print det.Total, Me.AlicuotaPercepcionesIIBB, (det.Total * (Me.AlicuotaPercepcionesIIBB - 1))
                If Me.AlicuotaPercepcionesIIBB > 1 Then
                    If Me.FechaEmision > #12/21/2010# Then    'desde el 21 de dic de 2010
                        ' bug #1 cambie > pór >= el 22-11-16 para poder hacer una factura solo de percepciones (ver FC grupo via del 22-11-16)
                        If Me.TotalNetoGravado >= MonedaConverter.Convertir(0, 0, Me.moneda.Id) Then
                            If det.IvaAplicado Then
                                totPercep = totPercep + funciones.RedondearDecimales(det.NetoGravado * det.Cantidad * (Me.AlicuotaPercepcionesIIBB - 1))
                            Else
                                totPercep = totPercep + funciones.RedondearDecimales(det.NetoGravado * det.Cantidad * (Me.AlicuotaPercepcionesIIBB - 1))
                            End If
                        End If
                    Else
                        totPercep = totPercep + funciones.RedondearDecimales(det.total * (Me.AlicuotaPercepcionesIIBB - 1))
                    End If
                End If
            End If
        Next det
    End If
    totalPercepciones = totPercep
End Property
Public Property Get AlicuotaCorrespondiente() As Double
    AlicuotaCorrespondiente = 1 + (Me.AlicuotaAplicada / 100)
End Property

Public Function DetallesMismaOT() As Long
    Dim d As FacturaDetalle
    Dim Id As Long
    For Each d In Detalles
        If Not d.OrigenEsConcepto Then
            If Id = 0 Then Id = d.detalleRemito.idpedido
            If Id <> d.Id Then GoTo no1
        End If
    Next
    DetallesMismaOT = Id
    Exit Function
no1:
    DetallesMismaOT = 0
End Function



Public Function AnticipoDescontado() As Boolean

    If IsSomething(Me.Detalles) Then
        Dim d As FacturaDetalle
        For Each d In Me.Detalles
            If d.DescuentoAnticipo Then GoTo e1
        Next

        AnticipoDescontado = False
    End If

    Exit Function
e1:
    AnticipoDescontado = True

End Function


Public Property Get DetalleAnticipoOT(OtId As Long) As FacturaDetalle
    Set DetalleAnticipoOT = Nothing
    Dim d As FacturaDetalle
    For Each d In Me.Detalles
        If (d.OtIdAnticipo = OtId) Then
            Set DetalleAnticipoOT = d
            Exit Property
        End If
    Next
End Property

Public Sub RemoveDetallesAnticipoOT()
'borro todos los anticipos
    Dim i As Long: i = 0
    Dim detalle As FacturaDetalle
    For Each detalle In Me.Detalles
        i = i + 1
        If detalle.OtIdAnticipo <> 0 Then
            Me.Detalles.remove i
        End If
    Next
End Sub

