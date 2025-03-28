VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFacturaProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vId As Long
Dim vNumero As String
Dim vProveedor As clsProveedor
Dim vFecha As Date
Dim vIvaAplicado As New Collection    'vIvaAplicado(0) creo que se usa por ejemplo para conceptos no gravados
Dim vPercep As New Collection
Dim vMonto As Double
Dim vImpuestoInterno As Double
Dim vRedondeo As Double
Dim vEstado As EstadoFacturaProveedor
Dim vConfigFactura As clsConfigFacturaProveedor
Dim vHistorial As New Collection
Dim vCuentasContables As New Collection
Public moneda As clsMoneda
Public tipoDocumentoContable As tipoDocumentoContable
Public FormaPagoCuentaCorriente As Boolean
'Public ConceptoNoGravado As Double
Public OtrosAbonadoGlobal As Double

Public NetoGravadoAbonadoGlobal As Double    'dto para saber cuanto se pago de la factura en todsa las OP

Public TotalAbonadoGlobalPendiente As Double    'dto para saber cuanto se pago de la factura en todsa las OP
Public NetoGravadoAbonadoGlobalPendiente As Double    'dto para saber cuanto se pago de la factura en todsa las OP
Public OtrosAbonadoGlobalPendiente As Double


Public OrdenesPagoId As String
Public LiquidacionesCajaId As String


Public OrdenPagoID As Long
Public LiquidacionCajaId As Long

Public NumeroLiqCaja As Long

Public TotalAbonado As Double
Dim vTipoCambio As Double
Public TotalPago As Double
Public TipoCambioPago As Double
Public UltimaActualizacion As Date
Public ImporteTotalAbonado As Double    'dto para poder hacer pagos parciales
Public NetoGravadoAbonado As Double    'dto para poder hacer pagos parciales
Public OtrosAbonado As Double

Public UsuarioCarga As clsUsuario


Public TotalCompensado As Double


Public Property Get TotalAbonadoGlobal() As Double    'dto para saber cuanto se pago de la factura en todsa las OP
    TotalAbonadoGlobal = funciones.FormatearDecimales(Me.OtrosAbonadoGlobal + Me.NetoGravadoAbonadoGlobal + Me.NetoGravadoAbonadoGlobalPendiente + Me.OtrosAbonadoGlobalPendiente)
End Property


Public Property Get ImporteTotalSaldo()
    ImporteTotalSaldo = funciones.FormatearDecimales(total - Me.TotalAbonadoGlobal - Me.TotalAbonadoGlobalPendiente)
End Property


Public Property Get ImporteNetoGravadoSaldo()
    ImporteNetoGravadoSaldo = funciones.FormatearDecimales(Me.NetoGravado - Me.NetoGravadoAbonadoGlobal - Me.NetoGravadoAbonadoGlobalPendiente)    'Me.NetoGravado - Me.NetoGravadoAbonadoGlobal - Me.NetoGravadoAbonadoGlobalPendiente)
End Property


Public Property Get ImporteOtrosSaldo()
    ImporteOtrosSaldo = funciones.FormatearDecimales(Me.TotalOtros - Me.OtrosAbonadoGlobal - Me.OtrosAbonadoGlobalPendiente)
End Property


Public Property Get TotalOtros()
    TotalOtros = Me.total - Me.NetoGravado
End Property

Public Function CalcularIvaTotalAbonadoParcial(ng As Double) As Double
    Dim tot_iva As Double
    tot_iva = 0
    Dim P As Long
    For P = 1 To vIvaAplicado.count
        'el iva 0% (exento), no suma por el valor 0 de la alicuota

        If (vIvaAplicado(P).alicuota.alicuota <> 0) Then
            tot_iva = tot_iva + ((ng * 1 + (vIvaAplicado(P).alicuota.alicuota / 100)) - ng)
        End If
    Next
    CalcularIvaTotalAbonadoParcial = tot_iva
End Function


Public Function CalcularTotalAbonadoParcial(ng As Double)
    Dim Iva As Double

    Iva = CalcularIvaTotalAbonadoParcial(ng)

    CalcularTotalAbonadoParcial = ng + Iva
End Function


Public Property Get DiferenciaPorTipoDeCambionIVA() As Double
    If TipoCambio > 0 Then
        DiferenciaPorTipoDeCambionIVA = funciones.RedondearDecimales(Me.TotalIVA - Me.TotalIVADiaPago)
    Else
        DiferenciaPorTipoDeCambionIVA = 0
    End If
End Property

Public Property Get DiferenciaPorTipoDeCambionTOTAL() As Double
    If TipoCambio > 0 Then
        DiferenciaPorTipoDeCambionTOTAL = funciones.RedondearDecimales(Me.total - Me.TotalDiaPago)
    Else
        DiferenciaPorTipoDeCambionTOTAL = 0
    End If
End Property

Public Property Get DiferenciaPorTipoDeCambionNG() As Double
    If TipoCambio > 0 Then
        '  DiferenciaPorTipoDeCambio = funciones.RedondearDecimales((TipoCambio - TipoCambioPago) * NetoGravado / TipoCambio, 2)
        DiferenciaPorTipoDeCambionNG = funciones.RedondearDecimales(Me.NetoGravado - Me.NetoGravadoDiaPago)
    Else
        DiferenciaPorTipoDeCambionNG = 0
    End If
End Property
Public Property Get NumeroFormateado() As String
    NumeroFormateado = enums.EnumTipoDocumentoContableShort(tipoDocumentoContable) & "-" & vConfigFactura.TipoFactura & "-" & vNumero
End Property

Public Property Get NumeroFormateadoCorto() As String
    NumeroFormateadoCorto = enums.EnumTipoDocumentoContableShort(tipoDocumentoContable) & "-" & vConfigFactura.TipoFactura
End Property


Public Property Let Historial(nvalue As Collection)
    Set vHistorial = nvalue
End Property

Public Property Get Historial() As Collection
    Set Historial = vHistorial
End Property

Public Property Let cuentasContables(nvalue As Collection)
    Set vCuentasContables = nvalue
End Property
Public Property Get cuentasContables() As Collection
    If vCuentasContables Is Nothing Then
        Set vCuentasContables = DAOCuentasFacturas.GetByFactura(vId)
    End If

    Set cuentasContables = vCuentasContables
End Property
Public Property Let configFactura(nvalue As clsConfigFacturaProveedor)
    Set vConfigFactura = nvalue
End Property
Public Property Get configFactura() As clsConfigFacturaProveedor
    Set configFactura = vConfigFactura
End Property
Public Property Let estado(nEstado As EstadoFacturaProveedor)
    vEstado = nEstado
End Property
Public Property Get estado() As EstadoFacturaProveedor
    estado = vEstado
End Property
Public Property Let redondeo(nRedondeo As Double)
    vRedondeo = nRedondeo
End Property
Public Property Get redondeo() As Double
    redondeo = vRedondeo
End Property
Public Property Let id(nId As Long)
    vId = nId
End Property
Public Property Get id() As Long
    id = vId
End Property
Public Property Let Proveedor(nProveedor As clsProveedor)
    Set vProveedor = nProveedor
End Property
Public Property Get Proveedor() As clsProveedor
    Set Proveedor = vProveedor
End Property
Public Property Let Monto(nMonto As Double)
    vMonto = nMonto
End Property
Public Property Get Monto() As Double
    Monto = vMonto
End Property
Public Property Let numero(nNumero As String)
    vNumero = nNumero
End Property
Public Property Get numero() As String
    numero = vNumero

End Property
Public Property Let FEcha(nFecha As Date)
    vFecha = nFecha
End Property
Public Property Get FEcha() As Date
    FEcha = vFecha
End Property

Public Property Let percepciones(nPercep As Collection)
    Set vPercep = nPercep
End Property
Public Property Get percepciones() As Collection
    Set percepciones = vPercep
End Property
Public Property Let ImpuestoInterno(nImpuesto As Double)
    vImpuestoInterno = nImpuesto
End Property

Public Property Get ImpuestoInterno() As Double
    ImpuestoInterno = vImpuestoInterno
End Property
Public Property Let IvaAplicado(nIvaAplicado As Collection)
    Set vIvaAplicado = nIvaAplicado
End Property
Public Property Get IvaAplicado() As Collection
    Set IvaAplicado = vIvaAplicado
End Property

Public Property Get totalPercepciones() As Double
    Dim tot_per As Double
    tot_per = 0
    Dim P As Long
    For P = 1 To vPercep.count
        tot_per = tot_per + vPercep(P).Monto
    Next
    totalPercepciones = tot_per
End Property

Public Property Get TotalPercepcionesDiscriminado(percepcionId As Long) As Double
    Dim tot_per As Double
    tot_per = 0
    Dim P As Long
    For P = 1 To vPercep.count
        If percepcionId = vPercep(P).Percepcion.id Then
            tot_per = tot_per + vPercep(P).Monto
        End If
    Next
    TotalPercepcionesDiscriminado = tot_per
End Property

Public Property Get TotalIVA() As Double
    Dim tot_iva As Double
    tot_iva = 0
    Dim P As Long
    For P = 1 To vIvaAplicado.count
        'el iva 0% (exento), no suma por el valor 0 de la alicuota

        If (vIvaAplicado(P).alicuota.alicuota <> 0) Then
            tot_iva = tot_iva + (vIvaAplicado(P).Monto * (vIvaAplicado(P).alicuota.alicuota / 100))
        End If
    Next
    TotalIVA = funciones.RedondearDecimales(tot_iva)



End Property


'no se usa creo
Public Property Get TotalNetoGravadoDiscriminado(alicuota As Double) As Double
    Dim tot_ng As Double
    tot_ng = 0
    Dim A As clsAlicuotaAplicada

    For Each A In Me.IvaAplicado
        If A.alicuota.alicuota = alicuota Then
            tot_ng = tot_ng + (A.Monto)
        End If
    Next

    TotalNetoGravadoDiscriminado = tot_ng

End Property

Public Property Get TotalIVADiscriminado(alicuota As Double) As Double
    Dim tot_iva As Double
    tot_iva = 0
    Dim A As clsAlicuotaAplicada

    For Each A In Me.IvaAplicado
        If A.alicuota.alicuota = alicuota Then
            If alicuota = 0 Then
                tot_iva = tot_iva + (A.Monto)
            Else
                tot_iva = tot_iva + (A.Monto * (A.alicuota.alicuota / 100))
            End If
        End If
    Next

    TotalIVADiscriminado = tot_iva
End Property

Public Property Get ObtenerTotalCompensado() As Double



End Property

Public Property Get total() As Double
    Dim tipo_cambio As Double
    Dim negativo As Integer
    Dim CalculoPercepciones As Double

    On Error GoTo err1

    If tipoDocumentoContable = 1 Then
        negativo = -1
    Else
        negativo = 1
    End If


    If Me.moneda.id = DAOMoneda.FindFirstByPatronOrDefault.id Then
        tipo_cambio = 1
    Else
        tipo_cambio = Me.TipoCambio
    End If

    'cambie la division de tipo de cambio por la operacion de arriba el 31-7-13
    'Total = funciones.RedondearDecimales(NetoGravado + vImpuestoInterno + vRedondeo + TotalIVA + (totalPercepciones / Me.TipoCambio)) '+ Me.ConceptoNoGravado) ' + Me.TotalIVADiscriminado(0))

    total = funciones.RedondearDecimales(NetoGravado + vImpuestoInterno + vRedondeo + TotalIVA + (totalPercepciones / tipo_cambio))     '+ Me.ConceptoNoGravado) ' + Me.TotalIVADiscriminado(0))
    'MsgBox (Total)
    Exit Property
err1:
    total = 0
End Property
Public Property Get TotalDiaPago() As Double
    TotalDiaPago = funciones.RedondearDecimales(((total / Me.TipoCambio) * Me.TipoCambioPago))
End Property

Public Property Get TotalDiaPagoAbonado() As Double
    TotalDiaPagoAbonado = funciones.RedondearDecimales(((Me.TotalAbonado / Me.TipoCambio) * Me.TipoCambioPago))
End Property

Public Property Get TipoCambio() As Double
    If vTipoCambio = 0 Then
        TipoCambio = 1
    Else
        TipoCambio = vTipoCambio
    End If
End Property

Public Property Let TipoCambio(nTipoCambio As Double)
    vTipoCambio = nTipoCambio
End Property


Public Property Get TotalIVADiaPago() As Double
    TotalIVADiaPago = funciones.RedondearDecimales(((TotalIVA / Me.TipoCambio) * Me.TipoCambioPago), 2)

End Property

Public Property Get NetoGravadoDiaPago() As Double
    NetoGravadoDiaPago = funciones.RedondearDecimales(((NetoGravado / Me.TipoCambio) * Me.TipoCambioPago), 2)
End Property
Public Property Get NetoGravadoAbonadoDiaPago() As Double
    NetoGravadoAbonadoDiaPago = funciones.RedondearDecimales(((NetoGravadoAbonado / Me.TipoCambio) * Me.TipoCambioPago), 2)
End Property

Public Property Get NetoNoGravado() As Double
    Dim tot_ng As Double
    tot_ng = 0
    Dim P As Long
    Dim A As clsAlicuotaAplicada

    For P = 1 To vIvaAplicado.count
        Set A = vIvaAplicado(P)
        If A.alicuota.alicuota = 0 Then
            tot_ng = tot_ng + vIvaAplicado(P).Monto
        End If
    Next
    NetoNoGravado = tot_ng
End Property
Public Property Get NetoGravado() As Double
    Dim tot_ng As Double
    tot_ng = 0
    Dim P As Long
    For P = 1 To vIvaAplicado.count
        tot_ng = tot_ng + vIvaAplicado(P).Monto
    Next
    NetoGravado = funciones.FormatearDecimales(tot_ng)



End Property

Public Property Get TotalAplicadoACuentas() As Double
    Dim vcta As clsCuentaFactura
    Dim c As Double
    For Each vcta In vCuentasContables
        c = c + vcta.Monto
    Next
    TotalAplicadoACuentas = c
End Property


Public Property Get TotalPendiente() As Double
    TotalPendiente = Me.total - Me.TotalAbonado + Me.TotalPago
End Property


Public Function AnalizarEstado() As EstadoFacturaProveedor
    If Me.total = Me.TotalAbonado Then AnalizarEstado = Saldada
    If Me.TotalAbonado > 0 And Me.TotalAbonado < Me.total Then AnalizarEstado = pagoParcial
End Function
