Attribute VB_Name = "enums"
Public forma_Entrega(2)
Public tipo_ot(2)
Public estados_Reques(11)
Dim Destinos(3)
Dim unidad(5)
Dim estados_material(2)
Dim estado_orden_entrega(3)
Dim estado_factura_proveedor(4)
Dim forma_de_pago_cta_cte(2)
Public estado_po(2)
Dim estado_presupuesto(8)
Dim estado_proceso_ot(3)
Dim estado_remito(3)
Dim estado_remito_facturado(4)
Dim meses(12) As String
Dim estado_recibo(3)
Dim estado_nnc(3)
Dim tipo_doc_contable(3)
Dim estado_doc_contable(5)
Dim tipos_Doc(2)
Dim tipo_complejidad(3)
Dim estado_saldado(5)
Dim estado_orden_pago(3)
Dim estado_liquidacion_caja(3)
Public tipoMateriales As New Dictionary
Public EstadosPeticionOfertaDetalle As New Dictionary
Public TiposAccidente As New Dictionary
Public TiposTratamiento As New Dictionary
Public TiposGravedad As New Dictionary
Public TiposCompensatorio As New Dictionary
Public EstadosCheque As New Dictionary

Dim estado_proveedor(3)

'Public Enum TipoOT
'    TipoOTHijaMarco = -1
'    TipoOTMarco = 1
'    TipoOTComun = 0
'End Enum

Public Enum TipoComplejidad
    ComplejidadAlta = 3
    ComplejidadMedia = 2
    ComplejidadBaja = 1
End Enum

Public Enum TipoPadron
    TipoPadronRetencion = 0
    TipoPadronPercepcion = 1
    'Agrega Nemer para poder cargar el Padron CABA
    TipoPadronUnificadoCABA = 2



End Enum

Public Enum TipoDocumento
    TipoDocumentoCuit = 0
    TipoDocumentoCuil = 1
End Enum

Public Enum TipoDocumentoDetalle
    TipoDocumentoDetalle_Fijo = 0
    TipoDocumentoDetalle_Dinamico = 1
End Enum

Public Enum TipoOt
    OT_TRADICIONAL = 1
    OT_STOCK = 2
    OT_ENTREGA = 3
End Enum


Public Enum EstadoProveedor
    EstadoProveedorCuentaCorriente = 1
    EstadoProveedorContado = 2
    EstadoProveedorEliminado = 0
End Enum

Public Enum ConceptoIncluir
    ConceptoProducto = 1
    ConceptoServicio = 2
    ConceptoProductoServicio = 3

End Enum

Public Enum EstadoOrdenPago
    EstadoOrdenPago_pendiente = 0
    EstadoOrdenPago_Aprobada = 1
    EstadoOrdenPago_Anulada = 2
    
End Enum

Public Enum EstadoLiquidacionCaja
    EstadoLiquidacionCaja_pendiente = 0
    EstadoLiquidacionCaja_Aprobada = 1
    EstadoLiquidacionCaja_Anulada = 2

End Enum


Public Enum FormaCaluloSuperficie
    FormaCaluloSuperficie_NoCalcula = 0
    FormaCaluloSuperficie_BaseXAltura = 1
    FormaCaluloSuperficie_Circulo = 2
    FormaCaluloSuperficie_Triangulo = 3
End Enum

Public Enum EstadoCheque
    ChequeAnulado = 0
    ChequeAceptado = 1


End Enum

Public Enum OperacionEntradaSalida
    OPEntrada = 1
    OPSalida = -1
End Enum

Public Enum TipoCompensatorio
    TC_Credito = 0
    TC_Debido = 1

End Enum


Public Enum EstadoNotaNoConformidad
    NNC_EnEdicion = 0
    NNC_Pendiente = 1
    NNC_Resuelta = 2
End Enum

Public Enum TipoEventoBroadcast
    TEB_RemitoAprobado = 1
    TEB_FacturaAprobada = 2
    TEB_OrdenTrabajoAprobada = 3
    TEB_OrdenEntregaAprobada = 4    'falta
    TEB_PresupuestoAprobado = 5
    TEB_OrdenTrabajoModificada = 6
    TEB_ArchivoOrdenTrabajo = 7
    TEB_ArchivoDetalleOrdenTrabajo = 8
    TEB_ArchivoPieza = 9
    TEB_OrdenTrabajoAnulada = 10
    TEB_OrdenTrabajoActivada = 11    'puesta en produccion
    TEB_PresupuestoEnviado = 12
    TEB_PresupuestoAnulado = 13
    TEB_PresupuestoCreado = 14
    TEB_RemitoAnulado = 15
    TEB_FacturaAnulada = 16
    TEB_FacturaCreada = 17
    TEB_RemitoCreado = 18
    TEB_IncidenciaOrdenTrabajo = 19
    TEB_IncidenciaDetalleOrdenTrabajo = 20
    TEB_IncidenciaPieza = 21
    TEB_RequerimientoCompraFinalizado = 22
    TEB_RequerimientoCompraAprobado = 23
    TEB_RequerimientoCompraAnulado = 24
    TEB_PeticionOfertaCreada = 25
    TEB_OrdenConAnticipoAprobada = 26
End Enum

Public Enum TipoSaldadoFactura
    NoSaldada = 0
    saldadoTotal = 1
    SaldadoParcial = 2
    notaCredito = 3
    notaCreditoParcial = 4
End Enum

Public Enum origenFacturado
    OrigenFacturadoConcepto = 1
    OrigenFacturadoRemito = 0
    OrigenFacturadoAnticipoOT = 2
End Enum

Public Enum EstadoDetalleFacturaCliente
    EstadoDetalleFacturaCliente_facturado = 1

End Enum


Public Enum TipoComprobanteUsado
    OrdenPago_ = 1
    Factura_ = 2
    FacturaProveedor_ = 3

    Recibo_ = 5
    Retencion_ = 6
    SaldoInicial_ = 7

    ReciboAnticipo_ = 8

End Enum


Public Enum tipoDocumentoContable
    Factura = 0
    notaCredito = 1
    notaDebito = 2
    DespachoAduana = 3
    LiquidacionBancaria = 4

    'e5re52- SE AGREGA ESTE TIPO DE COMPROBANTE NUEVO
    CompraBienesUsados = 5

End Enum



Public Enum EstadoOrdenEntrega
    Pendiente = 1
    Aprobado = 2
    FINALIZADO = 3
End Enum


Public Enum TipoCuentaBancaria
    CuentaCorriente = 1
    CajaAhorro = 2
End Enum

Public Enum TipoPeriodo
    TipoPeriodoFecha = 1
    TipoPeriodoMes = 2
    TipoPeriodoAño = 3
End Enum


Public Enum OrigenOperacion
    caja = 1
    Banco = 2
End Enum


Public Enum EstadoRecibo
    Pendiente = 1
    Aprobado = 2
    Reciboanulado = 3
End Enum

Public Enum OrigenRemito
    OrigenRemitoOt = 1
    OrigenRemitoConcepto = 3
    OrigenRemitooe = 2
    OrigenRemitoAplicado = 4
End Enum

Public Enum EstadoRemitoFacturado
    RemitoNoFacturado = 0
    RemitoFacturadoParcial = 1
    RemitoFacturadoTotal = 2
    RemitoNoFacturable = 3
End Enum

Public Enum TipoDeposito
    DepositoEfectivo = 1
    DepositoCheque = 2

End Enum

Public Enum EstadoRemito
    RemitoAnulado = 3
    RemitoAprobado = 2
    RemitoPendiente = 1
End Enum

Public Enum AdminNC_Motivo
    nc_Normal = 0
    nc_AjusteIIBB = 1
End Enum
Public Enum tipoEvento
    modificar_ = 1
    agregar_ = 2
    agregarColeccion_ = 3
End Enum
Public Enum EstadoProcesoDetalleOrdenTrabajo
    EstProcDetOT_AunNoDefinido = 0
    EstProcDetOT_ProcesoDefinido = 1
    EstProcDetOT_ProcesoNoDefinido = 2
End Enum
Public Enum EstadoOrdenTrabajo
    EstadoOT_Pendiente = 1
    EstadoOT_EnProceso = 2
    EstadoOT_ProcesoCompleto = 3
    EstadoOT_Finalizado = 4
    EstadoOT_EnEspera = 5
    EstadoOT_Desactivado = 6
End Enum
Public Enum EstadoPresupuesto
    Pendiente_ = 1
    Enviado_ = 2
    Procesado_ = 3
    ACotizar_ = 6
    NoCotizado = 7
    Desactivado = 8
End Enum
Public Enum FormaCotizar
    automatica_ = 0
    Cantidad_ = 1
    fabricados_ = 2
    fijo_ = 3
End Enum
Public Enum EstadoPO
    Pendiente_ = 0    'en edicion
    Finalizado_ = 1
    OrdenCompraCreada = 2
End Enum
Public Enum EstadoRequeCompra
    EnEdición_ = 0
    Finalizado_ = 1
    Aprobado_ = 2
    EnProceso_ = 3
    Procesado_ = 4
    EnPO_ = 5
    AprobadoParcial_ = 6
    EnProcesoParcial_ = 7
    ProcesadoParcial_ = 8
    EnPOParcial_ = 9
    Anulado = 10
    AnuladoParcial = 11
End Enum

Public Enum EstadoPeticionOfertaDetalle
    EPOD_Activo = 1
    EPOD_Anulado = 2
    EPOD_EnEspera = 3
    EPOD_comprado = 4
End Enum

Public Enum formaEntrega
    FormaEntrega_Retiramos = 1
    formaEntrega_Entregan = 2
End Enum

Public Enum Unidades
    kg_ = 1
    m2_ = 2
    Ml_ = 3
    un_ = 4
    litro_ = 5
End Enum

Public Enum EstadoUsuario
    activo = 1
    Inactivo = 2
End Enum

Public Enum tipoEntrega
    material_ = 1
    concepto_ = 2
End Enum

Public Enum TipoOperacionProveedor
    Alta = 1
    Modificacion = 2
    ver = 3
End Enum

Public Enum EstadoFacturaProveedor
    EnProceso = 1
    Aprobada = 2
    Saldada = 3
    pagoParcial = 4
End Enum

Public Enum FormadePagoFacturaProveedor
    PagoCuentaCorriente = 1
    PagoContado = 0
End Enum

Public Enum EstadoFacturaCliente
    EnProceso = 1
    Aprobada = 2
    Anulada = 3    'no se usa mas
    CanceladaNC = 4
    CanceladaNCParcial = 5
End Enum
Public Enum EstadoCliente
    activo = 1
    Inactivo = 2
End Enum
Public Enum TipoPersona
    cliente_ = 1
    proveedor_ = 2
End Enum
Public Enum EstadoMaterial
    activo = 1
    Inactivo = 2
End Enum
Public Enum destino
    ot_ = 1
    stock_ = 2
End Enum

Public Enum TipoMaterial
    TM_PerfilEspecial = 1
    TM_PerfilTubo = 2
    TM_PerfilCuadrado = 3
    TM_PerfilRectangular = 4
    TM_HojaPlancha = 5
    TM_UnidadKilo = 6
    TM_PerfilELE = 7
End Enum

Public Enum TipoAccidenteSiniestro
    TAS_DeTrabajo = 1
    TAS_ReAgravacion = 2
    TAS_InItinere = 3
End Enum

Public Enum TipoTratamientoSiniestro
    TTS_Ambulatorio = 1
End Enum

Public Enum TipoGravedadSiniestro
    TGS_Leve = 1
    TGS_Moderada = 2
End Enum

Public Function EnumTiposComplejidad(indice) As String
    EnumTiposComplejidad = tipo_complejidad(indice)
End Function

Public Function EnumTipoOT(indice) As String
    EnumTipoOT = tipo_ot(indice)
End Function
Public Function EnumTiposDoc(indice) As String
    enumtioposdoc = tipos_Doc(indice)
End Function
Public Function EnumTipoMaterial(indice As TipoMaterial) As String
    EnumTipoMaterial = tipoMateriales.item(CStr(indice))
End Function

Public Function EnumTipoDocumentoContable(indice) As String
    EnumTipoDocumentoContable = tipo_doc_contable(indice)
End Function
Public Function enumEstadoFacturaProveedor(indice) As String
    enumEstadoFacturaProveedor = estado_factura_proveedor(indice)
End Function

Public Function enumFormaDePagoFacturaProveedor(indice) As String
    enumFormaDePagoFacturaProveedor = forma_de_pago_cta_cte(indice)
End Function

Public Function enumEstadoOrdenEntrega(indice) As String
    enumEstadoOrdenEntrega = estado_orden_entrega(indice)
End Function


Public Function enumEstadoProcesoDetalleOrdenTrabajo(indice) As String
    enumEstadoProcesoDetalleOrdenTrabajo = estado_proceso_ot(indice)
End Function
Public Function enumEstadoPO(indice) As String
    enumEstadoPO = estado_po(indice)
End Function
Public Function enumEstadoMaterial(indice) As String
    enumEstadoMaterial = estados_material(indice)
End Function
Public Function enumDestino(indice) As String
    enumDestino = Destinos(indice)
End Function
Public Function enumEstadoRequeCompra(indice) As String
    enumEstadoRequeCompra = estados_Reques(indice)
End Function
Public Function enumUnidades(indice) As String
    enumUnidades = unidad(indice)
End Function
Public Function LlenarArrays()
    tipo_complejidad(ComplejidadBaja) = "Baja"
    tipo_complejidad(ComplejidadMedia) = "Media"
    tipo_complejidad(ComplejidadAlta) = "Alta"

    estado_doc_contable(EstadoFacturaCliente.Anulada) = "Anulada"
    estado_doc_contable(EstadoFacturaCliente.Aprobada) = "Aprobada"
    estado_doc_contable(EstadoFacturaCliente.CanceladaNC) = "Cancela NC"
    estado_doc_contable(EstadoFacturaCliente.CanceladaNCParcial) = "Cancela NC Parcial"
    estado_doc_contable(EstadoFacturaCliente.EnProceso) = "En Edición"

    tipo_doc_contable(tipoDocumentoContable.Factura) = "Factura"
    tipo_doc_contable(tipoDocumentoContable.notaCredito) = "N. Crédito"
    tipo_doc_contable(tipoDocumentoContable.notaDebito) = "N. Debito"


    tipos_Doc(TipoDocumento.TipoDocumentoCuit) = "CUIT"

    tipos_Doc(TipoDocumento.TipoDocumentoCuil) = "CUIL"


    estado_nnc(EstadoNotaNoConformidad.NNC_EnEdicion) = "En Edición"
    estado_nnc(EstadoNotaNoConformidad.NNC_Pendiente) = "Pendiente"
    estado_nnc(EstadoNotaNoConformidad.NNC_Resuelta) = "Resuelta"

    tipo_ot(0) = "Tradicional"
    tipo_ot(1) = "De Stock"
    tipo_ot(2) = "De Entrega"

    estado_proceso_ot(0) = "Aún no definido"
    estado_proceso_ot(1) = "Definido"
    estado_proceso_ot(2) = "No definido"

    estado_proveedor(EstadoProveedor.EstadoProveedorEliminado) = "Inactivo"
    estado_proveedor(EstadoProveedor.EstadoProveedorCuentaCorriente) = "Cuenta Corriente"
    estado_proveedor(EstadoProveedor.EstadoProveedorContado) = "Contado"

    forma_Entrega(formaEntrega.formaEntrega_Entregan) = "Entregan"
    forma_Entrega(formaEntrega.FormaEntrega_Retiramos) = "Retiramos"

    Destinos(1) = "OT"
    Destinos(2) = "Stock"

    estados_material(1) = "Activo"
    estados_material(2) = "Inactivo"

    estado_presupuesto(1) = "Pendiente"
    estado_presupuesto(2) = "Enviado"
    estado_presupuesto(3) = "Procesado"
    estado_presupuesto(6) = "A Cotizar"
    estado_presupuesto(7) = "No Cotizado"
    estado_presupuesto(8) = "Desactivado"

    forma_de_pago_cta_cte(0) = "Cta. Cte."
    forma_de_pago_cta_cte(1) = "Contado"


    estado_factura_proveedor(1) = "En Proceso"    'EstadoFacturaProveedor.EnProceso
    estado_factura_proveedor(2) = "Aprobada"    ' EstadoFacturaProveedor.Aprobada
    estado_factura_proveedor(3) = "Saldada"    'EstadoFacturaProveedor.Saldada
    estado_factura_proveedor(4) = "Pago Parcial"


    estado_orden_entrega(1) = "Pendiente"
    estado_orden_entrega(2) = "Aprobada"
    estado_orden_entrega(3) = "Finalizada"

    estado_saldado(TipoSaldadoFactura.NoSaldada) = "No Saldada"
    estado_saldado(TipoSaldadoFactura.notaCredito) = "N. Crédito"
    estado_saldado(TipoSaldadoFactura.notaCreditoParcial) = "N. Crédito Parcial"
    estado_saldado(TipoSaldadoFactura.SaldadoParcial) = "Parcial"
    estado_saldado(TipoSaldadoFactura.saldadoTotal) = "Total"

    estado_orden_pago(EstadoOrdenPago_pendiente) = "Pendiente"
    estado_orden_pago(EstadoOrdenPago.EstadoOrdenPago_Aprobada) = "Aprobada"
    estado_orden_pago(EstadoOrdenPago.EstadoOrdenPago_Anulada) = "Anulada"


    estado_po(0) = "Pendiente (en edicion)"
    estado_po(1) = "Finalizado"
    estado_po(2) = "Orden Compra creada"

    estado_recibo(EstadoRecibo.Aprobado) = "Aprobado"
    estado_recibo(EstadoRecibo.Pendiente) = "Pendiente"
    estado_recibo(EstadoRecibo.Reciboanulado) = "Anulado"

    unidad(1) = "Kg"
    unidad(2) = "M2"
    unidad(3) = "Ml"
    unidad(4) = "Un"
    unidad(5) = "Litro"

    estados_Reques(0) = "En Edición"
    estados_Reques(1) = "Finalizado"
    estados_Reques(2) = "Aprobado"
    estados_Reques(3) = "En Proceso"
    estados_Reques(4) = "Procesado"
    estados_Reques(5) = "En P.O."
    estados_Reques(6) = "Aprobado Parcial"
    estados_Reques(7) = "En Proceso Parcial"
    estados_Reques(8) = "Procesado Parcial"
    estados_Reques(9) = "En P.O. Parcial"
    estados_Reques(10) = "Anulado"
    estados_Reques(11) = "Anulado Parcial"


    estado_remito(EstadoRemito.RemitoAprobado) = "Aprobado"
    estado_remito(EstadoRemito.RemitoPendiente) = "Pendiente"
    estado_remito(EstadoRemito.RemitoAnulado) = "Anulado"
    estado_remito_facturado(EstadoRemitoFacturado.RemitoFacturadoParcial) = "Facturado Parcial"
    estado_remito_facturado(EstadoRemitoFacturado.RemitoFacturadoTotal) = "Facturado Total"
    estado_remito_facturado(EstadoRemitoFacturado.RemitoNoFacturable) = "No Facturable"
    estado_remito_facturado(EstadoRemitoFacturado.RemitoNoFacturado) = "No Facturado"


    meses(1) = "Enero"
    meses(2) = "Febrero"
    meses(3) = "Marzo"
    meses(4) = "Abril"
    meses(5) = "Mayo"
    meses(6) = "Junio"
    meses(7) = "Julio"
    meses(8) = "Agosto"
    meses(9) = "Septiembre"
    meses(10) = "Octubre"
    meses(11) = "Noviembre"
    meses(12) = "Diciembre"




    Set tipoMateriales = New Dictionary
    tipoMateriales.Add CStr(TipoMaterial.TM_PerfilCuadrado), "Perfil Cuadrado"
    tipoMateriales.Add CStr(TipoMaterial.TM_PerfilEspecial), "Perfil Especial"
    tipoMateriales.Add CStr(TipoMaterial.TM_HojaPlancha), "Hoja / Plancha"
    tipoMateriales.Add CStr(TipoMaterial.TM_PerfilRectangular), "Perfil Rectangular"
    tipoMateriales.Add CStr(TipoMaterial.TM_PerfilTubo), "Perfil Tubo"
    tipoMateriales.Add CStr(TipoMaterial.TM_UnidadKilo), "Unidad / Kilo / Litro"
    tipoMateriales.Add CStr(TipoMaterial.TM_PerfilELE), "Perfil L"

    EstadosPeticionOfertaDetalle.Add CStr(EstadoPeticionOfertaDetalle.EPOD_Activo), "Activo"
    EstadosPeticionOfertaDetalle.Add CStr(EstadoPeticionOfertaDetalle.EPOD_Anulado), "Anulado"
    EstadosPeticionOfertaDetalle.Add CStr(EstadoPeticionOfertaDetalle.EPOD_EnEspera), "En espera"
    EstadosPeticionOfertaDetalle.Add CStr(EstadoPeticionOfertaDetalle.EPOD_comprado), "Comprado"


    TiposAccidente.Add CStr(TipoAccidenteSiniestro.TAS_DeTrabajo), "De Trabajo"
    TiposAccidente.Add CStr(TipoAccidenteSiniestro.TAS_InItinere), "In Itinere"
    TiposAccidente.Add CStr(TipoAccidenteSiniestro.TAS_ReAgravacion), "Reagravación"

    TiposTratamiento.Add CStr(TipoTratamientoSiniestro.TTS_Ambulatorio), "Ambulatorio"

    TiposGravedad.Add CStr(TipoGravedadSiniestro.TGS_Leve), "Leve"
    TiposGravedad.Add CStr(TipoGravedadSiniestro.TGS_Moderada), "Moderada"

    TiposCompensatorio.Add CStr(TipoCompensatorio.TC_Credito), "Credito"
    TiposCompensatorio.Add CStr(TipoCompensatorio.TC_Debido), "Debito"


    EstadosCheque.Add CStr(EstadoCheque.ChequeAceptado), "Aceptado"
    EstadosCheque.Add CStr(EstadoCheque.ChequeAnulado), "Anulado"


End Function
Public Function EnumEstadoRecibo(indice) As String
    EnumEstadoRecibo = estado_recibo(indice)
End Function
Public Function EnumEstadoDocumentoContable(indice) As String
    EnumEstadoDocumentoContable = estado_doc_contable(indice)
End Function
Public Function EnumTipoDocumentoContableShort(indice) As String
' SE MUESTRAN EL DETALLE ABREVIADO PARA CADA TIPO DE COMPROBANTE
    Select Case indice
    Case 0: EnumTipoDocumentoContableShort = "FC"
    Case 1: EnumTipoDocumentoContableShort = "NC"
    Case 2: EnumTipoDocumentoContableShort = "ND"
    Case 3: EnumTipoDocumentoContableShort = "DA"
    Case 4: EnumTipoDocumentoContableShort = "LB"
    Case 5: EnumTipoDocumentoContableShort = "CBU"

    End Select
End Function
Public Function EnumPeriodo(indice) As String
    EnumPeriodo = meses(indice)
End Function
Public Function EnumEstadoRemito(indice) As String
    EnumEstadoRemito = estado_remito(indice)
End Function
Public Function EnumEstadoNNC(indice) As String
    EnumEstadoNNC = estado_nnc(indice)
End Function
Public Function EnumEstadoRemitoFacturado(indice) As String
    EnumEstadoRemitoFacturado = estado_remito_facturado(indice)
End Function
Public Function EnumTipoSaldadoFactura(indice)
    EnumTipoSaldadoFactura = estado_saldado(indice)
End Function
Public Function EnumEstadoPresupuesto(indice, Optional ByRef matriz)
    EnumEstadoPresupuesto = estado_presupuesto(indice)
    matriz = estado_presupuesto
End Function
Public Function EnumEstadoOrdenPago(indice) As String
    EnumEstadoOrdenPago = estado_orden_pago(indice)
End Function
Public Function EnumEstadoLiquidacionCaja(indice) As String
    EnumEstadoLiquidacionCaja = estado_liquidacion_caja(indice)
End Function


Public Function EnumEstadoProveedor(indice) As String
    EnumEstadoProveedor = estado_proveedor(indice)
End Function
