VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "CODEJO~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Signo Plast ERP"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   Icon            =   "principal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrInformeAccidentes 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6030
      Top             =   2430
   End
   Begin VB.Timer tmrEventos 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3945
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2190
      Top             =   1935
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2190
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2190
      Top             =   3495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   96
      ImageHeight     =   96
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "principal.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl Popup 
      Left            =   990
      Top             =   2505
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   1020
      Top             =   1830
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   7
   End
   Begin XtremeSuiteControls.TrayIcon TrayIcon 
      Left            =   4020
      Top             =   5415
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   16
      Text            =   "balblablbalba"
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   1320
      Top             =   3465
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "principal.frx":2AD8
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public servidorBBDD As New Collection
Public servidorActual As String
Dim classP As New classSignoplast
Private contMinutosInformesAccidente As Long
Dim statusBar As XtremeCommandBars.statusBar

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__NUEVO_EMPLEADO:
        Dim f32423 As New frmAltaEmpleados
        f32423.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__EMPLEADOS:
'        Dim f4333 As New frmListaEmpleados
        frmListaEmpleados.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__OS:

        ' Dim f43334 As New frmObraSocial
        frmObraSocial.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__SECTORIZAR:
        frmSectorizar.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__USUARIOS:
        frmUsuarios.Show

        'Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__SINIESTROS:
        '     Dim F939393 As New frmSiniestros
        '      F939393.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__LUGARES
'        Dim faa1 As New frmUbicaciones
        frmUbicaciones.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__VER_ACTUALIZACIONES
        frmTip.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__SISTEMA:
        frmConfigurarSistema.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__COTIZACIONES__GASTOS:
        frmConfigurarGastos.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__COTIZACIONES__DATOS_VARIOS:
        frmDatosVarios.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__TERMINACION__CONFIG_PINTURA:
        frmConfigurarPintura.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__TERMINACION__DEFINIR_CUENTAS:
        frmDefinirTerminacion.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__IVA:
        frmAdminIVA.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__FACTURAS:

        frmAdminFacturasEmisibles.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__RETENCIONES:
        frmAdminConfigRetenciones.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__PERCEPCIONES:
        frmAdminconfigPercepciones.Show

    Case ID_BUTTON.ID_BUTTON_CAJAYBANCOS__CONFIGURAR__ADMINISTRACION__BANCOS:

        frmAdminConfigBancos.Show

    Case ID_BUTTON.ID_BUTTON_CAJAYBANCOS__CONFIGURAR__ADMINISTRACION__CUENTAS:
        frmAdminconfigCuentas.Show


    Case ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__NUEVA:
        Dim frm2001 As frmVentasPresupuestoNuevo
        Set frm2001 = New frmVentasPresupuestoNuevo
        frm2001.Show

    Case ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__LISTADO:
        Dim frm2002 As frmVentasPresupuestoLista
        Set frm2002 = New frmVentasPresupuestoLista
        frm2002.Show
    Case ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__INFORME_HISTORICO:
        frmVentasHistoricosVentas.Show


    Case ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__INFORME_DINAMICO:
        frmVentasEstadisticasCotizaciones.Show

    Case ID_BUTTON.ID_BUTTON_VENTAS__PEDIDOS__NUEVO:
        Dim frm2006 As frmVentasPedidoNuevo
        Set frm2006 = New frmVentasPedidoNuevo
        frm2006.Show



    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__NUEVA:

        Dim frmNuevaOT As New frmNuevaOrdenTrabajo
        frmNuevaOT.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__LISTADO:
        Dim frmpp As New frmPlaneamientoPedidosPendientes
        frmpp.Show
    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__HISTORICO

        frmPlaneamientoResumenProduccion.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__MARCO_NUEVA:
        Dim fofo As New frmNuevoContratoMarco
        fofo.Show

    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_ENTREGA__NUEVA:

        frmPlaneamientoOENueva.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__PLANIFICAR:
        Dim frm_merla As New frmPlanificacionTemporal
        frm_merla.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_ENTREGA__LISTADO:

        frmPlaneamientoOELista.Show

    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__GLOBAL:

        frmPlaneamientoSeguimiento.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__DE_RUTAS:


        Dim fff As New frmPlaneamientoSeguimientoRutas3
        fff.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_PROCESOS:
        Dim frmT As New frmTiempoProcesoDetalle
        frmT.Show 1

    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_RUTAS:
        Dim frmSep As New frmSeguimientoEspecialPorRuta
        frmSep.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__VER_TIEMPOS
        frmPlaneamientoVerTiempos.Show
    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__OPER_PROCESO
        Dim f123433 As New frmPlaneamientoVerOperariosEnProceso
        f123433.Show
    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__VER_NNC
        Dim fzzz As New frmNotasNoConformidad
        fzzz.Show

    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__REMITOS__NUEVO:

        frmPlaneamientoRemitosNuevo.Show

    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__REMITOS__LISTADO:

        Dim frm3033 As New frmPlaneamientoRemitosLista
        frm3033.Show

    Case ID_BUTTON.ID_BUTTON_COMPRAS__REQUERIMIENTOS__NUEVO:
        Dim frm4001 As frmComprasRequesNuevo
        Set frm4001 = New frmComprasRequesNuevo
        frm4001.Show
    Case ID_BUTTON.ID_BUTTON_COMPRAS__REQUERIMIENTOS__LISTADO: frmComprasRequeLista.Show
    Case ID_BUTTON.ID_BUTTON_COMPRAS__PETICION_OFERTA__NUEVA: Dim f12312 As New frmComprasArmaPO: f12312.Show
    Case ID_BUTTON.ID_BUTTON_COMPRAS__PETICION_OFERTA__LISTADO: Dim f34242 As New frmComprasPeticionesLista: f34242.Show

    Case ID_BUTTON.ID_BUTTON_COMPRAS__PETICION_OFERTA__COMPRAR:
        Dim faaf As New frmComprasPOComprar
        faaf.Show

        '        Case ID_BUTTON.ID_BUTTON_COMPRAS__ORDEN_COMPRA__NUEVA: frmComprasOrdenesNueva.Show
    Case ID_BUTTON.ID_BUTTON_COMPRAS__ORDEN_COMPRA__LISTADO: frmComprasOrdenesLista.Show
    Case ID_BUTTON.ID_BUTTON_COMPRAS__PRECIOS__ADMINISTRAR: frmComprasPreciosPorRubro.Show
    Case ID_BUTTON.ID_BUTTON_COMPRAS__PRECIOS__HISTORICO: frmComprasPreciosHistorico.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__NUEVO_ELEMENTO:

        Dim frm5001 As frmNuevoElemento
        Set frm5001 = New frmNuevoElemento
        frm5001.btnModificar.Visible = False
        frm5001.lblidStock = Empty
        frm5001.Show


        'Dim frm9988 As New frmPieza
        ' frm9988.Show


    Case ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__NUEVO_CONJUNTO:

        frmDefinirConjunto.accion = 0
        frmDefinirConjunto.Show

    Case ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__LISTADO:

        frmListarStock.Show



    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__NUEVO_MATERIAL: frmMaterialesNuevo.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__RUBROS:

        frmComprasProveedoresRubros.Show


    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__GRUPOS: frmRubrosGrupos.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__ALMACENES: frmMaterialesAlmacenes.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__LISTADO:
        Dim frm1112 As New frmMaterialesLista2
        frm1112.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__HISTORIAL: frmComprasPreciosHistorico.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__NUEVA_TAREA:

        frmNuevaMDO.Show

    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__SECTORES:

        frmSectores.Show


    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__TAREAS:

        Dim ffff As New frmListaTareas
        ffff.Show




    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__CATEGORIA_SUELDOS:
        Dim F As New frmCategoriasSueldo
        F.Show
    Case ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__SUELDO:
        Dim cs As New CategoriaSueldo
        cs.EspecificarSueldo

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__REMITOS:
        Dim frm6001 As frmPlaneamientoRemitosLista
        Set frm6001 = New frmPlaneamientoRemitosLista

        frm6001.VerInfoAdministracion = True
        frm6001.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_FC:

        Dim f324 As New frmAdminFacturasEdicion
        f324.NuevoTipoDocumento = tipoDocumentoContable.Factura
        f324.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_NC:

        Dim f3241 As New frmAdminFacturasEdicion
        f3241.NuevoTipoDocumento = tipoDocumentoContable.notaCredito
        f3241.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_ND:

        Dim f32412 As New frmAdminFacturasEdicion
        f32412.NuevoTipoDocumento = tipoDocumentoContable.notaDebito
        f32412.Show



    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_ANTICIPO:

        Dim f324121 As New frmAdminFacturasEdicion
        f324121.NuevoTipoDocumento = tipoDocumentoContable.Factura
        f324121.EsAnticipo = True
        f324121.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__FACTURAS:

        frmAdminFacturasEmitidas.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COBRANZAS__RECIBOS:

        Dim fffff As New frmAdminCobranzasLista
        fffff.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COBRANZAS__NUEVO_RECIBO:

        frmAdminCobranzasReservarRecibo.Show



    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COBRANZAS__DEUDORES:

        Dim frm1122 As New frmAdminFacturasAdeudadas2
        frm1122.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__MOVIMIENTO__FONDOS:

        Dim frmMover As New frmMovimientoDeFondos
        frmMover.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__CREAR_LIQUIDACION_CAJA:
                    Dim f12323 As New frmAdminPagosLiquidaciondeCajaCrear
                    f12323.Show
                    
   Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__CREAR_LIQUIDACION_CAJA_DG:
                    Dim f12326 As New frmAdminPagosLiqCajaListaDG
                    f12326.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__LISTA_LIQUIDACION_CAJA:
                    Dim f12324 As New frmAdminPagosLiquidaciondeCajaLista
                    f12324.Show




    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__CREAR_ORDEN_PAGO:

        Dim f12322 As New frmAdminPagosCrearOrdenPago
        f12322.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__ORDEN_PAGO_LISTA:
        Dim fffffffffff As New frmAdminPagosOrdenesPagoLista
        fffffffffff.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__TRANSFERENCIAS:
        Dim f12325 As New frmAdminPagosTransferenciasBancarias
        f12325.Show





    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__COMPENSATORIOS
        Dim hdp As New frmCompensatorios
        hdp.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__RESUMEN_PAGOS

        Dim hdp1 As New frmResumenPagos
        hdp1.Show



    Case ID_BUTTON.ID_BUTTON_CAJAYBANCOS__CONFIGURAR__ADMINISTRACION__BANCOS:

        frmAdminConfigBancos.Show

    Case ID_BUTTON.ID_BUTTON_CAJAYBANCOS__CONFIGURAR__ADMINISTRACION__CUENTAS:
        frmAdminconfigCuentas.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_CLIENTE:

        Dim frm1111 As New frmAdminMasInfoCliente2
        frm1111.Show




    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_ORDEN_TRABAJO:

        frmAdminResumenOT.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_TOTAL:
        frmAdminResumenEstadoTotal.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_PERIODO:
        frmAdminResumenesFacturacion.Show

        '    '"Crear Recibo de Anticipo"
        '        Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__CREAR_RECIBO:
        '            frmAdminCobranzasReservarReciboAnticipo.Show
        '
        '   '"Ver Recibos de Anticipo"
        '       Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__VER_RECIBOS:
        '            frmAdminCobranzasListaAnticipo.Show

        '"Ver Detalle de Cta. Cte."
    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__MOVIMIENTOS:
        Dim frmcta As New frmCtaCte
        frmcta.Show

        '"Resúmen Saldos"
    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__RESUMEN__SALDOS:
        Dim frm1144 As New frmResumenSaldosProv
        frm1144.TipoPersonaCta = TipoPersona.proveedor_
        frm1144.caption = "Resúmen de saldos de Proveedores"
        frm1144.Show

    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__SALDOS:
        '            frmAdminCCResumenSaldos.Show
        Dim frm11441 As New frmResumenSaldosProv
        frm11441.TipoPersonaCta = TipoPersona.cliente_
        frm1144.caption = "Resúmen de saldos de Clientes"
        frm1144.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_VENTAS:
        Dim f3242 As New frmAdminSubdiariosVentasv2
        f3242.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_COBRANZAS:

        Dim f4444 As New frmAdminSubdiarioCobranzas
        f4444.Show



    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_RETENCIONES:

        Dim f4445 As New frmAdminSubdiarioRetenciones
        f4445.Show




    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_IVACOMPRAS

        Dim f4446 As New frmAdminSubdiarioCompras
        f4446.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__NUEVA:
        Dim frm1 As frmAdminComprasNuevaFCProveedor
        Set frm1 = New frmAdminComprasNuevaFCProveedor

        frm1.Factura = Nothing
        frm1.Show
    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__LISTADO: frmAdminComprasListaFCProveedor.Show
    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__PLAN_DE_CUENTAS_VER: frmAdminComprasPlanCuentasAdmin.Show
    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__PLAN_DE_CUENTAS_DEFINIR: frmAdminComprasCuentasDefinir.Show



    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__PADRON_IIBB:

        frmAdminIIBB.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__CENTRO_CAMBIO:

        frmAdminConfigCambio.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__EXTRAS__REPORTE_CMC:

        frmAdminExtrasReporteCMC.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CHEQUES:
        Dim cccfff As New frmAdminCheques
        cccfff.Show


    Case ID_BUTTON.ID_BUTTON_ADMINISTRACION__CHEQUES_DEPOSITAR:
        Dim dep As New frmDepositarCheque
        dep.Show

    Case ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__CLIENTES__NUEVO:
        Dim ff111 As New frmVentasClienteNuevo
        ff111.cliente = Nothing
        ff111.Show

    Case ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__CLIENTES__LISTADO: frmVentasClientesLista.Show


    Case ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__PROVEEDORES__NUEVO:

        Dim frmaa As New frmComprasProveedoresModifica
        frmaa.Show


    Case ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__PROVEEDORES__LISTADO:


        frmComprasProveedoresLista.Show


    Case ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__AGENDA__VER: frmSistemaAgendaGlobal.Show
    Case ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__TABLERO: frmSistemaTablero.Show
    Case ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__CAMBIAR_CONTRASEÑA:
        frmCambiarPassword.Frame1 = "[ " & funciones.GetUserObj.usuario & " ]"
        frmCambiarPassword.Show

        ' Desactivaciones dienemer 11.09.20
        '        Se desactiva AGENDA porque da error. Aparentemente no est? desarrollada.
        '        Case ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__AGENDA:
        '               frmUsuariosAgendaPersonal.Show

        '        Se desactiva EVENTOS porque da error. Para ver luego.
        '        Case ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__EVENTOS:
        '            Dim f2k2 As New frmEventos
        '            f2k2.Show
        '            f2k2.llenar

        '       Se desactiva ASIGNACION EVENTOS porque da error. No recuerdo si esto funcionaba
        '        Case ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS___ASIGNACION_EVENTOS
        '            Dim f212l As New frmUsuariosEventos
        '            f212l.Show


    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_INICIO_TAREAS_ASIGNADAS
        Dim F2221KMA As New frmTareasAsignadasInicio
        F2221KMA.Show

    Case ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_FIN_TAREAS_ASIGNADAS
        Dim f454 As New frmTareasIniciadasFin
        f454.Show

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__DOCUMENTOS
        MsgBox ("No desarrollado.")

    Case ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__TESTS
'        MsgBox ("Iniciando pruebas.")
        Dim f234 As New frmSistemasTests
        f234.Show
        
        
    Case ID_BUTTON_CLIENTES_PROVEEDORES__PROVEEDORES__ASOCRUBROS
        Dim f233 As New frmRubroProveedor
        f233.Show

    Case ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BUSQUEDA_PROGRAMAS
        Dim f22222 As New frmConsultaProgramasRadan
        f22222.Show

    Case ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_POSICION_IVA_MENSUAL
        DAOSubdiarios.PosicionIvaMensual

    Case ID_BUTTON_ADMINISTRACION__COMPRAS__CTA_CTE
        Dim famigo As New frmCtaCteProv
        famigo.Show

    End Select
End Sub



Private Sub MDIForm_Load()
    contMinutosInformesAccidente = 0

    Static pic As New StdPicture
    Set pic = Me.TrayIcon.Icon
    Set Me.TrayIcon.Icon = Nothing

    RegOCX.RegistrarOCXs


    Dim IdU As Long

    Dim tmpsrv As String
    'leo el .ini y verifico que est? configurado el servidor

    Dim jjj As Long
    servidorBBDD.Add LeerIni(App.path & "\config.ini", "Configurar", "ServidorBBDD", vbNullString)
    For jjj = 1 To 10
        tmpsrv = LeerIni(App.path & "\config.ini", "Configurar", "ServidorBBDD" & jjj, vbNullString)
        If LenB(tmpsrv) > 0 Then servidorBBDD.Add tmpsrv
    Next jjj


    If servidorBBDD.count = 0 Then
        MsgBox "Se produjo un error con el archivo config.ini! Verificar la existencia de servidor."
        End
    Else: frmLogin.Show 1
        'conectar.SetServidorBBDD  servidorBBDD 'ahora lohace el login
        If conectar.conectar Then

            IdU = funciones.getUser
            ChangeRegionalSettings

            Permisos.crearPermisos IdU
            llenar_vectores    'llena los vectores del modulo standar para estados
            enums.LlenarArrays    'llena los vectores del modulo enum
            Configurar.LoadConfiguration

            PrepararPopUp
            Me.Timer2.Enabled = True
            Me.tmrEventos.Enabled = Not funciones.InIDE

            frmPrincipal.Show

            Me.TrayIcon.text = "Signo Plast Event Handler"

            Set Me.TrayIcon.Icon = pic

            If LenB(LeerIni(App.path & "\config.ini", "Configurar", "puesto", vbNullString)) > 0 And InStr(1, funciones.GetUserObj.usuario, "puesto") Then
                Dim frmT As New frmTiempoProcesoDetalle
                frmT.Show 1
            End If

        Else
            MsgBox "No se puede establecer la conexion al servidor " & Me.servidorActual, vbCritical, "Error"
            End
        End If
    End If

    CreateRibbonBar

'         frmAdminPagosLiquidaciondeCajaCrear.Show
'         frmAdminPagosLiquidaciondeCajaLista.Show
            
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Está seguro de salir del sistema?", vbYesNo + vbQuestion) = vbNo Then Cancel = 1
End Sub

Public Sub mostrarTablero()
    Dim sp As New classSignoplast
    Dim marcado As Boolean

    Dim A As Long
    A = sp.cantidadGruposUsuario(funciones.getUser)
    If A > 0 Then
        If marcado Then
            frmSistemaTablero.Show
        Else
            Unload frmSistemaTablero
        End If
    Else
        MsgBox "No tiene grupos definidos, comuniquese con un supervisor!", vbInformation, "Información"
    End If
End Sub

'
'Private Sub stbar1_PanelClick(ByVal panel As MSComctlLib.panel)
'    If panel.Index = 5 Then
'        If MsgBox("Hay una nueva actualización, ?desea aplicarla ahora?", vbYesNo, "Confirmación") = vbYes Then
'
'        End If
'
'    ElseIf panel.Index = 2 Then
'        If MsgBox("?Desea cambiar el password ahora?", vbYesNo, "Confirmación") = vbYes Then
'            frmCambiarPassword.Show
'        End If
'
'    End If
'End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    SalirForzado
    Call UnInitializeAllContainer

End Sub

Private Sub Popup_ItemClick(ByVal item As Xtremesuitecontrols.IPopupControlItem)
    If item.Id = 666 Then
        Me.Popup.Close
    Else
        If MsgBox("Se va a aplicar una actualización." & vbNewLine & "Desea aplicarla ahora?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
            classP.actualizarSistema CLng(item.Id)
        End If
    End If
End Sub

Private Sub PrepararPopUp()
    Me.Popup.RemoveAllItems
    Me.Popup.Icons.RemoveAll

    Me.Popup.AnimateDelay = 255
    Me.Popup.Animation = xtpPopupAnimationFade
    Me.Popup.VisualTheme = xtpPopupThemeOffice2007
    Me.Popup.AllowMove = False
    Me.Popup.ShowDelay = 0    'para que quede fijo
    Me.Popup.Transparency = 100

    Dim item As PopupControlItem


    Set item = Popup.AddItem(135, 10, 170, 45, "Cerrar")
    item.Id = 666

    Set item = Popup.AddItem(10, 35, 160, 80, vbNullString)
    item.caption = "Actualización para el sistema."
    item.TextAlignment = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    item.Hyperlink = False

    Set item = Popup.AddItem(10, 75, 160, 95, vbNullString)
    item.Button = True
    item.caption = "Haga click aquí para actualizar."
    item.TextAlignment = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    item.Hyperlink = False
End Sub



Private Sub Timer2_Timer()

    Dim idNuevo As Long
    If Not funciones.InIDE Then
        If classP.VerificarSiHayActualizacion(idNuevo) Then
            statusBar(5).text = "** ACTUALIZACION DISPONIBLE **"
            If Permisos.SistemaVerUpdate Then
                Me.Popup.item(2).Id = idNuevo
                If Me.Popup.State = xtpPopupStateClosed Then
                    Me.Popup.Show
                End If
            End If
        Else
            statusBar(5).text = vbNullString
        End If
    End If
End Sub

Private Function AddButton(ribbonGroup As ribbonGroup, caption As String, Id As Long, Optional enabledCondition As Boolean = True, Optional iconId As Long = -1, Optional controlType As XTPControlType = xtpControlButton, Optional cmdBarCtrlParent As CommandBarControl = Nothing) As CommandBarControl
    Dim cmdControl As CommandBarControl

    If IsSomething(cmdBarCtrlParent) Then
        Set cmdControl = cmdBarCtrlParent.CommandBar.Controls.Add(controlType, Id, caption)
    Else
        Set cmdControl = ribbonGroup.Add(controlType, Id, caption)
    End If

    If iconId = -1 Then
        cmdControl.iconId = cmdControl.Id
    Else
        cmdControl.iconId = iconId
    End If
    cmdControl.Enabled = enabledCondition
    Set AddButton = cmdControl
End Function


Private Sub CreateRibbonBar()

    Set statusBar = CommandBars.statusBar
    Dim statusbarpane As statusbarpane

    statusBar.Visible = True
    statusBar.RibbonDividerIndex = 3    'para que me pinte todos azules

    Set statusbarpane = statusBar.AddPane(1)
    statusbarpane.text = " Usuario: " & GetUserObj.usuario
    statusbarpane.Style = SBPS_NOBORDERS
    statusbarpane.BeginGroup = True

    Set statusbarpane = statusBar.AddPane(2)    'vacio para que no se cague
    statusbarpane.text = vbNullString
    statusbarpane.Width = 0

    Set statusbarpane = statusBar.AddPane(3)
    statusbarpane.text = "Servidor: " & Me.servidorActual & " "
    'statusbarpane.BeginGroup = True

    Set statusbarpane = statusBar.AddPane(4)
    statusbarpane.text = "Versión: " & funciones.Version & " "
    'statusbarpane.BeginGroup = True

    Set statusbarpane = statusBar.AddPane(5)
    statusbarpane.text = vbNullString
    'statusbarpane.BeginGroup = True


    Set statusbarpane = statusBar.AddPane(6)    'vacio para que no se cague
    statusbarpane.text = vbNullString
    statusbarpane.Width = 0

'    Dim cmdControl As CommandBarControl
    Dim ribbonTab As ribbonTab
    Dim ribbonGroup As ribbonGroup

    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.FontHeight = 12
    RibbonBar.EnableDocking xtpFlagStretchedShared
    'RibbonBar.ShowQuickAccess = False

    Dim cmdBarCtrl As CommandBarControl
    CommandBars.Options.UseSharedImageList = False
    Set CommandBars.Icons = Me.ImageManager.Icons

    'PANEL DE CONTROL--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(0, "Panel de Control")
    ribbonTab.Id = ID_TAB.ID_TAB_PANEL_DE_CONTROL
    'Set ribbonGroup = ribbonTab.Groups.AddGroup("Usuarios y Empleados", ID_GROUP.ID_GROUP_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS)
    'AddButton ribbonGroup, "Nuevo empleado", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__NUEVO_EMPLEADO, Permisos.sistemaPanelControlGeneral

    'AddButton ribbonGroup, "Empleados", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__EMPLEADOS, Permisos.sistemaPanelControlGeneral
    'AddButton ribbonGroup, "Siniestros", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__SINIESTROS, (Permisos.RRHHInformeAccidente Or Permisos.RRHHSiniestros)
    'AddButton ribbonGroup, "Obras Sociales", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__OS, Permisos.sistemaPanelControlGeneral
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Configurar", ID_GROUP.ID_GROUP_PANEL_DE_CONTROL__CONFIGURAR)
    AddButton ribbonGroup, "Usuarios", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__USUARIOS, Permisos.sistemaPanelControlGeneral
    AddButton ribbonGroup, "Sistema", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__SISTEMA, Permisos.sistemaPanelControlGeneral
    Set cmdBarCtrl = AddButton(ribbonGroup, "Cotizaciones", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__COTIZACIONES, Permisos.sistemaPanelControlGeneral, , xtpControlPopup)
    AddButton ribbonGroup, "Gastos", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__COTIZACIONES__GASTOS, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl
    AddButton ribbonGroup, "Datos varios", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__COTIZACIONES__DATOS_VARIOS, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl

    Set cmdBarCtrl = AddButton(ribbonGroup, "Terminación", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__TERMINACION, Permisos.sistemaPanelControlGeneral, , xtpControlPopup)
    AddButton ribbonGroup, "Configurar pintura", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__TERMINACION__CONFIG_PINTURA, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl
    AddButton ribbonGroup, "Definir cuentas", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__TERMINACION__DEFINIR_CUENTAS, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl

    Set cmdBarCtrl = AddButton(ribbonGroup, "Administración", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION, Permisos.sistemaPanelControlGeneral, , xtpControlPopup)
    AddButton ribbonGroup, "IVA", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__IVA, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl
    AddButton ribbonGroup, "Facturas", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__FACTURAS, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl
    AddButton ribbonGroup, "Retenciones", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__RETENCIONES, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl
    AddButton ribbonGroup, "Percepciones", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__ADMINISTRACION__PERCEPCIONES, Permisos.sistemaPanelControlGeneral, , , cmdBarCtrl

    AddButton ribbonGroup, "Documentos", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__DOCUMENTOS, Permisos.sistemaPanelControlGeneral
    AddButton ribbonGroup, "Ubicaciones", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__CONFIGURAR__LUGARES, Permisos.sistemaPanelControlGeneral
    AddButton ribbonGroup, "Actualizaciones", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__VER_ACTUALIZACIONES, Permisos.sistemaPanelControlGeneral
    AddButton ribbonGroup, "Tests", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__TESTS, Permisos.sistemaPanelControlGeneral
  
    'RECURSOS HUMANOS--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(1, "Recursos Humanos")
    ribbonTab.Id = ID_TAB.ID_TAB_PANEL_DE_CONTROL
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Usuarios y Empleados", ID_GROUP.ID_GROUP_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS)
    AddButton ribbonGroup, "Nuevo empleado", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__NUEVO_EMPLEADO, Permisos.sistemaPanelControlGeneral
    AddButton ribbonGroup, "Empleados", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__EMPLEADOS, Permisos.sistemaPanelControlGeneral
    'AddButton ribbonGroup, "Siniestros", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__SINIESTROS, (Permisos.RRHHInformeAccidente Or Permisos.RRHHSiniestros)
    AddButton ribbonGroup, "Obras Sociales", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__OS, Permisos.sistemaPanelControlGeneral, 100


    'VENTAS--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(2, "Ventas")
    ribbonTab.Id = ID_TAB.ID_TAB_VENTAS

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Cotizaciones", ID_GROUP.ID_GROUP_VENTAS__COTIZACIONES)
    AddButton ribbonGroup, "Nueva", ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__NUEVA, Permisos.VentasCotizControl, ID_BUTTON.ID_BUTTON_ADD
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__LISTADO, Permisos.VentasCotizConsultas, ID_BUTTON.ID_BUTTON_SEARCH

    Set cmdBarCtrl = AddButton(ribbonGroup, "Informes", ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__INFORME, Permisos.VentasCotizConsultas, , xtpControlPopup)
    AddButton ribbonGroup, "Informe histórico", ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__INFORME_HISTORICO, Permisos.VentasCotizConsultas, , , cmdBarCtrl
    AddButton ribbonGroup, "Informe dinámico", ID_BUTTON.ID_BUTTON_VENTAS__COTIZACIONES__INFORME_DINAMICO, Permisos.VentasCotizConsultas, , , cmdBarCtrl


    Set ribbonGroup = ribbonTab.Groups.AddGroup("Pedidos", ID_GROUP.ID_GROUP_VENTAS__PEDIDOS)

    AddButton ribbonGroup, "Nuevo", ID_BUTTON.ID_BUTTON_VENTAS__PEDIDOS__NUEVO, Permisos.VentasPedidosControl
    'AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_VENTAS__PEDIDOS__LISTADO

    'COMPRAS--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(3, "Compras")
    ribbonTab.Id = ID_TAB.ID_TAB_COMPRAS
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Requerimientos", ID_GROUP.ID_GROUP_COMPRAS__REQUERIMIENTOS)
    AddButton ribbonGroup, "Nuevo", ID_BUTTON.ID_BUTTON_COMPRAS__REQUERIMIENTOS__NUEVO, Permisos.ComprasRequesControl
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_COMPRAS__REQUERIMIENTOS__LISTADO, Permisos.ComprasRequesConsultas, ID_BUTTON.ID_BUTTON_LISTADO
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Pet. Oferta", ID_GROUP.ID_GROUP_COMPRAS__PETICION_OFERTA)
    AddButton ribbonGroup, "Nueva", ID_BUTTON.ID_BUTTON_COMPRAS__PETICION_OFERTA__NUEVA, Permisos.ComprasPOCrear
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_COMPRAS__PETICION_OFERTA__LISTADO, Permisos.ComprasPOConsultar, ID_BUTTON.ID_BUTTON_LISTADO
    AddButton ribbonGroup, "Comprar", ID_BUTTON.ID_BUTTON_COMPRAS__PETICION_OFERTA__COMPRAR, Permisos.ComprasPOConsultar, ID_BUTTON.ID_BUTTON_LISTADO

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Orden Compra", ID_GROUP.ID_GROUP_COMPRAS__ORDEN_COMPRA)
    AddButton ribbonGroup, "Nueva", ID_BUTTON.ID_BUTTON_COMPRAS__ORDEN_COMPRA__NUEVA, Permisos.ComprasOCControl
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_COMPRAS__ORDEN_COMPRA__LISTADO, Permisos.ComprasOCConsultas


    Set ribbonGroup = ribbonTab.Groups.AddGroup("Precios", ID_GROUP.ID_GROUP_COMPRAS__PRECIOS)
    AddButton ribbonGroup, "Administrar", ID_BUTTON.ID_BUTTON_COMPRAS__PRECIOS__ADMINISTRAR, Permisos.ComprasAdminPrecios
    AddButton ribbonGroup, "Histórico", ID_BUTTON.ID_BUTTON_COMPRAS__PRECIOS__HISTORICO, Permisos.ComprasVerPrecios

    'PLANEAMIENTO--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(4, "Planeamiento")
    ribbonTab.Id = ID_TAB.ID_TAB_PLANEAMIENTO
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Orden de Trabajo", ID_GROUP.ID_GROUP_PLANEAMIENTO__ORDEN_TRABAJO)
    AddButton ribbonGroup, "Nueva", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__NUEVA, Permisos.PlanOTcontrol
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__LISTADO, Permisos.PlanOTconsultas, ID_BUTTON.ID_BUTTON_SEARCH
    AddButton ribbonGroup, "Nuevo Contr. Abierto", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__MARCO_NUEVA, Permisos.PlanOTcontrol
    AddButton ribbonGroup, "Planificación", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__PLANIFICAR, Permisos.PlanOTcontrol
    AddButton ribbonGroup, "Histórico", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_TRABAJO__HISTORICO, Permisos.PlanOTconsultas, ID_BUTTON.ID_BUTTON_HISTORICO

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Orden de Entrega", ID_GROUP.ID_GROUP_PLANEAMIENTO__ORDEN_ENTREGA)
    AddButton ribbonGroup, "Nueva", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_ENTREGA__NUEVA, Permisos.PlanOEcontrol
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__ORDEN_ENTREGA__LISTADO, Permisos.PlanOEconsultas, ID_BUTTON.ID_BUTTON_SEARCH

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Seguimiento", ID_GROUP.ID_GROUP_PLANEAMIENTO__SEGUIMIENTO)
    AddButton ribbonGroup, "Global", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__GLOBAL, Permisos.PlanSeguimientoGlobal
    AddButton ribbonGroup, "De rutas", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__DE_RUTAS, Permisos.PlanSeguimientoRutas

    Set cmdBarCtrl = AddButton(ribbonGroup, "Seguimiento Especial", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE, , , xtpControlPopup)
    AddButton ribbonGroup, "Por Tareas", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_PROCESOS, Permisos.PlanSeguimientoRutas, , , cmdBarCtrl
    AddButton ribbonGroup, "Por Rutas", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_RUTAS, Permisos.PlanSeguimientoRutas, , , cmdBarCtrl
    AddButton ribbonGroup, "Inicio tareas asignadas", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_INICIO_TAREAS_ASIGNADAS, Permisos.PlanSeguimientoRutas, , , cmdBarCtrl
    AddButton ribbonGroup, "Fin tareas iniciadas", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BARCODE_FIN_TAREAS_ASIGNADAS, Permisos.PlanSeguimientoRutas, , , cmdBarCtrl

    AddButton ribbonGroup, "Ver tiempos", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__VER_TIEMPOS, Permisos.PlanSeguimientoRutas
    AddButton ribbonGroup, "Operarios en Proceso", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__OPER_PROCESO, Permisos.PlanSeguimientoRutas
    AddButton ribbonGroup, "Ver NNC", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__VER_NNC, Permisos.PlanOTconsultas
    AddButton ribbonGroup, "Búsqueda de programas", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__SEGUIMIENTO__BUSQUEDA_PROGRAMAS, Permisos.PlanOTconsultas

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Remitos", ID_GROUP.ID_GROUP_PLANEAMIENTO__REMITOS)
    AddButton ribbonGroup, "Nuevo", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__REMITOS__NUEVO, Permisos.PlanRemitosControl
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_PLANEAMIENTO__REMITOS__LISTADO, Permisos.PlanRemitosConsultas, ID_BUTTON.ID_BUTTON_SEARCH

    'DESARROLLO--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(5, "Desarrollo")
    ribbonTab.Id = ID_TAB.ID_TAB_DESARROLLO
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Piezas", ID_GROUP.ID_GROUP_DESARROLLO__CENTRO_DE_COSTOS)
    AddButton ribbonGroup, "Nuevo elemento", ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__NUEVO_ELEMENTO, Permisos.DesaControl
    AddButton ribbonGroup, "Nuevo conjunto", ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__NUEVO_CONJUNTO, Permisos.DesaControl
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__LISTADO, Permisos.DesaConsultas, ID_BUTTON.ID_BUTTON_SEARCH
    'AddButton ribbonGroup, "Ver tiempos", ID_BUTTON.ID_BUTTON_DESARROLLO__CENTRO_DE_COSTOS__VER_TIEMPOS

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Materia Prima", ID_GROUP.ID_GROUP_DESARROLLO__MATERIA_PRIMA)
    AddButton ribbonGroup, "Nuevo material", ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__NUEVO_MATERIAL, Permisos.sistemaMaterialesConfig

    AddButton ribbonGroup, "Rubros", ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__RUBROS, Permisos.sistemaMaterialesConfig
    AddButton ribbonGroup, "Grupos", ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__GRUPOS, Permisos.sistemaMaterialesConfig

    AddButton ribbonGroup, "Almacenes", ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__ALMACENES, Permisos.sistemaPanelControlGeneral
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__LISTADO, Permisos.sistemaMaterialesConfig, ID_BUTTON.ID_BUTTON_SEARCH
    AddButton ribbonGroup, "Historial", ID_BUTTON.ID_BUTTON_DESARROLLO__MATERIA_PRIMA__HISTORIAL, Permisos.sistemaMaterialesConfig, ID_BUTTON.ID_BUTTON_HISTORICO


    Set ribbonGroup = ribbonTab.Groups.AddGroup("Mano de Obra", ID_GROUP.ID_GROUP_DESARROLLO__MANO_DE_OBRA)
    AddButton ribbonGroup, "Sectores", ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__SECTORES, Permisos.sistemaManoObraConfig
    AddButton ribbonGroup, "Sectorizar", ID_BUTTON.ID_BUTTON_PANEL_DE_CONTROL__USUARIOS_EMPLEADOS__SECTORIZAR, Permisos.sistemaPanelControlGeneral

    AddButton ribbonGroup, "Nueva tarea", ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__NUEVA_TAREA, Permisos.sistemaManoObraConfig

    AddButton ribbonGroup, "Tareas", ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__TAREAS, Permisos.sistemaManoObraConfig

    AddButton ribbonGroup, "Categoria sueldos", ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__CATEGORIA_SUELDOS, Permisos.sistemaManoObraConfig
    AddButton ribbonGroup, "Sueldo", ID_BUTTON.ID_BUTTON_DESARROLLO__MANO_DE_OBRA__SUELDO, Permisos.sistemaManoObraConfig

    'ADMINISTRAción--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(6, "Administración")
    ribbonTab.Id = ID_TAB.ID_TAB_ADMINISTRACION
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Ventas", ID_GROUP.ID_GROUP_ADMINISTRACION__FACTURACION)

    AddButton ribbonGroup, "Remitos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__REMITOS, Permisos.PlanRemitosConsultas

    Set cmdBarCtrl = AddButton(ribbonGroup, "Nueva", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA, Permisos.AdminFacturaControl, , xtpControlButtonPopup)
    AddButton ribbonGroup, "Factura", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_FC, Permisos.AdminFacturaControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Nota de Débito", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_ND, Permisos.AdminFacturaControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Nota de Crédito", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_NC, Permisos.AdminFacturaControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Factura Anticipo", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__NUEVA_ANTICIPO, Permisos.AdminFacturaControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Comprobantes", ID_BUTTON.ID_BUTTON_ADMINISTRACION__FACTURACION__FACTURAS, Permisos.AdminFacturaConsultas

    'Set ribbonGroup = ribbonTab.Groups.AddGroup("Ctas. Ctes. Clientes", ID_GROUP.ID_GROUP_ADMINISTRACION__CUENTAS_CORRIENTES)

    Set cmdBarCtrl = AddButton(ribbonGroup, "Cta. Cte. Clientes", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES, , , xtpControlButtonPopup)
    'AddButton ribbonGroup, "Crear Recibo de Anticipo", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__CREAR_RECIBO, Permisos.AdminCtaCteControl, , , cmdBarCtrl
    'AddButton ribbonGroup, "Ver Recibos de Anticipo", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__VER_RECIBOS, Permisos.AdminCtaCteControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Ver Detalle de Cta. Cte.", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__MOVIMIENTOS, Permisos.AdminCtaCteControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Resúmen Saldos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CTAS_CTES__SALDOS, Permisos.AdminCtaCteControl, , , cmdBarCtrl

    Set cmdBarCtrl = AddButton(ribbonGroup, "Informes", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES, , , xtpControlButtonPopup)
    AddButton ribbonGroup, "Por cliente", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_CLIENTE, Permisos.AdminInformesVarios, , , cmdBarCtrl
    AddButton ribbonGroup, "Por período", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_PERIODO, Permisos.AdminInformesVarios, , , cmdBarCtrl
    AddButton ribbonGroup, "Cashflow", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COBRANZAS__DEUDORES, Permisos.AdminFacturaConsultas, , , cmdBarCtrl

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Cobranzas", ID_GROUP.ID_GROUP_ADMINISTRACION__COBRANZAS)
    AddButton ribbonGroup, "Recibos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COBRANZAS__RECIBOS, Permisos.AdminCobroConsulta
    AddButton ribbonGroup, "Nuevo recibo", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COBRANZAS__NUEVO_RECIBO, Permisos.AdminCobroControl

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Caja y Bancos", ID_GROUP.ID_GROUP_ADMINISTRACION__CAJAYBANCOS)
    AddButton ribbonGroup, "Cheques", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CHEQUES, Permisos.AdminCajayBancos
    AddButton ribbonGroup, "Depositos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CHEQUES_DEPOSITAR, Permisos.AdminCajayBancos

    AddButton ribbonGroup, "Compensatorios", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__COMPENSATORIOS, Permisos.AdminOPControl
    AddButton ribbonGroup, "Resúmen de pagos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__RESUMEN_PAGOS, Permisos.AdminOPConsultas
'    AddButton ribbonGroup, "Movimiento de Fondos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__MOVIMIENTO__FONDOS, Permisos.AdminOPConsultas

    AddButton ribbonGroup, "Bancos", ID_BUTTON.ID_BUTTON_CAJAYBANCOS__CONFIGURAR__ADMINISTRACION__BANCOS, Permisos.AdminCajayBancos
    AddButton ribbonGroup, "Cuentas", ID_BUTTON.ID_BUTTON_CAJAYBANCOS__CONFIGURAR__ADMINISTRACION__CUENTAS, Permisos.AdminCajayBancos

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Pagos", ID_GROUP.ID_GROUP_ADMINISTRACION__VARIOS)

    Set cmdBarCtrl = AddButton(ribbonGroup, "Liquidaciones de Caja", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES, , , xtpControlButtonPopup)
    AddButton ribbonGroup, "Crear Liquidación", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__CREAR_LIQUIDACION_CAJA, , , , cmdBarCtrl
    AddButton ribbonGroup, "Crear Liquidación con DataGrid", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__CREAR_LIQUIDACION_CAJA_DG, , , , cmdBarCtrl
    AddButton ribbonGroup, "Ver Listado de Liquidaciones", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__LISTA_LIQUIDACION_CAJA, , , , cmdBarCtrl

    Set cmdBarCtrl = AddButton(ribbonGroup, "Ordenes de Pago", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES, , , xtpControlButtonPopup)
    AddButton ribbonGroup, "Crear Orden Pago", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__CREAR_ORDEN_PAGO, Permisos.AdminOPControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Ver Listado de Ordenes de Pago", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__ORDEN_PAGO_LISTA, Permisos.AdminOPConsultas, , , cmdBarCtrl

    AddButton ribbonGroup, "Transferencias", ID_BUTTON.ID_BUTTON_ADMINISTRACION__CAJABANCOS__TRANSFERENCIAS, Permisos.AdminCajayBancos
   
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Compras", ID_GROUP.ID_GROUP_ADMINISTRACION__COMPRAS)
    AddButton ribbonGroup, "Ingresar Comprobante", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__NUEVA
    AddButton ribbonGroup, "Listado Comprobante", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__LISTADO

    Set cmdBarCtrl = AddButton(ribbonGroup, "Plan de cuentas", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__PLAN_DE_CUENTAS, , , xtpControlButtonPopup)
    AddButton ribbonGroup, "Ver", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__PLAN_DE_CUENTAS_VER, , , , cmdBarCtrl
    AddButton ribbonGroup, "Definir", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__PLAN_DE_CUENTAS_DEFINIR, , , , cmdBarCtrl

    AddButton ribbonGroup, "Cta. Cte.", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__CTA_CTE
    AddButton ribbonGroup, "Resúmen Saldos", ID_BUTTON.ID_BUTTON_ADMINISTRACION__COMPRAS__RESUMEN__SALDOS

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Varios", ID_GROUP.ID_GROUP_ADMINISTRACION__VARIOS)
    AddButton ribbonGroup, "Padrones IIBB", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__PADRON_IIBB, Permisos.AdminIIBB
    AddButton ribbonGroup, "Centro de cambio", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__CENTRO_CAMBIO, Permisos.AdminCentroCambio

    Set cmdBarCtrl = AddButton(ribbonGroup, "Subdiarios", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS, , , xtpControlButtonPopup)
    AddButton ribbonGroup, "IVA Ventas", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_VENTAS, Permisos.AdminSubdiariosControl, , , cmdBarCtrl
    AddButton ribbonGroup, "IVA Compras", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_IVACOMPRAS, Permisos.AdminSubdiariosControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Cobranzas", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_COBRANZAS, Permisos.AdminSubdiariosControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Retenciones", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_RETENCIONES, Permisos.AdminSubdiariosControl, , , cmdBarCtrl
    AddButton ribbonGroup, "Posición IVA Mensual", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_POSICION_IVA_MENSUAL, Permisos.AdminSubdiariosControl, , , cmdBarCtrl

    ' REPORTE DE COMPARAción DE COMPROBANTES SIGNO VS AFIP
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Extras", ID_GROUP.ID_GROUP_ADMINISTRACION__EXTRAS)
    AddButton ribbonGroup, "Comparación Compras", ID_BUTTON_ADMINISTRACION__EXTRAS__REPORTE_CMC, Permisos.AdminSubdiariosControl
    'frmAdminExtrasReporteCMC

    '        Set cmdBarCtrl = AddButton(ribbonGroup, "Subdiarios Compras ", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_COMPRAS, , , xtpControlButtonPopup)
    '            AddButton ribbonGroup, "IVA Compras", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__SUBDIARIOS_IVACOMPRAS, Permisos.AdminSubdiariosControl, , , cmdBarCtrl
    'AddButton ribbonGroup, "Por OT", ID_BUTTON.ID_BUTTON_ADMINISTRACION__VARIOS__INFORMES_ORDEN_TRABAJO,Permisos.AdminInformesVarios , , , cmdBarCtrl


    'CLIENTES PROVEEDORES--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(7, "Clientes / Proveedores")
    ribbonTab.Id = ID_TAB.ID_TAB_CLIENTES_PROVEEDORES
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Clientes", ID_GROUP.ID_GROUP_CLIENTES_PROVEEDORES__CLIENTES)
    AddButton ribbonGroup, "Nuevo", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__CLIENTES__NUEVO, Permisos.VentasClientesControl, ID_BUTTON.ID_BUTTON_NUEVA_PERSONA
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__CLIENTES__LISTADO, Permisos.VentasClientesConsultas
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Proveedores", ID_GROUP.ID_GROUP_CLIENTES_PROVEEDORES__PROVEEDORES)
    AddButton ribbonGroup, "Nuevo", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__PROVEEDORES__NUEVO, Permisos.ComprasProveedorControl, ID_BUTTON.ID_BUTTON_NUEVA_PERSONA
    AddButton ribbonGroup, "Listado", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__PROVEEDORES__LISTADO, Permisos.ComprasProveedorConsultas
    AddButton ribbonGroup, "Asociación de Rubros", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__PROVEEDORES__ASOCRUBROS, Permisos.ComprasProveedorConsultas

    Set ribbonGroup = ribbonTab.Groups.AddGroup("Agenda", ID_GROUP.ID_GROUP_CLIENTES_PROVEEDORES__AGENDA)
    'AddButton ribbonGroup, "Nuevo", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__AGENDA__NUEVO
    AddButton ribbonGroup, "Ver", ID_BUTTON.ID_BUTTON_CLIENTES_PROVEEDORES__AGENDA__VER




    'USUARIO--------------------------------------------------------------------------------------------------------------------

    Set ribbonTab = RibbonBar.InsertTab(8, "Usuario")
    ribbonTab.Id = ID_TAB.ID_TAB_USUARIO
    Set ribbonGroup = ribbonTab.Groups.AddGroup("Herramientas", ID_GROUP.ID_GROUP_USUARIO__HERRAMIENTAS)
    AddButton ribbonGroup, "Tablero", ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__TABLERO
    AddButton ribbonGroup, "Agenda", ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__AGENDA
    AddButton ribbonGroup, "Cambiar contraseÑa", ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__CAMBIAR_CONTRASEÑA
    AddButton ribbonGroup, "Eventos", ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS__EVENTOS
    AddButton ribbonGroup, "Asignación Eventos", ID_BUTTON.ID_BUTTON_USUARIO__HERRAMIENTAS___ASIGNACION_EVENTOS, (funciones.GetUserObj.usuario = "marceloto" Or funciones.GetUserObj.usuario = "nicolasba" Or funciones.GetUserObj.usuario = "raulco")

    CommandBars.ShowTabWorkspace True
    RibbonBar.EnableFrameTheme
    CommandBars.ShowTabWorkspace True


End Sub

Private Sub tmrEventos_Timer()
    If Not Permisos.SistemaVerEventos Then Exit Sub
    On Error GoTo err1
    If Me.TrayIcon.Tag = 0 Then
        Dim col As New Collection
        Set col = DAOEvento.FindAllByUser(funciones.GetUserObj.Id, True)
        If col.count > 0 Then
            Me.TrayIcon.Tag = 1
            If col.count = 1 Then

                Dim E As EVENTO
                Set E = col.item(0)

                Me.TrayIcon.ShowBalloonTip 10, "Eventos en Signo Plast ERP", E.descripcion & vbNewLine & "Haga click aquí para leerlos.", 1
            Else

                Me.TrayIcon.ShowBalloonTip 10, "Eventos en Signo Plast ERP", "Han ocurrido nuevos eventos o tiene eventos sin leer." & vbNewLine & "Haga click aquí para leerlos.", 1
            End If
        End If
    End If
    Exit Sub
err1:
End Sub

Private Sub tmrInformeAccidentes_Timer()
    If Not Permisos.RRHHInformeAccidente Then Exit Sub

    contMinutosInformesAccidente = contMinutosInformesAccidente + 1


    If contMinutosInformesAccidente >= 15 Then

        'If Permisos.RRHHInformeAccidente Then
        If Me.WindowState <> vbMinimized Then
            If IsSomething(funciones.GetUserObj.Empleado) Then
                Dim sin As Collection
                Set sin = DAOSiniestroPersonal.FindAll("sp.id_accidente IS null AND sp.id_empleado_supervisor = " & funciones.GetUserObj.Empleado.Id)
                If sin.count > 0 Then
                    Dim nrosSiniestro As String
                    nrosSiniestro = funciones.JoinCollectionValues(sin, ", ", "NroSiniestro")
                    MsgBox "Tiene los siguientes siniestros pendientes a crear informe de accidente:" & vbNewLine & nrosSiniestro, vbInformation + vbOKOnly
                End If
            End If
        End If
        'End If

        contMinutosInformesAccidente = 0

    End If
End Sub
Private Sub TrayIcon_BalloonTipClicked()
    On Error GoTo err1
    Me.TrayIcon.Tag = 0
    Static formEventos As frmEventos
    If Not IsSomething(formEventos) Then Set formEventos = New frmEventos
    formEventos.Show
    formEventos.llenar
    Me.SetFocus
    formEventos.ZOrder 0
    Exit Sub
err1:
End Sub

Private Sub TrayIcon_BalloonTipClosed()
    Me.TrayIcon.Tag = 0
End Sub

Private Sub TrayIcon_DblClick()
    Select Case Me.WindowState
    Case FormWindowStateConstants.vbMinimized, FormWindowStateConstants.vbNormal
        Me.WindowState = FormWindowStateConstants.vbMaximized
    End Select
    Me.SetFocus
End Sub


