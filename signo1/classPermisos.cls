VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim clssp As New classSignoplast
Dim vPlanOTaprobaciones As Boolean
Dim vPlanOEaprobaciones As Boolean
Dim vPlanRemitosAprobaciones As Boolean
Dim vPlanOTmodificar As Boolean
Dim vPlanOEmodificar As Boolean
Dim vPlanInfoPantalla As Boolean
Dim vPlanOEcontrol As Boolean
Dim vPlanOTcontrol As Boolean
Dim vPlanOEconsultas As Boolean
Dim vPlanOTconsultas As Boolean
Dim vPlanSeguimientoGlobal As Boolean
Dim vPlanSeguimientoRutas As Boolean
Dim vPlanRemitosControl As Boolean
Dim vPlanRemitosConsultas As Boolean
Dim vSistemaGrupoDefault As Long
Dim vSistemaRootPanelControl As Boolean
Dim vSistemaUsuarioActivo As Boolean
Dim vSistemaTablero As Boolean
Dim vSistemaAgendaVer As Boolean
Dim vSistemaAgendaModificar As Boolean
Dim vSistemaManoObraConfig As Boolean
Dim vSistemaMaterialesConfig As Boolean
Dim vSistemaVerPreciosConfig As Boolean
Dim vSistemaPanelControlGeneral As Boolean
Dim vVentasRoot As Boolean
Dim vVentasCotizAprobaciones As Boolean
Dim vVentasCotizControl As Boolean
Dim vVentasCotizConsultas As Boolean
Dim vVentasPedidosControl As Boolean
Dim vVentasPedidosConsultas As Boolean
Dim vVentasClientesControl As Boolean
Dim vVentasClientesConsultas As Boolean
Dim vVentasInfoPantalla As Boolean
Dim vVentasCotizModificar As Boolean
Dim vDesaRoot As Boolean
Dim vDesaInfoPantalla As Boolean
Dim vDesaControl As Boolean
Dim vDesaConsultas As Boolean
Dim vDesaConsultaTiempo As Boolean
Dim vDesaManejoStock As Boolean

Dim vAdminFacturasAprobaciones As Boolean
Dim vAdminCobrosAprobaciones As Boolean
Dim vAdminIIBB As Boolean
Dim vAdminIIBBactualizar As Boolean

Dim vAdminInformesCashFlow As Boolean
Dim vAdminInformesVarios As Boolean
Dim vAdminSubdiariosControl As Boolean
Dim vAdminRoot As Boolean
Dim vAdminFacturaControl As Boolean
Dim vAdminFacturaConsultas As Boolean
Dim vAdminCobroControl As Boolean
Dim vAdminCobroConsulta As Boolean
Dim vAdminCentroCambio As Boolean
Dim vAdminCtaCteControl As Boolean
Dim vAdminInfoPantalla As Boolean

Public Property Get PlanInfoPantalla() As Boolean
    PlanInfoPantalla = vPlanInfoPantalla
End Property
Public Property Get PlanOEcontrol() As Boolean
    PlanOEcontrol = vPlanOEcontrol
End Property
Public Property Get PlanOTcontrol() As Boolean
    PlanOTcontrol = vPlanOTcontrol
End Property
Public Property Get PlanOEconsultas() As Boolean
    PlanOEconsultas = vPlanOEconsultas
End Property
Public Property Get PlanOTconsultas() As Boolean
    PlanOTconsultas = vPlanOTconsultas
End Property
Public Property Get PlanSeguimientoGlobal() As Boolean
    PlanSeguimientoGlobal = vPlanSeguimientoGlobal
End Property
Public Property Get PlanSeguimientoRutas() As Boolean
    PlanSeguimientoRutas = vPlanSeguimientoRutas
End Property
Public Property Get PlanRemitosControl() As Boolean
    PlanRemitosControl = vPlanRemitosControl
End Property
Public Property Get PlanRemitosConsultas() As Boolean
    PlanRemitosConsultas = vPlanRemitosConsultas
End Property


Public Property Get ventasCotizAprobaciones() As Boolean
    ventasCotizAprobaciones = vVentasCotizAprobaciones
End Property
Public Property Get VentasCotizModificar() As Boolean
    VentasCotizModificar = vVentasCotizModificar
End Property
Public Property Get planOTaprobaciones() As Boolean
    planOTaprobaciones = vPlanOTaprobaciones
End Property
Public Property Get planOEaprobaciones() As Boolean
    planOEaprobaciones = vPlanOEaprobaciones
End Property
Public Property Get planRemitosAprobaciones() As Boolean
    planRemitosAprobaciones = vPlanRemitosAprobaciones
End Property
Public Property Get planOEmodificar() As Boolean
    planOEmodificar = vPlanOEmodificar
End Property
Public Property Get planOTmodificar() As Boolean
    planOTmodificar = vPlanOTmodificar
End Property
Public Property Get planRoot() As Boolean
    planRoot = vPlanRoot
End Property
Public Property Get sistemaGrupoDefault() As Long
    sistemaGrupoDefault = vSistemaGrupoDefault
End Property
Public Property Get sistemaRootPanelControl() As Boolean
    sistemaRootPanelControl = vSistemaRootPanelControl
End Property
Public Property Get sistemaUsuarioActivo() As Boolean
    sistemaUsuarioActivo = vSistemaUsuarioActivo
End Property
Public Property Get sistemaTablero() As Boolean
    sistemaTablero = vSistemaTablero
End Property
Public Property Get sistemaAgendaVer() As Boolean
    sistemaAgendaVer = vSistemaAgendaVer
End Property
Public Property Get sistemaAgendaModificar() As Boolean
    sistemaAgendaModificar = vSistemaAgendaModificar
End Property
Public Property Get sistemaManoObraConfig() As Boolean
    sistemaManoObraConfig = vSistemaManoObraConfig
End Property
Public Property Get sistemaMaterialesConfig() As Boolean
    sistemaMaterialesConfig = vSistemaMaterialesConfig
End Property
Public Property Get sistemaVerPrecios() As Boolean
    sistemaVerPrecios = vSistemaVerPrecios
End Property
Public Property Get sistemaPanelControlGeneral() As Boolean
    sistemaPanelControlGeneral = vSistemaPanelControlGeneral
End Property
Public Property Get VentasRoot() As Boolean
    VentasRoot = vVentasRoot
End Property
Public Property Get VentasCotizControl() As Boolean
    VentasCotizControl = vVentasCotizControl
End Property
Public Property Get VentasCotizConsultas() As Boolean
    VentasCotizConsultas = vVentasCotizConsultas
End Property
Public Property Get VentasPedidosControl() As Boolean
    VentasPedidosControl = vVentasPedidosControl
End Property
Public Property Get VentasPedidosConsultas() As Boolean
    VentasPedidosConsultas = vVentasPedidosConsultas
End Property
Public Property Get VentasClientesControl() As Boolean
    VentasClientesControl = vVentasClientesControl
End Property
Public Property Get VentasClientesConsultas() As Boolean
    VentasClientesConsultas = vVentasClientesConsultas
End Property
Public Property Get VentasInfoPantalla() As Boolean
    VentasInfoPantalla = vVentasInfoPantalla
End Property
Public Property Get VentasCotizacionesModificar() As Boolean
    VentasCotizacionesModificar = vVentasCotizModificar
End Property


Public Property Get DesaRoot() As Boolean
    DesaRoot = vDesaRoot
End Property
Public Property Get DesaInfoPantalla() As Boolean
    DesaInfoPantalla = vDesaInfoPantalla
End Property
Public Property Get DesaControl() As Boolean
    DesaControl = vDesaControl
End Property
Public Property Get DesaConsultas() As Boolean
    DesaConsultas = vDesaConsultas
End Property
Public Property Get DesaConsultaTiempo() As Boolean
    DesaConsultaTiempo = vDesaConsultaTiempo
End Property
Public Property Get DesaManejoStock() As Boolean
    DesaManejoStock = vDesaManejoStock
End Property
Public Property Get AdminRoot() As Boolean
    AdminRoot = vAdminRoot
End Property
Public Property Get AdminFacturaControl() As Boolean
    AdminFacturaControl = vAdminFacturaControl
End Property
Public Property Get AdminFacturaConsultas() As Boolean
    AdminFacturaConsultas = vAdminFacturaConsultas
End Property
Public Property Get AdminCobroControl() As Boolean
    AdminCobroControl = vAdminCobroControl
End Property
Public Property Get AdminCobroConsulta() As Boolean
    AdminCobroConsulta = vAdminCobroConsulta
End Property
Public Property Get AdminCentroCambio() As Boolean
    AdminCentroCambio = vAdminCentroCambio
End Property
Public Property Get AdminCtaCteControl() As Boolean
    AdminCtaCteControl = vAdminCtaCteControl
End Property
Public Property Get AdminInfoPantalla() As Boolean
    AdminInfoPantalla = vAdminInfoPantalla
End Property
Public Property Get AdminIIBB() As Boolean
    AdminIIBB = vAdminIIBB
End Property
Public Property Get AdminIIBBactualizar() As Boolean
    AdminIIBBactualizar = vAdminIIBBactualizar
End Property
Public Property Get AdminFacturasAprobaciones() As Boolean
    AdminFacturasAprobaciones = vAdminFacturasAprobaciones
End Property
Public Property Get AdminCobrosAprobaciones() As Boolean
    AdminCobrosAprobaciones = vAdminCobrosAprobaciones
End Property

Public Property Get AdminInformesCashFlow() As Boolean
    AdminInformesCashFlow = vAdminInformesCashFlow
End Property
Public Property Get AdminInformesVarios() As Boolean
    AdminInformesVarios = vAdminInformesVarios
End Property
Public Property Get AdminSubdiariosControl() As Boolean
    AdminSubdiariosControl = vAdminSubdiariosControl
End Property

Public Function crearPermisos(idUsuario As Long) As Boolean

'Plan
    vPlanOTmodificar = clssp.verSeleccionado(strPermisos.ePlanOTmodificar, idUsuario)
    vPlanOEmodificar = clssp.verSeleccionado(strPermisos.ePlanOEmodificar, idUsuario)
    vPlanOTaprobaciones = clssp.verSeleccionado(strPermisos.ePlanOTaprobaciones, idUsuario)
    vPlanOEaprobaciones = clssp.verSeleccionado(strPermisos.ePlanOTaprobaciones, idUsuario)
    vPlanRemitosAprobaciones = clssp.verSeleccionado(strPermisos.ePlanOTaprobaciones, idUsuario)
    vPlanInfoPantalla = clssp.verSeleccionado(strPermisos.ePlanInfoPantalla, idUsuario)
    vPlanOEcontrol = clssp.verSeleccionado(strPermisos.ePlanOEcontrol, idUsuario)
    vPlanOTcontrol = clssp.verSeleccionado(strPermisos.ePlanOTcontrol, idUsuario)
    vPlanOEconsultas = clssp.verSeleccionado(strPermisos.ePlanOTconsultas, idUsuario)
    vPlanOTconsultas = clssp.verSeleccionado(strPermisos.ePlanOTconsultas, idUsuario)
    vPlanSeguimientoGlobal = clssp.verSeleccionado(strPermisos.ePlanSeguimientoGlobal, idUsuario)
    vPlanSeguimientoRutas = clssp.verSeleccionado(strPermisos.ePlanSeguimientoRutas, idUsuario)
    vPlanRemitosControl = clssp.verSeleccionado(strPermisos.ePlanRemitosControl, idUsuario)
    vPlanRemitosConsultas = clssp.verSeleccionado(strPermisos.ePlanRemitosConsultas, idUsuario)
    vPlanRoot = clssp.verSeleccionado(strPermisos.ePlanRoot, idUsuario)
    'ventas
    vVentasCotizAprobaciones = clssp.verSeleccionado(strPermisos.eVentasCotizacionesAprobaciones, idUsuario)
    vVentasCotizModificar = clssp.verSeleccionado(strPermisos.eVentasCotizModif, idUsuario)
    vVentasCotizControl = clssp.verSeleccionado(strPermisos.eVentasCotizacionesConsultas, idUsuario)
    vVentasCotizConsultas = clssp.verSeleccionado(strPermisos.eVentasCotizacionesConsultas, idUsuario)
    vVentasPedidosControl = clssp.verSeleccionado(strPermisos.eVentasPedidoControl, idUsuario)
    vVentasPedidosConsultas = clssp.verSeleccionado(strPermisos.eVentasPedidoConsultas, idUsuario)
    vVentasClientesControl = clssp.verSeleccionado(strPermisos.eVentasClienteControl, idUsuario)
    vVentasClientesConsultas = clssp.verSeleccionado(strPermisos.eVentasClienteConsultas, idUsuario)
    vVentasInfoPantalla = clssp.verSeleccionado(strPermisos.eVentasInfoPantalla, idUsuario)
    vVentasCotizModificar = clssp.verSeleccionado(strPermisos.eVentasCotizModif, idUsuario)
    vVentasRoot = clssp.verSeleccionado(strPermisos.eVentasRoot, idUsuario)

    'Sistema
    vSistemaGrupoDefault = CLng(clssp.verSeleccionado(strPermisos.eSistemaGrupoDefault, idUsuario))
    vSistemaRootPanelControl = clssp.verSeleccionado(strPermisos.eSistemaRootPanelControl, idUsuario)
    vSistemaUsuarioActivo = clssp.verSeleccionado(strPermisos.esistemaUsuarioActivo, idUsuario)
    vSistemaTablero = clssp.verSeleccionado(strPermisos.eSistemaTablero, idUsuario)
    vSistemaAgendaVer = clssp.verSeleccionado(strPermisos.eSistemaAgendaVer, idUsuario)
    vSistemaAgendaModificar = clssp.verSeleccionado(strPermisos.eSistemaAgendaModificar, idUsuario)
    vSistemaManoObraConfig = clssp.verSeleccionado(strPermisos.eSistemaManoObraConfig, idUsuario)
    vSistemaMaterialesConfig = clssp.verSeleccionado(strPermisos.eSistemaMaterialesConfig, idUsuario)
    vSistemaVerPrecios = clssp.verSeleccionado(strPermisos.eSistemaVerPrecios, idUsuario)
    vSistemaPanelControlGeneral = clssp.verSeleccionado(strPermisos.eSistemaPanelControlGeneral, idUsuario)

    vDesaRoot = clssp.verSeleccionado(strPermisos.eDesaRoot, idUsuario)
    vDesaInfoPantalla = clssp.verSeleccionado(strPermisos.eDesaInfoPantalla, idUsuario)
    vDesaControl = clssp.verSeleccionado(strPermisos.eDesaControl, idUsuario)
    vDesaConsultas = clssp.verSeleccionado(strPermisos.eDesaConsultas, idUsuario)
    vDesaConsultaTiempo = clssp.verSeleccionado(strPermisos.eDesaConsultaTiempo, idUsuario)
    vDesaManejoStock = clssp.verSeleccionado(strPermisos.eDesaManejoStock, idUsuario)

    vAdminRoot = clssp.verSeleccionado(strPermisos.eAdminRoot, idUsuario)
    vAdminFacturaControl = clssp.verSeleccionado(strPermisos.eAdminFacturaControl, idUsuario)
    vAdminFacturaConsultas = clssp.verSeleccionado(strPermisos.eAdminFacturaConsultas, idUsuario)
    vAdminCobroControl = clssp.verSeleccionado(strPermisos.eAdminCobroControl, idUsuario)
    vAdminCobroConsulta = clssp.verSeleccionado(strPermisos.eAdminCobroConsulta, idUsuario)
    vAdminCentroCambio = clssp.verSeleccionado(strPermisos.eAdminCentroCambio, idUsuario)
    vAdminCtaCteControl = clssp.verSeleccionado(strPermisos.eAdminCtaCteControl, idUsuario)
    vAdminInfoPantalla = clssp.verSeleccionado(strPermisos.eAdminInfoPantalla, idUsuario)
    vAdminFacturasAprobaciones = clssp.verSeleccionado(strPermisos.eAdminFacturaAprobaciones, idUsuario)
    vAdminCobrosAprobaciones = clssp.verSeleccionado(strPermisos.eAdminCobroAprobaciones, idUsuario)
    vAdminIIBB = clssp.verSeleccionado(strPermisos.eAdminIIBB, idUsuario)
    vAdminIIBBactualizar = clssp.verSeleccionado(strPermisos.eAdminIIBBactualizar, idUsuario)
    vAdminInformesCashFlow = clssp.verSeleccionado(strPermisos.eAdminInformesCashFlow, idUsuario)
    vAdminInformesVarios = clssp.verSeleccionado(strPermisos.eAdminInformesVarios, idUsuario)
    vAdminSubdiariosControl = clssp.verSeleccionado(strPermisos.eAdminSubdiariosControl, idUsuario)
End Function


