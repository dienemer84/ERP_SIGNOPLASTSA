Attribute VB_Name = "DAOLiquidacionCaja"
Option Explicit

Public Function FindAbonadoPendienteEnEstaOP(facid As Long, ocid As Long) As Collection

    Dim q As String

    q = "SELECT IFNULL( (SELECT SUM(total_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja = " & ocid & "),0 ) AS total_pendiente, " _
        & " IFNULL( (SELECT SUM(neto_gravado_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja = " & ocid & "),0 ) AS netogravado_pendiente, " _
        & " IFNULL( (SELECT SUM(otros_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja = " & ocid & "),0 ) AS otros_pendiente "

    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim tot As Double, ng As Double, Otros As Double
    tot = rs!total_pendiente
    ng = rs!netogravado_pendiente
    Otros = rs!otros_pendiente

    Dim C As New Collection
    C.Add tot
    C.Add ng
    C.Add Otros
    Set FindAbonadoPendienteEnEstaOP = C

End Function


Public Function FindAbonadoPendiente(facid As Long, ocid As Long) As Collection

    Dim q As String

    q = "SELECT IFNULL( (SELECT SUM(total_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja <> " & ocid & "),0 ) AS total_pendiente, " _
        & " IFNULL( (SELECT SUM(neto_gravado_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja <> " & ocid & "),0 ) AS netogravado_pendiente, " _
        & " IFNULL( (SELECT SUM(otros_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja <> " & ocid & "),0 ) AS otros_pendiente "

    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim tot As Double, ng As Double, Otros As Double
    tot = rs!total_pendiente
    ng = rs!netogravado_pendiente
    Otros = rs!otros_pendiente

    Dim C As New Collection
    C.Add tot
    C.Add ng
    C.Add Otros
    Set FindAbonadoPendiente = C
End Function


Public Function FindAbonadoFactura(facid As Long, ocid As Long) As Collection

    Dim q As String

    q = "SELECT IFNULL( (SELECT SUM(total_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja = " & ocid & " and op1.estado=1),0 ) AS total_pendiente, " _
        & " IFNULL( (SELECT SUM(neto_gravado_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja = " & ocid & " and op1.estado=1),0 ) AS netogravado_pendiente, " _
        & " IFNULL( (SELECT SUM(otros_liquidado) FROM liquidaciones_caja_facturas opf JOIN ordenes_pago op1 ON opf.id_liquidacion_caja=op1.id " _
        & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_liquidacion_caja = " & ocid & " and op1.estado=1),0 ) AS otros_pendiente "

    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim tot As Double, ng As Double, Otros As Double
    tot = rs!total_pendiente
    ng = rs!netogravado_pendiente
    Otros = rs!otros_pendiente

    Dim C As New Collection
    C.Add tot
    C.Add ng
    C.Add Otros
    Set FindAbonadoFactura = C

End Function


Public Function FindLast() As OrdenPago
    Set FindLast = FindAll("ordenes_pago.id = (SELECT MAX(id) FROM ordenes_pago)")(1)
End Function


Public Function FindByFacturaId(facid As Long) As OrdenPago
    Dim col As Collection
    Set col = FindAll("ordenes_pago.id = (SELECT DISTINCT id_liquidacion_caja from liquidaciones_caja_facturas opf inner join ordenes_pago op on opf.id_liquidacion_caja=op.id WHERE id_factura_proveedor = " & facid & " AND op.estado=1)")
    If col.count > 0 Then
        Set FindByFacturaId = col(1)
    Else
        Set FindByFacturaId = Nothing
    End If
End Function


Public Function FindAllByProveedor(provid As Long, Optional cond As String, Optional soloOp As Boolean = False) As Collection
    Dim q As String
    q = "ordenes_pago.id IN (SELECT DISTINCT opf.id_liquidacion_caja from liquidaciones_caja_facturas opf INNER JOIN AdminComprasFacturasProveedores cfp ON  cfp.id = opf.id_factura_proveedor WHERE cfp.id_proveedor = " & provid & " )"

    If LenB(cond) > 0 Then
        q = q & "  " & cond
        'ver aca
    End If

    If soloOp Then
        Set FindAllByProveedor = FindAllSoloOP(q)
        '  Debug.Print FindAllByProveedor.count

    Else
        Set FindAllByProveedor = FindAll(q)
        '  Debug.Print FindAllByProveedor.count
    End If
End Function


Public Function FindById(Id As Long) As clsLiquidacionCaja
    Dim col As Collection
    Set col = FindAll("liquidaciones_caja.id=" & Id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function


Public Function FindByNumeroLiq(NumeroLiq As Long) As clsLiquidacionCaja
    Dim col As Collection
    Set col = FindAll("liquidaciones_caja.numero_liq=" & NumeroLiq)
    If col.count = 0 Then
        Set FindByNumeroLiq = Nothing
    Else
        Set FindByNumeroLiq = col.item(1)
    End If
End Function


Public Function FindAllSoloOP(Optional filter As String = "1 = 1", Optional orderBy As String = "1") As Collection
    Dim q As String
    q = "SELECT * " _
        & " From ordenes_pago" _
        & " LEFT JOIN AdminConfigMonedas ON (AdminConfigMonedas.id = ordenes_pago.id_moneda)"

    q = q & " WHERE " & filter
    q = q & " ORDER BY " & orderBy
    Dim col As New Collection
    Dim op As OrdenPago
    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx
    While Not rs.EOF
        Set op = Map(rs, idx, "ordenes_pago", "AdminConfigMonedas")    ', "certificados_retencion")
        If funciones.BuscarEnColeccion(col, CStr(op.Id)) Then
            Set op = col.item(CStr(op.Id))
        Else
            col.Add op, CStr(op.Id)
        End If
        rs.MoveNext
    Wend

    Set FindAllSoloOP = col

End Function


Public Function FindAll(Optional filter As String = "1 = 1", Optional orderBy As String = "1") As Collection
    Dim q As String

    q = "SELECT *, (operaciones.pertenencia + 0) AS pertenencia2" _
        & " FROM liquidaciones_caja" _
        & " LEFT JOIN ordenes_pago_cheques ON (liquidaciones_caja.id = ordenes_pago_cheques.id_orden_pago)" _
        & " LEFT JOIN ordenes_pago_operaciones ON (liquidaciones_caja.id = ordenes_pago_operaciones.id_orden_pago)" _
        & " LEFT JOIN liquidaciones_caja_facturas ON (liquidaciones_caja.id = liquidaciones_caja_facturas.id_liquidacion_caja)" _
        & " LEFT JOIN AdminComprasCuentasContables cuentacontableordenpago ON (liquidaciones_caja.id_cuenta_contable = cuentacontableordenpago.id)" _
        & " LEFT JOIN operaciones ON (operaciones.id = ordenes_pago_operaciones.id_operacion)" _
        & " LEFT JOIN Cheques ON (Cheques.id = ordenes_pago_cheques.id_cheque)" _
        & " LEFT JOIN Chequeras ON (Chequeras.id = Cheques.id_chequera)" _
        & " LEFT JOIN AdminConfigBancos monbanco ON (monbanco.id = Chequeras.id_banco)" _
        & " LEFT JOIN AdminConfigMonedas monchequera ON (monchequera.id = Chequeras.id_moneda)" _
        & " LEFT JOIN AdminComprasFacturasProveedores ON (AdminComprasFacturasProveedores.id = liquidaciones_caja_facturas.id_factura_proveedor)" _
        & " LEFT JOIN AdminConfigMonedas ON (AdminConfigMonedas.id = liquidaciones_caja.id_moneda)" _
        & " LEFT JOIN AdminConfigMonedas monFacProv ON (monFacProv.id = AdminComprasFacturasProveedores.id_moneda)" _
        & " LEFT JOIN AdminConfigFacturasProveedor ON (AdminComprasFacturasProveedores.id_config_factura = AdminConfigFacturasProveedor.id)" _
        & " LEFT JOIN AdminConfigMonedas monedaoperacion ON (monedaoperacion.id = operaciones.moneda_id)" _
        & " LEFT JOIN AdminComprasCuentasContables ON (AdminComprasCuentasContables.id = operaciones.cuenta_contable_id)" _
        & " LEFT JOIN cajas ON (cajas.id = operaciones.cuentabanc_o_caja_id)" _
        & " LEFT JOIN AdminConfigCuentas ON (AdminConfigCuentas.id = operaciones.cuentabanc_o_caja_id)" _
        & " LEFT JOIN AdminConfigMonedas moncuentabancaria ON (moncuentabancaria.id = AdminConfigCuentas.moneda_id)" _
        & " LEFT JOIN AdminConfigMonedas moncheque ON (moncheque.id = Cheques.id_moneda)" _
        & " LEFT JOIN usuarios ON AdminComprasFacturasProveedores.id_usuario_creador=usuarios.id"
    q = q & " LEFT JOIN AdminConfigBancos ON (AdminConfigBancos.id = AdminConfigCuentas.idBanco)" _
        & " LEFT JOIN AdminConfigBancos bancocheque ON (bancocheque.id = Cheques.id_banco)" _
        & " LEFT JOIN proveedores ON (proveedores.id = AdminComprasFacturasProveedores.id_proveedor)"
    q = q & " LEFT JOIN ordenes_pago_retenciones opr ON opr.id_pago = liquidaciones_caja.id" _
        & " LEFT JOIN retenciones r ON r.id = opr.id_retencion "
    q = q & " WHERE " & filter
    q = q & " ORDER BY " & orderBy

    Dim col As New Collection
    Dim liq As clsLiquidacionCaja
    Dim fac As clsFacturaProveedor
    Dim oper As operacion

    Dim idx As Dictionary
    Dim rs As Recordset
    Dim ra As DTORetencionAlicuota

    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set liq = Map(rs, idx, "liquidaciones_caja", "AdminConfigMonedas", "cuentacontableordenpago", "retenciones")    ', "certificados_retencion")

        If funciones.BuscarEnColeccion(col, CStr(liq.Id)) Then
            Set liq = col.item(CStr(liq.Id))
        Else
            col.Add liq, CStr(liq.Id)
        End If

        Set fac = DAOFacturaProveedor.Map(rs, idx, "AdminComprasFacturasProveedores", "proveedores", "AdminConfigFacturasProveedor", , "monFacProv")

        If IsSomething(fac) Then
            If Not funciones.BuscarEnColeccion(liq.FacturasProveedor, CStr(fac.Id)) Then
                liq.FacturasProveedor.Add fac, CStr(fac.Id)
            End If

        End If

        Set oper = DAOOperacion.Map(rs, idx, "operaciones", "AdminComprasCuentasContables", "monedaoperacion", "AdminConfigCuentas", "cajas")
        If IsSomething(oper) Then
            If oper.Pertenencia = Banco Then
                If Not funciones.BuscarEnColeccion(liq.OperacionesBanco, CStr(oper.Id)) Then

                    liq.OperacionesBanco.Add oper, CStr(oper.Id)
                End If
            ElseIf oper.Pertenencia = caja Then
                If Not funciones.BuscarEnColeccion(liq.OperacionesCaja, CStr(oper.Id)) Then
                    liq.OperacionesCaja.Add oper, CStr(oper.Id)
                End If
            End If
        End If

        Set ra = MapAlicuotaRetencion(rs, idx, "opr", "r")
        If IsSomething(ra) Then
            If Not funciones.BuscarEnColeccion(liq.RetencionesAlicuota, CStr(ra.Retencion.Id)) Then
                liq.RetencionesAlicuota.Add ra, CStr(ra.Retencion.Id)
            End If
        End If

        rs.MoveNext
    Wend

    Set FindAll = col

End Function


Public Function Map(rs As Recordset, indice As Dictionary, _
                    tabla As String, _
                    Optional ByVal tablaMoneda As String = vbNullString, _
                    Optional ByVal tablaCuentaContable As String = vbNullString, _
                    Optional ByVal TablaRetenciones As String = vbNullString _
                    ) As clsLiquidacionCaja

    Dim liq As clsLiquidacionCaja

    'id_certificado_retencion
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set liq = New clsLiquidacionCaja
        liq.Id = Id

        liq.FEcha = GetValue(rs, indice, tabla, "fecha")
        liq.CuentaContableDescripcion = GetValue(rs, indice, tabla, "cuenta_contable_desc")
        liq.estado = GetValue(rs, indice, tabla, "estado")
        liq.alicuota = GetValue(rs, indice, tabla, "alicuota")

        liq.StaticTotalFacturas = GetValue(rs, indice, tabla, "static_total_facturas")
        liq.StaticTotalFacturasNG = GetValue(rs, indice, tabla, "static_total_factura_ng")
        liq.StaticTotalRetenido = GetValue(rs, indice, tabla, "static_total_a_retener")
        liq.StaticTotalOrigenes = GetValue(rs, indice, tabla, "static_total_origen")

        liq.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")
        liq.DiferenciaCambio = GetValue(rs, indice, tabla, "dif_cambio")
        liq.OtrosDescuentos = GetValue(rs, indice, tabla, "otros_descuentos")
        liq.DiferenciaCambioEnNG = GetValue(rs, indice, tabla, "dif_cambio_ng")
        liq.DiferenciaCambioEnTOTAL = GetValue(rs, indice, tabla, "dif_cambio_total")
        liq.NumeroLiq = GetValue(rs, indice, tabla, "numero_liq")
        If LenB(tablaCuentaContable) > 0 Then Set liq.CuentaContable = DAOCuentaContable.Map(rs, indice, tablaCuentaContable)
        If LenB(tablaMoneda) > 0 Then Set liq.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)

    End If

    Set Map = liq

End Function


Public Function MapAlicuotaRetencion(rs As Recordset, indice As Dictionary, _
                                     tabla As String, _
                                     ByVal TablaRetenciones As String) As DTORetencionAlicuota

    Dim ra As DTORetencionAlicuota

    'id_certificado_retencion
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, "id_retencion")

    If Id > 0 Then
        Set ra = New DTORetencionAlicuota
        ra.alicuotaRetencion = GetValue(rs, indice, tabla, "alicuota")
        Set ra.Retencion = DAORetenciones.Map(rs, indice, TablaRetenciones)
        ra.importe = GetValue(rs, indice, tabla, "total")

        'If LenB(tablaCertRetencion) > 0 Then Set op.CertificadoRetencion = DAOCertificadoRetencion.Map(rs, indice, tablaCertRetencion)
    End If

    Set MapAlicuotaRetencion = ra
End Function


Public Function Save(op As clsLiquidacionCaja, Optional cascada As Boolean = False) As Boolean
'Public Function Save(1=1, Optional cascada As Boolean = False) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    Save = Guardar(op, cascada)
    conectar.CommitTransaction
    Exit Function
err1:
    Save = False
    conectar.RollBackTransaction
End Function


Public Function aprobar(liq_mem As clsLiquidacionCaja, insideTransaction As Boolean) As Boolean

    On Error GoTo err1
    If insideTransaction Then conectar.BeginTransaction

    '3-10-2020 recargo la OP para que se actualicen los estados de las facturas y se validen bien

    Dim liq As clsLiquidacionCaja

    Set liq = DAOLiquidacionCaja.FindById(liq_mem.Id)

    If Not IsSomething(liq) Then
        GoTo err1
    End If

    'VALIDAR BIEN LOS TOTALES ANTES DE PODER APROBAR
    'verificar que las facturas esten todas aprobadsa...
    Dim F As clsFacturaProveedor
    Dim nopago As Double
    Dim nopago1 As Double

    Dim otrosvalores As Double

    Dim esf As EstadoFacturaProveedor
    For Each F In liq.FacturasProveedor

        '        Dim fac As clsFacturaProveedor
        '        Set fac = DAOFacturaProveedor.FindById(F.Id)
        '        Debug.Print (F.Id & "- " & F.NumeroFormateado & liq.Id)
        '
        '        If fac.estado = EstadoFacturaProveedor.EnProceso Then
        '            Err.Raise 44, "aprobar op", "La factura " & fac.NumeroFormateado & " no está aprobada. No se pudo aprobar la OP"
        '        End If
        '
        '        Dim X
        '
        '        Set X = DAOLiquidacionCaja.FindAbonadoPendienteEnEstaOP(fac.Id, liq.Id)
        '
        '        nopago1 = fac.total - fac.TotalAbonadoGlobal
        '        Debug.Print ("NoPago1: " & nopago1)
        '        otrosvalores = funciones.RedondearDecimales(funciones.RedondearDecimales(CDbl(X(1))) + funciones.RedondearDecimales(CDbl(X(2))) + funciones.RedondearDecimales(CDbl(X(3))))
        '        Debug.Print ("Otros Valores: " & otrosvalores)
        '        nopago = funciones.RedondearDecimales(nopago1) - otrosvalores - nopago1
        '        Debug.Print ("NoPago: " & nopago)
        '        esf = EstadoFacturaProveedor.Aprobada
        '        Debug.Print ("ESF: " & esf)
        '
        '        If nopago < 0 Then
        '            Err.Raise 44, "aprobar liquidacion", "El comprobante " & fac.NumeroFormateado & " tiene un error y no se pudo aprobar la OP"
        '        End If

        esf = EstadoFacturaProveedor.Saldada

        '        If nopago > 0 Then
        '            esf = EstadoFacturaProveedor.pagoParcial
        '        Else
        '            esf = EstadoFacturaProveedor.Saldada
        '        End If

        conectar.execute "UPDATE AdminComprasFacturasProveedores SET estado = " & esf & " WHERE id = " & F.Id

    Next F

    '    MsgBox (liq.Id)

    If liq.estado = EstadoLiquidacionCaja_pendiente Then
        Dim es As EstadoLiquidacionCaja
        es = liq.estado
        liq.estado = EstadoLiquidacionCaja_Aprobada

        If liq.EsParaFacturaProveedor Then


            If liq.FacturasProveedor.count > 0 Then
                If liq.FacturasProveedor(1).Proveedor.estado <> 2 Then
                    Dim d As New clsDTOPadronIIBB
                    'todo: cambiar validacion
                    Set d = DTOPadronIIBB.FindByCUIT(liq.FacturasProveedor(1).Proveedor.Cuit, TipoPadronRetencion)
                    Dim ret As Double

                    If IsSomething(d) Then
                        ret = d.alicuota
                    End If

                End If
            Else
                '                MsgBox "El proveedor es de tipo contado! " & vbNewLine & "No se le realizará ninguna retención!", vbInformation, "Información"
            End If
        End If
    End If


    'analizo las facturas de proveedores

    'TODO: debo verificar que los deudas por compensatorio no esten utilizadas en otra OP aprobada ni que esten ya canceladas en otro proceso

    If GuardarAprobada(liq) Then

        Dim fac1 As clsFacturaProveedor
        For Each fac1 In liq.FacturasProveedor
            If fac1.estado = EstadoFacturaProveedor.Saldada Then
                If Not DaoFacturaProveedorHistorial.agregar(fac1, "SALDADA") Then GoTo err1
            End If
            If fac1.estado = EstadoFacturaProveedor.pagoParcial Then
                If Not DaoFacturaProveedorHistorial.agregar(fac1, "PAGO PARCIAL") Then GoTo err1
            End If
        Next


        '        If liq.StaticTotalRetenido > 0 Then
        '
        '            Dim ra As DTORetencionAlicuota
        '            For Each ra In liq.RetencionesAlicuota
        '
        '
        '                If IsSomething(DAOCertificadoRetencion.Create(liq, ra.Retencion, ra.alicuotaRetencion, True)) Then
        '                    MsgBox "Se creo un certificado de Retenciones para la Orden de Pago. ", vbInformation
        '                Else
        '                    GoTo err1
        '                End If
        '            Next
        '
        '        End If
    Else
        GoTo err1
    End If


    DaoHistorico.Save "orden_pago_historial", "LC Aprobada", liq.Id
    aprobar = True
    If insideTransaction Then conectar.CommitTransaction
    Exit Function

err1:

    liq.estado = es
    If insideTransaction Then conectar.RollBackTransaction
    aprobar = False
End Function


Public Function Guardar(op As clsLiquidacionCaja, Optional cascada As Boolean = False) As Boolean

'TODO: tengo que revisar que las facturas no esten en otra op aprobada antes de continuar

    Dim q As String
    On Error GoTo E
    Dim Nueva As Boolean: Nueva = False

    If op.Id = 0 Then

        If IsSomething(DAOLiquidacionCaja.FindByNumeroLiq(CLng(op.NumeroLiq))) Then
            MsgBox "Ya existe una Liquidación con ese número!", vbCritical, "Error"
            Exit Function
        End If


        Nueva = True
        '        MsgBox ("Es nueva")
        q = "INSERT INTO liquidaciones_caja (id_moneda_pago,tipo_cambio,id_moneda, fecha, id_cuenta_contable,cuenta_contable_desc,estado,alicuota,static_total_facturas, static_total_factura_ng, static_total_a_retener, static_total_origen,dif_cambio, otros_descuentos,dif_cambio_ng,dif_cambio_total,numero_liq)" _
            & " VALUES ('id_moneda_pago','tipo_cambio','id_moneda', 'fecha', 'id_cuenta_contable', 'cuenta_contable_desc','0','alicuota','static_total_facturas', 'static_total_factura_ng', 'static_total_a_retener', 'static_total_origen', 'dif_cambio', 'otros_descuentos','dif_cambio_ng','dif_cambio_total','numero_liq')"

    Else
        '        MsgBox ("No es nueva, estoy aprobando una existente")
        q = "UPDATE liquidaciones_caja" _
            & " SET id_moneda = 'id_moneda'," _
            & " fecha = 'fecha'," _
            & " id_cuenta_contable = 'id_cuenta_contable'," _
            & " alicuota = 'alicuota'," _
            & " cuenta_contable_desc = 'cuenta_contable_desc'," _
            & " estado = 'estado'," _
            & " static_total_facturas = 'static_total_facturas'," _
            & " static_total_factura_ng = 'static_total_factura_ng'," _
            & " static_total_a_retener = 'static_total_a_retener'," _
            & " static_total_origen = 'static_total_origen'," _
            & " dif_cambio = 'dif_cambio'," _
            & " otros_descuentos = 'otros_descuentos'," _
            & " tipo_cambio = 'tipo_cambio'," _
            & " id_moneda_pago = 'id_moneda_pago'," _
            & " dif_cambio_ng = 'dif_cambio_ng'," _
            & " dif_cambio_total = 'dif_cambio_total'," _
            & " numero_liq = 'numero_liq'" _
            & " WHERE id = 'id'"
        q = Replace(q, "'id'", GetEntityId(op))
    End If

    q = Replace(q, "'id_moneda'", GetEntityId(op.moneda))
    q = Replace(q, "'alicuota'", Escape(op.alicuota))
    q = Replace(q, "'fecha'", Escape(op.FEcha))
    q = Replace(q, "'id_cuenta_contable'", GetEntityId(op.CuentaContable))
    q = Replace(q, "'cuenta_contable_desc'", Escape(op.CuentaContableDescripcion))
    q = Replace(q, "'estado'", Escape(op.estado))
    q = Replace(q, "'static_total_facturas'", Escape(op.StaticTotalFacturas))
    q = Replace(q, "'static_total_factura_ng'", Escape(op.StaticTotalFacturasNG))
    q = Replace(q, "'static_total_a_retener'", Escape(op.StaticTotalRetenido))
    q = Replace(q, "'static_total_origen'", Escape(op.StaticTotalOrigenes))
    q = Replace(q, "'dif_cambio'", Escape(op.DiferenciaCambio))
    q = Replace(q, "'otros_descuentos'", Escape(op.OtrosDescuentos))
    q = Replace(q, "'id_moneda_pago'", Escape(op.IdMonedaPago))
    q = Replace(q, "'tipo_cambio'", Escape(op.TipoCambio))
    q = Replace(q, "'dif_cambio_ng'", Escape(op.DiferenciaCambioEnNG))
    q = Replace(q, "'dif_cambio_total'", Escape(op.DiferenciaCambioEnTOTAL))
    q = Replace(q, "'numero_liq'", Escape(op.NumeroLiq))


    If Not conectar.execute(q) Then GoTo E

    If Nueva Then op.Id = conectar.UltimoId2()
    If op.Id = 0 Then GoTo E

    '------------------------------------------------------
    '------------------------------------------------------

    Dim fcp As clsFacturaProveedor

    For Each fcp In op.FacturasProveedor
        q = "UPDATE AdminComprasFacturasProveedores SET tipo_cambio_pago= " & fcp.TipoCambioPago & ", estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id = " & fcp.Id
        If Not conectar.execute(q) Then GoTo E
    Next


    q = "DELETE FROM liquidaciones_caja_facturas WHERE id_liquidacion_caja = " & op.Id
    If Not conectar.execute(q) Then GoTo E


    Dim es As EstadoFacturaProveedor
    Dim nopago As Double
    Dim fac As clsFacturaProveedor

    For Each fac In op.FacturasProveedor
        '        fac.ImporteTotalAbonado = fac.NetoGravado + fac.OtrosAbonado
        fac.ImporteTotalAbonado = fac.TotalAplicadoACuentas + fac.TotalOtros
        q = "INSERT INTO liquidaciones_caja_facturas VALUES (" & op.Id & ", " & fac.Id & "," & fac.ImporteTotalAbonado & "," & fac.TotalAplicadoACuentas & "," & fac.TotalOtros & ")"

        If Not conectar.execute(q) Then GoTo E

        nopago = 0

        '        fac.TotalAbonado = fac.NetoGravadoAbonado + fac.OtrosAbonado

        fac.TotalAbonado = fac.TotalPendiente

        'nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado

        nopago = fac.total - fac.TotalAbonadoGlobal - fac.TotalAbonado

        'nopago = fac.Total - fac.TotalPendiente

        '            MsgBox ("Acá como es nueva hace: fac.Total: " & fac.Total & " - fac.TotalPendiente :" & fac.TotalPendiente & " == " & nopago)

        '        q = "DELETE FROM orden_pago_deuda_compensatorios WHERE id_orden_pago = " & op.Id
        '        If Not conectar.execute(q) Then GoTo E

        'If op.estado = EstadoOrdenPago_Aprobada Then

        'nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado

        'es = EstadoFacturaProveedor.Aprobada
        es = EstadoFacturaProveedor.Saldada
        '
        '        If nopago > 0 Then
        '            es = EstadoFacturaProveedor.pagoParcial
        '        Else
        '            es = EstadoFacturaProveedor.Saldada
        '        End If

        q = "UPDATE AdminComprasFacturasProveedores SET estado = " & es & " WHERE id = " & fac.Id

        If Not conectar.execute(q) Then GoTo E


    Next fac


    '------------------------------------------------------


    q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.Id & ")"
    If Not conectar.execute(q) Then GoTo E
    q = "DELETE FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.Id
    If Not conectar.execute(q) Then GoTo E

    Dim oper As operacion
    For Each oper In op.OperacionesBanco
        'oper.IdPertenencia = op.Id no se usa creo
        oper.FechaCarga = Now
        If DAOOperacion.Save(oper) Then
            oper.Id = conectar.UltimoId2
            If oper.Id = 0 Then GoTo E
            q = "INSERT INTO ordenes_pago_operaciones VALUES (" & op.Id & ", " & oper.Id & ")"
            If Not conectar.execute(q) Then GoTo E
        Else
            GoTo E
        End If
    Next oper

    For Each oper In op.OperacionesCaja
        'oper.IdPertenencia = op.Id no se usa creo
        oper.FechaCarga = Now
        If DAOOperacion.Save(oper) Then
            oper.Id = conectar.UltimoId2
            If oper.Id = 0 Then GoTo E
            q = "INSERT INTO ordenes_pago_operaciones VALUES (" & op.Id & ", " & oper.Id & ")"
            If Not conectar.execute(q) Then GoTo E
        Else
            GoTo E
        End If
    Next oper

    Dim msg As String
    msg = "LIQ Creada"

    If Not Nueva Then msg = "LIQ Actualizada"
    DaoHistorico.Save "orden_pago_historial", msg, op.Id

    Guardar = True

    Exit Function
E:
    Guardar = False
    If Nueva Then op.Id = 0

End Function


Public Function GuardarAprobada(op As clsLiquidacionCaja, Optional cascada As Boolean = False) As Boolean

'TODO: tengo que revisar que las facturas no esten en otra op aprobada antes de continuar

    Dim q As String
    On Error GoTo E
    Dim Nueva As Boolean: Nueva = False

    If op.Id = 0 Then

        If IsSomething(DAOLiquidacionCaja.FindByNumeroLiq(CLng(op.NumeroLiq))) Then
            MsgBox "Ya existe una Liquidación con ese número!", vbCritical, "Error"
            Exit Function
        End If


        Nueva = True
        '        MsgBox ("Es nueva")
        q = "INSERT INTO liquidaciones_caja (id_moneda_pago,tipo_cambio,id_moneda, fecha, id_cuenta_contable,cuenta_contable_desc,estado,alicuota,static_total_facturas, static_total_factura_ng, static_total_a_retener, static_total_origen,dif_cambio, otros_descuentos,dif_cambio_ng,dif_cambio_total,numero_liq)" _
            & " VALUES ('id_moneda_pago','tipo_cambio','id_moneda', 'fecha', 'id_cuenta_contable', 'cuenta_contable_desc','0','alicuota','static_total_facturas', 'static_total_factura_ng', 'static_total_a_retener', 'static_total_origen', 'dif_cambio', 'otros_descuentos','dif_cambio_ng','dif_cambio_total','numero_liq')"

    Else
        '        MsgBox ("No es nueva, estoy aprobando una existente")
        q = "UPDATE liquidaciones_caja" _
            & " SET id_moneda = 'id_moneda'," _
            & " fecha = 'fecha'," _
            & " id_cuenta_contable = 'id_cuenta_contable'," _
            & " alicuota = 'alicuota'," _
            & " cuenta_contable_desc = 'cuenta_contable_desc'," _
            & " estado = 'estado'," _
            & " static_total_facturas = 'static_total_facturas'," _
            & " static_total_factura_ng = 'static_total_factura_ng'," _
            & " static_total_a_retener = 'static_total_a_retener'," _
            & " static_total_origen = 'static_total_origen'," _
            & " dif_cambio = 'dif_cambio'," _
            & " otros_descuentos = 'otros_descuentos'," _
            & " tipo_cambio = 'tipo_cambio'," _
            & " id_moneda_pago = 'id_moneda_pago'," _
            & " dif_cambio_ng = 'dif_cambio_ng'," _
            & " dif_cambio_total = 'dif_cambio_total'," _
            & " numero_liq = 'numero_liq'" _
            & " WHERE id = 'id'"
        q = Replace(q, "'id'", GetEntityId(op))
    End If

    q = Replace(q, "'id_moneda'", GetEntityId(op.moneda))
    q = Replace(q, "'alicuota'", Escape(op.alicuota))
    q = Replace(q, "'fecha'", Escape(op.FEcha))
    q = Replace(q, "'id_cuenta_contable'", GetEntityId(op.CuentaContable))
    q = Replace(q, "'cuenta_contable_desc'", Escape(op.CuentaContableDescripcion))
    q = Replace(q, "'estado'", Escape(op.estado))
    q = Replace(q, "'static_total_facturas'", Escape(op.StaticTotalFacturas))
    q = Replace(q, "'static_total_factura_ng'", Escape(op.StaticTotalFacturasNG))
    q = Replace(q, "'static_total_a_retener'", Escape(op.StaticTotalRetenido))
    q = Replace(q, "'static_total_origen'", Escape(op.StaticTotalOrigenes))
    q = Replace(q, "'dif_cambio'", Escape(op.DiferenciaCambio))
    q = Replace(q, "'otros_descuentos'", Escape(op.OtrosDescuentos))
    q = Replace(q, "'id_moneda_pago'", Escape(op.IdMonedaPago))
    q = Replace(q, "'tipo_cambio'", Escape(op.TipoCambio))
    q = Replace(q, "'dif_cambio_ng'", Escape(op.DiferenciaCambioEnNG))
    q = Replace(q, "'dif_cambio_total'", Escape(op.DiferenciaCambioEnTOTAL))
    q = Replace(q, "'numero_liq'", Escape(op.NumeroLiq))


    If Not conectar.execute(q) Then GoTo E

    If Nueva Then op.Id = conectar.UltimoId2()
    If op.Id = 0 Then GoTo E

    '------------------------------------------------------
    '------------------------------------------------------

    Dim fcp As clsFacturaProveedor

    '    For Each fcp In op.FacturasProveedor
    '        q = "UPDATE AdminComprasFacturasProveedores SET tipo_cambio_pago= " & fcp.TipoCambioPago & ", estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id = " & fcp.Id
    '        If Not conectar.execute(q) Then GoTo E
    '    Next

    '
    '    q = "DELETE FROM liquidaciones_caja_facturas WHERE id_liquidacion_caja = " & op.Id
    '    If Not conectar.execute(q) Then GoTo E


    '    Dim es As EstadoFacturaProveedor
    '    Dim nopago As Double
    '    Dim fac As clsFacturaProveedor
    '
    '    For Each fac In op.FacturasProveedor
    '        fac.ImporteTotalAbonado = fac.NetoGravadoAbonado + fac.OtrosAbonado
    '        q = "INSERT INTO liquidaciones_caja_facturas VALUES (" & op.Id & ", " & fac.Id & "," & fac.ImporteTotalAbonado & "," & fac.NetoGravadoAbonado & "," & fac.OtrosAbonado & ")"
    '
    '        If Not conectar.execute(q) Then GoTo E
    '
    '        nopago = 0
    '
    '
    '        fac.TotalAbonado = fac.NetoGravadoAbonado + fac.OtrosAbonado
    '
    '        'nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado
    '
    '        nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado
    'nopago = fac.Total - fac.TotalPendiente

    '            MsgBox ("Acá como es nueva hace: fac.Total: " & fac.Total & " - fac.TotalPendiente :" & fac.TotalPendiente & " == " & nopago)


    '        q = "DELETE FROM orden_pago_deuda_compensatorios WHERE id_orden_pago = " & op.Id
    '        If Not conectar.execute(q) Then GoTo E
    '
    '        'If op.estado = EstadoOrdenPago_Aprobada Then
    '
    '        'nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado
    '
    '        es = EstadoFacturaProveedor.Aprobada
    '        If nopago > 0 Then
    '            es = EstadoFacturaProveedor.pagoParcial
    '        Else
    '            es = EstadoFacturaProveedor.Saldada
    '        End If
    '
    '        q = "UPDATE AdminComprasFacturasProveedores SET estado = " & es & " WHERE id = " & fac.Id
    '
    '        If Not conectar.execute(q) Then GoTo E
    '
    '
    '    Next fac


    '------------------------------------------------------


    '    q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.Id & ")"
    '    If Not conectar.execute(q) Then GoTo E
    '    q = "DELETE FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.Id
    '    If Not conectar.execute(q) Then GoTo E

    '    Dim oper As operacion
    '    For Each oper In op.OperacionesBanco
    '        'oper.IdPertenencia = op.Id no se usa creo
    '        oper.FechaCarga = Now
    '        If DAOOperacion.Save(oper) Then
    '            oper.Id = conectar.UltimoId2
    '            If oper.Id = 0 Then GoTo E
    '            q = "INSERT INTO ordenes_pago_operaciones VALUES (" & op.Id & ", " & oper.Id & ")"
    '            If Not conectar.execute(q) Then GoTo E
    '        Else
    '            GoTo E
    '        End If
    '    Next oper
    '
    '    For Each oper In op.OperacionesCaja
    '        'oper.IdPertenencia = op.Id no se usa creo
    '        oper.FechaCarga = Now
    '        If DAOOperacion.Save(oper) Then
    '            oper.Id = conectar.UltimoId2
    '            If oper.Id = 0 Then GoTo E
    '            q = "INSERT INTO ordenes_pago_operaciones VALUES (" & op.Id & ", " & oper.Id & ")"
    '            If Not conectar.execute(q) Then GoTo E
    '        Else
    '            GoTo E
    '        End If
    '    Next oper
    '
    '    Dim msg As String
    '    msg = "LIQ Creada"
    '
    '    If Not Nueva Then msg = "LIQ Actualizada"
    '    DaoHistorico.Save "orden_pago_historial", msg, op.Id

    GuardarAprobada = True

    Exit Function
E:
    GuardarAprobada = False
    If Nueva Then op.Id = 0

End Function



Public Function RemoveFactura(opid As Long, facid As Long) As Boolean
    RemoveFactura = False

    Dim op As OrdenPago
    Set op = DAOOrdenPago.FindById(opid)
    If IsSomething(op) Then
        If op.estado = EstadoOrdenPago_pendiente Then    'si esta aprobada no se puede eliminar una factura
            Dim q As String

            q = "UPDATE AdminComprasFacturasProveedores SET estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id = " & opid
            RemoveFactura = conectar.execute(q)

            If RemoveFactura Then
                q = "DELETE FROM liquidaciones_caja_facturas WHERE id_factura_proveedor = " & facid
                RemoveFactura = conectar.execute(q)
            Else
                'vuelvo la factura al estado anterior, para no hacer una transaccion
                q = "UPDATE AdminComprasFacturasProveedores SET estado = " & EstadoFacturaProveedor.Saldada & " WHERE id = " & opid
                RemoveFactura = conectar.execute(q)
            End If
            'DaoHistorico.Save "orden_pago_historial", "Factura Id " & facid & " removida, opid"

        End If
    End If

End Function


Public Function Delete(liqid As Long, useInternalTransaction As Boolean) As Boolean
    On Error GoTo E

    Dim liq As clsLiquidacionCaja
    Set liq = DAOLiquidacionCaja.FindById(liqid)

    If useInternalTransaction Then conectar.BeginTransaction

    Dim q As String

    q = "UPDATE AdminComprasFacturasProveedores SET estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id IN (SELECT id_factura_proveedor FROM liquidaciones_caja_facturas WHERE id_liquidacion_caja = " & liqid & ")"
    Debug.Print (q)
    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM ordenes_pago_operaciones WHERE id_orden_pago = " & liqid & ")"
    Debug.Print (q)
    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM ordenes_pago_operaciones WHERE id_orden_pago = " & liqid
    Debug.Print (q)
    If Not conectar.execute(q) Then GoTo E


    If Not conectar.execute(q) Then GoTo E
    Dim estado_anterior As EstadoLiquidacionCaja
    estado_anterior = liq.estado
    liq.estado = EstadoLiquidacionCaja_Anulada

    If Not DAOLiquidacionCaja.Guardar(liq, False) Then GoTo E

    DaoHistorico.Save "orden_pago_historial", "LIQUIDACION Anulada", liq.NumeroLiq

    If useInternalTransaction Then conectar.CommitTransaction

    Delete = True
    Exit Function
E:
    liq.estado = estado_anterior
    If useInternalTransaction Then conectar.RollBackTransaction
    Delete = False
End Function


Public Function ResumenPagos(ByRef Cheques As Collection, ByRef caja As Collection, ByRef bancos As Collection, ByRef comp As Collection, ByRef retenciones As Collection, ByRef cheques3 As Collection, Optional filtro As String, Optional idProveedor As Long = -1) As Boolean
    On Error GoTo err1
    ResumenPagos = True
    Dim q As String
    Dim rs As Recordset

    '#'CHEQUES'
    q = "SELECT b.Nombre,SUM(monto * acm.cambio) as monto FROM ordenes_pago  op " _
        & " INNER JOIN ordenes_pago_cheques opc ON opc.id_liquidacion_caja=op.id " _
        & " LEFT JOIN Cheques c ON opc.id_cheque=c.id " _
        & " LEFT JOIN AdminConfigBancos b ON c.id_banco=b.id " _
        & " LEFT JOIN AdminConfigMonedas acm ON c.id_moneda=acm.id WHERE c.propio=1 and 1=1 " _

If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    q = q & " GROUP BY b.id "

    Dim d As DTONombreMonto

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        Cheques.Add d
        rs.MoveNext
    Wend

    '#OPERACIONES CAJA
    q = " SELECT ca.nombre,SUM(monto * acm.cambio ) as monto FROM ordenes_pago op " _
        & " INNER JOIN ordenes_pago_operaciones opo ON opo.id_liquidacion_caja=op.id " _
        & " LEFT JOIN operaciones o ON opo.id_operacion=o.id " _
        & " LEFT JOIN cajas ca ON ca.id=o.cuentabanc_o_caja_id " _
        & " LEFT JOIN AdminConfigMonedas acm ON o.moneda_id=acm.id " _
        & " WHERE o.pertenencia='caja' AND entrada_salida=-1 AND 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & " GROUP BY o.cuentabanc_o_caja_id"

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        caja.Add d
        rs.MoveNext
    Wend

    '#OPERACIONES BANCO
    q = "SELECT  ba.nombre,  SUM(monto * acm.cambio ) AS monto " _
        & " FROM ordenes_pago op   INNER JOIN ordenes_pago_operaciones opo     ON opo.id_liquidacion_caja = op.id " _
        & " LEFT JOIN operaciones o     ON opo.id_operacion = o.id " _
        & " LEFT JOIN AdminConfigCuentas cba     ON cba.id = o.cuentabanc_o_caja_id " _
        & " INNER JOIN AdminConfigBancos ba ON cba.idBanco=ba.id    LEFT JOIN AdminConfigMonedas acm     ON o.moneda_id = acm.id " _
        & " WHERE o.pertenencia='banco' AND entrada_salida=-1 AND 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & "GROUP BY o.cuentabanc_o_caja_id"
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto

        If Not IsNull(rs!nombre) Then d.nombre = rs!nombre Else d.nombre = vbNullString

        bancos.Add d
        rs.MoveNext
    Wend


    '#compensatorios
    q = "SELECT fp.numero_factura, (IF (com.tipo=1,(com.importe * acm.cambio),(com.importe * acm.cambio*-1))) AS monto  FROM ordenes_pago op " _
        & " INNER JOIN ordenes_pago_compensatorios com ON com.id_liquidacion_caja=op.id " _
        & " INNER JOIN AdminComprasFacturasProveedores fp ON com.id_comprobante=fp.id " _
        & " INNER JOIN AdminConfigMonedas acm ON fp.id_moneda=acm.id  where 1=1 "

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If


    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        If Not IsNull(rs!numero_factura) Then d.nombre = rs!numero_factura Else d.nombre = vbNullString

        comp.Add d
        rs.MoveNext
    Wend

    q = "SELECT 'IIBB' AS nombre, SUM(static_total_a_retener) AS monto FROM ordenes_pago op WHERE 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    'q = q & " GROUP BY id "

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        retenciones.Add d
        rs.MoveNext
    Wend

    '#'CHEQUES 3ROS'
    q = "SELECT b.Nombre,SUM(monto * acm.cambio) as monto FROM ordenes_pago  op " _
        & " INNER JOIN ordenes_pago_cheques opc ON opc.id_liquidacion_caja=op.id " _
        & " LEFT JOIN Cheques c ON opc.id_cheque=c.id " _
        & " LEFT JOIN AdminConfigBancos b ON c.id_banco=b.id " _
        & " LEFT JOIN AdminConfigMonedas acm ON c.id_moneda=acm.id WHERE c.propio=0 and 1=1 " _

If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    q = q & " GROUP BY b.id "

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        If Not IsNull(rs!nombre) Then d.nombre = rs!nombre Else d.nombre = vbNullString
        cheques3.Add d
        rs.MoveNext
    Wend
    Exit Function
err1:
    ResumenPagos = False

End Function


Public Function PrintLiq(LiquidacionCaja As clsLiquidacionCaja, pic As PictureBox) As Boolean
    Dim TAB1 As Integer
    Dim TAB2 As Integer
    Dim TAB3 As Integer
    Dim maxw As Single
    Dim C As Long
    Dim mtxt As String
    Dim textw As Single
    Dim lmargin As Integer

    pic.Picture = LoadResPicture(101, vbResBitmap)

    Dim A As Single
    lmargin = 720


    TAB1 = 300
    TAB2 = 300
    TAB3 = 300

    Printer.CurrentY = lmargin
    maxw = Printer.Width - lmargin * 2
    A = lmargin + (maxw - 3200) / 2
    Printer.PaintPicture pic.Picture, A, 100, 3200, 600

    Printer.FontBold = True
    Printer.FontSize = 12
    mtxt = "Liquidación de Caja Nº " & LiquidacionCaja.NumeroLiq
    textw = Printer.TextWidth(mtxt)

    Printer.CurrentX = lmargin + (maxw - textw) / 2
    Printer.Print mtxt
    Printer.FontSize = 10
    Printer.CurrentX = lmargin
    Printer.Print "Fecha: ";
    Printer.FontBold = False
    Printer.Print LiquidacionCaja.FEcha

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Moneda: ";
    Printer.FontBold = False
    Printer.Print LiquidacionCaja.moneda.NombreCorto & " " & LiquidacionCaja.moneda.NombreLargo

    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)

    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Comprobantes: "
    Printer.FontBold = False
    Printer.FontSize = 8
    Set LiquidacionCaja.FacturasProveedor = DAOFacturaProveedor.FindAllByLiquidacionCaja(LiquidacionCaja.Id)
    Dim F As clsFacturaProveedor
    C = 0
    For Each F In LiquidacionCaja.FacturasProveedor
        C = C + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2 + TAB3
        Printer.Print F.NumeroFormateado & String$(8, " del ") & F.FEcha & String$(8, " por ") & F.moneda.NombreCorto & " " & F.total & String$(20, " de "); UCase(F.Proveedor.RazonSocial)
    Next F
    If C = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2 + TAB3
        Printer.Print "NO POSEE FACTURAS ASOCIADAS"
    End If
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.FontSize = 10
    Printer.CurrentX = lmargin
    Printer.FontBold = True
    Printer.Print "Valores: "


    Dim tmpCol As New Collection

    Set tmpCol = New Collection

    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Transferencias: "
    Printer.FontBold = False

    Dim op As operacion
    Set tmpCol = New Collection
    C = 0
    For Each op In LiquidacionCaja.OperacionesBanco
        C = C + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print op.FechaOperacion & String$(8, " ") & op.moneda.NombreCorto & " " & op.Monto & " | Cta.Bancaria: " & op.CuentaBancaria.DescripcionFormateada & " | Nro. Cbte: " & op.Comprobante
    Next op
    If C = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE TRANSFERENCIAS"
    End If
    Printer.Print
    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Caja: "
    Printer.FontBold = False


    Set tmpCol = New Collection
    C = 0
    For Each op In LiquidacionCaja.OperacionesCaja
        C = C + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print op.FechaOperacion & String$(8, " ") & op.moneda.NombreCorto & " " & op.Monto
    Next op
    If C = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE OPERACIONES EN EFECTIVO"
    End If

    Printer.FontSize = 11
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Total Facturas: ";
    Printer.FontBold = False
    Printer.Print LiquidacionCaja.moneda.NombreCorto & " " & LiquidacionCaja.StaticTotalFacturas

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Total Abonado: ";
    Printer.FontBold = False
    Printer.Print LiquidacionCaja.moneda.NombreCorto & " " & LiquidacionCaja.StaticTotalOrigenes
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.EndDoc

    DaoHistorico.Save "orden_pago_historial", "OP Impresa", LiquidacionCaja.Id
End Function



