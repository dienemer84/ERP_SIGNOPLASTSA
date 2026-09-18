Attribute VB_Name = "DAOCuentaCorriente"
Option Explicit

Public Function ResumenSaldoProveedor() As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount
    On Error GoTo err1
    Dim col As Collection
    Dim dic As New Collection
    Dim deta As DTODetalleCuentaCorriente
    Dim proveedores As Collection

    Dim rs As Recordset

    Set rs = conectar.RSFactory("select id from proveedores")
    Dim saldo As Double

    While Not rs.EOF And Not rs.BOF

        Set col = DAOCuentaCorriente.FindAllDetallesProveedor(rs!Id)
        saldo = GetSaldo(col)

        If saldo > 0 Then dic.Add saldo, CStr(rs!Id)

        rs.MoveNext
        
    Wend

    Set ResumenSaldoProveedor = dic
    tickend = GetTickCount
    'Debug.Print tickend - tickStart

    Exit Function
err1:

End Function


Public Function GetSaldo(col As Collection) As Double
    
    Dim deta As DTODetalleCuentaCorriente


    Dim saldo As Double
    saldo = 0
    For Each deta In col
        'If IsNumeric(deta.Debe) Then     'decia deta>0  12-8-13
        If deta.Debe > 0 Or deta.Debe < 0 Then
            saldo = saldo + deta.Debe
        Else
            saldo = saldo - deta.Haber
        End If
    Next deta
    
    GetSaldo = saldo
    
End Function


Public Function FindResumenSaldosProveedoresRapido( _
    Optional ByVal FechaHasta As String = vbNullString, _
    Optional ByVal FechaDesde As String = vbNullString) As Collection

    On Error GoTo errHandler

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim resultado As Collection
    Dim dto As DTONombreMonto
    Dim q As String
    
    Dim fechaSQL As String
    Dim fechaDesdeSQL As String
    Dim mensajeError As String

    Set resultado = New Collection
    Set cn = conectar.obternerConexion
    
    'Preparar fecha Hasta.
    If LenB(Trim$(FechaHasta)) > 0 Then
        fechaSQL = conectar.Escape(FechaHasta)
    Else
        fechaSQL = conectar.Escape("9999-12-31")
    End If
    
    'Preparar fecha Desde.
    If LenB(Trim$(FechaDesde)) > 0 Then
        fechaDesdeSQL = conectar.Escape(FechaDesde)
    Else
        fechaDesdeSQL = conectar.Escape("1000-01-01")
    End If

    '=========================================================
    ' TABLA TEMPORAL CON UN REGISTRO POR PROVEEDOR
    '=========================================================

    cn.execute "DROP TEMPORARY TABLE IF EXISTS tmp_resumen_proveedores"
    cn.execute "DROP TEMPORARY TABLE IF EXISTS tmp_resumen_prov_mov"

    q = "CREATE TEMPORARY TABLE tmp_resumen_proveedores ("
    q = q & "id_proveedor BIGINT NOT NULL PRIMARY KEY, "
    q = q & "razon VARCHAR(255), "
    q = q & "saldo DOUBLE NOT NULL DEFAULT 0, "
    q = q & "fecha_cierre DATE NOT NULL DEFAULT '1990-01-01')"

    EjecutarPasoResumen cn, _
    "1 - Crear tabla temporal de proveedores", q

    q = "INSERT INTO tmp_resumen_proveedores "
    q = q & "(id_proveedor, razon, saldo, fecha_cierre) "
    q = q & "SELECT id, razon, 0, '1990-01-01' "
    q = q & "FROM proveedores"

    EjecutarPasoResumen cn, _
    "2 - Cargar proveedores", q

    '=========================================================
    ' SALDOS HISTÓRICOS
    '=========================================================

    q = "UPDATE tmp_resumen_proveedores r "
    q = q & "INNER JOIN ("
    
    q = q & "SELECT h.id_persona AS id_proveedor, "
    
    'El importe respeta Desde, pero fecha_cierre conserva
    'el último movimiento histórico para evitar duplicados.
    q = q & "SUM(CASE "
    q = q & "WHEN hd.fecha >= " & fechaDesdeSQL & " THEN "
    q = q & "IFNULL(hd.debe, 0) - IFNULL(hd.haber, 0) "
    q = q & "ELSE 0 END) AS importe, "
    
    q = q & "MAX(hd.fecha) AS fecha_cierre "
    
    q = q & "FROM cuenta_corriente_historic h "
    q = q & "INNER JOIN cuenta_corriente_historic_detalle hd "
    q = q & "ON hd.id_cuenta_corriente_historic = h.id "
    
    q = q & "WHERE h.tipo_persona = "
    q = q & CStr(TipoPersona.proveedor_) & " "
    
    q = q & "AND hd.tipo_comprobante <> "
    q = q & CStr(TipoComprobanteUsado.SaldoInicial_) & " "
    
    q = q & "AND hd.fecha <= " & fechaSQL & " "
    q = q & "GROUP BY h.id_persona"
    
    q = q & ") h ON h.id_proveedor = r.id_proveedor "
    
    q = q & "SET r.saldo = h.importe, "
    q = q & "r.fecha_cierre = "
    q = q & "IFNULL(h.fecha_cierre, '1990-01-01')"

    EjecutarPasoResumen cn, _
    "3 - Calcular saldos históricos", q

    '=========================================================
    ' TABLA TEMPORAL DE MOVIMIENTOS POSTERIORES AL HISTÓRICO
    '=========================================================

    q = "CREATE TEMPORARY TABLE tmp_resumen_prov_mov ("
    q = q & "id_proveedor BIGINT NOT NULL, "
    q = q & "importe DOUBLE NOT NULL DEFAULT 0, "
    q = q & "INDEX idx_proveedor (id_proveedor))"

    EjecutarPasoResumen cn, _
    "4 - Crear tabla temporal de movimientos", q

    '=========================================================
    ' FACTURAS, NOTAS DE DÉBITO Y NOTAS DE CRÉDITO
    '=========================================================

    '=========================================================
    ' FACTURAS, NOTAS DE DÉBITO Y NOTAS DE CRÉDITO
    '=========================================================

    '---------------------------------------------------------
    ' 5.1 - Crear tabla con las facturas que realmente entran
    '       en el reporte.
    '---------------------------------------------------------

    q = "CREATE TEMPORARY TABLE tmp_resumen_facturas ("
    q = q & "id_factura BIGINT NOT NULL PRIMARY KEY, "
    q = q & "id_proveedor BIGINT NOT NULL, "
    q = q & "tipo_doc_contable INT NOT NULL, "
    q = q & "impuesto_interno DOUBLE NOT NULL DEFAULT 0, "
    q = q & "redondeo_iva DOUBLE NOT NULL DEFAULT 0, "
    q = q & "tipo_cambio DOUBLE NOT NULL DEFAULT 1, "
    q = q & "id_moneda BIGINT, "
    q = q & "INDEX idx_tmp_fact_proveedor (id_proveedor))"

    EjecutarPasoResumen cn, _
        "5.1 - Crear temporal de facturas", q

    '---------------------------------------------------------
    ' 5.2 - Cargar solamente las facturas del período.
    '---------------------------------------------------------

    q = "INSERT INTO tmp_resumen_facturas ("
    q = q & "id_factura, "
    q = q & "id_proveedor, "
    q = q & "tipo_doc_contable, "
    q = q & "impuesto_interno, "
    q = q & "redondeo_iva, "
    q = q & "tipo_cambio, "
    q = q & "id_moneda) "

    q = q & "SELECT "
    q = q & "f.id, "
    q = q & "f.id_proveedor, "
    q = q & "f.tipo_doc_contable, "
    q = q & "IFNULL(f.impuesto_interno, 0), "
    q = q & "IFNULL(f.redondeo_iva, 0), "
    q = q & "IFNULL(NULLIF(f.tipo_cambio, 0), 1), "
    q = q & "f.id_moneda "

    q = q & "FROM AdminComprasFacturasProveedores f "

    q = q & "INNER JOIN tmp_resumen_proveedores r "
    q = q & "ON r.id_proveedor = f.id_proveedor "

    q = q & "WHERE f.estado IN ("
    q = q & CStr(EstadoFacturaProveedor.Aprobada) & ", "
    q = q & CStr(EstadoFacturaProveedor.Saldada) & ", "
    q = q & CStr(EstadoFacturaProveedor.pagoParcial) & ") "

    q = q & "AND f.fecha >= " & fechaDesdeSQL & " "
    q = q & "AND f.fecha > r.fecha_cierre "
    q = q & "AND f.fecha <= " & fechaSQL

    EjecutarPasoResumen cn, _
        "5.2 - Seleccionar facturas del reporte", q

    '---------------------------------------------------------
    ' 5.3 - Crear tabla para IVA y percepciones.
    '---------------------------------------------------------

    q = "CREATE TEMPORARY TABLE "
    q = q & "tmp_resumen_factura_importes ("
    q = q & "id_factura BIGINT NOT NULL, "
    q = q & "neto DOUBLE NOT NULL DEFAULT 0, "
    q = q & "iva DOUBLE NOT NULL DEFAULT 0, "
    q = q & "percepcion DOUBLE NOT NULL DEFAULT 0, "
    q = q & "INDEX idx_tmp_importe_factura (id_factura))"

    EjecutarPasoResumen cn, _
        "5.3 - Crear temporal de importes", q

    '---------------------------------------------------------
    ' 5.4 - Calcular neto e IVA únicamente para las facturas
    '       seleccionadas.
    '---------------------------------------------------------

    q = "INSERT INTO tmp_resumen_factura_importes "
    q = q & "(id_factura, neto, iva, percepcion) "

    q = q & "SELECT "
    q = q & "fi.id_factura_proveedor, "
    q = q & "ROUND(SUM(IFNULL(fi.valor, 0)), 2), "
    q = q & "ROUND(SUM("
    q = q & "IFNULL(fi.valor, 0) * "
    q = q & "(IFNULL(ali.alicuota, 0) / 100)"
    q = q & "), 2), "
    q = q & "0 "

    q = q & "FROM tmp_resumen_facturas tf "

    q = q & "INNER JOIN "
    q = q & "AdminComprasFacturasProveedoresIva fi "
    q = q & "ON fi.id_factura_proveedor = tf.id_factura "

    q = q & "LEFT JOIN AdminConfigIvaAlicuotas ali "
    q = q & "ON ali.id = fi.id_iva "

    q = q & "GROUP BY fi.id_factura_proveedor"

    EjecutarPasoResumen cn, _
        "5.4 - Calcular neto e IVA", q

    '---------------------------------------------------------
    ' 5.5 - Calcular percepciones de las facturas seleccionadas.
    '---------------------------------------------------------

    q = "INSERT INTO tmp_resumen_factura_importes "
    q = q & "(id_factura, neto, iva, percepcion) "

    q = q & "SELECT "
    q = q & "fp.id_factura_proveedor, "
    q = q & "0, "
    q = q & "0, "
    q = q & "SUM(IFNULL(fp.valor, 0)) "

    q = q & "FROM tmp_resumen_facturas tf "

    q = q & "INNER JOIN "
    q = q & "AdminComprasFacturasProveedoresPercepciones fp "
    q = q & "ON fp.id_factura_proveedor = tf.id_factura "

    q = q & "GROUP BY fp.id_factura_proveedor"

    EjecutarPasoResumen cn, _
        "5.5 - Calcular percepciones", q

    '---------------------------------------------------------
    ' 5.6 - Obtener el total de cada comprobante.
    '---------------------------------------------------------

    q = "INSERT INTO tmp_resumen_prov_mov "
    q = q & "(id_proveedor, importe) "

    q = q & "SELECT "
    q = q & "tf.id_proveedor, "

    q = q & "(CASE WHEN tf.tipo_doc_contable = "
    q = q & CStr(tipoDocumentoContable.notaCredito)
    q = q & " THEN -1 ELSE 1 END) * "

    q = q & "ROUND("
    q = q & "IFNULL(SUM(imp.neto), 0) + "
    q = q & "IFNULL(tf.impuesto_interno, 0) + "
    q = q & "IFNULL(tf.redondeo_iva, 0) + "
    q = q & "IFNULL(SUM(imp.iva), 0) + "

    q = q & "(IFNULL(SUM(imp.percepcion), 0) / "
    q = q & "CASE "
    q = q & "WHEN IFNULL(mon.patron, 0) = 1 THEN 1 "
    q = q & "ELSE IFNULL(NULLIF(tf.tipo_cambio, 0), 1) "
    q = q & "END), 2) "

    q = q & "FROM tmp_resumen_facturas tf "

    q = q & "LEFT JOIN tmp_resumen_factura_importes imp "
    q = q & "ON imp.id_factura = tf.id_factura "

    q = q & "LEFT JOIN AdminConfigMonedas mon "
    q = q & "ON mon.id = tf.id_moneda "

    q = q & "GROUP BY "
    q = q & "tf.id_factura, "
    q = q & "tf.id_proveedor, "
    q = q & "tf.tipo_doc_contable, "
    q = q & "tf.impuesto_interno, "
    q = q & "tf.redondeo_iva, "
    q = q & "tf.tipo_cambio, "
    q = q & "mon.patron"

    EjecutarPasoResumen cn, _
        "5.6 - Consolidar comprobantes", q

    ' Ya no necesitamos estas tablas.

    cn.execute _
        "DROP TEMPORARY TABLE IF EXISTS " & _
        "tmp_resumen_factura_importes"

    cn.execute _
        "DROP TEMPORARY TABLE IF EXISTS " & _
        "tmp_resumen_facturas"

    '=========================================================
    ' ÓRDENES DE PAGO
    '=========================================================

    q = "INSERT INTO tmp_resumen_prov_mov "
    q = q & "(id_proveedor, importe) "
    q = q & "SELECT f.id_proveedor, "
    q = q & "CASE WHEN op.estado = "
    q = q & CStr(EstadoOrdenPago.EstadoOrdenPago_Anulada)
    q = q & " THEN 0 ELSE -ROUND("
    q = q & "IFNULL(op.static_total_origen, 0) + "
    q = q & "IFNULL(op.static_total_a_retener, 0), 2) END "
    q = q & "FROM ordenes_pago op "
    q = q & "INNER JOIN ordenes_pago_facturas opf "
    q = q & "ON opf.id_orden_pago = op.id "
    q = q & "INNER JOIN AdminComprasFacturasProveedores f "
    q = q & "ON f.id = opf.id_factura_proveedor "
    q = q & "INNER JOIN tmp_resumen_proveedores r "
    q = q & "ON r.id_proveedor = f.id_proveedor "
    q = q & "WHERE op.fecha >= " & fechaDesdeSQL & " "
    q = q & "AND op.fecha > r.fecha_cierre "
    q = q & "AND op.fecha <= " & fechaSQL & " "
    q = q & "GROUP BY f.id_proveedor, op.id, op.estado, "
    q = q & "op.static_total_origen, op.static_total_a_retener"

    EjecutarPasoResumen cn, _
    "6 - Calcular órdenes de pago", q
    
    
    '=========================================================
    ' LIQUIDACIONES DE CAJA APLICADAS A FACTURAS
    '=========================================================

    q = "INSERT INTO tmp_resumen_prov_mov "
    q = q & "(id_proveedor, importe) "
    q = q & "SELECT f.id_proveedor, "
    q = q & "-ROUND(SUM("
    q = q & "IFNULL(lcf.neto_gravado_liquidado, 0) + "
    q = q & "IFNULL(lcf.otros_liquidado, 0)), 2) "
    q = q & "FROM liquidaciones_caja_facturas lcf "
    q = q & "INNER JOIN liquidaciones_caja lc "
    q = q & "ON lc.id = lcf.id_liquidacion_caja "
    q = q & "INNER JOIN AdminComprasFacturasProveedores f "
    q = q & "ON f.id = lcf.id_factura_proveedor "
    q = q & "INNER JOIN tmp_resumen_proveedores r "
    q = q & "ON r.id_proveedor = f.id_proveedor "
    q = q & "WHERE lc.estado = 1 "
    q = q & "AND lc.fecha >= " & fechaDesdeSQL & " "
    q = q & "AND lc.fecha > r.fecha_cierre "
    q = q & "AND lc.fecha <= " & fechaSQL & " "
    q = q & "GROUP BY f.id_proveedor"

    EjecutarPasoResumen cn, _
    "6.1 - Calcular liquidaciones de caja", q

    '=========================================================
    ' PAGOS A CUENTA
    '=========================================================

    q = "INSERT INTO tmp_resumen_prov_mov "
    q = q & "(id_proveedor, importe) "
    q = q & "SELECT p.id_proveedor, "
    q = q & "-ROUND(IFNULL(p.static_total_origen, 0), 2) "
    q = q & "FROM pagos_a_cuenta p "
    q = q & "INNER JOIN tmp_resumen_proveedores r "
    q = q & "ON r.id_proveedor = p.id_proveedor "
    q = q & "WHERE p.estado IN ("
    q = q & CStr(EstadoPagoACuenta.Disponible) & ", "
    q = q & CStr(EstadoPagoACuenta.Procesada) & ") "
    q = q & "AND p.fecha >= " & fechaDesdeSQL & " "
    q = q & "AND p.fecha > r.fecha_cierre "
    q = q & "AND p.fecha <= " & fechaSQL

    EjecutarPasoResumen cn, _
    "7 - Calcular pagos a cuenta", q

    '=========================================================
    ' SUMAR LOS MOVIMIENTOS AL SALDO HISTÓRICO
    '=========================================================

    q = "UPDATE tmp_resumen_proveedores r "
    q = q & "LEFT JOIN ("
    q = q & "SELECT id_proveedor, SUM(importe) AS total "
    q = q & "FROM tmp_resumen_prov_mov "
    q = q & "GROUP BY id_proveedor"
    q = q & ") mov ON mov.id_proveedor = r.id_proveedor "
    q = q & "SET r.saldo = r.saldo + IFNULL(mov.total, 0)"

    EjecutarPasoResumen cn, _
    "8 - Consolidar saldos", q

    '=========================================================
    ' DEVOLVER ÚNICAMENTE PROVEEDORES CON SALDO
    '=========================================================

    q = "SELECT razon, saldo "
    q = q & "FROM tmp_resumen_proveedores "
    q = q & "WHERE saldo >= 0.01 OR saldo < -0.01 "
    q = q & "ORDER BY razon"

    Debug.Print Format$(Now, "hh:nn:ss") & _
                " - INICIO: 9 - Leer resultados finales"
    
    DoEvents
    
    Set rs = cn.execute(q)
    
    Debug.Print Format$(Now, "hh:nn:ss") & _
                " - FIN: 9 - Leer resultados finales"
    
    DoEvents

    While Not rs.EOF

        Set dto = New DTONombreMonto
        dto.nombre = rs!razon
        dto.Monto = rs!saldo

        resultado.Add dto
        rs.MoveNext

    Wend

finalizar:

    On Error Resume Next

    If Not rs Is Nothing Then rs.Close

    If Not cn Is Nothing Then
    
        cn.execute _
            "DROP TEMPORARY TABLE IF EXISTS " & _
            "tmp_resumen_factura_importes"
    
        cn.execute _
            "DROP TEMPORARY TABLE IF EXISTS " & _
            "tmp_resumen_facturas"
    
        cn.execute _
            "DROP TEMPORARY TABLE IF EXISTS " & _
            "tmp_resumen_prov_mov"
    
        cn.execute _
            "DROP TEMPORARY TABLE IF EXISTS " & _
            "tmp_resumen_proveedores"
    
    End If

    Set rs = Nothing
    Set cn = Nothing

    If LenB(mensajeError) > 0 Then
        MsgBox mensajeError, vbCritical, "Resumen de saldos"
    End If

    Set FindResumenSaldosProveedoresRapido = resultado
    Exit Function

errHandler:

    mensajeError = "No se pudo generar el resumen de saldos." & _
                   vbCrLf & vbCrLf & Err.Description

    Set resultado = New Collection
    Resume finalizar

End Function


Private Sub EjecutarPasoResumen( _
    ByVal cn As ADODB.Connection, _
    ByVal nombrePaso As String, _
    ByVal consulta As String)

    Dim Inicio As Double

    Inicio = GetTickCount

    Debug.Print Format$(Now, "hh:nn:ss") & _
                " - INICIO: " & nombrePaso

    DoEvents

    cn.execute consulta

    Debug.Print Format$(Now, "hh:nn:ss") & _
                " - FIN: " & nombrePaso & _
                " - Tiempo: " & _
                Format$((GetTickCount - Inicio) / 1000, "0.00") & _
                " segundos"

    DoEvents

End Sub


Public Function CerrarPeriodoCtaCteProveedor(id_proveedor As Long, FechaHasta As String) As Boolean
    On Error GoTo Error
    Dim deta As DTODetalleCuentaCorriente
    Dim Periodo As Collection
    Dim condicion As String

    'chequear que fechaHasta no sea parte de alguna liquidacion!, es decir fecha hasta tiene que ser mayor a la mayor
    'fechaHasta almacenada

    If Not DAOCuentaCorrienteHistoric.IsValidFechaHasta(id_proveedor, proveedor_, Format(FechaHasta, "yyyy-mm-dd")) Then
        MsgBox "La fecha indicada es invalida para cerrar un periodo!", vbCritical
        Exit Function
    End If

    condicion = conectar.Escape(Format(FechaHasta, "yyyy-mm-dd"))
    Set Periodo = FindAllDetallesProveedor(id_proveedor, True, condicion)

    Dim strsql As String

    Dim cta As New CuentaCorrienteHistoric
    cta.id_persona = id_proveedor
    cta.Periodo = "HASTA " & Format(FechaHasta, "YYYY-MM-DD")
    cta.TipoPersona = proveedor_
    cta.FechaHasta = Format(FechaHasta, "YYYY-MM-DD")
    For Each deta In Periodo
        cta.Detalles.Add deta
    Next

    CerrarPeriodoCtaCteProveedor = DAOCuentaCorrienteHistoric.Save(cta)

    Exit Function
Error:
    MsgBox Err.Description, vbCritical
End Function


Public Function getMaxDesdeProveedor(id_proveedor As Long) As Date
    On Error GoTo err1
    Dim rs As Recordset
    Set rs = conectar.RSFactory("SELECT * FROM saldo_inicial_proveedor WHERE id_proveedor = " & id_proveedor)

    If Not rs.EOF And Not rs.BOF Then
        Dim maxDesde As Date
        maxDesde = CDate(rs!FEcha)
        getMaxDesdeProveedor = maxDesde
    Else
        maxDesde = CDate("2001-01-01")
    End If
    getMaxDesdeProveedor = maxDesde

    Exit Function
err1:
    getMaxDesdeProveedor = DateAdd("d", 1, Now)
    
End Function


Public Function FindAllDetallesProveedor(id_proveedor As Long, Optional sortCollection As Boolean = True, Optional condicion As String, Optional anteriores As Boolean = False, Optional soloOp As Boolean = False) As Collection
    Dim cond1 As String
    Dim detalle As DTODetalleCuentaCorriente
    Dim Detalles As New Collection

    Dim max_desde As String

    Dim max_fecha As Date
    max_fecha = "1990-01-01"
    If anteriores Then

        Dim olddetas As New Collection
        If (LenB(condicion) > 0) Then
            Set olddetas = DAOCuentaCorrienteHistoric.GetAllDetallesFromProveedor(id_proveedor, condicion)
        Else
            Set olddetas = DAOCuentaCorrienteHistoric.GetAllDetallesFromProveedor(id_proveedor)
        End If
        For Each detalle In olddetas
            Detalles.Add detalle
            If detalle.FEcha > max_fecha Then max_fecha = detalle.FEcha

        Next
    End If

    ' max_desde = conectar.Escape(DAOCuentaCorriente.getMaxDesdeProveedor(id_proveedor))

    max_desde = conectar.Escape(Format(max_fecha, "yyyy-mm-dd"))


    If Not anteriores Then
        Dim rs As Recordset
        Set rs = conectar.RSFactory("SELECT saldo_inicial,fecha FROM saldo_inicial_proveedor WHERE id_proveedor = " & id_proveedor)
        Set detalle = New DTODetalleCuentaCorriente

        detalle.Comprobante = "Saldo Inicial"
        detalle.tipoComprobante = SaldoInicial_
        detalle.IdComprobante = 0

        If Not rs.EOF Then
            Dim sald As Double
            sald = rs!saldo_inicial
            If sald < 0 Then
                detalle.Haber = rs!saldo_inicial
            Else
                detalle.Debe = rs!saldo_inicial
            End If
            If Not IsNull(rs!FEcha) Then detalle.FEcha = rs!FEcha
        Else
            detalle.saldo = 0
            detalle.FEcha = "2001-01-01"
        End If
        Detalles.Add detalle
    End If

    Dim ordenes As New Collection
    Dim Orden As OrdenPago

    If LenB(condicion) > 0 Then
        cond1 = "and ordenes_pago.fecha<=" & condicion
    End If


    Set ordenes = DAOOrdenPago.FindAllByProveedor( _
    id_proveedor, _
    cond1 & " AND ordenes_pago.fecha > " & max_desde, _
    soloOp)
    
    
    '------------------------------------------------------
    ' IMPORTE REALMENTE APLICADO POR CADA ORDEN DE PAGO
    '------------------------------------------------------
    
    qImportesOP = "SELECT opf.id_orden_pago, "
    
    qImportesOP = qImportesOP & _
        "SUM(CASE "
    
    qImportesOP = qImportesOP & _
        "WHEN f.tipo_doc_contable = " & _
        CStr(tipoDocumentoContable.notaCredito) & " "
    
    qImportesOP = qImportesOP & _
        "THEN -(IFNULL(opf.neto_gravado_abonado, 0) + " & _
        "IFNULL(opf.otros_abonado, 0)) "
    
    qImportesOP = qImportesOP & _
        "ELSE (IFNULL(opf.neto_gravado_abonado, 0) + " & _
        "IFNULL(opf.otros_abonado, 0)) "
    
    qImportesOP = qImportesOP & _
        "END) AS total_aplicado "
    
    qImportesOP = qImportesOP & _
        "FROM ordenes_pago_facturas opf "
    
    qImportesOP = qImportesOP & _
        "INNER JOIN ordenes_pago op " & _
        "ON op.id = opf.id_orden_pago "
    
    qImportesOP = qImportesOP & _
        "INNER JOIN AdminComprasFacturasProveedores f " & _
        "ON f.id = opf.id_factura_proveedor "
    
    qImportesOP = qImportesOP & _
        "WHERE f.id_proveedor = " & CStr(id_proveedor) & " "
    
    qImportesOP = qImportesOP & _
        "AND op.estado = " & _
        CStr(EstadoOrdenPago.EstadoOrdenPago_Aprobada) & " "
    
    qImportesOP = qImportesOP & _
        "AND op.fecha > " & max_desde & " "
    
    If LenB(condicion) > 0 Then
    
        qImportesOP = qImportesOP & _
            "AND op.fecha <= " & condicion & " "
    
    End If
    
    qImportesOP = qImportesOP & _
        "GROUP BY opf.id_orden_pago"
        
        
    Debug.Print "ID PROVEEDOR: " & CStr(id_proveedor)
    Debug.Print "MAX_DESDE: " & max_desde
    Debug.Print "CONDICION: " & condicion
    Debug.Print "SQL IMPORTES OP:"
    Debug.Print qImportesOP
    
    Set rsImportesOP = conectar.RSFactory(qImportesOP)
    
    Do While Not rsImportesOP.EOF
    
        importesAplicadosOP.Add _
            CStr(rsImportesOP!id_orden_pago), _
            CDbl(rsImportesOP!total_aplicado)
    
        rsImportesOP.MoveNext
    
    Loop
    
    Set rsImportesOP = Nothing
    
    
    For Each Orden In ordenes
        'ver si solo mostrar las aprobadas (revisado) muestra las pendientes indicandolo en el estado

        ' If Orden.estado <> EstadoOrdenPago_Anulada Then
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = "OP-" & Orden.Id

        '#178
        If (Orden.estado = EstadoOrdenPago_pendiente) Then
            detalle.Comprobante = detalle.Comprobante & " (Pendiente)"
        End If

        If (Orden.estado = EstadoOrdenPago_Anulada) Then
            detalle.Comprobante = detalle.Comprobante & " (Anulada)"
        End If

        detalle.tipoComprobante = OrdenPago_
        detalle.IdComprobante = Orden.Id

        If (Orden.estado = EstadoOrdenPago_Anulada) Then

            detalle.Haber = 0
        Else
            detalle.Haber = funciones.RedondearDecimales( _
            Orden.StaticTotalOrigenes + _
            Orden.StaticTotalRetenido)
        End If
        
        detalle.FEcha = Orden.FEcha

        Detalles.Add detalle
        ' End If
    Next Orden

    Dim facturas As Collection
    Dim fac As clsFacturaProveedor

    Dim cond2 As String

    Dim qq As String
    cond2 = "AdminComprasFacturasProveedores.id_proveedor = " & id_proveedor & " AND AdminComprasFacturasProveedores.estado IN (" & EstadoFacturaProveedor.Aprobada & ", " & EstadoFacturaProveedor.Saldada & ", " & EstadoFacturaProveedor.pagoParcial & ") and  AdminComprasFacturasProveedores.fecha > " & max_desde
    If LenB(condicion) > 0 Then
        cond2 = cond2 & " and AdminComprasFacturasProveedores.fecha<=" & condicion
    End If

    Set facturas = DAOFacturaProveedor.FindAll(cond2)
    For Each fac In facturas
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = fac.NumeroFormateado
        '#234
        If fac.estado = pagoParcial Then
            detalle.Comprobante = fac.NumeroFormateado & " (P.Parcial)"
        Else
            detalle.Comprobante = fac.NumeroFormateado
        End If

        detalle.tipoComprobante = TipoComprobanteUsado.FacturaProveedor_
        detalle.IdComprobante = fac.Id

        If InStr(fac.OrdenesPagoId, ",") > 0 Then
            detalle.Comprobante = detalle.Comprobante & " (Ops." & fac.OrdenesPagoId & ")"
        Else
            If fac.OrdenPagoID > 0 Then
                If BuscarEnColeccion(ordenes, CStr(fac.OrdenPagoID)) Then
                    detalle.Comprobante = detalle.Comprobante & " (Op." & fac.OrdenPagoID & " " & ordenes.item(CStr(fac.OrdenPagoID)).FEcha & ")"
                End If
            End If
        End If

        If fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then
            detalle.Haber = fac.total
        Else
            detalle.Debe = fac.total
        End If

        detalle.FEcha = fac.FEcha

        detalle.AtributoExtra = False
        For Each Orden In ordenes
            detalle.AtributoExtra = funciones.BuscarEnColeccion(Orden.FacturasProveedor, CStr(fac.Id))
            If detalle.AtributoExtra = True Then Exit For
        Next Orden

        Detalles.Add detalle
        
    Next fac
    
    '=========================================================
    ' PAGOS A CUENTA DE PROVEEDORES
    '=========================================================
    
    Dim qPago As String
    Dim rsPago As Recordset
    
    qPago = "SELECT " _
          & "p.id, " _
          & "p.fecha, " _
          & "p.estado, " _
          & "IFNULL(p.static_total_origen, 0) AS importe " _
          & "FROM pagos_a_cuenta p " _
          & "WHERE p.id_proveedor = " & id_proveedor & " " _
          & "AND p.fecha > " & max_desde & " " _
          & "AND p.estado IN (0, 1) "
    
    If LenB(condicion) > 0 Then
    
        qPago = qPago _
              & "AND p.fecha <= " & condicion & " "
    
    End If
    
    qPago = qPago & "ORDER BY p.fecha, p.id"
    
    Set rsPago = conectar.RSFactory(qPago)
    
    While Not rsPago.EOF
    
        Set detalle = New DTODetalleCuentaCorriente
    
        detalle.Comprobante = _
            "PAGO A CUENTA-" & CStr(rsPago!Id)
    
        If CLng(rsPago!estado) = EstadoPagoACuenta.Disponible Then
    
            detalle.Comprobante = _
                detalle.Comprobante & " (Disponible)"
    
        Else
    
            detalle.Comprobante = _
                detalle.Comprobante & " (Procesado)"
    
        End If
    
        detalle.IdComprobante = CLng(rsPago!Id)
    
        detalle.FEcha = CDate(rsPago!FEcha)
    
        detalle.Debe = 0
    
        detalle.Haber = funciones.RedondearDecimales( _
            CDbl(rsPago!Importe))
    
        Detalles.Add detalle
    
        rsPago.MoveNext
    
    Wend
    
    Set rsPago = Nothing
    

    If sortCollection And Detalles.count > 0 Then
        Dim q As String

        q = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp_cta_cte_sort (fecha DATE, comprobante VARCHAR(50), debe DOUBLE, haber DOUBLE, extra TINYINT, id_comprobante BIGINT, tipo_comprobante INT) TYPE=HEAP"
        conectar.execute q
        conectar.execute "TRUNCATE tmp_cta_cte_sort"


        For Each detalle In Detalles
            q = "INSERT INTO tmp_cta_cte_sort VALUES ('fecha', 'comprobante', 'debe', 'haber', 'extra','id_comprobante', 'tipo_comprobante')"
            q = Replace$(q, "'fecha'", Escape(detalle.FEcha))
            q = Replace$(q, "'comprobante'", Escape(detalle.Comprobante))
            q = Replace$(q, "'debe'", Escape(detalle.Debe))
            q = Replace$(q, "'haber'", Escape(detalle.Haber))
            q = Replace$(q, "'extra'", Escape(detalle.AtributoExtra))
            q = Replace$(q, "'id_comprobante'", Escape(detalle.IdComprobante))
            q = Replace$(q, "'tipo_comprobante'", Escape(detalle.tipoComprobante))

            conectar.execute q
        Next detalle

        Set Detalles = New Collection
        Dim Id As Long
        Id = 0
        Set rs = conectar.RSFactory("SELECT * FROM tmp_cta_cte_sort ORDER BY fecha ASC")
        While Not rs.EOF
            Id = Id + 1
            Set detalle = New DTODetalleCuentaCorriente
            detalle.tmpId = Id
            detalle.Comprobante = rs!Comprobante
            If Not IsNull(rs!FEcha) Then detalle.FEcha = rs!FEcha
            detalle.Debe = rs!Debe
            detalle.Haber = rs!Haber
            detalle.AtributoExtra = rs!extra
            detalle.tipoComprobante = rs!tipo_comprobante
            detalle.IdComprobante = rs!id_comprobante
            Detalles.Add detalle
            rs.MoveNext
        Wend
    End If

    Set FindAllDetallesProveedor = Detalles
End Function


Public Function FindAllDetallesProveedor2(id_proveedor As Long, Optional sortCollection As Boolean = True, Optional condicion As String, Optional anteriores As Boolean = False, Optional soloOp As Boolean = False) As Collection

    Dim cond1 As String
    Dim detalle As DTODetalleCuentaCorriente
    Dim Detalles As New Collection
    Dim max_desde As String
    Dim max_fecha As Date

    max_fecha = "1990-01-01"

    If anteriores Then

        Dim olddetas As New Collection
        If (LenB(condicion) > 0) Then
            Set olddetas = DAOCuentaCorrienteHistoric.GetAllDetallesFromProveedor(id_proveedor, condicion)
        Else
            Set olddetas = DAOCuentaCorrienteHistoric.GetAllDetallesFromProveedor(id_proveedor)
        End If
        For Each detalle In olddetas
            Detalles.Add detalle
            If detalle.FEcha > max_fecha Then max_fecha = detalle.FEcha

        Next
    End If

    max_desde = conectar.Escape(Format(max_fecha, "yyyy-mm-dd"))

    If Not anteriores Then
        Dim rs As Recordset
        Set rs = conectar.RSFactory("SELECT saldo_inicial, fecha FROM saldo_inicial_proveedor WHERE id_proveedor = " & id_proveedor)
        Set detalle = New DTODetalleCuentaCorriente

        detalle.Comprobante = "Saldo Inicial"
        detalle.tipoComprobante = SaldoInicial_
        detalle.IdComprobante = 0

        If Not rs.EOF Then
            Dim sald As Double
            sald = rs!saldo_inicial
            If sald < 0 Then
                detalle.Haber = rs!saldo_inicial
            Else
                detalle.Debe = rs!saldo_inicial
            End If
            If Not IsNull(rs!FEcha) Then detalle.FEcha = rs!FEcha
        Else
            detalle.saldo = 0
            detalle.FEcha = "2001-01-01"
        End If
        Detalles.Add detalle
    End If


    Dim ordenes As New Collection
    Dim Orden As OrdenPago
    Dim importesAplicadosOP As New Dictionary
    Dim rsImportesOP As Recordset
    Dim qImportesOP As String

    If LenB(condicion) > 0 Then
        cond1 = "and ordenes_pago.fecha<=" & condicion
    End If


    Set ordenes = DAOOrdenPago.FindAllByProveedor( _
    id_proveedor, _
    cond1 & _
    " AND ordenes_pago.fecha > " & max_desde & _
    " AND ordenes_pago.estado = " & _
    CStr(EstadoOrdenPago.EstadoOrdenPago_Aprobada), _
    soloOp)
    
    
    For Each Orden In ordenes
        'ver si solo mostrar las aprobadas (revisado) muestra las pendientes indicandolo en el estado

        ' If Orden.estado <> EstadoOrdenPago_Anulada Then
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = "OP-" & Orden.Id

        '#178
        If (Orden.estado = EstadoOrdenPago_pendiente) Then
            detalle.Comprobante = detalle.Comprobante & " (Pendiente)"
        End If

        If (Orden.estado = EstadoOrdenPago_Anulada) Then
            detalle.Comprobante = detalle.Comprobante & " (Anulada)"
        End If

        detalle.tipoComprobante = OrdenPago_
        detalle.IdComprobante = Orden.Id

        If (Orden.estado = EstadoOrdenPago_Anulada) Then

            detalle.Haber = 0
            
        Else
        
            detalle.Haber = 0
        
            Debug.Print "OP=" & CStr(Orden.Id) & _
                        " | EXISTE=" & _
                        CStr(importesAplicadosOP.Exists(CStr(Orden.Id))) & _
                        " | CANTIDAD DICCIONARIO=" & _
                        CStr(importesAplicadosOP.count)
        
            If importesAplicadosOP.Exists(CStr(Orden.Id)) Then
        
                detalle.Haber = funciones.RedondearDecimales( _
                    CDbl(importesAplicadosOP.item(CStr(Orden.Id))))
        
            End If
        
        End If
        
        detalle.FEcha = Orden.FEcha

        Detalles.Add detalle
        ' End If
    Next Orden


    Dim facturas As Collection
    Dim fac As clsFacturaProveedor

    Dim cond2 As String

    Dim qq As String
    cond2 = "AdminComprasFacturasProveedores.id_proveedor = " & id_proveedor & " AND AdminComprasFacturasProveedores.estado IN (" & EstadoFacturaProveedor.Aprobada & ", " & EstadoFacturaProveedor.Saldada & ", " & EstadoFacturaProveedor.pagoParcial & ") and  AdminComprasFacturasProveedores.fecha > " & max_desde
    If LenB(condicion) > 0 Then
        cond2 = cond2 & " and AdminComprasFacturasProveedores.fecha<=" & condicion
    End If


    Set facturas = DAOFacturaProveedor.FindAll(cond2)
    For Each fac In facturas
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = fac.NumeroFormateado
        '#234
        If fac.estado = pagoParcial Then
            detalle.Comprobante = fac.NumeroFormateado & " (P.Parcial)"
        Else
            detalle.Comprobante = fac.NumeroFormateado
        End If

        detalle.tipoComprobante = TipoComprobanteUsado.FacturaProveedor_
        detalle.IdComprobante = fac.Id

        If InStr(fac.OrdenesPagoId, ",") > 0 Then
            detalle.Comprobante = detalle.Comprobante & " (Ops." & fac.OrdenesPagoId & ")"
        Else

            If fac.OrdenPagoID > 0 Then
                If BuscarEnColeccion(ordenes, CStr(fac.OrdenPagoID)) Then
                    detalle.Comprobante = detalle.Comprobante & " (Op." & fac.OrdenPagoID & " " & ordenes.item(CStr(fac.OrdenPagoID)).FEcha & ")"

                End If

            End If
        End If

        If fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then
            detalle.Haber = fac.total
        Else
            detalle.Debe = fac.total
        End If

        detalle.FEcha = fac.FEcha

        detalle.AtributoExtra = False
        For Each Orden In ordenes
            detalle.AtributoExtra = funciones.BuscarEnColeccion(Orden.FacturasProveedor, CStr(fac.Id))
            If detalle.AtributoExtra = True Then Exit For
        Next Orden

        Detalles.Add detalle
    Next fac

    '=========================================================
    ' LIQUIDACIONES DE CAJA APROBADAS
    '=========================================================

    Dim qLiq As String
    Dim rsLiq As ADODB.Recordset

    qLiq = "SELECT "
    qLiq = qLiq & "lc.id, "
    qLiq = qLiq & "lc.numero_liq, "
    qLiq = qLiq & "lc.fecha, "
    qLiq = qLiq & "ROUND(SUM("
    qLiq = qLiq & "IFNULL(lcf.neto_gravado_liquidado, 0) + "
    qLiq = qLiq & "IFNULL(lcf.otros_liquidado, 0)"
    qLiq = qLiq & "), 2) AS importe "
    qLiq = qLiq & "FROM liquidaciones_caja_facturas lcf "
    qLiq = qLiq & "INNER JOIN liquidaciones_caja lc "
    qLiq = qLiq & "ON lc.id = lcf.id_liquidacion_caja "
    qLiq = qLiq & "INNER JOIN AdminComprasFacturasProveedores f "
    qLiq = qLiq & "ON f.id = lcf.id_factura_proveedor "
    qLiq = qLiq & "WHERE f.id_proveedor = "
    qLiq = qLiq & CStr(id_proveedor) & " "
    qLiq = qLiq & "AND lc.estado = "
    qLiq = qLiq & _
        CStr(EstadoLiquidacionCaja.EstadoLiquidacionCaja_Aprobada) & " "

    'No usamos max_desde porque los cierres históricos actuales
    'todavía no guardan liquidaciones de caja.
    If LenB(condicion) > 0 Then
        qLiq = qLiq & "AND lc.fecha <= " & condicion & " "
    End If

    qLiq = qLiq & _
        "GROUP BY lc.id, lc.numero_liq, lc.fecha "

    qLiq = qLiq & _
        "ORDER BY lc.fecha, lc.id"

    Set rsLiq = conectar.RSFactory(qLiq)

    While Not rsLiq.EOF

        Set detalle = New DTODetalleCuentaCorriente

        detalle.Comprobante = _
            "LIQ.CAJA-" & CStr(rsLiq!numero_liq)

        detalle.IdComprobante = CLng(rsLiq!Id)

        detalle.tipoComprobante = _
            TipoComprobanteUsado.LiquidacionCajaProveedor_

        detalle.FEcha = CDate(rsLiq!FEcha)

        detalle.Debe = 0

        detalle.Haber = funciones.RedondearDecimales( _
            CDbl(rsLiq!Importe))

        Detalles.Add detalle

        rsLiq.MoveNext

    Wend

    Set rsLiq = Nothing


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim pagosacta As New Collection
    Dim PagoACta As clsPagoACta
    
   If LenB(condicion) > 0 Then
        cond1 = "and pagos_a_cuenta.fecha<=" & condicion
    End If
    
    'Set pagosacta = DAOOrdenPago.FindAllByProveedor(id_proveedor, cond1 & "  and pagos_a_cuenta.fecha> " & max_desde, soloOp)
    Set pagosacta = DAOPagoACta.FindAllByProveedor(id_proveedor, cond1 & "  and pagos_a_cuenta.fecha> " & max_desde, soloOp)
    For Each PagoACta In pagosacta
        'ver si solo mostrar las aprobadas (revisado) muestra las pendientes indicandolo en el estado

        ' If Orden.estado <> EstadoOrdenPago_Anulada Then
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = "PAGO A CUENTA-" & PagoACta.Id

        '#178
        If (PagoACta.estado = EstadoOrdenPago_pendiente) Then
            detalle.Comprobante = detalle.Comprobante & " (Disponible)"
        End If

        If (PagoACta.estado = EstadoOrdenPago_Anulada) Then
            detalle.Comprobante = detalle.Comprobante & " (Anulada)"
        End If

        detalle.tipoComprobante = OrdenPago_
        detalle.IdComprobante = PagoACta.Id

        If (PagoACta.estado = EstadoOrdenPago_Anulada) Then

            detalle.Haber = 0
        Else
            detalle.Haber = funciones.RedondearDecimales(PagoACta.TotalOrdenPago)          '.StaticTotalFacturas + Orden.TotalCompensatorios)
        End If
        
        detalle.FEcha = PagoACta.FEcha

        Detalles.Add detalle
        ' End If
    Next PagoACta


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   If sortCollection And Detalles.count > 0 Then
        Dim q As String

        ' Agrego esto 7 y 8
        Dim saldo As Double
        saldo = 0

        q = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp_cta_cte_sort (fecha DATE, comprobante VARCHAR(50), debe DOUBLE, haber DOUBLE, extra TINYINT, id_comprobante BIGINT, tipo_comprobante INT) ENGINE=MEMORY;"
        conectar.execute q
        conectar.execute "TRUNCATE tmp_cta_cte_sort"



        For Each detalle In Detalles
            q = "INSERT INTO tmp_cta_cte_sort VALUES ('fecha', 'comprobante', 'debe', 'haber', 'extra','id_comprobante', 'tipo_comprobante')"
            ' Agrego esto 6
            saldo = saldo + detalle.Debe - detalle.Haber

            q = Replace$(q, "'fecha'", Escape(detalle.FEcha))
            q = Replace$(q, "'comprobante'", Escape(detalle.Comprobante))
            q = Replace$(q, "'debe'", Escape(detalle.Debe))
            q = Replace$(q, "'haber'", Escape(detalle.Haber))
            q = Replace$(q, "'extra'", Escape(detalle.AtributoExtra))
            'Agrego esto 5
            q = Replace$(q, "'saldo'", Escape(saldo))

            q = Replace$(q, "'id_comprobante'", Escape(detalle.IdComprobante))
            q = Replace$(q, "'tipo_comprobante'", Escape(detalle.tipoComprobante))

            conectar.execute q
        Next detalle
        'Agregp esto 4
        saldo = 0

        Set Detalles = New Collection
        Dim Id As Long
        Id = 0
        Set rs = conectar.RSFactory("SELECT * FROM tmp_cta_cte_sort ORDER BY fecha ASC")

        While Not rs.EOF
            Id = Id + 1

            'Agrego esto 2
            saldo = saldo + rs!Debe - rs!Haber

            Set detalle = New DTODetalleCuentaCorriente
            detalle.tmpId = Id
            detalle.Comprobante = rs!Comprobante
            If Not IsNull(rs!FEcha) Then detalle.FEcha = rs!FEcha
            detalle.Debe = rs!Debe
            detalle.Haber = rs!Haber
            detalle.AtributoExtra = rs!extra
            detalle.tipoComprobante = rs!tipo_comprobante
            detalle.IdComprobante = rs!id_comprobante

            'Agrego esto 3
            detalle.saldo = saldo

            Detalles.Add detalle
            rs.MoveNext
        Wend
    End If

    Set FindAllDetallesProveedor2 = Detalles
End Function


Public Function FindAllDetalles(id_cliente As Long, Optional sortCollection As Boolean = True, Optional fecha_hasta As String) As Collection
'si se llama desde resumen de saldo no se necesita que este ordenado y me ahorro el overhead del ordenado por la base de datos

    Dim detalle As DTODetalleCuentaCorriente
    Dim Detalles As New Collection
    Dim q As String
    Dim rs As Recordset


    Set rs = conectar.RSFactory("SELECT saldo_inicial FROM saldo_inicial_cliente WHERE id_cliente = " & id_cliente)
    If Not rs.EOF Then
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Haber = rs!saldo_inicial
        detalle.Comprobante = "Saldo Inicial"
        Detalles.Add detalle
    End If


    Dim facturas As Collection
    Dim fac As Factura

    '12:03 AGREGO QUE TAMBIEN TENGA EN CUENTA LOS COMPROBANTES QUE ESTAN CANCELADOS PARCIALMENTE (AdminFacturas.estado = 5)
    q = "AdminFacturas.idCliente = " & id_cliente & _
        " AND (AdminFacturas.estado = " & EstadoFacturaCliente.Aprobada & _
        " OR AdminFacturas.estado = " & EstadoFacturaCliente.CanceladaNC & _
        " OR AdminFacturas.estado = " & EstadoFacturaCliente.CanceladaNCParcial & _
        " OR AdminFacturas.estado = " & EstadoFacturaCliente.AplicadaND & _
        " OR AdminFacturas.estado = " & EstadoFacturaCliente.AplicadaACbte & ")"


    If LenB(fecha_hasta) > 0 Then
        q = q & " and  AdminFacturas.FechaEmision <=" & conectar.Escape(fecha_hasta)
    End If

    Dim recs As String

    Set facturas = DAOFactura.FindAll(q)

    For Each fac In facturas
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = fac.GetShortDescription(False, True)

        Debug.Print (fac.NumeroFormateado & "-  Estado: " & fac.estado)

        If fac.Saldado Then
            recs = vbNullString
            Set rs = RSFactory("SELECT * FROM AdminRecibosDetalleFacturas WHERE idFactura =" & fac.Id)
            While Not rs.EOF
                recs = recs & "RC-" & rs!idRecibo & " "
                rs.MoveNext
            Wend
            detalle.Comprobante = detalle.Comprobante & " ( " & recs & ")"

        End If

        If fac.Cancelada Then
            detalle.Comprobante = detalle.Comprobante & " (cancelada NC)"

        End If

        If fac.Tipo.TipoDoc = tipoDocumentoContable.notaCredito Then
            detalle.Debe = 0
            detalle.Haber = fac.TotalEstatico.total * fac.CambioAPatron

        Else
            detalle.Debe = fac.TotalEstatico.total * fac.CambioAPatron
            detalle.Haber = 0

        End If

        detalle.FEcha = fac.FechaEmision
        detalle.AtributoExtra = (fac.Saldado = TipoSaldadoFactura.saldadoTotal) Or (fac.Saldado = TipoSaldadoFactura.notaCredito)
        detalle.tipoComprobante = Factura_
        detalle.IdComprobante = fac.Id

        Detalles.Add detalle

    Next fac

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim recibos As Collection
    Dim ret As retencionRecibo

    q = "rec.idCliente = " & id_cliente & " AND rec.estado = " & EstadoRecibo.Aprobado

    If LenB(fecha_hasta) Then
        q = q & "  AND rec.fecha<=" & conectar.Escape(fecha_hasta)
    End If

    Set recibos = DAORecibo.FindAll(q)

    Dim rec As Recibo

    For Each rec In recibos
        Set detalle = New DTODetalleCuentaCorriente

        detalle.Comprobante = "RC-" & rec.Id

        'detalle.Haber = funciones.RedondearDecimales(MonedaConverter.Convertir(rec.TotalEstatico.TotalReciboEstatico, rec.Moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.Redondeo, DAOMoneda.MONEDA_PESO_ID, rec.Moneda.Id), 2)
        If rec.TotalEstatico.TotalRecibidoEstatico > 0 Then
            detalle.Haber = funciones.RedondearDecimales(MonedaConverter.Convertir(rec.TotalEstatico.TotalReciboEstatico, rec.moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.redondeo, DAOMoneda.MONEDA_PESO_ID, rec.moneda.Id), 2) + funciones.RedondearDecimales(MonedaConverter.Convertir(rec.aCuenta, rec.moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.redondeo, DAOMoneda.MONEDA_PESO_ID, rec.moneda.Id), 2)
        Else
            'Set rec.facturas = DAOFactura.FindAll("AdminFacturas.id IN (SELECT idFactura FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.Id & ")")

            Set rec.Cheques = DAOCheques.FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & " IN (SELECT idCheque FROM AdminRecibosCheques WHERE idRecibo = " & rec.Id & ")")
            detalle.Haber = funciones.RedondearDecimales(MonedaConverter.Convertir(rec.TotalRecibido, rec.moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.redondeo, DAOMoneda.MONEDA_PESO_ID, rec.moneda.Id), 2)

            'comentado el 3-7-13
            'detalle.Haber = detalle.Haber - rec.TotalRetenciones
            ''MsgBox rec.Id & " Error"
        End If


        detalle.FEcha = rec.FEcha
        detalle.tipoComprobante = Recibo_
        detalle.IdComprobante = rec.Id
        Detalles.Add detalle

        For Each ret In rec.retenciones
            Set detalle = New DTODetalleCuentaCorriente
            detalle.tipoComprobante = Retencion_
            detalle.IdComprobante = rec.Id
            detalle.Comprobante = "RET-" & ret.NroRetencion
            detalle.Haber = ret.valor
            detalle.FEcha = rec.FEcha

            Detalles.Add detalle
        Next ret
    Next rec


    If sortCollection And Detalles.count > 0 Then

        Dim saldo As Double
        saldo = 0
        q = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp_cta_cte_sort ( fecha DATE, comprobante VARCHAR(50), debe DOUBLE, haber DOUBLE, extra INT, idComprobante INT, tipoComprobante INT) "    'TYPE=HEAP"
        conectar.execute q
        conectar.execute "TRUNCATE tmp_cta_cte_sort"



        For Each detalle In Detalles
            q = "INSERT INTO tmp_cta_cte_sort (fecha,comprobante,debe,haber,extra,tipoComprobante,idComprobante) VALUES ('fecha', 'comprobante', 'debe', 'haber','extra','tipoComprobante','idComprobante')"
            saldo = saldo + detalle.Debe - detalle.Haber

            If detalle.Comprobante = "Saldo Inicial" Then
                q = Replace$(q, "'fecha'", "'2000-01-01'")
            Else
                q = Replace$(q, "'fecha'", Escape(detalle.FEcha))
            End If

            q = Replace$(q, "'comprobante'", Escape(detalle.Comprobante))
            q = Replace$(q, "'debe'", Escape(detalle.Debe))
            q = Replace$(q, "'haber'", Escape(detalle.Haber))
            q = Replace$(q, "'extra'", Escape(detalle.AtributoExtra))
            q = Replace$(q, "'saldo'", Escape(saldo))
            q = Replace$(q, "'tipoComprobante'", Escape(detalle.tipoComprobante))
            q = Replace$(q, "'idComprobante'", Escape(detalle.IdComprobante))

            conectar.execute q

        Next detalle
        saldo = 0
        Set Detalles = New Collection
        Set rs = conectar.RSFactory("SELECT * FROM tmp_cta_cte_sort ORDER BY fecha ASC")
        While Not rs.EOF
            saldo = saldo + rs!Debe - rs!Haber
            Set detalle = New DTODetalleCuentaCorriente
            detalle.Comprobante = rs!Comprobante


            If Not IsNull(rs!FEcha) Then detalle.FEcha = rs!FEcha
            detalle.Debe = rs!Debe
            detalle.Haber = rs!Haber
            If IsNull(rs!extra) Then
                detalle.AtributoExtra = False
            Else
                detalle.AtributoExtra = Abs(rs!extra)
            End If

            detalle.tipoComprobante = Abs(rs!tipoComprobante)
            detalle.IdComprobante = Abs(rs!IdComprobante)
            detalle.saldo = saldo
            Detalles.Add detalle
            rs.MoveNext
        Wend
    End If

    Set FindAllDetalles = Detalles
End Function




Public Function ResumenSaldo() As Double
'pide los detalles los recorre y calcula el saldo, traer todos o por cliente?
End Function

