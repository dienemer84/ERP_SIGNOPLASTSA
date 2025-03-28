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
        cta.detalles.Add deta
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
    Dim detalles As New Collection

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
            detalles.Add detalle
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
        detalles.Add detalle
    End If

    Dim ordenes As New Collection
    Dim Orden As OrdenPago

    If LenB(condicion) > 0 Then
        cond1 = "and ordenes_pago.fecha<=" & condicion
    End If


    Set ordenes = DAOOrdenPago.FindAllByProveedor(id_proveedor, cond1 & "  and ordenes_pago.fecha> " & max_desde, soloOp)
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
            detalle.Haber = funciones.RedondearDecimales(Orden.TotalOrdenPago)          '.StaticTotalFacturas + Orden.TotalCompensatorios)
        End If
        
        detalle.FEcha = Orden.FEcha

        detalles.Add detalle
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

        detalles.Add detalle
    Next fac


    If sortCollection And detalles.count > 0 Then
        Dim q As String

        q = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp_cta_cte_sort (fecha DATE, comprobante VARCHAR(50), debe DOUBLE, haber DOUBLE, extra TINYINT, id_comprobante BIGINT, tipo_comprobante INT) TYPE=HEAP"
        conectar.execute q
        conectar.execute "TRUNCATE tmp_cta_cte_sort"


        For Each detalle In detalles
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

        Set detalles = New Collection
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
            detalles.Add detalle
            rs.MoveNext
        Wend
    End If

    Set FindAllDetallesProveedor = detalles
End Function


Public Function FindAllDetallesProveedor2(id_proveedor As Long, Optional sortCollection As Boolean = True, Optional condicion As String, Optional anteriores As Boolean = False, Optional soloOp As Boolean = False) As Collection

    Dim cond1 As String
    Dim detalle As DTODetalleCuentaCorriente
    Dim detalles As New Collection
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
            detalles.Add detalle
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
        detalles.Add detalle
    End If


    Dim ordenes As New Collection
    Dim Orden As OrdenPago

    If LenB(condicion) > 0 Then
        cond1 = "and ordenes_pago.fecha<=" & condicion
    End If


    Set ordenes = DAOOrdenPago.FindAllByProveedor(id_proveedor, cond1 & "  and ordenes_pago.fecha> " & max_desde, soloOp)
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
            detalle.Haber = funciones.RedondearDecimales(Orden.TotalOrdenPago)          '.StaticTotalFacturas + Orden.TotalCompensatorios)
        End If
        
        detalle.FEcha = Orden.FEcha

        detalles.Add detalle
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

        detalles.Add detalle
    Next fac


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

        detalles.Add detalle
        ' End If
    Next PagoACta


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   If sortCollection And detalles.count > 0 Then
        Dim q As String

        ' Agrego esto 7 y 8
        Dim saldo As Double
        saldo = 0

        q = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp_cta_cte_sort (fecha DATE, comprobante VARCHAR(50), debe DOUBLE, haber DOUBLE, extra TINYINT, id_comprobante BIGINT, tipo_comprobante INT) ENGINE=MEMORY;"
        conectar.execute q
        conectar.execute "TRUNCATE tmp_cta_cte_sort"



        For Each detalle In detalles
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

        Set detalles = New Collection
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

            detalles.Add detalle
            rs.MoveNext
        Wend
    End If

    Set FindAllDetallesProveedor2 = detalles
End Function


Public Function FindAllDetalles(id_cliente As Long, Optional sortCollection As Boolean = True, Optional fecha_hasta As String) As Collection
'si se llama desde resumen de saldo no se necesita que este ordenado y me ahorro el overhead del ordenado por la base de datos

    Dim detalle As DTODetalleCuentaCorriente
    Dim detalles As New Collection
    Dim q As String
    Dim rs As Recordset


    Set rs = conectar.RSFactory("SELECT saldo_inicial FROM saldo_inicial_cliente WHERE id_cliente = " & id_cliente)
    If Not rs.EOF Then
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Haber = rs!saldo_inicial
        detalle.Comprobante = "Saldo Inicial"
        detalles.Add detalle
    End If


    Dim facturas As Collection
    Dim fac As Factura

    '12:03 AGREGO QUE TAMBIEN TENGA EN CUENTA LOS COMPROBANTES QUE ESTAN CANCELADOS PARCIALMENTE (AdminFacturas.estado = 5)
    q = "AdminFacturas.idCliente = " & id_cliente & "  and (AdminFacturas.estado = " & EstadoFacturaCliente.Aprobada & " or AdminFacturas.estado = " & EstadoFacturaCliente.CanceladaNC & " or AdminFacturas.estado = " & EstadoFacturaCliente.CanceladaNCParcial & ")"


    If LenB(fecha_hasta) > 0 Then
        q = q & " and  AdminFacturas.FechaEmision <=" & conectar.Escape(fecha_hasta)
    End If

    Dim recs As String

    Set facturas = DAOFactura.FindAll(q)

    For Each fac In facturas
        Set detalle = New DTODetalleCuentaCorriente
        detalle.Comprobante = fac.GetShortDescription(False, True)

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

        detalles.Add detalle

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
            detalle.Haber = funciones.RedondearDecimales(MonedaConverter.Convertir(rec.TotalEstatico.TotalReciboEstatico, rec.moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.Redondeo, DAOMoneda.MONEDA_PESO_ID, rec.moneda.Id), 2) + funciones.RedondearDecimales(MonedaConverter.Convertir(rec.ACuenta, rec.moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.Redondeo, DAOMoneda.MONEDA_PESO_ID, rec.moneda.Id), 2)
        Else
            'Set rec.facturas = DAOFactura.FindAll("AdminFacturas.id IN (SELECT idFactura FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.Id & ")")

            Set rec.cheques = DAOCheques.FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & " IN (SELECT idCheque FROM AdminRecibosCheques WHERE idRecibo = " & rec.Id & ")")
            detalle.Haber = funciones.RedondearDecimales(MonedaConverter.Convertir(rec.TotalRecibido, rec.moneda.Id, DAOMoneda.MONEDA_PESO_ID) + MonedaConverter.Convertir(rec.Redondeo, DAOMoneda.MONEDA_PESO_ID, rec.moneda.Id), 2)

            'comentado el 3-7-13
            'detalle.Haber = detalle.Haber - rec.TotalRetenciones
            ''MsgBox rec.Id & " Error"
        End If


        detalle.FEcha = rec.FEcha
        detalle.tipoComprobante = Recibo_
        detalle.IdComprobante = rec.Id
        detalles.Add detalle

        For Each ret In rec.retenciones
            Set detalle = New DTODetalleCuentaCorriente
            detalle.tipoComprobante = Retencion_
            detalle.IdComprobante = rec.Id
            detalle.Comprobante = "RET-" & ret.NroRetencion
            detalle.Haber = ret.Valor
            detalle.FEcha = rec.FEcha

            detalles.Add detalle
        Next ret
    Next rec


    If sortCollection And detalles.count > 0 Then
        
        Dim saldo As Double
        saldo = 0
        q = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp_cta_cte_sort ( fecha DATE, comprobante VARCHAR(50), debe DOUBLE, haber DOUBLE, extra INT, idComprobante INT, tipoComprobante INT) "    'TYPE=HEAP"
        conectar.execute q
        conectar.execute "TRUNCATE tmp_cta_cte_sort"



        For Each detalle In detalles
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
        Set detalles = New Collection
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
            detalles.Add detalle
            rs.MoveNext
        Wend
    End If

    Set FindAllDetalles = detalles
End Function

Public Function ResumenSaldo() As Double
'pide los detalles los recorre y calcula el saldo, traer todos o por cliente?
End Function

