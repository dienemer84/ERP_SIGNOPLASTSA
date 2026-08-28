Attribute VB_Name = "DAOResumenBancario"
Option Explicit


Public Function FindAll( _
    ByVal fechaDesde As Date, _
    ByVal fechaHasta As Date, _
    Optional ByVal IdCuentaBancaria As Long = 0, _
    Optional ByVal IdMoneda As Long = 0, _
    Optional ByVal TipoMovimiento As String = vbNullString, _
    Optional ByVal Origen As String = vbNullString _
) As Collection

    On Error GoTo err1


    Dim q As String
    Dim filtro As String
    Dim rs As Recordset

    Dim col As New Collection
    Dim mov As DTOResumenBancario


    Dim etapa As String
    Dim numeroFila As Long


    q = "SELECT "
    
    q = q & " CASE "
    q = q & " WHEN movimientos.fecha IS NULL THEN NULL "
    q = q & " WHEN LEFT(CAST(movimientos.fecha AS CHAR), 10) = '0000-00-00' THEN NULL "
    q = q & " ELSE CAST(movimientos.fecha AS DATETIME) "
    q = q & " END AS fecha,"
    
    q = q & " CASE "
    q = q & " WHEN movimientos.fecha_carga IS NULL THEN NULL "
    q = q & " WHEN LEFT(CAST(movimientos.fecha_carga AS CHAR), 10) = '0000-00-00' THEN NULL "
    q = q & " ELSE CAST(movimientos.fecha_carga AS DATETIME) "
    q = q & " END AS fecha_carga,"
    
    q = q & " movimientos.id_banco,"
    q = q & " movimientos.banco,"
    q = q & " movimientos.id_cuenta_bancaria,"
    q = q & " movimientos.cuenta_bancaria,"
    
    '----------------------------------------------------------
    ' CUENTA ORIGEN DE TRANSFERENCIA INTERBANCARIA
    '----------------------------------------------------------
    q = q & " CASE "
    
    q = q & " WHEN movimientos.origen = " _
          & " 'TRANSFERENCIA INTERBANCARIA' " _
          & " AND movimientos.tipo_movimiento = 'INGRESO' "
    
    q = q & " THEN IFNULL(("
    
    q = q & " SELECT CONCAT(" _
          & " IFNULL(bo.nombre, '')," _
          & " ' | N° '," _
          & " IFNULL(co.cuenta, '')" _
          & " ) "
    
    q = q & " FROM movimientos_caja_bancos_operaciones mo "
    
    q = q & " INNER JOIN operaciones oo "
    q = q & " ON oo.id = mo.id_operacion "
    
    q = q & " LEFT JOIN AdminConfigCuentas co "
    q = q & " ON co.id = oo.cuentabanc_o_caja_id "
    
    q = q & " LEFT JOIN AdminConfigBancos bo "
    q = q & " ON bo.id = co.idBanco "
    
    q = q & " WHERE mo.id_movimiento_caja_bancos = " _
          & " movimientos.id_origen "
    
    q = q & " AND oo.pertenencia = 'banco' "
    q = q & " AND oo.entrada_salida = -1 "
    
    q = q & " LIMIT 1"
    
    q = q & " ), '-') "
    
    q = q & " ELSE '-' "
    
    q = q & " END AS cuenta_origen,"
    
    q = q & " movimientos.cbu,"
    
    q = q & " movimientos.id_moneda,"
    q = q & " movimientos.tipo_movimiento,"
    q = q & " movimientos.origen,"
    q = q & " movimientos.id_origen,"
    q = q & " movimientos.numero_origen,"
    q = q & " movimientos.id_operacion,"
    q = q & " movimientos.comprobante,"
    q = q & " movimientos.detalle,"
    q = q & " movimientos.ingreso,"
    q = q & " movimientos.egreso "
    
    q = q & "FROM ("
    
    q = q & SQLRecibos()
    q = q & " UNION ALL "
    
    q = q & SQLOrdenesPago()
    q = q & " UNION ALL "
    
    q = q & SQLOrdenesPagoCaja()
    q = q & " UNION ALL "
    
    q = q & SQLLiquidacionesCaja()
    q = q & " UNION ALL "
    
    q = q & SQLLiquidacionesCajaEfectivo()
    q = q & " UNION ALL "
    
    q = q & SQLPagosACuentaBanco()
    q = q & " UNION ALL "
    
    q = q & SQLPagosACuentaCaja()
    q = q & " UNION ALL "
    
    q = q & SQLBoletasDeposito()
    q = q & " UNION ALL "
    
    q = q & SQLChequesIngresadosBanco()
    q = q & " UNION ALL "

    q = q & SQLMovimientosManuales()
    
    ' Cerrar el SELECT interno
    q = q & ") movimientos "
    q = q & "WHERE 1 = 1 "

    q = q & "AND movimientos.fecha >= " _
          & conectar.Escape(fechaDesde) & " "

    q = q & "AND movimientos.fecha <= " _
          & conectar.Escape(fechaHasta) & " "

    If IdCuentaBancaria > 0 Then
        q = q & "AND movimientos.id_cuenta_bancaria = " _
              & IdCuentaBancaria & " "
    End If

    If IdMoneda > 0 Then
        q = q & "AND movimientos.id_moneda = " _
              & IdMoneda & " "
    End If
    
    If LenB(TipoMovimiento) > 0 Then
    q = q & "AND movimientos.tipo_movimiento = " _
              & conectar.Escape(TipoMovimiento) & " "
    End If
    
    If LenB(Origen) > 0 Then
        q = q & "AND movimientos.origen = " _
              & conectar.Escape(Origen) & " "
    End If


    q = q & "ORDER BY "
    q = q & " movimientos.banco,"
    q = q & " movimientos.cuenta_bancaria,"
    q = q & " movimientos.id_moneda,"
    q = q & " movimientos.fecha,"
    q = q & " movimientos.fecha_carga,"
    q = q & " movimientos.id_operacion"

    Debug.Print String$(100, "-")
    Debug.Print q
    Debug.Print String$(100, "-")

    etapa = "Abrir recordset"
    Set rs = conectar.RSFactory(q)
    
    numeroFila = 0

While Not rs.EOF

    numeroFila = numeroFila + 1

    Set mov = New DTOResumenBancario

    etapa = "fecha"
    mov.FEcha = ValorFecha(rs, "fecha")

    etapa = "fecha_carga"
    mov.FechaCarga = ValorFecha(rs, "fecha_carga")

    etapa = "id_banco"
    mov.IdBanco = ValorLong(rs, "id_banco")

    etapa = "banco"
    mov.Banco = ValorTexto(rs, "banco")

    etapa = "id_cuenta_bancaria"
    mov.IdCuentaBancaria = _
        ValorLong(rs, "id_cuenta_bancaria")

    etapa = "cuenta_bancaria"
    mov.CuentaBancaria = _
        ValorTexto(rs, "cuenta_bancaria")
    
    etapa = "cuenta_origen"
    mov.CuentaOrigen = _
        ValorTexto(rs, "cuenta_origen")

    etapa = "cbu"
    mov.CBU = ValorTexto(rs, "cbu")

    etapa = "id_moneda"
    mov.IdMoneda = ValorLong(rs, "id_moneda")

    etapa = "tipo_movimiento"
    mov.TipoMovimiento = _
        ValorTexto(rs, "tipo_movimiento")

    etapa = "origen"
    mov.Origen = ValorTexto(rs, "origen")

    etapa = "id_origen"
    mov.IdOrigen = ValorLong(rs, "id_origen")

    etapa = "numero_origen"
    mov.NumeroOrigen = ValorTexto(rs, "numero_origen")

    etapa = "id_operacion"
    mov.IdOperacion = ValorLong(rs, "id_operacion")

    etapa = "comprobante"
    mov.Comprobante = ValorTexto(rs, "comprobante")

    etapa = "detalle"
    mov.detalle = ValorTexto(rs, "detalle")

    etapa = "ingreso"
    mov.Ingreso = ValorDouble(rs, "ingreso")

    etapa = "egreso"
    mov.Egreso = ValorDouble(rs, "egreso")

    etapa = "agregar movimiento"
    col.Add mov

    rs.MoveNext

Wend

etapa = "calcular saldos"
CalcularSaldos col
    Set FindAll = col
    Exit Function

err1:
    Debug.Print "DAOResumenBancario.FindAll"
    Debug.Print "Error: " & Err.Number
    Debug.Print "Descripción: " & Err.Description
    Debug.Print "SQL: " & q

MsgBox "Error en DAOResumenBancario.FindAll" & vbCrLf & _
       "Número: " & Err.Number & vbCrLf & _
       "Descripción: " & Err.Description & vbCrLf & _
       "Fila: " & numeroFila & vbCrLf & _
       "Campo o etapa: " & etapa, _
       vbCritical, _
       "Reporte bancario"

    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing
    Set FindAll = Nothing

End Function


Private Function SQLRecibos() As String

    Dim q As String

    q = "SELECT "
    q = q & " o.fecha_operacion AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(c.moneda_id, 0)," _
          & " NULLIF(rec.idMoneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " 'INGRESO' AS tipo_movimiento,"
    q = q & " 'RECIBO' AS origen,"

    q = q & " rec.id AS id_origen,"
    q = q & " CAST(rec.id AS CHAR) AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"
    q = q & " IFNULL(cli.razon, '') AS detalle,"

    q = q & " IFNULL(o.monto, 0) AS ingreso,"
    q = q & " 0 AS egreso "

    q = q & "FROM AdminRecibos rec "

    q = q & "INNER JOIN operaciones_recibos opr "
    q = q & " ON opr.reciboId = rec.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = opr.operacionId "

    q = q & "LEFT JOIN AdminConfigCuentas c "
    q = q & " ON c.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "LEFT JOIN clientes cli "
    q = q & " ON cli.id = rec.idCliente "

    q = q & "WHERE rec.estado = 2 "
    q = q & "AND o.pertenencia = 'banco' "
    q = q & "AND o.entrada_salida = 1 "

    SQLRecibos = q

End Function


Private Function SQLOrdenesPago() As String

    Dim q As String

    q = "SELECT "
    q = q & " o.fecha_operacion AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(c.moneda_id, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " 'EGRESO' AS tipo_movimiento,"
    q = q & " 'ORDEN DE PAGO' AS origen,"

    q = q & " op.id AS id_origen,"
    q = q & " CAST(op.id AS CHAR) AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"
    q = q & " IFNULL(op.cuenta_contable_desc, '') AS detalle,"

    q = q & " 0 AS ingreso,"
    q = q & " IFNULL(o.monto, 0) AS egreso "

    q = q & "FROM ordenes_pago op "

    q = q & "INNER JOIN ordenes_pago_operaciones opo "
    q = q & " ON opo.id_orden_pago = op.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = opo.id_operacion "

    q = q & "LEFT JOIN AdminConfigCuentas c "
    q = q & " ON c.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "WHERE op.estado = 1 "
    q = q & "AND o.pertenencia = 'banco' "
    q = q & "AND o.entrada_salida = -1 "

    SQLOrdenesPago = q

End Function


Private Function SQLLiquidacionesCaja() As String

    Dim q As String

    q = "SELECT "
    q = q & " o.fecha_operacion AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(c.moneda_id, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " 'EGRESO' AS tipo_movimiento,"
    q = q & " 'LIQUIDACION DE CAJA' AS origen,"

    q = q & " lc.id AS id_origen,"
    q = q & " CAST(lc.numero_liq AS CHAR) AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"
    q = q & " CONCAT('Liquidacion Nro ', lc.numero_liq)"
    q = q & " AS detalle,"

    q = q & " 0 AS ingreso,"
    q = q & " IFNULL(o.monto, 0) AS egreso "

    q = q & "FROM liquidaciones_caja lc "

    q = q & "INNER JOIN liquidaciones_caja_operaciones lco "
    q = q & " ON lco.id_liquidacion_caja = lc.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = lco.id_operacion "

    q = q & "LEFT JOIN AdminConfigCuentas c "
    q = q & " ON c.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "WHERE lc.estado = 1 "
    q = q & "AND o.pertenencia = 'banco' "
    q = q & "AND o.entrada_salida = -1 "

    SQLLiquidacionesCaja = q

End Function


Private Function SQLMovimientosManuales() As String

    Dim q As String

    q = "SELECT "
    q = q & " o.fecha_operacion AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(c.moneda_id, 0)," _
          & " NULLIF(mov.id_moneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = 1 THEN 'INGRESO' "
    q = q & " ELSE 'EGRESO' "
    q = q & " END AS tipo_movimiento,"

    ' Para transferencias se muestra un origen separado.
    q = q & " CASE "
    q = q & " WHEN UPPER(IFNULL(mov.tipo_movimiento, '')) = " _
          & " 'TRANSFERENCIA' "
    q = q & " THEN 'TRANSFERENCIA INTERBANCARIA' "
    q = q & " ELSE 'MOVIMIENTO CAJA/BANCOS' "
    q = q & " END AS origen,"

    q = q & " mov.id AS id_origen,"
    q = q & " CAST(mov.id AS CHAR) AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"
    q = q & " IFNULL(mov.observaciones, '') AS detalle,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = 1 "
    q = q & " THEN ABS(IFNULL(o.monto, 0)) "
    q = q & " ELSE 0 "
    q = q & " END AS ingreso,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = -1 "
    q = q & " THEN IFNULL(o.monto, 0) "
    q = q & " ELSE 0 "
    q = q & " END AS egreso "

    q = q & "FROM movimientos_caja_bancos mov "

    q = q & "INNER JOIN movimientos_caja_bancos_operaciones movop "
    q = q & " ON movop.id_movimiento_caja_bancos = mov.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = movop.id_operacion "

    '----------------------------------------------------------
    ' CUENTA BANCARIA
    '
    ' INGRESO/EGRESO:
    '   primero se usa la cuenta principal de la cabecera.
    '
    ' TRANSFERENCIA:
    '   se usa la cuenta particular de cada operación.
    '----------------------------------------------------------
    q = q & "LEFT JOIN AdminConfigCuentas c "
    q = q & " ON c.id = CASE "

    q = q & " WHEN UPPER(IFNULL(mov.tipo_movimiento, '')) = " _
          & " 'TRANSFERENCIA' "

    q = q & " THEN NULLIF(o.cuentabanc_o_caja_id, 0) "

    q = q & " ELSE COALESCE(" _
          & " NULLIF(mov.id_cuenta_bancaria_principal, 0)," _
          & " NULLIF(o.cuentabanc_o_caja_id, 0)" _
          & " ) "

    q = q & " END "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "WHERE mov.estado = 1 "
    
    '----------------------------------------------------------
    ' EVITAR DOBLE IMPACTO:
    '
    ' Si una TRANSFERENCIA tiene cheques propios asociados
    ' desde la misma cuenta origen y dichos cheques conforman
    ' el importe de la operación de salida, NO mostrar esa
    ' salida bancaria.
    '
    ' El egreso real aparecerá cuando el cheque ingrese al
    ' banco mediante fecha_ingreso_banco.
    '----------------------------------------------------------
    
    q = q & "AND NOT ("
    
    q = q & " UPPER(IFNULL(mov.tipo_movimiento, '')) " _
        & " IN ('TRANSFERENCIA', 'EGRESO') "
    
    q = q & " AND o.entrada_salida = -1 "
    
    q = q & " AND EXISTS ("
    
    q = q & " SELECT 1 "
    
    q = q & " FROM movimientos_caja_bancos_cheques mcc "
    
    q = q & " INNER JOIN Cheques chx "
    q = q & " ON chx.id = mcc.id_cheque "
    
    q = q & " INNER JOIN Chequeras chqx "
    q = q & " ON chqx.id = chx.id_chequera "
    
    q = q & " WHERE mcc.id_movimiento_caja_bancos = mov.id "
    
    q = q & " AND IFNULL(chx.propio, 0) = 1 "
    
    'El cheque debe pertenecer a la misma cuenta bancaria
    'que estamos intentando debitar.
    q = q & " AND chqx.id_cuenta_bancaria = " _
          & " o.cuentabanc_o_caja_id "
    
    q = q & " GROUP BY " _
          & " mcc.id_movimiento_caja_bancos," _
          & " chqx.id_cuenta_bancaria "
    
    'El total de los cheques debe coincidir con el egreso.
    q = q & " HAVING ABS(SUM(IFNULL(chx.monto, 0))) = " _
          & " ABS(IFNULL(o.monto, 0)) "
    
    q = q & " )"
    
    q = q & ") "

    SQLMovimientosManuales = q

End Function


Private Sub CalcularSaldos(ByRef Movimientos As Collection)

    Dim mov As DTOResumenBancario

    Dim saldo As Double
    Dim claveAnterior As String
    Dim claveActual As String

    saldo = 0
    claveAnterior = vbNullString

    For Each mov In Movimientos

        claveActual = CStr(mov.IdCuentaBancaria) _
                    & "|" _
                    & CStr(mov.IdMoneda)

        If claveActual <> claveAnterior Then
            saldo = 0
            claveAnterior = claveActual
        End If

        saldo = saldo + mov.Ingreso - mov.Egreso
        mov.SaldoAcumulado = saldo

    Next mov

End Sub


Private Function ValorTexto( _
    ByRef rs As Recordset, _
    ByVal campo As String _
) As String

    If IsNull(rs.Fields(campo).value) Then
        ValorTexto = vbNullString
    Else
        ValorTexto = CStr(rs.Fields(campo).value)
    End If

End Function


Private Function ValorLong( _
    ByRef rs As Recordset, _
    ByVal campo As String _
) As Long

    If IsNull(rs.Fields(campo).value) Then
        ValorLong = 0
    Else
        ValorLong = CLng(rs.Fields(campo).value)
    End If

End Function


Private Function ValorDouble( _
    ByRef rs As Recordset, _
    ByVal campo As String _
) As Double

    If IsNull(rs.Fields(campo).value) Then
        ValorDouble = 0
    Else
        ValorDouble = CDbl(rs.Fields(campo).value)
    End If

End Function


Private Function ValorFecha( _
    ByRef rs As Recordset, _
    ByVal campo As String _
) As Date

    On Error GoTo fechaInvalida

    Dim valor As Variant
    Dim texto As String

    Dim anio As Integer
    Dim mes As Integer
    Dim dia As Integer

    Dim hora As Integer
    Dim minuto As Integer
    Dim segundo As Integer

    valor = rs.Fields(campo).value

    If IsNull(valor) Or IsEmpty(valor) Then
        ValorFecha = 0
        Exit Function
    End If

    ' ADO ya entregó una fecha real.
    If VarType(valor) = vbDate Then
        ValorFecha = CDate(valor)
        Exit Function
    End If

    ' Algunos drivers MySQL entregan el resultado de un UNION
    ' como un arreglo de bytes.
    If VarType(valor) = (vbArray Or vbByte) Then
        texto = StrConv(valor, vbUnicode)
    Else
        texto = CStr(valor)
    End If

    texto = Trim$(texto)

    ' Eliminar posibles caracteres nulos.
    texto = Replace(texto, Chr$(0), vbNullString)

    If LenB(texto) = 0 Then
        ValorFecha = 0
        Exit Function
    End If

    If Left$(texto, 10) = "0000-00-00" Then
        ValorFecha = 0
        Exit Function
    End If

    ' Formato ISO de MySQL:
    ' yyyy-mm-dd
    ' yyyy-mm-dd hh:mm:ss
    If Len(texto) >= 10 _
       And Mid$(texto, 5, 1) = "-" _
       And Mid$(texto, 8, 1) = "-" Then

        anio = CInt(Left$(texto, 4))
        mes = CInt(Mid$(texto, 6, 2))
        dia = CInt(Mid$(texto, 9, 2))

        ValorFecha = DateSerial(anio, mes, dia)

        If Len(texto) >= 19 Then

            hora = CInt(Mid$(texto, 12, 2))
            minuto = CInt(Mid$(texto, 15, 2))
            segundo = CInt(Mid$(texto, 18, 2))

            ValorFecha = ValorFecha + _
                         TimeSerial(hora, minuto, segundo)

        End If

        Exit Function
    End If

    ' Fecha entregada con formato regional.
    If IsDate(texto) Then
        ValorFecha = CDate(texto)
        Exit Function
    End If

    GoTo fechaInvalida

fechaInvalida:
    Debug.Print String$(60, "-")
    Debug.Print "Fecha inválida"
    Debug.Print "Campo: "; campo
    Debug.Print "Tipo Variant: "; VarType(valor)
    Debug.Print "Tipo: "; TypeName(valor)
    Debug.Print "Texto: "; texto
    Debug.Print String$(60, "-")

    ValorFecha = 0

End Function



Private Function SQLOrdenesPagoCaja() As String

    Dim q As String

    q = "SELECT "
    q = q & " o.fecha_operacion AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'CAJA') AS banco,"

    q = q & " IFNULL(cf.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(cf.cuenta, ca.nombre) AS cuenta_bancaria,"
    q = q & " IFNULL(cf.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(cf.moneda_id, 0)," _
          & " NULLIF(op.id_moneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = 1 THEN 'INGRESO' "
    q = q & " ELSE 'EGRESO' "
    q = q & " END AS tipo_movimiento,"

    q = q & " 'ORDEN DE PAGO' AS origen,"

    q = q & " op.id AS id_origen,"
    q = q & " CAST(op.id AS CHAR) AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"
    q = q & " IFNULL(op.cuenta_contable_desc, '') AS detalle,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = 1 "
    q = q & " THEN ABS(IFNULL(o.monto, 0)) "
    q = q & " ELSE 0 END AS ingreso,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = -1 "
    q = q & " THEN IFNULL(o.monto, 0) "
    q = q & " ELSE 0 END AS egreso "

    q = q & "FROM ordenes_pago op "

    q = q & "INNER JOIN ordenes_pago_operaciones opo "
    q = q & " ON opo.id_orden_pago = op.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = opo.id_operacion "

    q = q & "INNER JOIN cajas ca "
    q = q & " ON ca.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigCuentas cf "
    q = q & " ON cf.id = ca.id_cuenta_financiera "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = cf.idBanco "

    q = q & "WHERE op.estado = 1 "
    q = q & "AND o.pertenencia = 'caja' "

    SQLOrdenesPagoCaja = q

End Function


Private Function SQLLiquidacionesCajaEfectivo() As String

    Dim q As String

    q = "SELECT "
    q = q & " o.fecha_operacion AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'CAJA') AS banco,"

    q = q & " IFNULL(cf.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(cf.cuenta, ca.nombre) AS cuenta_bancaria,"
    q = q & " IFNULL(cf.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(cf.moneda_id, 0)," _
          & " NULLIF(lc.id_moneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = 1 THEN 'INGRESO' "
    q = q & " ELSE 'EGRESO' "
    q = q & " END AS tipo_movimiento,"

    q = q & " 'LIQUIDACION DE CAJA' AS origen,"

    q = q & " lc.id AS id_origen,"
    q = q & " CAST(lc.numero_liq AS CHAR) AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"
    q = q & " CONCAT('Liquidacion Nro ', lc.numero_liq) " _
          & " AS detalle,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = 1 "
    q = q & " THEN ABS(IFNULL(o.monto, 0)) "
    q = q & " ELSE 0 END AS ingreso,"

    q = q & " CASE "
    q = q & " WHEN o.entrada_salida = -1 "
    q = q & " THEN IFNULL(o.monto, 0)"
    q = q & " ELSE 0 END AS egreso "

    q = q & "FROM liquidaciones_caja lc "

    q = q & "INNER JOIN liquidaciones_caja_operaciones lco "
    q = q & " ON lco.id_liquidacion_caja = lc.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = lco.id_operacion "

    q = q & "INNER JOIN cajas ca "
    q = q & " ON ca.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigCuentas cf "
    q = q & " ON cf.id = ca.id_cuenta_financiera "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = cf.idBanco "

    q = q & "WHERE lc.estado = 1 "
    q = q & "AND o.pertenencia = 'caja' "

    SQLLiquidacionesCajaEfectivo = q

End Function


Private Function SQLPagosACuentaBanco() As String

    Dim q As String

    q = "SELECT "

    q = q & " COALESCE(o.fecha_operacion, pcta.fecha) AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(c.moneda_id, 0)," _
          & " NULLIF(pcta.id_moneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " 'EGRESO' AS tipo_movimiento,"
    q = q & " 'PAGO A CUENTA' AS origen,"

    q = q & " pcta.id AS id_origen,"
    
    q = q & " CASE "
    q = q & " WHEN IFNULL(opcta.id_orden_pago, 0) > 0 "
    q = q & " THEN CONCAT(" _
          & " CAST(pcta.id AS CHAR)," _
          & " ' / OP '," _
          & " CAST(opcta.id_orden_pago AS CHAR)" _
          & " ) "
    q = q & " ELSE CAST(pcta.id AS CHAR) "
    q = q & " END AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"

    q = q & " CONCAT(" _
          & " 'Proveedor: '," _
          & " IFNULL(prov.razon, '')" _
          & " ) AS detalle,"

    q = q & " 0 AS ingreso,"
    q = q & " IFNULL(o.monto, 0) AS egreso "

    q = q & "FROM pagos_a_cuenta pcta "
    
    ' OP donde fue aplicado el Pago a Cuenta.
    q = q & "LEFT JOIN (" _
          & " SELECT id_pago_a_cuenta," _
          & " MAX(id_orden_pago) AS id_orden_pago " _
          & " FROM ordenes_pago_pagos_a_cuenta " _
          & " GROUP BY id_pago_a_cuenta" _
          & ") opcta "
    
    q = q & " ON opcta.id_pago_a_cuenta = pcta.id "

    ' Se usa DISTINCT como protección ante relaciones repetidas.
    q = q & "INNER JOIN (" _
          & " SELECT DISTINCT " _
          & " id_pago_a_cuenta," _
          & " id_operacion " _
          & " FROM pagos_a_cuenta_operaciones" _
          & ") pco "

    q = q & " ON pco.id_pago_a_cuenta = pcta.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = pco.id_operacion "

    q = q & "LEFT JOIN AdminConfigCuentas c "
    q = q & " ON c.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "LEFT JOIN proveedores prov "
    q = q & " ON prov.id = pcta.id_proveedor "

    q = q & "WHERE o.pertenencia = 'banco' "
    q = q & "AND o.entrada_salida = -1 "

    SQLPagosACuentaBanco = q

End Function


Private Function SQLPagosACuentaCaja() As String

    Dim q As String

    q = "SELECT "

    q = q & " COALESCE(o.fecha_operacion, pcta.fecha) AS fecha,"
    q = q & " o.fecha_carga AS fecha_carga,"

    ' La Caja se muestra con el banco de su cuenta financiera.
    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'CAJA') AS banco,"

    ' Se devuelve el ID 30, no el ID 1 de la tabla cajas.
    q = q & " IFNULL(cf.id, 0) AS id_cuenta_bancaria,"

    q = q & " CASE "
    q = q & " WHEN cf.id IS NOT NULL "
    q = q & " THEN cf.cuenta "
    q = q & " ELSE ca.nombre "
    q = q & " END AS cuenta_bancaria,"

    q = q & " IFNULL(cf.cbu, '') AS cbu,"

    q = q & " COALESCE(" _
          & " NULLIF(o.moneda_id, 0)," _
          & " NULLIF(cf.moneda_id, 0)," _
          & " NULLIF(pcta.id_moneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    q = q & " 'EGRESO' AS tipo_movimiento,"
    q = q & " 'PAGO A CUENTA' AS origen,"

    q = q & " pcta.id AS id_origen,"
    
    q = q & " CASE "
    q = q & " WHEN IFNULL(opcta.id_orden_pago, 0) > 0 "
    q = q & " THEN CONCAT(" _
          & " CAST(pcta.id AS CHAR)," _
          & " ' / OP '," _
          & " CAST(opcta.id_orden_pago AS CHAR)" _
          & " ) "
    q = q & " ELSE CAST(pcta.id AS CHAR) "
    q = q & " END AS numero_origen,"

    q = q & " o.id AS id_operacion,"
    q = q & " IFNULL(o.comprobante, '-') AS comprobante,"

    q = q & " CONCAT(" _
          & " 'Proveedor: '," _
          & " IFNULL(prov.razon, '')" _
          & " ) AS detalle,"

    q = q & " 0 AS ingreso,"
    q = q & " IFNULL(o.monto, 0) AS egreso "

    q = q & "FROM pagos_a_cuenta pcta "
    
    ' OP donde fue aplicado el Pago a Cuenta.
    q = q & "LEFT JOIN (" _
          & " SELECT id_pago_a_cuenta," _
          & " MAX(id_orden_pago) AS id_orden_pago " _
          & " FROM ordenes_pago_pagos_a_cuenta " _
          & " GROUP BY id_pago_a_cuenta" _
          & ") opcta "
    
    q = q & " ON opcta.id_pago_a_cuenta = pcta.id "

    q = q & "INNER JOIN (" _
          & " SELECT DISTINCT " _
          & " id_pago_a_cuenta," _
          & " id_operacion " _
          & " FROM pagos_a_cuenta_operaciones" _
          & ") pco "

    q = q & " ON pco.id_pago_a_cuenta = pcta.id "

    q = q & "INNER JOIN operaciones o "
    q = q & " ON o.id = pco.id_operacion "

    q = q & "INNER JOIN cajas ca "
    q = q & " ON ca.id = o.cuentabanc_o_caja_id "

    q = q & "LEFT JOIN AdminConfigCuentas cf "
    q = q & " ON cf.id = ca.id_cuenta_financiera "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = cf.idBanco "

    q = q & "LEFT JOIN proveedores prov "
    q = q & " ON prov.id = pcta.id_proveedor "

    q = q & "WHERE o.pertenencia = 'caja' "
    q = q & "AND o.entrada_salida = -1 "

    SQLPagosACuentaCaja = q

End Function


Private Function SQLBoletasDeposito() As String

    Dim q As String

    q = "SELECT "

    '----------------------------------------------------------
    ' FECHAS
    '----------------------------------------------------------
    q = q & " bd.fecha_deposito AS fecha,"
    q = q & " bd.fecha_deposito AS fecha_carga,"

    '----------------------------------------------------------
    ' BANCO
    '----------------------------------------------------------
    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    '----------------------------------------------------------
    ' CUENTA BANCARIA
    '----------------------------------------------------------
    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    '----------------------------------------------------------
    ' MONEDA
    '----------------------------------------------------------
    q = q & " IFNULL(c.moneda_id, 0) AS id_moneda,"

    '----------------------------------------------------------
    ' TIPO
    '----------------------------------------------------------
    q = q & " 'INGRESO' AS tipo_movimiento,"
    q = q & " 'DEPOSITO' AS origen,"

    '----------------------------------------------------------
    ' IDENTIFICACION
    '----------------------------------------------------------
    q = q & " bd.id AS id_origen,"
    q = q & " CAST(bd.numero_boleta AS CHAR) AS numero_origen,"

    ' No necesitamos la operación individual del cheque.
    ' Usamos el ID de la boleta para mantener un valor único.
    q = q & " bd.id AS id_operacion,"

    q = q & " CAST(bd.numero_boleta AS CHAR) AS comprobante,"

    '----------------------------------------------------------
    ' DETALLE
    '----------------------------------------------------------
    q = q & " CONCAT(" _
          & " 'Boleta de deposito Nro ', " _
          & " bd.numero_boleta" _
          & " ) AS detalle,"

    '----------------------------------------------------------
    ' IMPORTE
    '----------------------------------------------------------
    q = q & " IFNULL(bd.monto, 0) AS ingreso,"
    q = q & " 0 AS egreso "

    '----------------------------------------------------------
    ' TABLAS
    '----------------------------------------------------------
    q = q & "FROM boleta_deposito bd "

    q = q & "LEFT JOIN AdminConfigCuentas c "
    q = q & " ON c.id = bd.id_cuenta "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "WHERE bd.numero_boleta IS NOT NULL "

    SQLBoletasDeposito = q

End Function

Private Function SQLChequesIngresadosBanco() As String

    Dim q As String

    q = "SELECT "

    '----------------------------------------------------------
    ' FECHA REAL DE IMPACTO BANCARIO
    '----------------------------------------------------------
    q = q & " ch.fecha_ingreso_banco AS fecha,"
    q = q & " ch.fecha_ingreso_banco AS fecha_carga,"

    '----------------------------------------------------------
    ' BANCO / CUENTA DE LA CHEQUERA
    '----------------------------------------------------------
    q = q & " IFNULL(b.id, 0) AS id_banco,"
    q = q & " IFNULL(b.nombre, 'SIN BANCO') AS banco,"

    q = q & " IFNULL(c.id, 0) AS id_cuenta_bancaria,"
    q = q & " IFNULL(c.cuenta, 'SIN CUENTA') AS cuenta_bancaria,"
    q = q & " IFNULL(c.cbu, '') AS cbu,"

    '----------------------------------------------------------
    ' MONEDA
    '----------------------------------------------------------
    q = q & " COALESCE(" _
          & " NULLIF(ch.id_moneda, 0)," _
          & " NULLIF(c.moneda_id, 0)," _
          & " NULLIF(chq.id_moneda, 0)," _
          & " 0" _
          & " ) AS id_moneda,"

    '----------------------------------------------------------
    ' MOVIMIENTO
    '----------------------------------------------------------
    q = q & " 'EGRESO' AS tipo_movimiento,"
    q = q & " 'CHEQUE PROPIO' AS origen,"

    ' ID interno = cheque.
    q = q & " ch.id AS id_origen,"

    ' Lo que ve el usuario = número del cheque.
    q = q & " CAST(ch.numero AS CHAR) AS numero_origen,"

    q = q & " ch.id AS id_operacion,"

    '----------------------------------------------------------
    ' COMPROBANTE QUE ORIGINÓ EL CHEQUE
    '----------------------------------------------------------
    q = q & " CASE "

    q = q & " WHEN IFNULL(ch.orden_pago_origen, 0) > 0 "
    q = q & " THEN CONCAT('OP ', ch.orden_pago_origen) "

    q = q & " WHEN IFNULL(ch.pago_a_cuenta_origen, 0) > 0 "
    q = q & " THEN CONCAT('PCTA ', ch.pago_a_cuenta_origen) "

    q = q & " WHEN IFNULL(ch.liquidacion_caja_origen, 0) > 0 "
    q = q & " THEN CONCAT(" _
          & " 'LIQ ', " _
          & " IFNULL(lc.numero_liq, ch.liquidacion_caja_origen)" _
          & " ) "

    q = q & " WHEN IFNULL(ch.movimiento_origen, 0) > 0 "
    q = q & " THEN CONCAT('MOV ', ch.movimiento_origen) "

    q = q & " ELSE '-' "

    q = q & " END AS comprobante,"

    '----------------------------------------------------------
    ' DETALLE
    '----------------------------------------------------------
    q = q & " CONCAT(" _
          & " 'Cheque Nro ', ch.numero" _
          & " ) AS detalle,"

    '----------------------------------------------------------
    ' IMPORTE
    '----------------------------------------------------------
    q = q & " 0 AS ingreso,"
    q = q & " ABS(IFNULL(ch.monto, 0)) AS egreso "

    '----------------------------------------------------------
    ' TABLAS
    '----------------------------------------------------------
    q = q & "FROM Cheques ch "

    q = q & "INNER JOIN Chequeras chq "
    q = q & " ON chq.id = ch.id_chequera "

    q = q & "INNER JOIN AdminConfigCuentas c "
    q = q & " ON c.id = chq.id_cuenta_bancaria "

    q = q & "LEFT JOIN AdminConfigBancos b "
    q = q & " ON b.id = c.idBanco "

    q = q & "LEFT JOIN liquidaciones_caja lc "
    q = q & " ON lc.id = ch.liquidacion_caja_origen "

    '----------------------------------------------------------
    ' SOLAMENTE CHEQUES PROPIOS QUE YA INGRESARON AL BANCO
    '----------------------------------------------------------
    q = q & "WHERE IFNULL(ch.propio, 0) = 1 "
    q = q & "AND IFNULL(ch.ingresado, 0) = 1 "
    q = q & "AND ch.fecha_ingreso_banco IS NOT NULL "
    q = q & "AND ch.id_chequera IS NOT NULL "

    SQLChequesIngresadosBanco = q

End Function


