Attribute VB_Name = "DAOConciliacionBancaria"
Option Explicit


Public Function ExistePeriodoCerrado( _
    ByVal IdCuentaBancaria As Long, _
    ByVal FechaDesde As Date, _
    ByVal FechaHasta As Date _
) As Boolean

    On Error GoTo err1

    Dim q As String
    Dim rs As Recordset

    ExistePeriodoCerrado = False

    q = "SELECT id " _
      & "FROM conciliaciones_bancarias " _
      & "WHERE estado = 1 " _
      & "AND id_cuenta_bancaria = " _
      & conectar.Escape(IdCuentaBancaria) & " " _
      & "AND fecha_desde <= " _
      & conectar.Escape(FechaHasta) & " " _
      & "AND fecha_hasta >= " _
      & conectar.Escape(FechaDesde) & " " _
      & "LIMIT 1"

    Set rs = conectar.RSFactory(q)

    ExistePeriodoCerrado = Not rs.EOF

    Exit Function

err1:
    ExistePeriodoCerrado = True

End Function


Public Function Guardar( _
    ByRef Conciliacion As clsConciliacionBancaria _
) As Boolean

    On Error GoTo err1

    Dim q As String
    Dim Movimiento As DTOResumenBancario

    Guardar = False

    If Conciliacion Is Nothing Then Exit Function

    If Conciliacion.IdCuentaBancaria <= 0 Then Exit Function

    If Conciliacion.FechaDesde > Conciliacion.FechaHasta Then
        Exit Function
    End If

    If ExistePeriodoCerrado( _
            Conciliacion.IdCuentaBancaria, _
            Conciliacion.FechaDesde, _
            Conciliacion.FechaHasta) Then

        Exit Function

    End If

    conectar.BeginTransaction

    '------------------------------------------------------
    ' CABECERA
    '------------------------------------------------------
    q = "INSERT INTO conciliaciones_bancarias (" _
      & "id_cuenta_bancaria, fecha_desde," _
      & "fecha_hasta, fecha_cierre," _
      & "id_usuario_cierre, saldo_inicial," _
      & "total_ingresos, total_egresos," _
      & "saldo_final, cantidad_movimientos," _
      & "estado, observaciones" _
      & ") VALUES (" _
      & "'id_cuenta_bancaria'," _
      & "'fecha_desde','fecha_hasta'," _
      & "'fecha_cierre', 'id_usuario_cierre'," _
      & "'saldo_inicial','total_ingresos'," _
      & "'total_egresos','saldo_final'," _
      & "'cantidad_movimientos'," _
      & "1, 'observaciones' )"

    q = Replace$(q, _
        "'id_cuenta_bancaria'", _
        conectar.Escape(Conciliacion.IdCuentaBancaria))

    q = Replace$(q, _
        "'fecha_desde'", _
        conectar.Escape(Conciliacion.FechaDesde))

    q = Replace$(q, _
        "'fecha_hasta'", _
        conectar.Escape(Conciliacion.FechaHasta))

    q = Replace$(q, _
        "'fecha_cierre'", _
        conectar.Escape(Conciliacion.FechaCierre))

    q = Replace$(q, _
        "'id_usuario_cierre'", _
        conectar.Escape(Conciliacion.IdUsuarioCierre))

    q = Replace$(q, _
        "'saldo_inicial'", _
        conectar.Escape(Conciliacion.saldoInicial))

    q = Replace$(q, _
        "'total_ingresos'", _
        conectar.Escape(Conciliacion.totalIngresos))

    q = Replace$(q, _
        "'total_egresos'", _
        conectar.Escape(Conciliacion.totalEgresos))

    q = Replace$(q, _
        "'saldo_final'", _
        conectar.Escape(Conciliacion.saldoFinal))

    q = Replace$(q, _
        "'cantidad_movimientos'", _
        conectar.Escape(Conciliacion.CantidadMovimientos))

    q = Replace$(q, _
        "'observaciones'", _
        conectar.Escape(Conciliacion.Observaciones))

    If Not conectar.execute(q) Then GoTo err1

    Conciliacion.Id = conectar.UltimoId2()

    If Conciliacion.Id <= 0 Then GoTo err1

    '------------------------------------------------------
    ' SNAPSHOT
    '------------------------------------------------------
    For Each Movimiento In Conciliacion.Detalles

        If Not GuardarDetalle( _
                Conciliacion.Id, _
                Movimiento) Then

            GoTo err1

        End If

    Next Movimiento

    conectar.CommitTransaction

    Guardar = True
    Exit Function

err1:

    conectar.RollBackTransaction
    Guardar = False

End Function


Private Function GuardarDetalle( _
    ByVal IdConciliacion As Long, _
    ByRef Movimiento As DTOResumenBancario _
) As Boolean

    On Error GoTo err1

    Dim q As String

    GuardarDetalle = False

    q = "INSERT INTO conciliaciones_bancarias_detalles (" _
      & "id_conciliacion, fecha," _
      & "fecha_carga, id_banco," _
      & "banco, id_cuenta_bancaria," _
      & "cuenta_bancaria, cuenta_origen," _
      & "id_moneda, tipo_movimiento," _
      & "origen, id_origen," _
      & "numero_origen, id_operacion," _
      & "comprobante, detalle," _
      & "ingreso, egreso, saldo_acumulado" _
      & ") VALUES (" _
      & "'id_conciliacion', 'fecha'," _
      & "'fecha_carga', 'id_banco', 'banco'," _
      & "'id_cuenta_bancaria', 'cuenta_bancaria'," _
      & "'cuenta_origen', 'id_moneda'," _
      & "'tipo_movimiento', 'origen'," _
      & "'id_origen', 'numero_origen'," _
      & "'id_operacion', 'comprobante'," _
      & "'detalle', 'ingreso'," _
      & "'egreso', 'saldo_acumulado'" _
      & ")"

    q = Replace$(q, "'id_conciliacion'", _
        conectar.Escape(IdConciliacion))

    q = Replace$(q, "'fecha'", _
        conectar.Escape(Movimiento.FEcha))

    q = Replace$(q, "'fecha_carga'", _
        conectar.Escape(Movimiento.FechaCarga))

    q = Replace$(q, "'id_banco'", _
        conectar.Escape(Movimiento.IdBanco))

    q = Replace$(q, "'banco'", _
        conectar.Escape(Movimiento.Banco))

    q = Replace$(q, "'id_cuenta_bancaria'", _
        conectar.Escape(Movimiento.IdCuentaBancaria))

    q = Replace$(q, "'cuenta_bancaria'", _
        conectar.Escape(Movimiento.CuentaBancaria))

    q = Replace$(q, "'cuenta_origen'", _
        conectar.Escape(Movimiento.CuentaOrigen))

    q = Replace$(q, "'id_moneda'", _
        conectar.Escape(Movimiento.IdMoneda))

    q = Replace$(q, "'tipo_movimiento'", _
        conectar.Escape(Movimiento.TipoMovimiento))

    q = Replace$(q, "'origen'", _
        conectar.Escape(Movimiento.Origen))

    q = Replace$(q, "'id_origen'", _
        conectar.Escape(Movimiento.IdOrigen))

    q = Replace$(q, "'numero_origen'", _
        conectar.Escape(Movimiento.NumeroOrigen))

    q = Replace$(q, "'id_operacion'", _
        conectar.Escape(Movimiento.IdOperacion))

    q = Replace$(q, "'comprobante'", _
        conectar.Escape(Movimiento.Comprobante))

    q = Replace$(q, "'detalle'", _
        conectar.Escape(Movimiento.detalle))

    q = Replace$(q, "'ingreso'", _
        conectar.Escape(Movimiento.Ingreso))

    q = Replace$(q, "'egreso'", _
        conectar.Escape(Movimiento.Egreso))

    q = Replace$(q, "'saldo_acumulado'", _
        conectar.Escape(Movimiento.SaldoAcumulado))

    GuardarDetalle = conectar.execute(q)

    Exit Function

err1:
    GuardarDetalle = False

End Function


Public Function ObtenerIdConciliacionCerrada( _
    ByVal IdCuentaBancaria As Long, _
    ByVal FechaMovimiento As Date _
) As Long

    On Error GoTo err1

    Dim q As String
    Dim rs As Recordset
    Dim fechaSolo As Date

    ObtenerIdConciliacionCerrada = 0

    If IdCuentaBancaria <= 0 Then Exit Function

    fechaSolo = DateSerial( _
                    Year(FechaMovimiento), _
                    Month(FechaMovimiento), _
                    Day(FechaMovimiento))

    q = "SELECT id " _
      & "FROM conciliaciones_bancarias " _
      & "WHERE estado = 1 " _
      & "AND id_cuenta_bancaria = " _
      & conectar.Escape(IdCuentaBancaria) & " " _
      & "AND " & conectar.Escape(fechaSolo) _
      & " BETWEEN fecha_desde AND fecha_hasta " _
      & "ORDER BY id DESC " _
      & "LIMIT 1"

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF Then
        ObtenerIdConciliacionCerrada = CLng(rs!Id)
    End If

    Exit Function

err1:
    ObtenerIdConciliacionCerrada = 0

End Function


Public Function PeriodoCerrado( _
    ByVal IdCuentaBancaria As Long, _
    ByVal FechaMovimiento As Date _
) As Boolean

    PeriodoCerrado = _
        (ObtenerIdConciliacionCerrada( _
            IdCuentaBancaria, _
            FechaMovimiento) > 0)

End Function
