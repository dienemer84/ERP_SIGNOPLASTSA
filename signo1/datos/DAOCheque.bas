Attribute VB_Name = "DAOCheques"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NUMERO As String = "numero"
Public Const CAMPO_FECHA_RECIBIDO As String = "fecha_recibido"
Public Const CAMPO_FECHA_VENCIMIENTO As String = "fecha_vencimiento"
Public Const CAMPO_MONTO As String = "monto"
Public Const CAMPO_ID_CHEQUERA As String = "id_chequera"
Public Const CAMPO_ID_BANCO As String = "id_banco"
Public Const CAMPO_ORIGEN As String = "origen"
Public Const CAMPO_EN_CARTERA As String = "en_cartera"
Public Const CAMPO_PROPIO As String = "propio"
Public Const CAMPO_ID_MONEDA As String = "id_moneda"
Public Const CAMPO_OBSERVACIONES As String = "observaciones"
Public Const CAMPO_TERCEROS_PROPIO As String = "teceros_propio"
Public Const TABLA_CHEQUE As String = "cheq"
Public Const TABLA_RECIBO_CHEQUE As String = "admincheq"


Public Function FindAll(Optional ByRef filter As String = vbNullString, Optional ByRef filter2 As String, Optional orderBy As String) As Collection
    On Error GoTo err1

    Dim rs As ADODB.Recordset
    
    Dim q As String
    
    q = "SELECT *, " _
      & "CASE " _
      & "WHEN COALESCE(cheq.movimiento_origen, 0) > 0 THEN " _
      & "COALESCE(NULLIF(CONCAT_WS(' | ', cta_mov.codigo, cta_mov.nombre), ''), " _
      & "'MOVIMIENTO SIN CUENTA CONTABLE') " _
      & "WHEN COALESCE(cheq.liquidacion_caja_origen, 0) > 0 THEN 'PROVEEDORES VARIOS' " _
      & "ELSE COALESCE(cheq.origen, '') " _
      & "END AS destino_mostrado, liq.numero_liq AS numero_liquidacion_caja, " _
      & "rec.fecha AS fecha_rec FROM Cheques cheq" _
      & " LEFT JOIN Chequeras cheqs ON cheqs.id = cheq.id_chequera" _
      & " LEFT JOIN AdminConfigBancos banc ON banc.id = cheq.id_banco" _
      & " LEFT JOIN AdminConfigMonedas mon ON mon.id = cheq.id_moneda" _
      & " LEFT JOIN AdminConfigMonedas mon2 ON mon2.id = cheqs.id_moneda" _
      & " LEFT JOIN AdminConfigBancos banc2 ON banc2.id = cheqs.id_banco" _
      & " LEFT JOIN ordenes_pago op ON op.id = cheq.orden_pago_origen" _
      & " LEFT JOIN liquidaciones_caja liq " _
      & " ON liq.id = cheq.liquidacion_caja_origen" _
      & " LEFT JOIN pagos_a_cuenta pac ON pac.id = cheq.pago_a_cuenta_origen" _
      & " LEFT JOIN movimientos_caja_bancos mov " _
      & " ON mov.id = cheq.movimiento_origen" _
      & " LEFT JOIN AdminComprasCuentasContables cta_mov " _
      & " ON cta_mov.id = mov.id_cuentacontable" _
      & " LEFT JOIN AdminRecibosCheques reccheq ON cheq.id = reccheq.idCheque" _
      & " LEFT JOIN AdminRecibos rec ON reccheq.idRecibo = rec.id" _
      & " WHERE 1 = 1 "

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If
    
    If LenB(filter2) > 0 Then
        q = q & " AND " & filter2
    End If

    If LenB(orderBy) > 0 Then
        q = q & " ORDER BY " & orderBy
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim Cheques As New Collection

    Dim tmpCheque As cheque


    While Not rs.EOF
        Set tmpCheque = DAOCheques.Map(rs, fieldsIndex, TABLA_CHEQUE, "banc", "mon", "cheqs", "mon2", "banc2", "ordenesp", "liq", "facturasp", "prov", "rec", "reccheq")
        Cheques.Add tmpCheque, CStr(tmpCheque.Id)

        rs.MoveNext

    Wend

    Set FindAll = Cheques
    Exit Function

err1:
    Set FindAll = Nothing
End Function



Public Function FindAllTercerosUti(Optional ByRef filter As String = vbNullString, _
                                   Optional ByRef filter2 As String = vbNullString, _
                                   Optional orderBy As String = vbNullString) As Collection
    On Error GoTo ErrorHandler

    Dim rs As ADODB.Recordset
    Dim q As String
    Dim fieldsIndex As Dictionary
    Dim Cheques As New Collection
    Dim tmpCheque As cheque

    ' Construir la consulta SQL
    '----------------------------------------------------------
    ' CONSTRUIR CONSULTA SQL
    '----------------------------------------------------------
    
    q = "SELECT *, "
    q = q & "liq.numero_liq AS numero_liquidacion_caja, "
    
    '----------------------------------------------------------
    ' DESTINO DEL CHEQUE
    '----------------------------------------------------------
    
    q = q & "CASE "
    
    'Cheque depositado en una cuenta bancaria
    q = q & "WHEN IFNULL(cheq.depositado, 0) = 1 "
    q = q & "OR dep.id IS NOT NULL THEN "
    
    q = q & "CONCAT("
    q = q & "'DEPOSITADO EN BANCO ', "
    q = q & "COALESCE("
    q = q & "NULLIF(TRIM(banc_dep.nombre), ''), "
    q = q & "'SIN DEFINIR'"
    q = q & "), "
    
    q = q & "CASE "
    q = q & "WHEN cta_dep.cuenta IS NOT NULL "
    q = q & "AND TRIM(cta_dep.cuenta) <> '' THEN "
    q = q & "CONCAT(' | CTA. ', cta_dep.cuenta) "
    q = q & "ELSE '' "
    q = q & "END"
    
    q = q & ") "
    
    'Cheque utilizado en un movimiento de caja y bancos
    q = q & "WHEN IFNULL(cheq.movimiento_origen, 0) > 0 THEN "
    
    q = q & "COALESCE("
    q = q & "NULLIF("
    q = q & "CONCAT_WS(' | ', cta_mov.codigo, cta_mov.nombre), "
    q = q & "''"
    q = q & "), "
    q = q & "'MOVIMIENTO SIN CUENTA CONTABLE'"
    q = q & ") "
    
    'Cheque utilizado en una liquidación de caja
    q = q & "WHEN IFNULL(cheq.liquidacion_caja_origen, 0) > 0 THEN "
    q = q & "'PROVEEDORES VARIOS' "
    
    'Cheque utilizado en un pago a cuenta
    q = q & "WHEN IFNULL(cheq.pago_a_cuenta_origen, 0) > 0 THEN "
    q = q & "COALESCE(prov_pcta.razon, '') "
    
    'Cheque utilizado en una orden de pago
    q = q & "WHEN IFNULL(cheq.orden_pago_origen, 0) > 0 THEN "
    q = q & "COALESCE(prov.razon, '') "
    
    'Cheque sin destino determinado
    q = q & "ELSE '' "
    
    q = q & "END AS razon_proveedor "
    
    '----------------------------------------------------------
    ' TABLA PRINCIPAL
    '----------------------------------------------------------
    
    q = q & "FROM Cheques cheq "
    
    '----------------------------------------------------------
    ' CHEQUERA, BANCO Y MONEDA
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN Chequeras cheqs "
    q = q & "ON cheqs.id = cheq.id_chequera "
    
    q = q & "LEFT JOIN AdminConfigBancos banc "
    q = q & "ON banc.id = cheq.id_banco "
    
    q = q & "LEFT JOIN AdminConfigMonedas mon "
    q = q & "ON mon.id = cheq.id_moneda "
    
    q = q & "LEFT JOIN AdminConfigMonedas mon2 "
    q = q & "ON mon2.id = cheqs.id_moneda "
    
    q = q & "LEFT JOIN AdminConfigBancos banc2 "
    q = q & "ON banc2.id = cheqs.id_banco "
    
    '----------------------------------------------------------
    ' ORDEN DE PAGO
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN ordenes_pago op "
    q = q & "ON op.id = cheq.orden_pago_origen "
    
    q = q & "LEFT JOIN ordenes_pago_facturas opf "
    q = q & "ON op.id = opf.id_orden_pago "
    
    q = q & "LEFT JOIN AdminComprasFacturasProveedores acfp "
    q = q & "ON acfp.id = opf.id_factura_proveedor "
    
    q = q & "LEFT JOIN proveedores prov "
    q = q & "ON prov.id = acfp.id_proveedor "
    
    '----------------------------------------------------------
    ' LIQUIDACIÓN DE CAJA
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN liquidaciones_caja liq "
    q = q & "ON liq.id = cheq.liquidacion_caja_origen "
    
    '----------------------------------------------------------
    ' PAGO A CUENTA
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN pagos_a_cuenta pcta "
    q = q & "ON pcta.id = cheq.pago_a_cuenta_origen "
    
    q = q & "LEFT JOIN proveedores prov_pcta "
    q = q & "ON prov_pcta.id = pcta.id_proveedor "
    
    '----------------------------------------------------------
    ' MOVIMIENTO DE CAJA Y BANCOS
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN movimientos_caja_bancos mov "
    q = q & "ON mov.id = cheq.movimiento_origen "
    
    q = q & "LEFT JOIN AdminComprasCuentasContables cta_mov "
    q = q & "ON cta_mov.id = mov.id_cuentacontable "
    
    '----------------------------------------------------------
    ' RECIBO DE ORIGEN
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN AdminRecibosCheques admincheq "
    q = q & "ON admincheq.idCheque = cheq.id "
    
    '----------------------------------------------------------
    ' DEPÓSITO DEL CHEQUE
    '----------------------------------------------------------
    
    q = q & "LEFT JOIN cheques_depositos dep "
    q = q & "ON dep.id_cheque = cheq.id "
    
    q = q & "LEFT JOIN boleta_deposito bol_dep "
    q = q & "ON bol_dep.id = dep.id_boleta "
    
    q = q & "LEFT JOIN operaciones op_dep "
    q = q & "ON op_dep.id = dep.id_operacion "
    
    'La cuenta puede provenir de la boleta o de la operación
    q = q & "LEFT JOIN AdminConfigCuentas cta_dep "
    q = q & "ON cta_dep.id = COALESCE("
    q = q & "bol_dep.id_cuenta, "
    q = q & "op_dep.cuentabanc_o_caja_id"
    q = q & ") "
    
    q = q & "LEFT JOIN AdminConfigBancos banc_dep "
    q = q & "ON banc_dep.id = cta_dep.idBanco "
    
    '----------------------------------------------------------
    ' CONDICIÓN BASE
    '----------------------------------------------------------
    
    q = q & "WHERE 1 = 1 "

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    If LenB(filter2) > 0 Then
        q = q & " AND " & filter2
    End If

    If LenB(orderBy) > 0 Then
        q = q & " ORDER BY " & orderBy
    End If

    ' Ejecutar la consulta
    Set rs = conectar.RSFactory(q)

    ' Construir el índice de campos
    BuildFieldsIndex rs, fieldsIndex

    ' Procesar los registros
    While Not rs.EOF
        Set tmpCheque = DAOCheques.Map2(rs, fieldsIndex, TABLA_CHEQUE, "banc", "mon", "cheqs", "mon2", "banc2", "ordenesp", "facturasp", "prov", "admincheq")
        
        ' Verificar si la clave ya existe en la colección
        If Not funciones.BuscarEnColeccion(Cheques, CStr(tmpCheque.Id)) Then
            Cheques.Add tmpCheque, CStr(tmpCheque.Id)
        End If
    
        rs.MoveNext
    Wend

    ' Devolver la colección de cheques
    Set FindAllTercerosUti = Cheques
    Exit Function

ErrorHandler:
    ' Manejo de errores
    Dim errMsg As String
    errMsg = "Error en FindAllTercerosUti: " & vbCrLf & _
             "Número de error: " & Err.Number & vbCrLf & _
             "Descripción: " & Err.Description & vbCrLf

    ' Mostrar el error en un mensaje (opcional)
    MsgBox errMsg, vbCritical, "Error"

    ' Devolver Nothing en caso de error
    Set FindAllTercerosUti = Nothing
End Function




Public Function FindAllDisponiblesByChequera(chequeraId As Long) As Collection
    Set FindAllDisponiblesByChequera = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID_CHEQUERA & "=" & chequeraId & " AND " & TABLA_CHEQUE & "." & DAOCheques.CAMPO_FECHA_VENCIMIENTO & " IS NULL AND " & TABLA_CHEQUE & "." & DAOCheques.CAMPO_EN_CARTERA & " = 0")

End Function

Public Function FindByChequeraAndId(chequeraId As Long, Id As Long) As cheque
    Dim col As Collection
    Set col = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID_CHEQUERA & "=" & chequeraId & " AND " & TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & " = " & Id)
    If col.count = 0 Then
        Set FindByChequeraAndId = Nothing
    Else
        Set FindByChequeraAndId = col.item(1)
    End If

End Function

Public Function FindByChequeraAndNro(chequeraId As Long, NRO As String) As cheque
    Dim col As Collection
    Set col = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID_CHEQUERA & "=" & chequeraId & " AND " & TABLA_CHEQUE & "." & DAOCheques.CAMPO_NUMERO & " = " & Escape(NRO))
    If col.count = 0 Then
        Set FindByChequeraAndNro = Nothing
    Else
        Set FindByChequeraAndNro = col.item(1)
    End If

End Function

Public Function FindById(Id As Long) As cheque
    Dim col As Collection
    Set col = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & "=" & Id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If

End Function

Public Function FindAllByChequeraId(chequeraId As Long, Optional filter2 As String) As Collection
    Set FindAllByChequeraId = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID_CHEQUERA & "=" & chequeraId, filter2)
End Function


Public Function Map(ByRef rs As Recordset, _
                    ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef bancoTableNameOrAlias As String = vbNullString, _
                    Optional ByRef monedaTableNameOrAlias As String = vbNullString, _
                    Optional ByRef chequeraTableNameOrAlias As String = vbNullString, _
                    Optional ByRef monedaChequeraTableNameOrAlias As String = vbNullString, _
                    Optional ByRef bancoChequeraTableNameOrAlias As String = vbNullString, _
                    Optional ByRef OrdenesP As String = vbNullString, _
                    Optional ByRef LiquidacionesC As String = vbNullString, _
                    Optional ByRef FacturasP As String = vbNullString, _
                    Optional ByRef proveedores As String = vbNullString, _
                    Optional ByRef rec As String = vbNullString, _
                    Optional ByRef reccheq As String = vbNullString _
                    ) As cheque

    Dim tmpCheque As cheque
    Dim Id As Variant
    Dim FechaIngreso As Variant
    
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID)

    If Id > 0 Then

        Set tmpCheque = New cheque

        tmpCheque.Observaciones = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_OBSERVACIONES)
        tmpCheque.Id = Id
        tmpCheque.EnCartera = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_EN_CARTERA)
        tmpCheque.FechaRecibido = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_FECHA_RECIBIDO)
        tmpCheque.FechaVencimiento = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_FECHA_VENCIMIENTO)
        tmpCheque.Monto = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_MONTO)
        tmpCheque.numero = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_NUMERO)
        If IsNull(rs.Fields("destino_mostrado").value) Then
            tmpCheque.OrigenDestino = vbNullString
        Else
            tmpCheque.OrigenDestino = _
                CStr(rs.Fields("destino_mostrado").value)
        End If
        tmpCheque.Propio = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_PROPIO)
        tmpCheque.IdChequera = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID_CHEQUERA)
        tmpCheque.TercerosPropio = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_TERCEROS_PROPIO)
        tmpCheque.FechaEmision = GetValue(rs, fieldsIndex, tableNameOrAlias, "fecha_emision")

        tmpCheque.IdOrdenPagoOrigen = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "orden_pago_origen")
            
        tmpCheque.IdLiquidacionCajaOrigen = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "liquidacion_caja_origen")

        If IsNull(rs.Fields("numero_liquidacion_caja").value) Then
            tmpCheque.NumeroLiquidacionCaja = 0
        Else
            tmpCheque.NumeroLiquidacionCaja = _
                CLng(rs.Fields("numero_liquidacion_caja").value)
        End If

        tmpCheque.NumeroPagoACuenta = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "pago_a_cuenta_origen")

        tmpCheque.NumeroMovimiento = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "movimiento_origen")

        tmpCheque.entro = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "ingresado")

        FechaIngreso = GetValue( _
                            rs, _
                            fieldsIndex, _
                            tableNameOrAlias, _
                            "fecha_ingreso_banco")
        
        If IsNull(FechaIngreso) Or IsEmpty(FechaIngreso) Then
            tmpCheque.FechaIngresoBanco = 0
        Else
            tmpCheque.FechaIngresoBanco = CDate(FechaIngreso)
        End If

        tmpCheque.Depositado = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "depositado")

        tmpCheque.estado = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "estado")

        tmpCheque.FechaRecibo = GetValue( _
            rs, fieldsIndex, rec, "fecha_rec")

        If LenB(bancoTableNameOrAlias) > 0 Then
            Set tmpCheque.Banco = DAOBancos.Map( _
                rs, fieldsIndex, bancoTableNameOrAlias)
        End If

        If LenB(monedaTableNameOrAlias) > 0 Then
            Set tmpCheque.moneda = DAOMoneda.Map( _
                rs, fieldsIndex, monedaTableNameOrAlias)
        End If

        If LenB(chequeraTableNameOrAlias) > 0 Then
            Set tmpCheque.chequera = DAOChequeras.Map( _
                rs, _
                fieldsIndex, _
                chequeraTableNameOrAlias, _
                monedaChequeraTableNameOrAlias, _
                bancoChequeraTableNameOrAlias)
        End If

    End If

    Set Map = tmpCheque

End Function


Public Function Map2(ByRef rs As Recordset, _
                     ByRef fieldsIndex As Dictionary, _
                     ByRef tableNameOrAlias As String, _
                     Optional ByRef bancoTableNameOrAlias As String = vbNullString, _
                     Optional ByRef monedaTableNameOrAlias As String = vbNullString, _
                     Optional ByRef chequeraTableNameOrAlias As String = vbNullString, _
                     Optional ByRef monedaChequeraTableNameOrAlias As String = vbNullString, _
                     Optional ByRef bancoChequeraTableNameOrAlias As String = vbNullString, _
                     Optional ByRef OrdenesP As String = vbNullString, _
                     Optional ByRef FacturasP As String = vbNullString, _
                     Optional ByRef proveedores As String = vbNullString, _
                     Optional ByRef recibosChequesTableNameOrAlias As String = vbNullString _
                     ) As cheque

    Dim tmpCheque As cheque
    Dim Id As Variant
    Dim FechaIngreso As Variant
    
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID)

    If Id > 0 Then

        Set tmpCheque = New cheque
        
        tmpCheque.Observaciones = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_OBSERVACIONES)
        tmpCheque.Id = Id
        tmpCheque.EnCartera = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_EN_CARTERA)
        tmpCheque.FechaRecibido = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_FECHA_RECIBIDO)
        tmpCheque.FechaVencimiento = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_FECHA_VENCIMIENTO)
        tmpCheque.Monto = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_MONTO)
        tmpCheque.numero = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_NUMERO)
        tmpCheque.OrigenDestino = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ORIGEN)
        tmpCheque.Propio = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_PROPIO)
        tmpCheque.IdChequera = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID_CHEQUERA)
        tmpCheque.TercerosPropio = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_TERCEROS_PROPIO)
        tmpCheque.FechaEmision = GetValue(rs, fieldsIndex, tableNameOrAlias, "fecha_emision")
        
        tmpCheque.IdOrdenPagoOrigen = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "orden_pago_origen")
        
        tmpCheque.IdLiquidacionCajaOrigen = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "liquidacion_caja_origen")
        
        If IsNull(rs.Fields("numero_liquidacion_caja").value) Then
            tmpCheque.NumeroLiquidacionCaja = 0
        Else
            tmpCheque.NumeroLiquidacionCaja = _
                CLng(rs.Fields("numero_liquidacion_caja").value)
        End If
       
        tmpCheque.NumeroPagoACuenta = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "pago_a_cuenta_origen")

        tmpCheque.NumeroMovimiento = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "movimiento_origen")
        
        tmpCheque.entro = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "ingresado")

        FechaIngreso = GetValue( _
                            rs, _
                            fieldsIndex, _
                            tableNameOrAlias, _
                            "fecha_ingreso_banco")
        
        If IsNull(FechaIngreso) Or IsEmpty(FechaIngreso) Then
            tmpCheque.FechaIngresoBanco = 0
        Else
            tmpCheque.FechaIngresoBanco = CDate(FechaIngreso)
        End If

        tmpCheque.Depositado = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "depositado")

        tmpCheque.estado = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "estado")
        
        If Not IsNull(rs.Fields("razon_proveedor").value) Then
            tmpCheque.destino = rs.Fields("razon_proveedor").value
        Else
            tmpCheque.destino = vbNullString
        End If
        
        tmpCheque.Recibo = GetValue( _
            rs, fieldsIndex, recibosChequesTableNameOrAlias, "idRecibo")
        
        If LenB(bancoTableNameOrAlias) > 0 Then
            Set tmpCheque.Banco = DAOBancos.Map( _
                rs, fieldsIndex, bancoTableNameOrAlias)
        End If

        If LenB(monedaTableNameOrAlias) > 0 Then
            Set tmpCheque.moneda = DAOMoneda.Map( _
                rs, fieldsIndex, monedaTableNameOrAlias)
        End If

        If LenB(chequeraTableNameOrAlias) > 0 Then
            Set tmpCheque.chequera = DAOChequeras.Map( _
                rs, _
                fieldsIndex, _
                chequeraTableNameOrAlias, _
                monedaChequeraTableNameOrAlias, _
                bancoChequeraTableNameOrAlias)
        End If

    End If

    Set Map2 = tmpCheque

End Function

Public Function Map3(ByRef rs As Recordset, _
                     ByRef fieldsIndex As Dictionary, _
                     ByRef tableNameOrAlias As String, _
                     Optional ByRef bancoTableNameOrAlias As String = vbNullString, _
                     Optional ByRef monedaTableNameOrAlias As String = vbNullString, _
                     Optional ByRef chequeraTableNameOrAlias As String = vbNullString, _
                     Optional ByRef monedaChequeraTableNameOrAlias As String = vbNullString, _
                     Optional ByRef bancoChequeraTableNameOrAlias As String = vbNullString, _
                     Optional ByRef OrdenesP As String = vbNullString, _
                     Optional ByRef LiquidacionesC As String = vbNullString, _
                     Optional ByRef FacturasP As String = vbNullString, _
                     Optional ByRef proveedores As String = vbNullString, _
                     Optional ByRef rec As String = vbNullString, _
                     Optional ByRef reccheq As String = vbNullString _
                     ) As cheque

    Dim tmpCheque As cheque
    Dim Id As Variant
    Dim FechaIngreso As Variant
    
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID)

    If Id > 0 Then

        Set tmpCheque = New cheque

        tmpCheque.Observaciones = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_OBSERVACIONES)

        tmpCheque.Id = Id

        tmpCheque.EnCartera = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_EN_CARTERA)

        tmpCheque.FechaRecibido = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_FECHA_RECIBIDO)

        tmpCheque.FechaVencimiento = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_FECHA_VENCIMIENTO)

        tmpCheque.Monto = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_MONTO)

        tmpCheque.numero = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_NUMERO)

        tmpCheque.OrigenDestino = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ORIGEN)

        tmpCheque.Propio = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_PROPIO)

        tmpCheque.IdChequera = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID_CHEQUERA)

        tmpCheque.TercerosPropio = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_TERCEROS_PROPIO)

        tmpCheque.FechaEmision = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "fecha_emision")

        tmpCheque.IdOrdenPagoOrigen = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "orden_pago_origen")
        
        tmpCheque.IdLiquidacionCajaOrigen = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "liquidacion_caja_origen")
        
        If LenB(LiquidacionesC) > 0 Then
            tmpCheque.NumeroLiquidacionCaja = GetValue( _
                rs, fieldsIndex, LiquidacionesC, "numero_liq")
        Else
            tmpCheque.NumeroLiquidacionCaja = 0
        End If

        tmpCheque.NumeroPagoACuenta = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "pago_a_cuenta_origen")

        tmpCheque.NumeroMovimiento = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "movimiento_origen")

        tmpCheque.entro = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "ingresado")

        FechaIngreso = GetValue( _
                            rs, _
                            fieldsIndex, _
                            tableNameOrAlias, _
                            "fecha_ingreso_banco")
        
        If IsNull(FechaIngreso) Or IsEmpty(FechaIngreso) Then
            tmpCheque.FechaIngresoBanco = 0
        Else
            tmpCheque.FechaIngresoBanco = CDate(FechaIngreso)
        End If

        tmpCheque.Depositado = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "depositado")

        tmpCheque.estado = GetValue( _
            rs, fieldsIndex, tableNameOrAlias, "estado")

        If LenB(bancoTableNameOrAlias) > 0 Then
            Set tmpCheque.Banco = DAOBancos.Map( _
                rs, fieldsIndex, bancoTableNameOrAlias)
        End If

        If LenB(monedaTableNameOrAlias) > 0 Then
            Set tmpCheque.moneda = DAOMoneda.Map( _
                rs, fieldsIndex, monedaTableNameOrAlias)
        End If

        If LenB(chequeraTableNameOrAlias) > 0 Then
            Set tmpCheque.chequera = DAOChequeras.Map( _
                rs, _
                fieldsIndex, _
                chequeraTableNameOrAlias, _
                monedaChequeraTableNameOrAlias, _
                bancoChequeraTableNameOrAlias)
        End If

    End If

    Set Map3 = tmpCheque

End Function


Public Function Guardar(cheque As cheque) As Boolean
    Dim q As String
    
    If Not PuedeModificarIngresoBancoCheque(cheque) Then
        Guardar = False
        Exit Function
    End If

    If cheque.Id = 0 Then
        q = "INSERT INTO Cheques" _
          & "(numero," _
          & "fecha_recibido," _
          & "fecha_vencimiento," _
          & "monto," _
          & "id_chequera," _
          & "id_banco," _
          & "origen," _
          & "en_cartera," _
          & "propio," _
          & "id_moneda," _
          & "observaciones, teceros_propio, ingresado, fecha_emision, orden_pago_origen, liquidacion_caja_origen, pago_a_cuenta_origen, movimiento_origen, depositado" _
          & ") Values " _
          & "('numero'," _
          & "'fecha_recibido'," _
          & "'fecha_vencimiento'," _
          & "'monto'," _
          & "'id_chequera'," _
          & "'id_banco'," _
          & "'origen'," _
          & "'en_cartera'," _
          & "'propio'," _
          & "'id_moneda'," _
          & "'observaciones','teceros_propio','ingresado','fecha_emision','orden_pago_origen','liquidacion_caja_origen', 'pago_a_cuenta_origen', 'movimiento_origen', 'depositado' " _
          & ")"

    Else

        q = "UPDATE Cheques" _
          & " SET " _
          & "numero = 'numero' , " _
          & "fecha_recibido = 'fecha_recibido' ," _
          & "fecha_vencimiento = 'fecha_vencimiento' ," _
          & "monto = 'monto' ," _
          & "id_chequera = 'id_chequera' ," _
          & "id_banco = 'id_banco' ," _
          & "origen = 'origen' ," _
          & "en_cartera = 'en_cartera' ," _
          & "propio = 'propio' ," _
          & "id_moneda = 'id_moneda' ," _
          & "observaciones = 'observaciones' ," _
          & "teceros_propio='teceros_propio', " _
          & "ingresado='ingresado', " _
          & "fecha_emision='fecha_emision', " _
          & "orden_pago_origen = 'orden_pago_origen', " _
          & "liquidacion_caja_origen = 'liquidacion_caja_origen', " _
          & "pago_a_cuenta_origen = 'pago_a_cuenta_origen', " _
          & "movimiento_origen = 'movimiento_origen', " _
          & "estado = 'estado', " _
          & "depositado = 'depositado' " _
          & " Where " _
          & "id = 'id' " _

q = Replace(q, "'id'", cheque.Id)
    End If


    q = Replace(q, "'numero'", conectar.Escape(cheque.numero))
    q = Replace(q, "'fecha_recibido'", conectar.Escape(cheque.FechaRecibido))
    q = Replace(q, "'fecha_vencimiento'", conectar.Escape(cheque.FechaVencimiento))
    q = Replace(q, "'monto'", conectar.Escape(cheque.Monto))
    q = Replace(q, "'id_chequera'", conectar.Escape(cheque.IdChequera))
    q = Replace(q, "'id_banco'", conectar.Escape(cheque.Banco.Id))
    q = Replace(q, "'origen'", conectar.Escape(cheque.OrigenDestino))
    q = Replace(q, "'en_cartera'", conectar.Escape(cheque.EnCartera))
    q = Replace(q, "'propio'", conectar.Escape(cheque.Propio))
    q = Replace(q, "'id_moneda'", conectar.Escape(cheque.moneda.Id))
    q = Replace(q, "'observaciones'", conectar.Escape(cheque.Observaciones))
    q = Replace(q, "'teceros_propio'", conectar.Escape(cheque.TercerosPropio))
    q = Replace(q, "'ingresado'", conectar.Escape(Abs(cheque.entro)))
    q = Replace(q, "'orden_pago_origen'", conectar.Escape(cheque.IdOrdenPagoOrigen))
    q = Replace(q, "'pago_a_cuenta_origen'", conectar.Escape(cheque.NumeroPagoACuenta))
    q = Replace(q, "'movimiento_origen'", conectar.Escape(cheque.NumeroMovimiento))
    q = Replace(q, "'liquidacion_caja_origen'", conectar.Escape(cheque.IdLiquidacionCajaOrigen))
    q = Replace(q, "'estado'", conectar.Escape(cheque.estado))
    q = Replace(q, "'depositado'", conectar.Escape(cheque.Depositado))
    
    Dim sqlFechaEmision As String

    If CDbl(cheque.FechaEmision) > 0 Then
    
        sqlFechaEmision = conectar.Escape( _
            Format$(cheque.FechaEmision, "yyyy-mm-dd"))
    
    Else
    
        sqlFechaEmision = "NULL"
    
    End If
    
    q = Replace(q, "'fecha_emision'", sqlFechaEmision)
    
    Guardar = conectar.execute(q)
    If Not Guardar Then Exit Function

    If cheque.Id = 0 Then
        Dim idche As Long
        Guardar = conectar.UltimoId("Cheques", idche)
        cheque.Id = idche
    End If

End Function

Public Function FindAllEnCartera(Optional ByRef filter2 As String, Optional ByRef orderBy As String) As Collection
    Set FindAllEnCartera = FindAll(DAOCheques.CAMPO_EN_CARTERA & " = 1", filter2, orderBy)
End Function


Public Function FindAllEnCarteraDeTerceros() As Collection
    Set FindAllEnCarteraDeTerceros = FindAll(DAOCheques.CAMPO_EN_CARTERA & " = 1 and " & DAOCheques.CAMPO_PROPIO & " = 0")
End Function


Public Function AnularCheque(ByVal idCheque As Long) As Boolean

    On Error GoTo err1

    Dim q As String
    Dim rsVerificacion As ADODB.Recordset
    Dim che As cheque

    AnularCheque = False

    If idCheque <= 0 Then Exit Function

    '======================================================
    ' CARGAR EL CHEQUE ACTUAL
    '======================================================
    Set che = FindById(idCheque)

    If Not IsSomething(che) Then

        MsgBox "No se encontró el cheque que se intenta anular.", _
               vbExclamation, _
               "Anular cheque"

        Exit Function

    End If

    '======================================================
    ' VALIDAR CONCILIACION BANCARIA CERRADA
    '
    ' Si el cheque ya ingresó al banco y esa fecha
    ' pertenece a una conciliación cerrada, no se puede
    ' eliminar ese impacto bancario.
    '======================================================
    If Not PuedeModificarIngresoBancoCheque(che) Then
        Exit Function
    End If

    '======================================================
    ' ANULAR
    '======================================================
    q = "UPDATE Cheques SET " & _
        "estado = " & CLng(ChequeAnulado) & ", " & _
        "en_cartera = 0, " & _
        "ingresado = 0, " & _
        "fecha_ingreso_banco = NULL, " & _
        "depositado = 0 " & _
        "WHERE id = " & idCheque & " " & _
        "AND propio = 1 " & _
        "AND IFNULL(estado, 1) <> " & _
            CLng(ChequeAnulado) & " " & _
        "AND IFNULL(orden_pago_origen, 0) = 0 " & _
        "AND IFNULL(liquidacion_caja_origen, 0) = 0 " & _
        "AND IFNULL(pago_a_cuenta_origen, 0) = 0 " & _
        "AND IFNULL(movimiento_origen, 0) = 0 " & _
        "AND NOT EXISTS (" & _
            "SELECT 1 " & _
            "FROM ordenes_pago_cheques opc " & _
            "WHERE opc.id_cheque = Cheques.id" & _
        ")"

    If Not conectar.execute(q) Then Exit Function

    '======================================================
    ' VERIFICAR QUE REALMENTE HAYA QUEDADO ANULADO
    '======================================================
    Set rsVerificacion = conectar.RSFactory( _
        "SELECT estado FROM Cheques " & _
        "WHERE id = " & idCheque)

    If rsVerificacion Is Nothing Then Exit Function
    If rsVerificacion.EOF Then Exit Function
    If IsNull(rsVerificacion!estado) Then Exit Function

    AnularCheque = _
        (CLng(rsVerificacion!estado) = _
         CLng(ChequeAnulado))

    Exit Function


err1:

    AnularCheque = False

    MsgBox "No se pudo anular el cheque." & _
           vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbCritical, _
           "Anular cheque"

End Function


Public Function ActualizarIngresoBanco( _
    ByVal idCheque As Long, _
    ByVal Ingresado As Boolean, _
    ByVal FechaIngresoBanco As Variant _
) As Boolean

    On Error GoTo err1

    Dim che As cheque
    Dim q As String
    Dim sqlFecha As String

    ActualizarIngresoBanco = False

    Set che = FindById(idCheque)

    If Not IsSomething(che) Then Exit Function

    '------------------------------------------------------
    ' PREPARAR NUEVO ESTADO
    '------------------------------------------------------
    che.entro = Ingresado

    If Ingresado Then

        If Not IsDate(FechaIngresoBanco) Then

            MsgBox "Debe indicar una fecha de ingreso válida.", _
                   vbExclamation, _
                   "Ingreso bancario"

            Exit Function

        End If

        If CDbl(CDate(FechaIngresoBanco)) <= 0 Then

            MsgBox "Debe indicar una fecha de ingreso válida.", _
                   vbExclamation, _
                   "Ingreso bancario"

            Exit Function

        End If

        che.FechaIngresoBanco = _
            CDate(FechaIngresoBanco)

        sqlFecha = _
            conectar.Escape( _
                Format$( _
                    che.FechaIngresoBanco, _
                    "yyyy-mm-dd"))

    Else

        che.FechaIngresoBanco = 0
        sqlFecha = "NULL"

    End If

    '------------------------------------------------------
    ' VALIDAR CONCILIACION CERRADA
    '------------------------------------------------------
    If Not PuedeModificarIngresoBancoCheque(che) Then
        Exit Function
    End If

    '------------------------------------------------------
    ' ACTUALIZAR
    '------------------------------------------------------
    q = "UPDATE Cheques SET " _
      & "ingresado = " & Abs(CInt(Ingresado)) & ", " _
      & "fecha_ingreso_banco = " & sqlFecha & " " _
      & "WHERE id = " & idCheque

    ActualizarIngresoBanco = _
        conectar.execute(q)

    Exit Function

err1:

    ActualizarIngresoBanco = False

End Function


Public Function FindAllPropiosConciliacion( _
    Optional ByRef filter As String = vbNullString, _
    Optional ByVal mostrarIngresados As Boolean = False) As Collection

    On Error GoTo err1

    Dim rs As ADODB.Recordset
    Dim resultado As New Collection
    Dim tmpCheque As cheque
    Dim q As String

    q = "SELECT "
    q = q & "cheq.id AS cheque_id, "
    q = q & "cheq.numero AS cheque_numero, "
    q = q & "cheq.monto AS cheque_monto, "
    q = q & "cheq.fecha_vencimiento AS cheque_vencimiento, "
    q = q & "cheq.fecha_emision AS cheque_emision, "
    q = q & "cheq.origen AS cheque_origen, "
    q = q & "cheq.ingresado AS cheque_ingresado, "
    q = q & "cheq.fecha_ingreso_banco AS cheque_fecha_ingreso, "
    q = q & "cheq.estado AS cheque_estado, "
    q = q & "cheq.orden_pago_origen AS cheque_op, "
    q = q & "cheq.liquidacion_caja_origen AS cheque_liquidacion, "
    q = q & "cheq.pago_a_cuenta_origen AS cheque_pago_cuenta, "
    q = q & "cheq.movimiento_origen AS cheque_movimiento, "
    q = q & "cheqs.id AS chequera_id, "
    q = q & "cheqs.numero AS chequera_numero, "
    q = q & "banco.id AS banco_id, "
    q = q & "banco.Nombre AS banco_nombre, "
    q = q & "cta.id AS cuenta_id, "
    q = q & "cta.cuenta AS cuenta_numero, "
    q = q & "liq.numero_liq AS numero_liquidacion "

    q = q & "FROM Cheques cheq "

    q = q & "INNER JOIN Chequeras cheqs "
    q = q & "ON cheqs.id = cheq.id_chequera "

    q = q & "LEFT JOIN AdminConfigBancos banco "
    q = q & "ON banco.id = cheqs.id_banco "

    q = q & "LEFT JOIN AdminConfigCuentas cta "
    q = q & "ON cta.id = cheqs.id_cuenta_bancaria "

    q = q & "LEFT JOIN liquidaciones_caja liq "
    q = q & "ON liq.id = cheq.liquidacion_caja_origen "

    q = q & "WHERE cheq.propio = 1 "
    q = q & "AND cheq.id_chequera > 0 "

    'Solamente cheques que realmente fueron utilizados
    q = q & "AND ("
    q = q & "COALESCE(cheq.orden_pago_origen, 0) > 0 "
    q = q & "OR COALESCE(cheq.liquidacion_caja_origen, 0) > 0 "
    q = q & "OR COALESCE(cheq.pago_a_cuenta_origen, 0) > 0 "
    q = q & "OR COALESCE(cheq.movimiento_origen, 0) > 0"
    q = q & ") "

    If Not mostrarIngresados Then
        q = q & "AND COALESCE(cheq.ingresado, 0) = 0 "
    End If

    If LenB(filter) > 0 Then
        q = q & "AND (" & filter & ") "
    End If

    q = q & "ORDER BY "
    q = q & "banco.Nombre, "
    q = q & "cta.cuenta, "
    q = q & "cheqs.numero, "
    q = q & "CAST(cheq.numero AS UNSIGNED), "
    q = q & "cheq.id"

    Set rs = conectar.RSFactory(q)

    While Not rs.EOF

        Set tmpCheque = New cheque

        tmpCheque.Id = CLng(rs.Fields("cheque_id").value)
        tmpCheque.numero = CStr(rs.Fields("cheque_numero").value)
        tmpCheque.Propio = True

        If Not IsNull(rs.Fields("cheque_monto").value) Then
            tmpCheque.Monto = CDbl( _
                rs.Fields("cheque_monto").value)
        End If

        If Not IsNull(rs.Fields("cheque_vencimiento").value) Then
            tmpCheque.FechaVencimiento = CDate( _
                rs.Fields("cheque_vencimiento").value)
        End If

        If Not IsNull(rs.Fields("cheque_emision").value) Then
            tmpCheque.FechaEmision = CDate( _
                rs.Fields("cheque_emision").value)
        End If

        If Not IsNull(rs.Fields("cheque_origen").value) Then
            tmpCheque.OrigenDestino = CStr( _
                rs.Fields("cheque_origen").value)
        End If

        If Not IsNull(rs.Fields("cheque_ingresado").value) Then
            tmpCheque.entro = CBool( _
                rs.Fields("cheque_ingresado").value)
        End If

        If Not IsNull(rs.Fields("cheque_fecha_ingreso").value) Then
            tmpCheque.FechaIngresoBanco = CDate( _
                rs.Fields("cheque_fecha_ingreso").value)
        End If

        If Not IsNull(rs.Fields("cheque_estado").value) Then
            tmpCheque.estado = CLng( _
                rs.Fields("cheque_estado").value)
        End If

        If Not IsNull(rs.Fields("cheque_op").value) Then
            tmpCheque.IdOrdenPagoOrigen = CLng( _
                rs.Fields("cheque_op").value)
        End If

        If Not IsNull(rs.Fields("cheque_liquidacion").value) Then
            tmpCheque.IdLiquidacionCajaOrigen = CLng( _
                rs.Fields("cheque_liquidacion").value)
        End If

        If Not IsNull(rs.Fields("numero_liquidacion").value) Then
            tmpCheque.NumeroLiquidacionCaja = CLng( _
                rs.Fields("numero_liquidacion").value)
        End If

        If Not IsNull(rs.Fields("cheque_pago_cuenta").value) Then
            tmpCheque.NumeroPagoACuenta = CLng( _
                rs.Fields("cheque_pago_cuenta").value)
        End If

        If Not IsNull(rs.Fields("cheque_movimiento").value) Then
            tmpCheque.NumeroMovimiento = CLng( _
                rs.Fields("cheque_movimiento").value)
        End If

        Set tmpCheque.chequera = New chequera

        tmpCheque.chequera.Id = CLng( _
            rs.Fields("chequera_id").value)

        tmpCheque.IdChequera = tmpCheque.chequera.Id

        If Not IsNull(rs.Fields("chequera_numero").value) Then
            tmpCheque.chequera.numero = CLng( _
                rs.Fields("chequera_numero").value)
        End If

        If Not IsNull(rs.Fields("banco_id").value) Then

            Set tmpCheque.chequera.Banco = New Banco

            tmpCheque.chequera.Banco.Id = CLng( _
                rs.Fields("banco_id").value)

            If Not IsNull(rs.Fields("banco_nombre").value) Then
                tmpCheque.chequera.Banco.nombre = CStr( _
                    rs.Fields("banco_nombre").value)
            End If

        End If

        If Not IsNull(rs.Fields("cuenta_id").value) Then

            Set tmpCheque.chequera.CuentaBancaria = _
                New CuentaBancaria

            tmpCheque.chequera.CuentaBancaria.Id = CLng( _
                rs.Fields("cuenta_id").value)

            If Not IsNull(rs.Fields("cuenta_numero").value) Then
                tmpCheque.chequera.CuentaBancaria.numero = CStr( _
                    rs.Fields("cuenta_numero").value)
            End If

        End If

        resultado.Add tmpCheque, CStr(tmpCheque.Id)

        rs.MoveNext

    Wend

    Set FindAllPropiosConciliacion = resultado
    Exit Function

err1:
    Debug.Print "FindAllPropiosConciliacion: " & _
                Err.Number & " - " & Err.Description

    Set FindAllPropiosConciliacion = Nothing

End Function


Public Function ActualizarIngresosBancoLote( _
    ByVal idsCheques As Collection, _
    ByVal FechaIngreso As Date) As Boolean

    On Error GoTo err1

    Dim q As String
    Dim listaIds As String
    Dim valor As Variant

    Dim che As cheque
    Dim idCheque As Long

    ActualizarIngresosBancoLote = False

    '------------------------------------------------------
    ' VALIDACIONES GENERALES
    '------------------------------------------------------
    If idsCheques Is Nothing Then Exit Function

    If idsCheques.count = 0 Then Exit Function

    If CDbl(FechaIngreso) <= 0 Then

        MsgBox "Debe indicar una fecha de ingreso válida.", _
               vbExclamation, _
               "Conciliación de cheques"

        Exit Function

    End If

    If FechaIngreso > Date Then

        MsgBox "La fecha de ingreso no puede ser posterior a hoy.", _
               vbExclamation, _
               "Conciliación de cheques"

        Exit Function

    End If

    listaIds = vbNullString

    '======================================================
    ' PRIMERA ETAPA
    '
    ' VALIDAR TODOS LOS CHEQUES.
    '
    ' IMPORTANTE:
    ' todavía NO modificamos nada en la base.
    ' Si uno falla, no se modifica ninguno.
    '======================================================
    For Each valor In idsCheques

        If Not IsNumeric(valor) Then

            MsgBox "Se encontró un identificador de cheque inválido." & _
                   vbCrLf & vbCrLf & _
                   "No se realizó ninguna modificación.", _
                   vbExclamation, _
                   "Conciliación de cheques"

            Exit Function

        End If

        idCheque = CLng(valor)

        If idCheque <= 0 Then

            MsgBox "Se encontró un identificador de cheque inválido." & _
                   vbCrLf & vbCrLf & _
                   "No se realizó ninguna modificación.", _
                   vbExclamation, _
                   "Conciliación de cheques"

            Exit Function

        End If

        Set che = FindById(idCheque)

        If Not IsSomething(che) Then

            MsgBox "No se encontró el cheque ID " & _
                   idCheque & "." & _
                   vbCrLf & vbCrLf & _
                   "No se realizó ninguna modificación.", _
                   vbExclamation, _
                   "Conciliación de cheques"

            Exit Function

        End If

        '--------------------------------------------------
        ' SOLO CHEQUES PROPIOS
        '--------------------------------------------------
        If Not che.Propio Then

            MsgBox "El cheque N° " & che.numero & _
                   " no es un cheque propio." & _
                   vbCrLf & vbCrLf & _
                   "No se realizó ninguna modificación.", _
                   vbExclamation, _
                   "Conciliación de cheques"

            Exit Function

        End If

        '--------------------------------------------------
        ' NO DEBE ESTAR YA INGRESADO
        '--------------------------------------------------
        If che.entro Then

            MsgBox "El cheque N° " & che.numero & _
                   " ya se encuentra ingresado al banco." & _
                   vbCrLf & vbCrLf & _
                   "No se realizó ninguna modificación.", _
                   vbExclamation, _
                   "Conciliación de cheques"

            Exit Function

        End If

        '--------------------------------------------------
        ' SIMULAR EL NUEVO ESTADO
        '--------------------------------------------------
        che.entro = True
        che.FechaIngresoBanco = FechaIngreso

        '--------------------------------------------------
        ' MISMA VALIDACION QUE EL INGRESO INDIVIDUAL
        '--------------------------------------------------
        If Not PuedeModificarIngresoBancoCheque(che) Then

            'PuedeModificarIngresoBancoCheque ya muestra
            'el motivo exacto del bloqueo.
            Exit Function

        End If

        '--------------------------------------------------
        ' AGREGAR A LA LISTA SQL
        '--------------------------------------------------
        If LenB(listaIds) > 0 Then
            listaIds = listaIds & ","
        End If

        listaIds = listaIds & CStr(idCheque)

    Next valor

    If LenB(listaIds) = 0 Then Exit Function


    '======================================================
    ' SEGUNDA ETAPA
    '
    ' TODOS PASARON LAS VALIDACIONES.
    ' AHORA SI MODIFICAMOS TODOS JUNTOS.
    '======================================================
    q = "UPDATE Cheques SET " _
      & "ingresado = 1, " _
      & "fecha_ingreso_banco = " _
      & conectar.Escape( _
            Format$(FechaIngreso, "yyyy-mm-dd")) & " " _
      & "WHERE id IN (" & listaIds & ") " _
      & "AND propio = 1 " _
      & "AND (ingresado IS NULL OR ingresado = 0)"

    ActualizarIngresosBancoLote = _
        conectar.execute(q)

    Exit Function


err1:

    ActualizarIngresosBancoLote = False

    MsgBox "No se pudieron actualizar los cheques." & _
           vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbCritical, _
           "Conciliación de cheques"

End Function


Private Function ObtenerIdCuentaBancariaChequera(ByVal IdChequera As Long) As Long

    On Error GoTo err1

    Dim chq As chequera

    ObtenerIdCuentaBancariaChequera = 0

    If IdChequera <= 0 Then Exit Function

    Set chq = DAOChequeras.GetById(IdChequera)

    If Not IsSomething(chq) Then Exit Function
    If Not IsSomething(chq.CuentaBancaria) Then Exit Function

    ObtenerIdCuentaBancariaChequera = chq.CuentaBancaria.Id
    Exit Function

err1:
    ObtenerIdCuentaBancariaChequera = 0

End Function


Private Function PuedeModificarIngresoBancoCheque( _
    ByRef che As cheque _
) As Boolean

    On Error GoTo err1

    Dim chequeActual As cheque
    Dim idCuenta As Long
    Dim IdConciliacion As Long

    PuedeModificarIngresoBancoCheque = False

    '======================================================
    ' 1. PROTEGER EL ESTADO HISTORICO ACTUAL
    '======================================================
    If che.Id > 0 Then

        Set chequeActual = FindById(che.Id)

        If IsSomething(chequeActual) Then

            If chequeActual.Propio Then

               If chequeActual.entro And _
                    CDbl(chequeActual.FechaIngresoBanco) > 0 Then

                    idCuenta = ObtenerIdCuentaBancariaChequera( _
                                    chequeActual.IdChequera)

                    If idCuenta > 0 Then

                        IdConciliacion = _
                            DAOConciliacionBancaria. _
                                ObtenerIdConciliacionCerrada( _
                                    idCuenta, _
                                    chequeActual.FechaIngresoBanco)

                        If IdConciliacion > 0 Then

                            MsgBox _
                                "No se puede modificar el cheque N° " & _
                                chequeActual.numero & "." & vbCrLf & _
                                vbCrLf & _
                                "La fecha de ingreso bancario actual (" & _
                                Format$( _
                                    chequeActual.FechaIngresoBanco, _
                                    "dd/mm/yyyy") & _
                                ") pertenece a la Conciliación " & _
                                "Bancaria Nro " & IdConciliacion & "." & _
                                vbCrLf & vbCrLf & _
                                "Los cheques ya incluidos en un período " & _
                                "cerrado no pueden modificarse.", _
                                vbExclamation, _
                                "Período bancario cerrado"

                            Exit Function

                        End If

                    End If

                End If

            End If

        End If

    End If

    '======================================================
    ' 2. VALIDAR EL NUEVO VALOR A GUARDAR
    '======================================================
    If che.Propio Then
    
        If che.entro And _
           CDbl(che.FechaIngresoBanco) > 0 Then

            idCuenta = ObtenerIdCuentaBancariaChequera(che.IdChequera)

            If idCuenta <= 0 Then

                MsgBox _
                    "La chequera del cheque N° " & che.numero & _
                    " no tiene asociada una cuenta bancaria." & _
                    vbCrLf & vbCrLf & _
                    "Debe asignar la cuenta bancaria a la chequera " & _
                    "antes de registrar el ingreso al banco.", _
                    vbExclamation, _
                    "Chequera sin cuenta bancaria"

                Exit Function

            End If

            IdConciliacion = _
                DAOConciliacionBancaria. _
                    ObtenerIdConciliacionCerrada( _
                        idCuenta, _
                        che.FechaIngresoBanco)

            If IdConciliacion > 0 Then

                MsgBox _
                    "No se puede registrar el ingreso bancario del " & _
                    "cheque N° " & che.numero & "." & vbCrLf & _
                    vbCrLf & _
                    "Fecha: " & _
                    Format$(che.FechaIngresoBanco, "dd/mm/yyyy") & _
                    vbCrLf & _
                    "La cuenta bancaria de la chequera se encuentra " & _
                    "cerrada por la Conciliación Bancaria Nro " & _
                    IdConciliacion & ".", _
                    vbExclamation, _
                    "Período bancario cerrado"

                Exit Function

            End If

        End If

    End If

    PuedeModificarIngresoBancoCheque = True
    Exit Function

err1:

    PuedeModificarIngresoBancoCheque = False

    MsgBox _
        "No se pudo validar el ingreso bancario del cheque." & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "Conciliación bancaria"

End Function
