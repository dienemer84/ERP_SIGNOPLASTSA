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
    
    q = "SELECT *, rec.fecha AS fecha_rec" _
      & " FROM Cheques cheq" _
      & " LEFT JOIN Chequeras cheqs ON cheqs.id = cheq.id_chequera" _
      & " LEFT JOIN AdminConfigBancos banc ON banc.id = cheq.id_banco" _
      & " LEFT JOIN AdminConfigMonedas mon ON mon.id = cheq.id_moneda" _
      & " LEFT JOIN AdminConfigMonedas mon2 ON mon2.id = cheqs.id_moneda" _
      & " LEFT JOIN AdminConfigBancos banc2 ON banc2.id = cheqs.id_banco" _
      & " LEFT JOIN ordenes_pago op ON op.id = cheq.orden_pago_origen" _
      & " LEFT JOIN liquidaciones_caja liq ON liq.id = cheq.liquidacion_caja_origen" _
      & " LEFT JOIN pagos_a_cuenta pac ON pac.id = cheq.pago_a_cuenta_origen" _
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
    q = "SELECT *" _
      & " FROM Cheques cheq" _
      & " LEFT JOIN Chequeras cheqs ON cheqs.id = cheq.id_chequera" _
      & " LEFT JOIN AdminConfigBancos banc ON banc.id = cheq.id_banco" _
      & " LEFT JOIN AdminConfigMonedas mon ON mon.id = cheq.id_moneda" _
      & " LEFT JOIN AdminConfigMonedas mon2 ON mon2.id = cheqs.id_moneda" _
      & " LEFT JOIN AdminConfigBancos banc2 ON banc2.id = cheqs.id_banco" _
      & " LEFT JOIN ordenes_pago op ON op.id = cheq.orden_pago_origen" _
      & " LEFT JOIN liquidaciones_caja liq ON liq.id = cheq.orden_pago_origen" _
      & " LEFT JOIN ordenes_pago_facturas opf ON op.id = opf.id_orden_pago" _
      & " LEFT JOIN AdminComprasFacturasProveedores acfp ON acfp.id = opf.id_factura_proveedor" _
      & " LEFT JOIN proveedores prov ON prov.id = acfp.id_proveedor" _
      & " LEFT JOIN AdminRecibosCheques admincheq ON admincheq.idCheque = cheq.id" _
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
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID)

    If Id > 0 Then
        Set tmpCheque = New cheque
        tmpCheque.observaciones = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_OBSERVACIONES)
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
        tmpCheque.IdOrdenPagoOrigen = GetValue(rs, fieldsIndex, tableNameOrAlias, "orden_pago_origen")
        tmpCheque.NumeroLiquidacionCaja = GetValue(rs, fieldsIndex, tableNameOrAlias, "liquidacion_caja_origen")
        tmpCheque.NumeroPagoACuenta = GetValue(rs, fieldsIndex, tableNameOrAlias, "pago_a_cuenta_origen")
        tmpCheque.entro = GetValue(rs, fieldsIndex, tableNameOrAlias, "ingresado")
        tmpCheque.Depositado = GetValue(rs, fieldsIndex, tableNameOrAlias, "depositado")
        tmpCheque.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, "estado")

        tmpCheque.FechaRecibo = GetValue(rs, fieldsIndex, rec, "fecha_rec")

        
        ' Verificar si la tabla existe y obtener el valor de "numero_liq"
        On Error Resume Next ' Ignorar errores temporalmente
        tmpCheque.NumeroLiquidacionCaja = GetValue(rs, fieldsIndex, LiquidacionesC, "numero_liq")
        If Err.Number <> 0 Then
            ' Si hay un error, el campo no existe
            tmpCheque.NumeroLiquidacionCaja = "" ' O un valor por defecto
            Err.Clear ' Limpiar el error
        End If
        On Error GoTo 0 ' Restaurar el manejo de errores
        
        
                
        If LenB(bancoTableNameOrAlias) > 0 Then Set tmpCheque.Banco = DAOBancos.Map(rs, fieldsIndex, bancoTableNameOrAlias)
        If LenB(monedaTableNameOrAlias) > 0 Then Set tmpCheque.moneda = DAOMoneda.Map(rs, fieldsIndex, monedaTableNameOrAlias)
        If LenB(chequeraTableNameOrAlias) > 0 Then Set tmpCheque.chequera = DAOChequeras.Map(rs, fieldsIndex, chequeraTableNameOrAlias, monedaChequeraTableNameOrAlias, bancoChequeraTableNameOrAlias)
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
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID)

    If Id > 0 Then
        Set tmpCheque = New cheque
        tmpCheque.observaciones = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_OBSERVACIONES)
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
        tmpCheque.IdOrdenPagoOrigen = GetValue(rs, fieldsIndex, tableNameOrAlias, "orden_pago_origen")
        tmpCheque.entro = GetValue(rs, fieldsIndex, tableNameOrAlias, "ingresado")
        tmpCheque.Depositado = GetValue(rs, fieldsIndex, tableNameOrAlias, "depositado")
        tmpCheque.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, "estado")
        tmpCheque.destino = GetValue(rs, fieldsIndex, proveedores, "razon")
        tmpCheque.Recibo = GetValue(rs, fieldsIndex, recibosChequesTableNameOrAlias, "idRecibo")
        
        If LenB(bancoTableNameOrAlias) > 0 Then Set tmpCheque.Banco = DAOBancos.Map(rs, fieldsIndex, bancoTableNameOrAlias)
        If LenB(monedaTableNameOrAlias) > 0 Then Set tmpCheque.moneda = DAOMoneda.Map(rs, fieldsIndex, monedaTableNameOrAlias)
        If LenB(chequeraTableNameOrAlias) > 0 Then Set tmpCheque.chequera = DAOChequeras.Map(rs, fieldsIndex, chequeraTableNameOrAlias, monedaChequeraTableNameOrAlias, bancoChequeraTableNameOrAlias)
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
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_ID)

    If Id > 0 Then
        Set tmpCheque = New cheque
        tmpCheque.observaciones = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOCheques.CAMPO_OBSERVACIONES)
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
        tmpCheque.IdOrdenPagoOrigen = GetValue(rs, fieldsIndex, tableNameOrAlias, "orden_pago_origen")
        tmpCheque.NumeroLiquidacionCaja = GetValue(rs, fieldsIndex, tableNameOrAlias, "liquidacion_caja_origen")
        tmpCheque.NumeroPagoACuenta = GetValue(rs, fieldsIndex, tableNameOrAlias, "pago_a_cuenta_origen")
        tmpCheque.entro = GetValue(rs, fieldsIndex, tableNameOrAlias, "ingresado")
        tmpCheque.Depositado = GetValue(rs, fieldsIndex, tableNameOrAlias, "depositado")
        tmpCheque.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, "estado")

        
        ' Verificar si la tabla existe y obtener el valor de "numero_liq"
        On Error Resume Next ' Ignorar errores temporalmente
        tmpCheque.NumeroLiquidacionCaja = GetValue(rs, fieldsIndex, LiquidacionesC, "numero_liq")
        If Err.Number <> 0 Then
            ' Si hay un error, el campo no existe
            tmpCheque.NumeroLiquidacionCaja = "" ' O un valor por defecto
            Err.Clear ' Limpiar el error
        End If
        On Error GoTo 0 ' Restaurar el manejo de errores
        
        
                
        If LenB(bancoTableNameOrAlias) > 0 Then Set tmpCheque.Banco = DAOBancos.Map(rs, fieldsIndex, bancoTableNameOrAlias)
        If LenB(monedaTableNameOrAlias) > 0 Then Set tmpCheque.moneda = DAOMoneda.Map(rs, fieldsIndex, monedaTableNameOrAlias)
        If LenB(chequeraTableNameOrAlias) > 0 Then Set tmpCheque.chequera = DAOChequeras.Map(rs, fieldsIndex, chequeraTableNameOrAlias, monedaChequeraTableNameOrAlias, bancoChequeraTableNameOrAlias)
    End If

    Set Map3 = tmpCheque

End Function

Public Function Guardar(cheque As cheque) As Boolean
    Dim q As String

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
          & "observaciones, teceros_propio,ingresado,fecha_emision,orden_pago_origen,depositado" _
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
          & "'observaciones', 'teceros_propio','ingresado','fecha_emision','orden_pago_origen','depositado' " _
          & ")"

    Else

        q = "Update Cheques" _
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
          & "orden_pago_origen='orden_pago_origen', " _
          & "estado='estado', " _
          & "depositado='depositado' " _
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
    q = Replace(q, "'observaciones'", conectar.Escape(cheque.observaciones))
    q = Replace(q, "'teceros_propio'", conectar.Escape(cheque.TercerosPropio))
    q = Replace(q, "'ingresado'", conectar.Escape(Abs(cheque.entro)))
    q = Replace(q, "'orden_pago_origen'", conectar.Escape(cheque.IdOrdenPagoOrigen))
    q = Replace(q, "'estado'", conectar.Escape(cheque.estado))
    q = Replace(q, "'depositado'", conectar.Escape(cheque.Depositado))
    q = Replace(q, "'fecha_emision'", conectar.Escape(Format(cheque.FechaEmision, "yyyy-mm-dd")))
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
