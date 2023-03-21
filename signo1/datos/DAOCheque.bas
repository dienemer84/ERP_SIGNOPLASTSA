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

Public Function FindAll(Optional ByRef filter As String = vbNullString, Optional orderBy As String) As Collection
    On Error GoTo err1

    'Dim tickStart As Double
    'Dim tickEnd As Double
    'tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

'    q = "SELECT *" _
'        & " FROM Cheques cheq" _
'        & " LEFT JOIN Chequeras cheqs ON cheqs.id = cheq.id_chequera" _
'        & " LEFT JOIN AdminConfigBancos banc ON banc.id = cheq.id_banco" _
'        & " LEFT JOIN AdminConfigMonedas mon ON mon.id = cheq.id_moneda" _
'        & " LEFT JOIN AdminConfigMonedas mon2 ON mon2.id = cheqs.id_moneda" _
'        & " LEFT JOIN AdminConfigBancos banc2 ON banc2.id = cheqs.id_banco" _
'        & " LEFT JOIN ordenes_pago_facturas ordenesp ON cheq.orden_pago_origen=ordenesp.id_orden_pago" _
'        & " LEFT JOIN AdminComprasFacturasProveedores facturasp ON ordenesp.id_factura_proveedor=facturasp.id" _
'        & " LEFT JOIN proveedores prov ON facturasp.id_proveedor=prov.id" _
'        & " WHERE 1 = 1 "

    q = "SELECT *" _
        & " FROM Cheques cheq" _
        & " LEFT JOIN Chequeras cheqs ON cheqs.id = cheq.id_chequera" _
        & " LEFT JOIN AdminConfigBancos banc ON banc.id = cheq.id_banco" _
        & " LEFT JOIN AdminConfigMonedas mon ON mon.id = cheq.id_moneda" _
        & " LEFT JOIN AdminConfigMonedas mon2 ON mon2.id = cheqs.id_moneda" _
        & " LEFT JOIN AdminConfigBancos banc2 ON banc2.id = cheqs.id_banco" _
        & " WHERE 1 = 1 "

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    If LenB(orderBy) > 0 Then
        q = q & " ORDER BY " & orderBy
    End If


    Set rs = conectar.RSFactory(q)
    
'    'debug.print (q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim Cheques As New Collection

    Dim tmpCheque As cheque

       
    While Not rs.EOF
        Set tmpCheque = DAOCheques.Map(rs, fieldsIndex, TABLA_CHEQUE, "banc", "mon", "cheqs", "mon2", "banc2", "ordenesp", "facturasp", "prov")
        Cheques.Add tmpCheque, CStr(tmpCheque.Id)

        rs.MoveNext

    Wend

    'tickEnd = GetTickCount

    'Debug.Print tickEnd - tickStart, "ms elapsed"
    
    
    Set FindAll = Cheques
    Exit Function

err1:
    Set FindAll = Nothing
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

Public Function FindByChequeraAndNro(chequeraId As Long, nro As String) As cheque
    Dim col As Collection
    Set col = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID_CHEQUERA & "=" & chequeraId & " AND " & TABLA_CHEQUE & "." & DAOCheques.CAMPO_NUMERO & " = " & Escape(nro))
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

Public Function FindAllByChequeraId(chequeraId As Long) As Collection
    Set FindAllByChequeraId = FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID_CHEQUERA & "=" & chequeraId)
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
                    Optional ByRef FacturasP As String = vbNullString, _
                    Optional ByRef proveedores As String = vbNullString _
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
        If LenB(bancoTableNameOrAlias) > 0 Then Set tmpCheque.Banco = DAOBancos.Map(rs, fieldsIndex, bancoTableNameOrAlias)
        If LenB(monedaTableNameOrAlias) > 0 Then Set tmpCheque.moneda = DAOMoneda.Map(rs, fieldsIndex, monedaTableNameOrAlias)
        If LenB(chequeraTableNameOrAlias) > 0 Then Set tmpCheque.chequera = DAOChequeras.Map(rs, fieldsIndex, chequeraTableNameOrAlias, monedaChequeraTableNameOrAlias, bancoChequeraTableNameOrAlias)
    End If

    Set Map = tmpCheque

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

Public Function FindAllEnCartera() As Collection
    Set FindAllEnCartera = FindAll(DAOCheques.CAMPO_EN_CARTERA & " = 1")
End Function

Public Function FindAllEnCarteraDeTerceros() As Collection
    Set FindAllEnCarteraDeTerceros = FindAll(DAOCheques.CAMPO_EN_CARTERA & " = 1 and " & DAOCheques.CAMPO_PROPIO & " = 0")
End Function
