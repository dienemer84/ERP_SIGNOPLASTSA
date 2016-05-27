Attribute VB_Name = "DAOCuentaCorrienteHistoric"
Option Explicit


Public Function GetById(TipoPersona As TipoPersona, id As Long) As CuentaCorrienteHistoric
    Dim col As New Collection
    Dim strsql As String
    Dim rs As Recordset
    Dim cta As CuentaCorrienteHistoric
    Dim rsdeta As Recordset
    Dim idx As Dictionary
    strsql = "select * from cuenta_corriente_historic where id = " & id

    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx

    If Not rs.EOF And Not rs.BOF Then
        Set cta = New CuentaCorrienteHistoric
        Set cta = Map(rs, idx, "cuenta_corriente_historic", True)


    End If

    col.Add cta
    rs.MoveNext

    Set GetById = cta



End Function

Public Function Map(ByRef rs As Recordset, ByRef idx As Dictionary, ByRef tableNameOrAlias As String, withDetalles As Boolean) As CuentaCorrienteHistoric

    Dim cta As CuentaCorrienteHistoric
    Dim id As Variant
    Dim rsdeta As Recordset
    id = GetValue(rs, idx, tableNameOrAlias, "id")

    If id >= 0 Then
        Set cta = New CuentaCorrienteHistoric
        Set cta.Detalles = New Collection
        cta.id = id

        cta.FechaHasta = GetValue(rs, idx, tableNameOrAlias, "fecha_hasta")
        cta.id_persona = GetValue(rs, idx, tableNameOrAlias, "id_persona")
        cta.Periodo = GetValue(rs, idx, tableNameOrAlias, "periodo")
        cta.TipoPersona = GetValue(rs, idx, tableNameOrAlias, "tipo_persona")
        If withDetalles Then
            Dim deta As DTODetalleCuentaCorriente
            Set rsdeta = conectar.RSFactory(" select * from cuenta_corriente_historic_detalle where id_cuenta_corriente_historic = " & cta.id)
            While Not rsdeta.EOF
                Set deta = New DTODetalleCuentaCorriente
                deta.Comprobante = rsdeta!detalle
                deta.Debe = rsdeta!Debe
                deta.Haber = rsdeta!Haber
                deta.IdComprobante = rsdeta!id_comprobante
                deta.saldo = rsdeta!saldo
                deta.tipoComprobante = rsdeta!tipo_comprobante
                deta.FEcha = rsdeta!FEcha
                cta.Detalles.Add deta
                rsdeta.MoveNext
            Wend


        End If
    End If
    Set Map = cta






End Function

Public Function IsValidFechaHasta(id As Long, TipoPersona As TipoPersona, FechaHasta As String) As Boolean

    Dim col As New Collection
    Dim cta As CuentaCorrienteHistoric
    Set col = GetAll(TipoPersona, id, False)
    Dim valir As Boolean
    IsValidFechaHasta = True
    For Each cta In col

        If CDate(FechaHasta) <= cta.FechaHasta Then
            IsValidFechaHasta = False
            Exit Function
        End If
    Next


End Function

Public Function GetAllDetallesFromProveedor(id As Long, Optional condicion As String = "1=1") As Collection

    Dim col As New Collection
    Dim deta As DTODetalleCuentaCorriente
    Dim rsdeta As Recordset
    Dim strsql As String

    strsql = " SELECT   hd.* From   cuenta_corriente_historic h LEFT JOIN cuenta_corriente_historic_detalle hd " _
             & " ON hd.id_cuenta_corriente_historic = h.`id`  where id_persona = " & id & " and tipo_persona= " & TipoPersona.proveedor_

    strsql = strsql & " and  hd.fecha<= " & condicion



    Set rsdeta = conectar.RSFactory(strsql)
    If IsSomething(rsdeta) Then
        While Not rsdeta.EOF
            Set deta = New DTODetalleCuentaCorriente
            If rsdeta!tipo_comprobante <> TipoComprobanteUsado.SaldoInicial_ Then
                deta.Comprobante = rsdeta!detalle
                deta.Debe = rsdeta!Debe
                deta.Haber = rsdeta!Haber
                deta.IdComprobante = rsdeta!id_comprobante
                deta.saldo = rsdeta!saldo
                deta.tipoComprobante = rsdeta!tipo_comprobante
                deta.FEcha = rsdeta!FEcha
                col.Add deta
            End If
            rsdeta.MoveNext
        Wend
    End If

    Set GetAllDetallesFromProveedor = col
End Function

Public Function GetAll(TipoPersona As TipoPersona, id As Long, withDetalles As Boolean, Optional filtro As String = "1 = 1 ") As Collection
    Dim col As New Collection
    Dim strsql As String
    Dim rs As Recordset
    Dim cta As CuentaCorrienteHistoric
    Dim rsdeta As Recordset
    Dim idx As Dictionary
    strsql = "select * from cuenta_corriente_historic where id_persona=" & id & " and tipo_persona=" & TipoPersona & " and 1 = 1 and " & filtro

    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set cta = New CuentaCorrienteHistoric
        Set cta = Map(rs, idx, "cuenta_corriente_historic", withDetalles)




        col.Add cta
        rs.MoveNext
    Wend
    Set GetAll = col

End Function

Public Function Save(cta As CuentaCorrienteHistoric) As Boolean
    On Error GoTo error1
    Dim strsql As String
    conectar.BeginTransaction
    strsql = "INSERT INTO cuenta_corriente_historic (periodo,id_persona,tipo_persona,fecha_hasta) VALUES " _
             & " ('periodo', 'id_persona', 'tipo_persona', 'fecha_hasta') "

    strsql = Replace$(strsql, "'periodo'", conectar.Escape(cta.Periodo))
    strsql = Replace$(strsql, "'id_persona'", conectar.Escape(cta.id_persona))
    strsql = Replace$(strsql, "'tipo_persona'", conectar.Escape(cta.TipoPersona))
    strsql = Replace$(strsql, "'fecha_hasta'", conectar.Escape(cta.FechaHasta))


    If Not conectar.execute(strsql) Then GoTo error1
    Dim id_cuenta As Long

    id_cuenta = conectar.UltimoId2
    Dim deta As DTODetalleCuentaCorriente

    For Each deta In cta.Detalles
        strsql = " INSERT INTO `sp`.`cuenta_corriente_historic_detalle`    (`id_cuenta_corriente_historic`, `detalle`,`debe`,`haber`,`saldo`,`id_comprobante`,`tipo_comprobante`,fecha) " _
                 & " VALUES ('id_cuenta_corriente_historic','detalle', 'debe','haber','saldo', 'id_comprobante', 'tipo_comprobante','fecha') "
        strsql = Replace$(strsql, "'id_cuenta_corriente_historic'", conectar.Escape(id_cuenta))
        strsql = Replace$(strsql, "'detalle'", conectar.Escape(deta.Comprobante))
        strsql = Replace$(strsql, "'debe'", conectar.Escape(deta.Debe))
        strsql = Replace$(strsql, "'haber'", conectar.Escape(deta.Haber))
        strsql = Replace$(strsql, "'saldo'", conectar.Escape(deta.saldo))
        strsql = Replace$(strsql, "'id_comprobante'", conectar.Escape(deta.IdComprobante))
        strsql = Replace$(strsql, "'tipo_comprobante'", conectar.Escape(deta.tipoComprobante))
        strsql = Replace$(strsql, "'fecha'", conectar.Escape(deta.FEcha))
        If Not conectar.execute(strsql) Then GoTo error1






    Next deta

    'tengo q definir el saldo inicial del proveedor entonces
    Dim rs As Recordset
    Dim saldo_inicial As Double
    Dim count As Long
    saldo_inicial = 0
    Set rs = conectar.RSFactory("SELECT saldo_inicial FROM saldo_inicial_proveedor WHERE id_proveedor = " & cta.id_persona)
    While Not rs.EOF And Not rs.BOF
        count = count + 1
        saldo_inicial = rs!saldo_inicial

        rs.MoveNext
    Wend
    If count = 1 Then
        'update
        saldo_inicial = DAOCuentaCorriente.GetSaldo(cta.Detalles)
        strsql = "update saldo_inicial_proveedor set saldo_inicial =" & funciones.RedondearDecimales(saldo_inicial, 2) & ", fecha = " & conectar.Escape(cta.FechaHasta) & " where id_proveedor = " & cta.id_persona

        If Not conectar.execute(strsql) Then GoTo error1

    ElseIf count = 0 Then
        saldo_inicial = DAOCuentaCorriente.GetSaldo(cta.Detalles)

        strsql = "INSERT INTO `sp`.`saldo_inicial_proveedor` (`id_proveedor`, `saldo_inicial`,    `fecha`) " _
                 & " VALUES ('id_proveedor', 'saldo_inicial', 'fecha')"

        strsql = Replace$(strsql, "'id_proveedor'", conectar.Escape(cta.id_persona))
        strsql = Replace$(strsql, "'saldo_inicial'", conectar.Escape(saldo_inicial))
        strsql = Replace$(strsql, "'fecha'", conectar.Escape(Format(cta.FechaHasta, "YYYY-MM-DD")))
        If Not conectar.execute(strsql) Then GoTo error1


    Else
        MsgBox "Se produjo un error con la carga de los saldos iniciales"

    End If

    Save = True
    conectar.CommitTransaction
    Exit Function
error1:
    Save = False
    conectar.RollBackTransaction
End Function

