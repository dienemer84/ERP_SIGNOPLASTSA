Attribute VB_Name = "DAOPresupuestos"
Dim strsql As String
Dim P As clsPresupuesto
Dim rs As ADODB.Recordset
Public Const tablaCliente As String = "c"
Public Const tablaPresu As String = "p"
Public Const tablaU1 As String = "u1"
Public Const tablaU2 As String = "u2"
Public Const tablaU3 As String = "u3"
Public Const tablaMoneda As String = "m"
Public Const formaPagoAnticipo_ As String = "forma_pago_anticipo"
Public Const anticipo_ As String = "anticipo"
Public Const formaPagoSaldo_ As String = "forma_pago_saldo"
Public Const saldo_ As String = "saldo"
Public Const estado_ As String = "estado"
Public Const fechaEntrega_ As String = "fechaEntrega"
Public Const Detalle_ As String = "detalle"
Public Const Fecha_ As String = "fecha"
Public Const fechaFinalizado_ As String = "fechaFinalizado"
Public Const fechaModificado_ As String = "fechaModificado"
Public Const fechaProcesado_ As String = "fechaProcesado"
Public Const diasPagoAnticipo_ As String = "dias_pago_anticipo"
Public Const diasPagoSaldo_ As String = "dias_pago_saldo"
Public Const PorcentajeManoObraMuerta_ As String = "porcentaje_mano_obra_muerta"
Public Const gastos_ As String = "gastos"
Public Const manteOferta_ As String = "ManteOferta"
Public Const PorcMen15_ As String = "PorcMen15"
Public Const PorcMDO_ As String = "PorcMDO"
Public Const PorcMen10_ As String = "PorcMen10"
Public Const PorcMas15_ As String = "PorcMas15"
Public Const VencimientoPresupuesto_ As String = "vencimientoPresupuesto"
Public Const Descuento_ As String = "descuento"
Public Function aprobar(T As clsPresupuesto) As Boolean
    On Error GoTo err1








    conectar.BeginTransaction
    fecha_anterior = T.FechaModificado
    Set usuario_anterior = T.UsuarioFinalizado
    If T.EstadoPresupuesto = ACotizar_ Then
        T.EstadoPresupuesto = EstadoPresupuesto.Pendiente_
        Set T.UsuarioFinalizado = funciones.GetUserObj
        T.FechaFinalizado = Now
        aprobar = True
        If Not Guardar(T) Then GoTo err1
        If Not DAOPresupuestoDetalleHistorico.Create(T) Then GoTo err1
    Else
        aprobar = False
    End If
    DAOPresupuestoHistorial.agregar T, "Presupuesto aprobado"
    conectar.CommitTransaction
    DAOEvento.Publish T.Id, TipoEventoBroadcast.TEB_PresupuestoAprobado
    Exit Function
err1:
    T.EstadoPresupuesto = ACotizar_
    T.FechaModificado = fecha_anterior
    Set T.UsuarioFinalizado = Nothing
    conectar.RollBackTransaction

End Function

Private Function CambiarEstado(T As clsPresupuesto, nuevoEstado As EstadoPresupuesto)
    estado_ant = T.EstadoPresupuesto
    fecha_anterior = T.FechaModificado
    Set usuario_anterior = T.UsuarioModificado
    T.EstadoPresupuesto = nuevoEstado
    Set T.UsuarioModificado = funciones.GetUserObj
    If Not Guardar(T) Then
        GoTo err1
    Else
        CambiarEstado = Guardar(T)
    End If

    Exit Function
err1:
    T.EstadoPresupuesto = estado_ant
    Set T.UsuarioModificado = usuario_anterior
End Function
Public Function desactivar(T As clsPresupuesto) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    desactivar = CambiarEstado(T, Desactivado)
    If desactivar Then
        DAOPresupuestoHistorial.agregar T, "Presupuesto desactivado"
        DAOEvento.Publish T.Id, TipoEventoBroadcast.TEB_PresupuestoAnulado
        conectar.CommitTransaction
    Else
        GoTo err1
    End If
    Exit Function
err1:
    conectar.RollBackTransaction
End Function
Public Function NoCotizar(T As clsPresupuesto) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    NoCotizar = CambiarEstado(T, EstadoPresupuesto.NoCotizado)
    If NoCotizar Then
        DAOPresupuestoHistorial.agregar T, "Presupuesto no cotizado"
        conectar.CommitTransaction
    Else
        GoTo err1
    End If
    Exit Function
err1:
    conectar.RollBackTransaction
End Function
Public Function enviar(T As clsPresupuesto) As Boolean
    conectar.BeginTransaction
    enviar = True
    If exporta(T, True) Then
        If Not CambiarEstado(T, Enviado_) Then
            DAOPresupuestoHistorial.agregar T, "Presupuesto enviado"

        End If
    Else
        GoTo err1
    End If
    conectar.CommitTransaction

    DAOEvento.Publish T.Id, TipoEventoBroadcast.TEB_PresupuestoEnviado
    Exit Function
err1:
    enviar = False
    conectar.RollBackTransaction
End Function
Public Function procesar(T As clsPresupuesto) As Boolean
    conectar.BeginTransaction
    procesar = CambiarEstado(T, EstadoPresupuesto.Procesado_)
    If procesar Then
        DAOPresupuestoHistorial.agregar T, "Presupuesto procesado"
        conectar.CommitTransaction
    Else
        GoTo err1
    End If
    Exit Function
err1:
    conectar.RollBackTransaction

End Function
Public Function GetById(Id As Long) As clsPresupuesto
    strsql = "SELECT " _
           & "p.*, u1.*, u2.*, u3.*,m.*, c.* " _
           & "FROM presupuestos p " _
           & "INNER JOIN clientes c " _
           & "ON p.idCliente = c.id " _
           & "LEFT JOIN usuarios u1 " _
           & "ON p.idVendedor = u1.id " _
           & "LEFT JOIN  usuarios u2 " _
           & "ON p.idUsuarioFinalizado=u2.id " _
           & "LEFT JOIN usuarios u3 " _
           & "ON p.idUsuarioModificado=u3.id " _
           & "LEFT JOIN AdminConfigMonedas m " _
           & "ON p.idMoneda=m.id " _
           & " where p.id=" & Id
    Set rs = conectar.RSFactory(strsql)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    If Not rs.EOF And Not rs.BOF Then
        Set GetById = Map(rs, fieldsIndex, tablaPresu, tablaCliente, tablaMoneda, tablaU1, tablaU2, tablaU3)

    End If
End Function
Public Function GetAll(Optional filtro As String = Empty, Optional estados As Collection = Nothing, Optional cliente As Long = -1, Optional numero As Long = 0) As Collection
    Dim col As Collection
    strsql = "SELECT " _
           & "p.*, u1.*, u2.*, u3.*,m.*, c.* " _
           & "FROM presupuestos p " _
           & "INNER JOIN clientes c " _
           & "ON p.idCliente = c.id " _
           & "LEFT JOIN usuarios u1 " _
           & "ON p.idVendedor = u1.id " _
           & "LEFT JOIN  usuarios u2 " _
           & "ON p.idUsuarioFinalizado=u2.id " _
           & "LEFT JOIN usuarios u3 " _
           & "ON p.idUsuarioModificado=u3.id " _
           & "LEFT JOIN AdminConfigMonedas m " _
           & "ON p.idMoneda=m.id " _
           & " where 1=1"

    If Not estados Is Nothing And estados.count > 0 Then
        strsql = strsql & " and ("
        For i = 1 To estados.count
            If i = 1 Then
                filtro_estados = "p.estado=" & estados.item(i)
            Else
                filtro_estados = " or p.estado=" & estados.item(i)
            End If
            strsql = strsql & filtro_estados
        Next i
        strsql = strsql & ")"
    End If
    If cliente > 0 Then
        strsql = strsql & " and idCliente=" & cliente
    End If
    If filtro <> Empty Then
        filtro = " and p.detalle like '%" & Trim(filtro) & "%' or c.razon LIKE '%" & Trim(filtro) & "%'"
        strsql = strsql & filtro
    End If
    If numero > 0 Then
        strsql = strsql & " and  p.id=" & numero
    End If



    Set col = New Collection
    Set rs = conectar.RSFactory(strsql)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim tmp As clsPresupuesto
    While Not rs.EOF
        Set tmp = Map(rs, fieldsIndex, tablaPresu, tablaCliente, tablaMoneda, tablaU1, tablaU2, tablaU3)

        col.Add tmp, CStr(tmp.Id)

        rs.MoveNext
    Wend




    Set GetAll = col
End Function

Public Function Map(ByRef rs As Recordset, Index As Dictionary, ByRef tableNameOrAlias As String, Optional ByRef tableNameOrAliasCliente As String = vbNullString, Optional ByRef tableNameOrAliasMoneda As String = vbNullString, Optional tableUsuarioCreado As String = vbNullString, Optional tableUsuarioModificado As String = vbNullString, Optional tableUsuarioFinalizado As String = vbNullString) As clsPresupuesto
'Dim P As clsPresupuesto
    Dim Id As Variant
    Id = GetValue(rs, Index, tableNameOrAlias, "id")
    If Id > 0 Then
        Set P = New clsPresupuesto
        P.Id = Id
        P.EstadoPresupuesto = GetValue(rs, Index, tableNameOrAlias, estado_)
        P.FechaEntrega = GetValue(rs, Index, tableNameOrAlias, fechaEntrega_)
        P.Anticipo = GetValue(rs, Index, tableNameOrAlias, anticipo_)
        P.FormaPagoAnticipo = GetValue(rs, Index, tableNameOrAlias, formaPagoAnticipo_)
        P.FormaPagoSaldo = GetValue(rs, Index, tableNameOrAlias, formaPagoSaldo_)
        P.detalle = GetValue(rs, Index, tableNameOrAlias, Detalle_)
        P.fechaCreado = GetValue(rs, Index, tableNameOrAlias, Fecha_)
        P.FechaFinalizado = GetValue(rs, Index, tableNameOrAlias, fechaFinalizado_)
        P.FechaModificado = GetValue(rs, Index, tableNameOrAlias, fechaModificado_)
        P.FechaProcesado = GetValue(rs, Index, tableNameOrAlias, fechaProcesado_)
        P.Gastos = GetValue(rs, Index, tableNameOrAlias, gastos_)
        P.manteOferta = GetValue(rs, Index, tableNameOrAlias, manteOferta_)
        P.PorcMen15 = GetValue(rs, Index, tableNameOrAlias, PorcMen15_)
        P.PorcMDO = GetValue(rs, Index, tableNameOrAlias, PorcMDO_)
        P.PorcMen10 = GetValue(rs, Index, tableNameOrAlias, PorcMen10_)
        P.PorcMas15 = GetValue(rs, Index, tableNameOrAlias, PorcMas15_)
        P.PorcentajeManoObraMuerta = GetValue(rs, Index, tableNameOrAlias, PorcentajeManoObraMuerta_)
        P.VencimientoPresupuesto = GetValue(rs, Index, tableNameOrAlias, VencimientoPresupuesto_)
        P.Descuento = GetValue(rs, Index, tableNameOrAlias, Descuento_)
        P.DiasPagoAnticipo = GetValue(rs, Index, tableNameOrAlias, diasPagoAnticipo_)
        P.DiasPagoSaldo = GetValue(rs, Index, tableNameOrAlias, diasPagoSaldo_)

        If LenB(tableNameOrAliasCliente) > 0 Then Set P.cliente = DAOCliente.Map(rs, Index, tableNameOrAliasCliente)
        If LenB(tableNameOrAliasMoneda) > 0 Then Set P.moneda = DAOMoneda.Map(rs, Index, tableNameOrAliasMoneda)
        If LenB(tableUsuarioModificado) > 0 Then Set P.UsuarioModificado = DAOUsuarios.Map(rs, Index, tableUsuarioModificado)
        If LenB(tableUsuarioFinalizado) > 0 Then Set P.UsuarioFinalizado = DAOUsuarios.Map(rs, Index, tableUsuarioFinalizado)
        If LenB(tableUsuarioCreado) > 0 Then Set P.UsuarioCreado = DAOUsuarios.Map(rs, Index, tableUsuarioCreado)

    End If

    Set Map = P
End Function
Public Function Guardar(T As clsPresupuesto, Optional Cascade As Boolean = False) As Boolean
    On Error GoTo err1
    Dim u_anterior As clsUsuario
    Dim f_anterior As Date
    Dim Id As Long
    Dim tmp As clsPresupuesto
    Guardar = True
    Dim insert As Boolean

    If T.Id > 0 Then
        'update
        Set tmp = DAOPresupuestos.GetById(T.Id)
        Set T.UsuarioModificado = funciones.GetUserObj
        If T.FechaModificado = tmp.FechaModificado Then
            Set u_anterior = T.UsuarioModificado
            f_anterior = T.FechaModificado
            T.FechaModificado = Now
            fecha_procesada = IIf(CDbl(T.FechaProcesado) = 0, 0, funciones.datetimeFormateada(T.FechaProcesado))
            fecha_finalizada = IIf(CDbl(T.FechaFinalizado) = 0, 0, funciones.datetimeFormateada(T.FechaFinalizado))
            fecha_modif = IIf(CDbl(T.FechaModificado) = 0, 0, funciones.datetimeFormateada(T.FechaModificado))
            If T.UsuarioFinalizado Is Nothing Then
                id_u_fin = 0
            Else
                id_u_fin = T.UsuarioFinalizado.Id
            End If
            strsql = "update presupuestos" _
                   & " SET " _
                   & "fecha = '" & funciones.datetimeFormateada(T.fechaCreado) & "' ," _
                   & "fechaEntrega = " & T.FechaEntrega & " ," _
                   & "idCliente = " & T.cliente.Id & " ," _
                   & "idVendedor = '" & T.UsuarioCreado.Id & "'," _
                   & "estado = " & T.EstadoPresupuesto & "," _
                   & "detalle = '" & T.detalle & "' ," _
                   & "descuento = " & T.Descuento & "," _
                   & "PorcMDO = " & T.PorcMDO & " ," _
                   & "PorcMen10 = " & T.PorcMen10 & " ," _
                   & "PorcMen15 = " & T.PorcMen15 & "," _
                   & "PorcMas15 = " & T.PorcMas15 & " ," _
                   & "gastos = " & T.Gastos & " ," _
                   & "porcentaje_mano_obra_muerta = " & T.PorcentajeManoObraMuerta & " ," _
                   & "ManteOferta = " & T.manteOferta & " ," _
                   & "vencimientoPresupuesto = '" & funciones.dateFormateada(T.VencimientoPresupuesto) & "'," _
                   & "fechaProcesado = '" & fecha_procesada & "'," _
                   & "idUsuarioFinalizado = " & id_u_fin & " ," _
                   & "fechaFinalizado = '" & fecha_finalizada & "' ," _
                   & "idMoneda = " & T.moneda.Id & " ," & "fechaModificado = '" & fecha_modif & "' ," _
                   & "anticipo= " & T.Anticipo & "," _
                   & "forma_pago_anticipo= '" & T.FormaPagoAnticipo & "', " & "forma_pago_saldo= '" & T.FormaPagoSaldo & "', " _
                   & "dias_pago_anticipo =" & T.DiasPagoAnticipo & ", dias_pago_saldo = " & T.DiasPagoSaldo & ", " _
                   & "idUsuarioModificado = " & T.UsuarioModificado.Id & " Where id =" & T.Id
            If Not conectar.execute(strsql) Then GoTo err1
            If Not T.DetallePresupuesto Is Nothing Then

                If Cascade Then If Not DAOPresupuestosDetalle.Save(T) Then GoTo err1
                DAOPresupuestoHistorial.agregar T, "Presupuesto guardado"
            End If
        Else
            MsgBox "Presupuesto ya guardado en otra sesión!", vbInformation, "Información"
            Guardar = False
            GoTo err1
        End If
    Else
        'insert
        insert = True
        strsql = "insert into `sp`.`presupuestos` " _
               & "(fecha, fechaEntrega, idCliente, idVendedor, estado, " _
               & " detalle, descuento, PorcMDO, PorcMen10, PorcMen15, PorcMas15, " _
               & "gastos, ManteOferta, vencimientoPresupuesto, idMoneda, dias_pago_saldo,forma_pago_saldo,porcentaje_mano_obra_muerta) " _
               & " values " _
               & " ('" & funciones.datetimeFormateada(T.fechaCreado) & "','" & T.FechaEntrega & "'," & T.cliente.Id & " ," & T.UsuarioCreado.Id & "," & T.EstadoPresupuesto & ",'" & T.detalle & "'," & T.Descuento & ", " _
               & " " & T.PorcMDO & "," & T.PorcMen10 & "," & T.PorcMen15 & "," & T.PorcMas15 & "," & T.Gastos & "," & T.manteOferta & ",'" & funciones.datetimeFormateada(T.VencimientoPresupuesto) & "' " _
               & "," & T.moneda.Id & "," & T.cliente.FP & ",'" & T.cliente.FormaPago & "'," & T.PorcentajeManoObraMuerta & ")"
        If Not conectar.execute(strsql) Then GoTo err1
        conectar.UltimoId "presupuestos", Id
        T.Id = Id
        DAOPresupuestoHistorial.agregar T, "Presupuesto ingresado"
        If Not T.DetallePresupuesto Is Nothing Then

            If Cascade Then If Not DAOPresupuestosDetalle.Save(T) Then GoTo err1
            DAOPresupuestoHistorial.agregar T, "Presupuesto guardado"
        End If

    End If

    If insert Then
        DAOEvento.Publish T.Id, TipoEventoBroadcast.TEB_PresupuestoCreado
    End If

    Guardar = True
    Exit Function
err1:
    Set T.UsuarioModificado = u_anterior
    T.FechaModificado = f_anterior
    Guardar = False
End Function


Public Function Save(T As clsPresupuesto) As Boolean
    conectar.BeginTransaction
    Save = Guardar(T, True)
    If Save Then
        conectar.CommitTransaction

    Else
        conectar.RollBackTransaction
    End If
End Function

Public Function ExisteDetalle(detalle As String) As Boolean
    Set rs = conectar.RSFactory("select count(id) as cant from presupuestos where detalle='" & detalle & "'")
    If rs!Cant > 0 Then
        buscarDetalle = True
    Else
        buscarDetalle = False
    End If
End Function
Public Function ProximoPresupuesto() As Long
    ProximoPresupuesto = conectar.ProximoId("presupuestos")
End Function
Public Function ImprimirPresupuesto(T As clsPresupuesto) As Boolean    '1- enviar 2-imprimir
    On Error GoTo err2
    Dim Total As Double
    Dim totaldto As Double
    Dim SubTotal As Double
    SubTotal = T.Total(Manual)
    Total = SubTotal * 1
    presu1.Sections("header").Controls("lblCliente").caption = T.cliente.razon
    presu1.Sections("header").Controls("lblRef").caption = T.detalle
    presu1.Sections("header").Controls("lblfechaPresu").caption = T.fechaCreado
    presu1.Sections("header").Controls("lblnroPresu").caption = T.IdFormateada
    presu1.Sections("fondo").Controls("lbldto").caption = funciones.FormatearDecimales(T.Descuento)
    presu1.Sections("fondo").Controls("lblsubtotal").caption = funciones.FormatearDecimales(T.SubTotal(Manual))
    totaldto = SubTotal * (T.Descuento / 100)
    Total = SubTotal - totaldto
    presu1.Sections("fondo").Controls("lbltotal").caption = funciones.FormatearDecimales(T.Total(Manual), 2)
    presu1.Sections("fondo").Controls("lblmoneda").caption = T.moneda.NombreCorto
    presu1.Sections("fondo").Controls("lblfecEnt").caption = T.FechaEntrega & " días desde la recepción de la O/C."
    presu1.Sections("fondo").Controls("lblManteOferta").caption = T.manteOferta & " días"

    fp_anticipo = T.Anticipo & "% " & T.DiasPagoAnticipo & " días"


    If Trim(T.FormaPagoAnticipo) <> Empty Then
        fp_anticipo = fp_anticipo & " - Forma de pago: " & T.FormaPagoAnticipo
    End If
    presu1.Sections("fondo").Controls("lblAnticipo").caption = fp_anticipo



    fp_saldo = T.DiasPagoSaldo & " días"
    If Trim(T.FormaPagoSaldo) <> Empty Then
        fp_saldo = fp_saldo & " - Forma de Pago: " & T.FormaPagoSaldo
    End If
    presu1.Sections("fondo").Controls("lblSaldo").caption = fp_saldo
    If T.Anticipo <= 0 Then
        presu1.Sections("fondo").Controls("lblAnticipo").Visible = False
        presu1.Sections("fondo").Controls("lblAnticipo_header").Visible = False
    Else
        presu1.Sections("fondo").Controls("lblAnticipo").Visible = True
        presu1.Sections("fondo").Controls("lblAnticipo_header").Visible = True
    End If

    If T.DiasPagoSaldo <= 0 Then
        presu1.Sections("fondo").Controls("lblSaldo").Visible = False
        presu1.Sections("fondo").Controls("lblSaldo_header").Visible = False
    Else
        presu1.Sections("fondo").Controls("lblSaldo").Visible = True
        presu1.Sections("fondo").Controls("lblSaldo_header").Visible = True
    End If
    ImprimirPresupuesto = True
    strsql = "select concat(s.detalle,' ',dp.masdetalles) as deta, dp.entregaitem,dp.masdetalles,dp.item as item,dp.cantidad as cantidad, dp.valorunitario as unitario, s.detalle as detalle, dp.cantidad*dp.valorunitario as total from detalle_presupuesto dp, stock s where dp.idpieza=s.id and dp.idpresupuesto=" & T.Id
    Set presu1.DataSource = conectar.RSFactory(strsql)
    presu1.Show 1
    Exit Function
err2:
    MsgBox Err.Description
    ImprimirPresupuesto = False
End Function
Public Function exporta(T As clsPresupuesto, Optional enviar As Boolean = False) As Boolean
    On Error GoTo errEXCEL
    Dim P As Long
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application

    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    exporta = True

    xla.Activate


    Set T.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(T)

    With xla
        .Columns("A").HorizontalAlignment = xlHAlignCenter
        .Columns("B").HorizontalAlignment = xlHAlignCenter
        .Columns("C").HorizontalAlignment = xlHAlignCenter
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").HorizontalAlignment = xlHAlignCenter
        .Columns("F").HorizontalAlignment = xlHAlignCenter
        .Columns("G").HorizontalAlignment = xlHAlignCenter
        .Range("A1:A4").HorizontalAlignment = xlHAlignRight
        .Range("a7:G7").Interior.Color = &HC0C0C0
        .Range("A7:G7").Font.Bold = True
        .Range("A1:A4").Font.Bold = True
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 18
        .Columns("G").ColumnWidth = 25
        .Range("C1:E1").Merge
        .Range("C2:E2").Merge
        .Range("C3:E3").Merge
        .Range("C4:E4").Merge
        .Range("A1:B1").Merge
        .Range("A2:B2").Merge
        .Range("A3:B3").Merge
        .Range("A4:B4").Merge
        .Range("C1:E1").HorizontalAlignment = xlHAlignLeft
        .Range("C2:E2").HorizontalAlignment = xlHAlignLeft
        .Range("C3:E3").HorizontalAlignment = xlHAlignLeft
        .Range("C4:E4").HorizontalAlignment = xlHAlignLeft
        .Cells(1, 1).value = "Presupuesto Nro."
        .Cells(1, 3).value = str(Format(T.Id, "0000"))
        .Cells(2, 1).value = "Fecha presupuesto"
        .Cells(2, 3).value = str(T.fechaCreado)
        .Cells(3, 1).value = "Cliente"
        .Cells(3, 3).value = T.cliente.razon
        .Cells(4, 1).value = "Referencia"
        .Cells(4, 3).value = T.detalle
        .Cells(7, 1).value = "Item"
        .Cells(7, 2).value = "Cantidad"
        .Cells(7, 3).value = "Detalle"
        .Columns("c").ColumnWidth = 30
        .Cells(7, 4).value = "P.U."
        .Cells(7, 5).value = "Total"
        .Cells(7, 6).value = "Entrega"
        .Cells(7, 7).value = "Observaciones"
        c = 0
        'encambezado
        Dim max_NOMBRE_pieza As Integer
        max_NOMBRE_pieza = 0
        Dim max_DETALLES As Integer
        max_DETALLES = 0
        c = T.DetallePresupuesto.count
        For P = 1 To T.DetallePresupuesto.count
            Set d = T.DetallePresupuesto(P)
            .Cells(P + 7, 1).value = d.item
            .Cells(P + 7, 2).value = d.Cantidad
            .Cells(P + 7, 3).value = d.Pieza.nombre
            If max_NOMBRE_pieza < Len(.Cells(P + 7, 3).text) Then
                max_NOMBRE_pieza = Len(.Cells(P + 7, 3).text)
            End If
            .Cells(P + 7, 4).value = CStr(funciones.RedondearDecimales(d.ValorManual, 2))
            .Cells(P + 7, 4).NumberFormat = "0.00"
            .Cells(P + 7, 5).value = CStr(funciones.RedondearDecimales(d.ValorManual * d.Cantidad, 2))
            .Cells(P + 7, 5).NumberFormat = "0.00"
            .Cells(P + 7, 6).value = d.entrega & " días"
            .Cells(P + 7, 7).value = d.Detalles
            If max_DETALLES < Len(.Cells(P + 7, 7).text) Then
                max_DETALLES = Len(.Cells(P + 7, 7).text)
            End If
        Next P
        .Columns("C").ColumnWidth = max_NOMBRE_pieza + 5
        If max_DETALLES = 0 Then
            .Columns("g").Hidden = True
        Else
            .Columns("G").ColumnWidth = max_DETALLES + 5
        End If

        .Columns("D").ColumnWidth = 17
        .Columns("E").ColumnWidth = 17
        .Columns("D").HorizontalAlignment = xlHAlignRight
        .Columns("E").HorizontalAlignment = xlHAlignRight
        .Columns("C").HorizontalAlignment = xlHAlignLeft
        .Cells(P + 7, 5).value = funciones.RedondearDecimales(T.SubTotal(Manual), 2)
        .Cells(P + 7, 4).value = "SubTotal " & T.moneda.NombreCorto
        .Cells(P + 8, 4).value = "Desc "
        .Cells(P + 9, 4).value = "Total " & T.moneda.NombreCorto
        .Cells(P + 7, 5).Formula = "=SUM(E8:E" & 8 + (P - 2) & ")"   'funciones.RedondearDecimales(T.Total(Manual), 2)
        .Cells(P + 8, 5).value = T.Descuento & "%"
        .Cells(P + 9, 5).value = funciones.RedondearDecimales(T.Total(Manual), 2)
        .Cells(P + 9, 5).NumberFormat = "0.00"
        .Range("D" & P + 7 & ":D" & P + 9).Font.Bold = True
        .Range(.Cells(7, 1), .Cells(c + 7, 7)).Borders.LineStyle = xlContinuous
        .Cells(c + 11, 1).value = "Fecha de Entrega total"
        .Cells(c + 12, 1).value = "Mant. de oferta"
        .Cells(c + 11, 3).value = str(T.FechaEntrega) & " días desde la recepción de la O/C"
        .Cells(c + 12, 3).value = str(T.manteOferta) & " días"
        .Cells(c + 13, 1).value = "Condición de Pago"
        .Cells(c + 13, 3).value = ""
        aa = 14
        If T.Anticipo > 0 Then
            .Cells(c + aa, 1).value = "Anticipo " & T.Anticipo & "%"
            max_cel_anti = Len(.Cells(c + aa, 1))
            .Columns("A").ColumnWidth = max_cel_anti
            .Cells(c + aa, 2).value = " a " & T.DiasPagoAnticipo & "Días"
            If T.FormaPagoAnticipo <> Empty Then
                .Cells(c + aa, 3).value = "FP: " & T.FormaPagoAnticipo
            End If
            aa = aa + 1
        End If

        .Cells(c + aa, 1).value = "Saldo"
        If T.DiasPagoSaldo > 0 Then
            .Cells(c + aa, 2).value = " a " & T.DiasPagoSaldo & " días"
        End If
        If T.FormaPagoSaldo <> Empty Then
            .Cells(c + aa, 3).value = "FP: " & T.FormaPagoSaldo
        End If
        .Cells(c + 18, 1).value = "Todos los precios son MAS IVA."
        .Range("A" & c + 10 & ":A" & c + 14).Font.Bold = True
        .Range("A" & c + 10 & ":B" & c + 15).HorizontalAlignment = xlHAlignLeft
        .Range("A" & c + 10 & ":B" & c + 10).HorizontalAlignment = xlHAlignLeft
        .Range("A" & c + 12 & ":B" & c + 12).HorizontalAlignment = xlHAlignLeft
        .Range("A" & c + 18 & ":C" & c + 18).HorizontalAlignment = xlHAlignLeft
        .Range("A" & c + 19 & ":B" & c + 19).HorizontalAlignment = xlHAlignLeft
        .Range("C" & c + 12).HorizontalAlignment = xlHAlignLeft
        .Range("A" & c + 11).HorizontalAlignment = xlHAlignLeft
        .Range("C" & c + 11).HorizontalAlignment = xlHAlignLeft
        .Cells(c + 19, 1).value = "Por favor, incluír en la Orden de Compra el número de presupuesto."
        .Cells(c + 21, 4).value = "SIGNOPLAST S.A."
        .Cells(c + 22, 4).value = "Administracion: Arieta 4720 - Planta: Almafuerte 4670"
        .Cells(c + 23, 4).value = "Telefonos: +54(11)4651.0051/9 - Fax: +54(11)4651.0050"
        .Cells(c + 24, 4).value = "Mail: ventas@signoplast.com.ar - Web: www.signoplast.com.ar"
        .Range("A" & c + 21 & ":g" & c + 21).Merge
        .Range("A" & c + 22 & ":g" & c + 22).Merge
        .Range("A" & c + 23 & ":g" & c + 23).Merge
        .Range("A" & c + 24 & ":g" & c + 24).Merge
        .Range("A" & c + 21 & ":g" & c + 21).Font.Bold = True
        .Range("A" & c + 24 & ":g" & c + 24).Font.Bold = True
        .Range("A" & c + 21 & ":g" & c + 24).HorizontalAlignment = xlHAlignCenter
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.CD
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter
        Dim refe As String
        refe = ref
        archi = "PRES" & Format(T.IdFormateada, "00000") & ".xlsx  "
        frmPrincipal.CD.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave
        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbExclamation, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            exporta = False
        End If
        xlb.Saved = True
        xlb.Close
        xls.Quit
        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing
        exporta = True



    End With
    'preguntar por exporta
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then

    Else
        'conectar.RollBackTransaction
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        exporta = False
    End If
    exporta = False
    'xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing
End Function


Public Function CrearOT(T As clsPresupuesto, idpedido As Long, DetalleOt As String) As OrdenTrabajo
    On Error GoTo err1
    Dim Ot As New OrdenTrabajo
    Dim detalle_ot As DetalleOrdenTrabajo

    If idpedido <= 0 Then
        Set Ot.Detalles = New Collection
    Else
        Set Ot = DAOOrdenTrabajo.FindById(idpedido)
        Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.Id)
    End If


    Ot.FechaEntrega = DateAdd("d", T.FechaEntrega, Now)
    Ot.Activa = True
    Ot.Anticipo = T.Anticipo
    Ot.AnticipoFacturado = False
    Ot.CantDiasAnticipo = T.DiasPagoAnticipo
    Ot.CantDiasSaldo = T.DiasPagoSaldo
    Set Ot.cliente = T.cliente
    Set Ot.ClienteFacturar = T.cliente
    Ot.FormaDePagoAnticipo = T.FormaPagoAnticipo
    Ot.FormaDePagoSaldo = T.FormaPagoSaldo
    Ot.descripcion = DetalleOt    'T.detalle
    Ot.Descuento = T.Descuento
    Ot.TipoOrden = OT_TRADICIONAL  'cambiar?
    Ot.Entregada = False
    Ot.estado = EstadoOT_Pendiente
    Ot.MismaFechaEntregaParaDetalles = True
    Set Ot.moneda = T.moneda
    Ot.NroPresupuesto = T.Id
    Ot.StockDescontado = False
    'creo el detalle
    For Each detalle In T.DetallePresupuesto
        Set detalle_ot = New DetalleOrdenTrabajo

        detalle_ot.CantidadEntregada = 0
        detalle_ot.CantidadFabricados = 0
        detalle_ot.CantidadFacturada = 0
        detalle_ot.CantidadImpresionesDeRuta = 0
        detalle_ot.CantidadPedida = detalle.Cantidad
        detalle_ot.EstadoProceso = EstProcDetOT_AunNoDefinido
        detalle_ot.FechaEntrega = DateAdd("d", detalle.entrega, Now)
        detalle_ot.item = detalle.item
        detalle_ot.NombrePiezaHistorico = detalle.Pieza.nombre
        detalle_ot.Nota = detalle.Detalles
        detalle_ot.NotaProduccion = vbEmpty
        Set detalle_ot.OrdenTrabajo = Ot
        Set detalle_ot.Pieza = detalle.Pieza
        detalle_ot.Precio = detalle.ValorManual
        detalle_ot.PrecioModificado = False
        detalle_ot.ReservaStock = 0
        detalle_ot.Retirado = False
        detalle_ot.idPresupuestoOrigen = T.Id
        Ot.Detalles.Add detalle_ot
    Next detalle
    Dim ptmp As New clsPresupuesto

    If T.EstadoPresupuesto = EstadoPresupuesto.Procesado_ Then
        MsgBox "No puede procesar un presupuesto ya procesado.", vbCritical, "Error"
    Else
        If DAOOrdenTrabajo.Save(Ot) Then
            procesar T
            Set CrearOT = Ot
        End If
    End If
    Exit Function

err1:
    Set CrearOT = Nothing
End Function


Public Function ReCotizar(PresupuestoOriginal As clsPresupuesto) As Boolean
    Dim P As New clsPresupuesto
    Dim deta As clsPresupuestoDetalle
    Dim d As clsPresupuestoDetalle
    Set PresupuestoOriginal.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(PresupuestoOriginal)
    With PresupuestoOriginal
        P.Anticipo = .Anticipo
        Set P.cliente = .cliente
        P.Descuento = .Descuento
        P.detalle = "RE: " & .detalle
        Set P.DetallePresupuesto = New Collection

        For Each deta In PresupuestoOriginal.DetallePresupuesto
            With deta
                Set d = New clsPresupuestoDetalle
                d.Amortizacion = .Amortizacion
                d.Cantidad = .Cantidad
                d.Detalles = .Detalles
                d.entrega = 0
                d.FormaCotizar = .FormaCotizar
                d.item = .item
                Set d.Pieza = .Pieza
                Set d.presupuesto = P
                d.ValorManual = .ValorManual
                d.ValorSistema = .ValorSistema

            End With
            P.DetallePresupuesto.Add d
        Next
        P.DiasPagoAnticipo = .DiasPagoAnticipo
        P.DiasPagoSaldo = .DiasPagoSaldo
        P.EstadoPresupuesto = ACotizar_
        P.fechaCreado = Now
        P.FechaEntrega = 0
        P.FormaPagoAnticipo = .FormaPagoAnticipo
        P.FormaPagoSaldo = .FormaPagoSaldo
        P.Gastos = .Gastos
        P.Id = 0
        P.manteOferta = .manteOferta
        Set P.moneda = .moneda
        P.PorcentajeManoObraMuerta = .PorcentajeManoObraMuerta
        P.PorcMas15 = .PorcMas15
        P.PorcMDO = .PorcMDO
        P.PorcMen10 = .PorcMen10
        P.PorcMen15 = .PorcMen15
        Set P.UsuarioCreado = funciones.GetUserObj
        P.VencimientoPresupuesto = P.VencimientoPresupuesto
        If DAOPresupuestos.Save(P) Then
            ReCotizar = True
        End If

    End With


End Function
