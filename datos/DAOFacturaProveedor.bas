Attribute VB_Name = "DAOFacturaProveedor"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Public Function Save(fc As clsFacturaProveedor) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    If Not Guardar(fc) Then GoTo err1
    conectar.CommitTransaction
    Exit Function
err1:
    Save = False
    conectar.RollBackTransaction
End Function


Public Function Guardar(fc As clsFacturaProveedor) As Boolean
    On Error GoTo err1

    Dim strsql As String
    If fc.id = 0 Then
        'guardo la factura
        strsql = "insert into AdminComprasFacturasProveedores  (id_usuario_creador,tipo_cambio,id_config_factura,estado,id_proveedor, fecha, impuesto_interno,  monto_neto, numero_factura, redondeo_iva, id_moneda,tipo_doc_contable, forma_de_pago_cta_cte) values (" & funciones.GetUserObj.id & " , " & fc.TipoCambio & ", " & fc.configFactura.id & "," & fc.estado & "," & fc.Proveedor.id & ", " & Escape(fc.FEcha) & "," & Escape(fc.ImpuestoInterno) & "," & Escape(fc.Monto) & "," & Escape(fc.numero) & "," & Escape(fc.Redondeo) & ", " & GetEntityId(fc.Moneda) & "," & fc.tipoDocumentoContable & ", " & Escape(fc.FormaPagoCuentaCorriente) & ")"
        conectar.execute strsql
        fc.id = conectar.UltimoId2
        A = DAOPercepcionesAplicadas.Save(fc)
        b = DAOIvaAplicado.Save(fc)
        c = DAOCuentasFacturas.Save(fc)
        If Not A Or Not b Or Not c Then
            Err.Raise 100
        End If


        Set obj1 = Nothing
        Set Obj = Nothing
        Set obj2 = Nothing
        DaoFacturaProveedorHistorial.agregar fc, "Factura creada"

    Else
        strsql = "update AdminComprasFacturasProveedores set tipo_cambio_pago=" & fc.TipoCambioPago & ", tipo_cambio=" & fc.TipoCambio & ", id_config_factura=" & fc.configFactura.id & ",estado=" & fc.estado & ",id_proveedor=" & fc.Proveedor.id & ",fecha=" & Escape(fc.FEcha) & ",impuesto_interno=" & Escape(fc.ImpuestoInterno) & ",monto_neto=" & Escape(fc.Monto) & ",numero_factura=" & Escape(fc.numero) & ",redondeo_iva=" & Escape(fc.Redondeo) & ", id_moneda =" & GetEntityId(fc.Moneda) & ", tipo_doc_contable=" & fc.tipoDocumentoContable & ", forma_de_pago_cta_cte = " & Escape(fc.FormaPagoCuentaCorriente) & " where id=" & fc.id
        If Not conectar.execute(strsql) Then GoTo err1
        b = DAOPercepcionesAplicadas.Save(fc)
        A = DAOIvaAplicado.Save(fc)
        c = DAOCuentasFacturas.Save(fc)



        If Not A Or Not b Or Not c Then
            Err.Raise 100
        End If

        Set obj1 = Nothing
        Set Obj = Nothing
        Set obj2 = Nothing
        DaoFacturaProveedorHistorial.agregar fc, "Factura modificada"
    End If
    Guardar = True
    Exit Function
err1:
    Guardar = False
    If Err.Number = 100 Then MsgBox "Se produjo algun error, no se  guadar�n los cambios!"
End Function
Public Function existeFactura(Factura As clsFacturaProveedor) As Boolean
    On Error GoTo err4
    Dim q As String
    q = "select count(id) as cantidad from AdminComprasFacturasProveedores where id_proveedor=" & Factura.Proveedor.id & " and numero_factura=" & Escape(Factura.numero) & " and id_config_factura=" & Escape(Factura.configFactura.id) & "  AND tipo_doc_contable=" & Escape(Factura.tipoDocumentoContable)


    If Factura.id <> 0 Then q = q & " and AdminComprasFacturasProveedores.id <> " & Factura.id

    Set rs = conectar.RSFactory(q)
    If Not rs.EOF And Not rs.BOF Then
        existeFactura = rs!Cantidad > 0

    End If
    Exit Function
err4:
End Function
Public Function GetByDate(desde As Date, Optional hasta As Date) As Collection
    Dim Factura As clsFacturaProveedor
    Dim col As New Collection

    Set GetByDate = DAOFacturaProveedor.FindAll("fecha>='" & Format(desde, "yyyy-mm-dd") & "' and fecha<= '" & Format(hasta, "yyyy-mm-dd") & "'", False)
End Function
Public Function aprobar(fc As clsFacturaProveedor) As Boolean
    Set cn = conectar.obternerConexion
    On Error GoTo err121
    cn.BeginTrans
    Dim estadoAnterior As EstadoFacturaProveedor
    aprobar = True
    If fc.estado = EstadoFacturaProveedor.EnProceso Then



        fc.estado = EstadoFacturaProveedor.Aprobada
        cn.execute "update AdminComprasFacturasProveedores SET  estado=2 where id=" & fc.id
        DaoFacturaProveedorHistorial.agregar fc, "Factura aprobada"

        If Not fc.FormaPagoCuentaCorriente Then
            If Not DAOFacturaProveedor.PagarEnEfectivo(fc, fc.FEcha, False) Then GoTo err121
        End If
    Else
        MsgBox "No puede cambiar el estado de la factura, ya fue aprobada!", vbInformation, "Informaci�n"
    End If
    cn.CommitTrans
    Exit Function
err121:
    cn.RollbackTrans
    aprobar = False
    fc.estado = estadoAnterior
End Function



Public Function FindById(id As Long) As clsFacturaProveedor
    On Error GoTo err1
    Set FindById = FindAll(" AdminComprasFacturasProveedores.Id=" & id, True)(1)
    Exit Function
err1:
    Set FindById = Nothing

End Function

Public Function FindAll(Optional filtro As String = vbNullString, Optional withHistorial As Boolean = False, Optional orderBy As String = vbNullString, Optional soloPropias As Boolean = False) As Collection
    Dim indice As New Dictionary
    Dim q As String
    Dim rs As Recordset
    Dim col As New Collection
    q = "SELECT *, (SELECT max(id_orden_pago) FROM ordenes_pago_facturas inner join ordenes_pago on ordenes_pago_facturas.id_orden_pago=ordenes_pago.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND ordenes_pago.estado<>2  limit 1) as nro_orden" _
        & " From" _
        & " AdminComprasFacturasProveedores" _
        & " LEFT JOIN AdminConfigFacturasProveedor ON (AdminComprasFacturasProveedores.id_config_factura = AdminConfigFacturasProveedor.id)" _
        & " LEFT JOIN proveedores ON (AdminComprasFacturasProveedores.id_proveedor = proveedores.id)" _
        & " LEFT JOIN AdminConfigMonedas ON (AdminComprasFacturasProveedores.id_moneda = AdminConfigMonedas.id)" _
        & " LEFT JOIN AdminConfigIVAProveedor ON (AdminConfigFacturasProveedor.id_iva = AdminConfigIVAProveedor.id)" _
        & " LEFT JOIN AdminConfigIvaAlicuotas ON (AdminConfigFacturasProveedor.id = AdminConfigIvaAlicuotas.id_config_factura)" _
        & " LEFT JOIN AdminComprasFacturasProveedoresIva ON AdminComprasFacturasProveedoresIva.id_factura_proveedor=    AdminComprasFacturasProveedores.id " _
        & " LEFT JOIN AdminComprasFacturasProveedoresPercepciones ON AdminComprasFacturasProveedoresPercepciones.id_factura_proveedor=AdminComprasFacturasProveedores.id  " _
        & " LEFT JOIN AdminConfigPercepciones ON AdminComprasFacturasProveedoresPercepciones.id_percepcion=AdminConfigPercepciones.id " _
        & " LEFT JOIN AdminConfigIvaAlicuotas AS a1 ON AdminComprasFacturasProveedoresIva.id_iva=a1.id " _
        & " LEFT JOIN AdminComprasCuentasFacturas  ON (AdminComprasFacturasProveedores.id = AdminComprasCuentasFacturas.id_factura) " _
        & " LEFT JOIN AdminComprasCuentasContables  ON (AdminComprasCuentasFacturas.id_cuenta = AdminComprasCuentasContables.id) " _
        & " WHERE 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    If soloPropias Then
        q = q & " and AdminComprasFacturasProveedores.id_usuario_creador=" & funciones.GetUserObj.id

    End If

    If LenB(orderBy) > 0 Then
        q = q & " ORDER BY " & orderBy
    End If


    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, indice
    Dim F As clsFacturaProveedor
    Dim per As clsPercepcionesAplicadas
    Dim Iva As clsAlicuotaAplicada
    Dim cta As clsCuentaFactura

    While Not rs.EOF
        Set F = Map(rs, indice, "AdminComprasFacturasProveedores", "proveedores", "AdminConfigFacturasProveedor", "AdminConfigIVAProveedor", "AdminConfigMonedas")
        If funciones.BuscarEnColeccion(col, CStr(F.id)) Then
            Set F = col.item(CStr(F.id))
        Else
            If withHistorial Then
                F.Historial = DaoFacturaProveedorHistorial.getAllByIdFactura(F.id)
            End If
        End If

        If IsSomething(F.configFactura) Then
            F.configFactura.alicuotas.Add DAOAlicuotas.Map(rs, indice, "AdminConfigIvaAlicuotas")
        End If

        Set per = DAOPercepcionesAplicadas.Map(rs, indice, "AdminComprasFacturasProveedoresPercepciones", "AdminConfigPercepciones")
        If IsSomething(per) Then
            If Not funciones.BuscarEnColeccion(F.percepciones, CStr(per.id)) Then
                If per.id <> 0 Then F.percepciones.Add per, CStr(per.id)
            End If
        End If

        Set cta = DAOCuentasFacturas.Map(rs, indice, "AdminComprasCuentasFacturas", "AdminComprasCuentasContables")
        If IsSomething(cta) Then
            If Not funciones.BuscarEnColeccion(F.cuentasContables, CStr(cta.id)) Then
                F.cuentasContables.Add cta, CStr(cta.id)
            End If
        End If

        Set Iva = DAOIvaAplicado.Map(rs, indice, "AdminComprasFacturasProveedoresIva", "a1")
        If IsSomething(Iva) Then
            If Not funciones.BuscarEnColeccion(F.IvaAplicado, CStr(Iva.id)) Then
                F.IvaAplicado.Add Iva, CStr(Iva.id)
            End If
        End If

        If Not funciones.BuscarEnColeccion(col, CStr(F.id)) Then col.Add F, CStr(F.id)
        rs.MoveNext
    Wend
    Set FindAll = col
End Function
Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaProveedor As String = vbNullString, _
                    Optional tablaAdminConfigFacturasProveedor As String = vbNullString, _
                    Optional tablaAdminConfigIVAProveedor As String = vbNullString, _
                    Optional tablaMoneda As String = vbNullString) As clsFacturaProveedor

    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim fc As clsFacturaProveedor

    If id > 0 Then
        Set fc = New clsFacturaProveedor
        fc.id = id
        fc.tipoDocumentoContable = GetValue(rs, indice, tabla, "tipo_doc_contable")
        fc.estado = GetValue(rs, indice, tabla, "estado")
        fc.FEcha = GetValue(rs, indice, tabla, "fecha")
        fc.ImpuestoInterno = GetValue(rs, indice, tabla, "impuesto_interno")
        fc.Monto = GetValue(rs, indice, tabla, "monto_neto")
        fc.numero = GetValue(rs, indice, tabla, "numero_factura")
        fc.Redondeo = GetValue(rs, indice, tabla, "redondeo_iva")
        'fc.ConceptoNoGravado = GetValue(rs, indice, tabla, "no_gravado")
        fc.FormaPagoCuentaCorriente = GetValue(rs, indice, tabla, "forma_de_pago_cta_cte")
        fc.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")
        fc.TipoCambioPago = GetValue(rs, indice, tabla, "tipo_cambio_pago")
        fc.TotalAbonado = GetValue(rs, indice, tabla, "total_abonado")
        fc.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")

        If LenB(tablaMoneda) > 0 Then Set fc.Moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        fc.Proveedor = DAOProveedor.Map2(rs, indice, tablaProveedor)
        If LenB(tablaAdminConfigFacturasProveedor) > 0 Then fc.configFactura = DAOConfigFacturaProveedor.Map(rs, indice, tablaAdminConfigFacturasProveedor, tablaAdminConfigIVAProveedor)

        If indice.Exists(".nro_orden") Then fc.OrdenPagoId = GetValue(rs, indice, vbNullString, "nro_orden")

    End If

    Set Map = fc
End Function


Public Function FindAllAlicuotasIVA() As Collection
    Dim q As String
    q = "SELECT DISTINCT alicuota FROM AdminConfigIvaAlicuotas ORDER BY alicuota DESC"
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim col As New Collection

    While Not rs.EOF
        col.Add rs.Fields("alicuota").value
        rs.MoveNext
    Wend

    Set FindAllAlicuotasIVA = col
End Function

Public Function Delete(facid As Long) As Boolean
    On Error GoTo E

    Dim q As String
    conectar.BeginTransaction
    Dim facProv As clsFacturaProveedor
    Set facProv = DAOFacturaProveedor.FindById(facid)
    Dim facsOrphan As New Collection

    'para borrar fijarse que no este en ninguna orden de pago, si esta en alguna, la op no debe estar aprobada para poder sacarla de ahi
    Dim op As OrdenPago
    Set op = DAOOrdenPago.FindByFacturaId(facid)

    If IsSomething(op) Then
        If op.estado = EstadoOrdenPago_Aprobada Then
            Delete = False
            MsgBox "La factura ya se encuentra incluida en una orden de pago aprobada, no se puede eliminar.", vbExclamation
            Exit Function
        ElseIf op.estado = EstadoOrdenPago_pendiente Or op.estado = EstadoOrdenPago_Anulada Then

            Dim facs As Collection
            Dim F As clsFacturaProveedor
            Set facs = op.FacturasProveedor
            For Each F In facs
                If F.id <> facid Then
                    facsOrphan.Add F
                End If
            Next F

            If Not DAOOrdenPago.Delete(op.id, False) Then GoTo E

        End If
    End If


    q = "DELETE FROM AdminComprasFacturasProveedoresPercepciones WHERE id_factura_proveedor = " & facid
    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM AdminComprasFacturasProveedores WHERE id = " & facid
    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM AdminComprasFacturasProveedoresHistorial WHERE id_factura = " & facid
    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM AdminComprasFacturasProveedoresIva WHERE id_factura_proveedor = " & facid
    If Not conectar.execute(q) Then GoTo E


    conectar.CommitTransaction

    Delete = True
    If facsOrphan.count > 0 Then
        MsgBox "El comprobante " & facProv.NumeroFormateado & " fue eliminado as� como tambi�n la Orden de Pago N� " & op.id & " a la que pertenecia." & vbNewLine & "Los siguientes comprobantes formaban parte de esa orden de pago y ahora se encuentran en estado [Pendiente] para ser incluidos en alguna otra orden de pago:" & vbNewLine & funciones.JoinCollectionValues(facsOrphan, vbNewLine, "NumeroFormateado"), vbInformation + vbOKOnly
    Else
        If IsSomething(op) Then
            MsgBox "El comprobante " & facProv.NumeroFormateado & " fue eliminado as� como tambi�n la Orden de Pago N� " & op.id & " a la que pertenecia.", vbInformation + vbOKOnly
        End If
    End If

    Exit Function
E:
    conectar.RollBackTransaction
    Delete = False
End Function

Public Function FindAllByOrdenPago(ByVal opid As Long) As Collection
    Set FindAllByOrdenPago = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id IN (SELECT id_factura_proveedor FROM ordenes_pago_facturas WHERE id_orden_pago = " & opid & ")")
End Function

Public Function PagarEnEfectivo(fac As clsFacturaProveedor, fechaPago As Date, insideTransaction As Boolean) As Boolean

    On Error GoTo eh

    If insideTransaction Then conectar.BeginTransaction

    Dim op As New OrdenPago
    op.FacturasProveedor.Add fac
    fac.TotalAbonado = fac.Total
    op.FEcha = fechaPago
    op.estado = EstadoOrdenPago_pendiente
    Set op.Moneda = fac.Moneda



    Dim opeCaja As New operacion
    opeCaja.Pertenencia = OrigenOperacion.caja
    opeCaja.Monto = fac.Total
    Set opeCaja.Moneda = fac.Moneda
    opeCaja.FechaOperacion = op.FEcha


    opeCaja.FechaCarga = Now
    Set opeCaja.caja = DAOCaja.FindById(1)
    opeCaja.EntradaSalida = OPSalida
    op.OperacionesCaja.Add opeCaja



    Dim d As New clsDTOPadronIIBB
    Set d = DTOPadronIIBB.FindByCUIT(fac.Proveedor.Cuit, TipoPadronRetencion)
    op.Alicuota = d.Alicuota



    'cuando se paga en efectivo no hay retenciones
    'Dim colRet As Collection
    'Set colRet = DAORetenciones.FindAllEsAgente
    'Dim d2 As Dictionary
    'Set d2 = DAOCertificadoRetencion.VerPosibleRetenciones(op.FacturasProveedor, colRet, op.Alicuota)
    'Dim totRet As Double
    'totRet = 0
    'For Each ret In colRet
    '    totRet = totRet + d2.Item(CStr(ret.Id))
    'Next ret

    op.StaticTotalFacturas = funciones.RedondearDecimales(MonedaConverter.Convertir(IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.Total * -1, fac.Total), fac.Moneda.id, op.Moneda.id))
    op.StaticTotalFacturasNG = funciones.RedondearDecimales(MonedaConverter.Convertir(IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.NetoGravado * -1, fac.NetoGravado), fac.Moneda.id, op.Moneda.id))
    op.StaticTotalOrigenes = op.TotalOrigenes
    op.StaticTotalRetenido = 0    'funciones.RedondearDecimales(totRet)

    PagarEnEfectivo = DAOOrdenPago.Guardar(op, True)

    If PagarEnEfectivo Then
        PagarEnEfectivo = DAOOrdenPago.aprobar(op, False)
    Else
        GoTo eh
    End If

    If insideTransaction Then conectar.CommitTransaction

    PagarEnEfectivo = True

    fac.estado = EstadoFacturaProveedor.Saldada

    Exit Function
eh:
    PagarEnEfectivo = False
    If insideTransaction Then conectar.RollBackTransaction
End Function


Public Function ExportarColeccion(col As Collection) As Boolean
    On Error GoTo err1
    ExportarColeccion = True
    Dim detalle As DetalleOrdenTrabajo
    Dim Entregas As Collection
    Dim remitoDetalle As remitoDetalle

    Dim xlWorkbook As New Excel.Workbook
    Dim xlWorksheet As New Excel.Worksheet
    Dim xlApplication As New Excel.Application

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    'fila, columna

    Dim offset As Long
    offset = 3
    xlWorksheet.Cells(offset, 1).value = "Raz�n Social"
    xlWorksheet.Cells(offset, 2).value = "Comprobante"
    xlWorksheet.Cells(offset, 3).value = "Fecha"
    xlWorksheet.Cells(offset, 4).value = "Moneda"
    xlWorksheet.Cells(offset, 5).value = "NG"
    xlWorksheet.Cells(offset, 6).value = "IVA"
    xlWorksheet.Cells(offset, 7).value = "No Gravado"
    xlWorksheet.Cells(offset, 8).value = "Percepciones"
    xlWorksheet.Cells(offset, 9).value = "Imp. Total"
    xlWorksheet.Cells(offset, 10).value = "Total"
    xlWorksheet.Cells(offset, 11).value = "Cta. Contable"
    xlWorksheet.Cells(offset, 12).value = "Estado"
    xlWorksheet.Cells(offset, 13).value = "Forma de Pago"
    xlWorksheet.Cells(offset, 14).value = "Orden de Pago"
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 14)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 14)).Interior.Color = &HC0C0C0


    '.Borders.LineStyle = xlContinuous

    Dim fac As clsFacturaProveedor
    Dim initoffset As Long
    initoffset = offset
    Dim c As Integer
    Dim Total As Double
    Dim totalneto As Double
    Dim totalno As Double
    Dim totIva As Double


    For Each fac In col
        If fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        Total = Total + MonedaConverter.Convertir(fac.Total * c, fac.Moneda.id, MonedaConverter.Patron.id)
        totalneto = totalneto + MonedaConverter.Convertir(fac.Monto * c - fac.TotalNetoGravadoDiscriminado(0) * c, fac.Moneda.id, MonedaConverter.Patron.id)
        totalno = totalno + MonedaConverter.Convertir(fac.TotalNetoGravadoDiscriminado(0) * c, fac.Moneda.id, MonedaConverter.Patron.id)
        totIva = totIva + MonedaConverter.Convertir(fac.TotalIVA * c, fac.Moneda.id, MonedaConverter.Patron.id)



        offset = offset + 1
        xlWorksheet.Cells(offset, 1).value = fac.Proveedor.RazonSocial
        xlWorksheet.Cells(offset, 2).value = fac.NumeroFormateado
        xlWorksheet.Cells(offset, 3).value = fac.FEcha
        xlWorksheet.Cells(offset, 4).value = fac.Moneda.NombreCorto
        xlWorksheet.Cells(offset, 5).value = fac.TotalNetoGravadoDiscriminado(0)
        xlWorksheet.Cells(offset, 6).value = fac.TotalIVA
        xlWorksheet.Cells(offset, 7).value = fac.TotalNetoGravadoDiscriminado(0)
        xlWorksheet.Cells(offset, 8).value = fac.totalPercepciones
        xlWorksheet.Cells(offset, 9).value = fac.ImpuestoInterno
        xlWorksheet.Cells(offset, 10).value = fac.Total
        If fac.cuentasContables.count > 0 Then xlWorksheet.Cells(offset, 11).value = fac.cuentasContables.item(1).cuentas.codigo
        xlWorksheet.Cells(offset, 12).value = enums.enumEstadoFacturaProveedor(fac.estado)
        If fac.FormaPagoCuentaCorriente Then xlWorksheet.Cells(offset, 13).value = "Cta. Cte." Else xlWorksheet.Cells(offset, 13).value = "Contado"
        xlWorksheet.Cells(offset, 14).value = fac.OrdenPagoId


        xlWorksheet.Range(xlWorksheet.Cells(initoffset, 1), xlWorksheet.Cells(offset, 14)).Borders.LineStyle = xlContinuous
    Next


    xlWorksheet.Cells(offset + 3, 2).value = "Total NG"
    xlWorksheet.Cells(offset + 4, 2).value = "Total NNG"
    xlWorksheet.Cells(offset + 5, 2).value = "Total Neto"
    xlWorksheet.Cells(offset + 6, 2).value = "Tota IVA"
    xlWorksheet.Cells(offset + 7, 2).value = "Total Filtrado"

    xlWorksheet.Cells(offset + 3, 3).value = totalneto
    xlWorksheet.Cells(offset + 4, 3).value = totalno
    xlWorksheet.Cells(offset + 5, 3).value = totalneto + totalno
    xlWorksheet.Cells(offset + 6, 3).value = totIva
    xlWorksheet.Cells(offset + 7, 3).value = Total
    xlWorksheet.Range(xlWorksheet.Cells(offset + 3, 2), xlWorksheet.Cells(offset + 7, 3)).Borders.LineStyle = xlContinuous
    xlWorksheet.Range(xlWorksheet.Cells(offset + 3, 2), xlWorksheet.Cells(offset + 7, 2)).Interior.Color = &HC0C0C0



    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True
    ''

    Dim ruta As String
    ruta = Environ$("TEMP")
    If LenB(ruta) = 0 Then ruta = Environ$("TMP")
    If LenB(ruta) = 0 Then ruta = App.path
    ruta = ruta & "\" & funciones.CreateGUID() & ".xls"

    xlWorkbook.SaveAs ruta

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    ShellExecute -1, "open", ruta, "", "", 4

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    Exit Function
err1:
    ExportarColeccion = False
End Function


