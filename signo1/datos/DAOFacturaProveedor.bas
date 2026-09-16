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

    '#209
    If DAOSubdiarios.ComprobanteComprasLiquidado(fc.Id) Then
        MsgBox "El comprobante se encuentra liquidado, no se puede volver a modificar o eliminar.", vbCritical
        Exit Function
    End If

    '#209
    Dim fecha_liqui_max As Date
    fecha_liqui_max = DAOSubdiarios.MaxFechaLiqui(False)
    If fc.FEcha <= fecha_liqui_max Then
        MsgBox "La fecha del comprobante es inválida ya que corresponde a un periodo ya liquidado", vbCritical, "Error"
        Exit Function
    End If


    Dim strsql As String
    If fc.Id = 0 Then
        'guardo la factura
        strsql = "insert into AdminComprasFacturasProveedores  (id_usuario_creador,tipo_cambio,id_config_factura,estado,id_proveedor, fecha, impuesto_interno,  monto_neto, numero_factura, redondeo_iva, id_moneda,tipo_doc_contable, forma_de_pago_cta_cte,ultima_actualizacion, ARCA) values (" & funciones.GetUserObj.Id & " , " & fc.TipoCambio & ", " & fc.configFactura.Id & "," & fc.estado & "," & fc.Proveedor.Id & ", " & Escape(fc.FEcha) & "," & Escape(fc.ImpuestoInterno) & "," & Escape(fc.Monto) & "," & Escape(fc.numero) & "," & Escape(fc.redondeo) & ", " & GetEntityId(fc.moneda) & "," & fc.tipoDocumentoContable & ", " & Escape(fc.FormaPagoCuentaCorriente) & "," & Escape(Now) & ", " & IIf(fc.EsArca, 1, 0) & ")"
        conectar.execute strsql
        fc.Id = conectar.UltimoId2
        A = DAOPercepcionesAplicadas.Save(fc)
        B = DAOIvaAplicado.Save(fc)
        c = DAOCuentasFacturas.Save(fc)
        If Not A Or Not B Or Not c Then
            Err.Raise 100
        End If


        Set obj1 = Nothing
        Set Obj = Nothing
        Set obj2 = Nothing
        DaoFacturaProveedorHistorial.agregar fc, "Factura creada"

    Else
        ' #195
        Dim fca As clsFacturaProveedor
        Set fca = DAOFacturaProveedor.FindById(fc.Id)
        If fca.UltimaActualizacion > Now Then Err.Raise 104, "fc", "La factura fué guardada en otra sesión, por favor actualice y vuelva a realizar la operación"

        strsql = "update AdminComprasFacturasProveedores set ultima_actualizacion=" & Escape(Now) & ", tipo_cambio_pago=" & fc.TipoCambioPago & ", tipo_cambio=" & fc.TipoCambio & ", id_config_factura=" & fc.configFactura.Id & ",estado=" & fc.estado & ",id_proveedor=" & fc.Proveedor.Id & ",fecha=" & Escape(fc.FEcha) & ",impuesto_interno=" & Escape(fc.ImpuestoInterno) & ",monto_neto=" & Escape(fc.Monto) & ",numero_factura=" & Escape(fc.numero) & ",redondeo_iva=" & Escape(fc.redondeo) & ", id_moneda =" & GetEntityId(fc.moneda) & ", tipo_doc_contable=" & fc.tipoDocumentoContable & ", forma_de_pago_cta_cte = " & Escape(fc.FormaPagoCuentaCorriente) & ", ARCA=" & IIf(fc.EsArca, 1, 0) & " where id=" & fc.Id
        If Not conectar.execute(strsql) Then GoTo err1
        B = DAOPercepcionesAplicadas.Save(fc)
        A = DAOIvaAplicado.Save(fc)
        c = DAOCuentasFacturas.Save(fc)

        If Not A Or Not B Or Not c Then
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
    If Err.Number = 100 Then MsgBox "Se produjo algun error, no se  guadarán los cambios!"
    If Err.Number = 104 Then MsgBox Err.Description
End Function


Public Function existeFactura(Factura As clsFacturaProveedor) As Boolean
    On Error GoTo err4
    Dim q As String
    q = "select count(id) as cantidad from AdminComprasFacturasProveedores where id_proveedor=" & Factura.Proveedor.Id & " and numero_factura=" & Escape(Factura.numero) & " and id_config_factura=" & Escape(Factura.configFactura.Id) & "  AND tipo_doc_contable=" & Escape(Factura.tipoDocumentoContable)


    If Factura.Id <> 0 Then q = q & " and AdminComprasFacturasProveedores.id <> " & Factura.Id

    Set rs = conectar.RSFactory(q)
    If Not rs.EOF And Not rs.BOF Then
        existeFactura = rs!cantidad > 0

    End If
    Exit Function
err4:
End Function

Public Function GetByDate(desde As Date, Optional hasta As Date) As Collection
    Set GetByDate = DAOFacturaProveedor.FindAll("fecha>='" & Format(desde, "yyyy-mm-dd") & "' and fecha<= '" & Format(hasta, "yyyy-mm-dd") & "'", False)

End Function


Public Function aprobar(fc As clsFacturaProveedor) As Boolean

    Set fc = DAOFacturaProveedor.FindById(fc.Id)


    Set cn = conectar.obternerConexion
    On Error GoTo err121
    cn.BeginTrans
    Dim fca As clsFacturaProveedor
    Set fca = DAOFacturaProveedor.FindById(fc.Id)
    If fca.UltimaActualizacion > Now Then Err.Raise 104, "fc", "La factura fué guardada en otra sesión, por favor actualice y vuelva a realizar la operación"

    Dim estadoAnterior As EstadoFacturaProveedor
    aprobar = True
    If fc.estado = EstadoFacturaProveedor.EnProceso Then



        fc.estado = EstadoFacturaProveedor.Aprobada
        cn.execute "update AdminComprasFacturasProveedores SET ultima_actualizacion= " & Escape(Now) & ", estado=2 where id=" & fc.Id
        DaoFacturaProveedorHistorial.agregar fc, "Factura aprobada"

        If Not fc.FormaPagoCuentaCorriente Then
            If Not DAOFacturaProveedor.PagarEnEfectivo(fc, fc.FEcha, False) Then GoTo err121
        End If
    Else
        '  MsgBox "No puede cambiar el estado de la factura, ya fue aprobada!", vbInformation, "Información"
        Err.Raise 4431, "Aprobar factura", "Error: La factura fué aprobada en otra sesión "

    End If
    cn.CommitTrans
    Exit Function
err121:
    If Err.Number = 104 Or Err.Number = 4431 Then
        MsgBox Err.Description
    Else
        MsgBox "Se produjo un error y no se pudo aprobar la factura", vbCritical
    End If
    cn.RollbackTrans
    aprobar = False
    fc.estado = estadoAnterior
End Function

Public Function ForzarEstadoAprobado(fc As clsFacturaProveedor) As Boolean

    Set fc = DAOFacturaProveedor.FindById(fc.Id)

    Set cn = conectar.obternerConexion
    On Error GoTo err121
    cn.BeginTrans
    Dim fca As clsFacturaProveedor
    Set fca = DAOFacturaProveedor.FindById(fc.Id)
    If fca.UltimaActualizacion > Now Then Err.Raise 104, "fc", "La factura fué guardada en otra sesión, por favor actualice y vuelva a realizar la operación"

    Dim estadoAnterior As EstadoFacturaProveedor
    ForzarEstadoAprobado = True
    If fc.estado = EstadoFacturaProveedor.EnProceso Then

        fc.estado = EstadoFacturaProveedor.Aprobada
        cn.execute "update AdminComprasFacturasProveedores SET ultima_actualizacion= " & Escape(Now) & ", estado=2 where id=" & fc.Id
        DaoFacturaProveedorHistorial.agregar fc, "Factura aprobada"

        If Not fc.FormaPagoCuentaCorriente Then
            If Not DAOFacturaProveedor.PagarEnEfectivo(fc, fc.FEcha, False) Then GoTo err121
        End If
    Else
        
        Err.Raise 4431, "Aprobar factura", "Error: La factura fué aprobada en otra sesión "

    End If
    cn.CommitTrans
    Exit Function
err121:
    If Err.Number = 104 Or Err.Number = 4431 Then
        MsgBox Err.Description
    Else
        MsgBox "Se produjo un error y no se pudo aprobar la factura", vbCritical
    End If
    cn.RollbackTrans
    ForzarEstadoAprobado = False
    fc.estado = estadoAnterior
End Function


Public Function FindById(Id As Long) As clsFacturaProveedor
    On Error GoTo err1
    Set FindById = FindAll(" AdminComprasFacturasProveedores.Id=" & Id, True)(1)
    Exit Function
err1:
    Set FindById = Nothing

End Function


Public Function FindAll(Optional filtro As String = vbNullString, _
                        Optional withHistorial As Boolean = False, _
                        Optional orderBy As String = vbNullString, _
                        Optional soloPropias As Boolean = False, _
                        Optional widhCompensatorios As Boolean = False, _
                        Optional preguntarSiSuperaLimite As Boolean = False) As Collection
                        
    On Error Resume Next
    On Error GoTo err1
    Dim indice As New Dictionary
    Dim q As String
    Dim rs As Recordset
    Dim col As New Collection
    
    If preguntarSiSuperaLimite Then
    
        ' 1. Primero contar registros
        Dim qCount As String
        qCount = "SELECT COUNT(*) AS total FROM AdminComprasFacturasProveedores " _
            & " LEFT JOIN AdminComprasCuentasFacturas  ON (AdminComprasFacturasProveedores.id = AdminComprasCuentasFacturas.id_factura) " _
            & " LEFT JOIN AdminComprasCuentasContables  ON (AdminComprasCuentasFacturas.id_cuenta = AdminComprasCuentasContables.id) " _
            & "WHERE 1=1 "
    
        If LenB(filtro) > 0 Then
            qCount = qCount & " AND " & filtro
        End If
    
        If soloPropias Then
            qCount = qCount & " AND AdminComprasFacturasProveedores.id_usuario_creador=" & funciones.GetUserObj.Id
        End If
    
        Set rs = conectar.RSFactory(qCount)
    
        totalRegistros = rs!total
    
        rs.Close
    
        ' 2. Verificar si supera el l�mite
        If totalRegistros > 1000 Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("Se encontraron " & totalRegistros & " registros. �Desea continuar?", vbYesNo + vbQuestion, "Confirmaci�n")
            If resp = vbNo Then
                Set FindAll = New Collection
                Exit Function
            End If
        End If
    
    End If
    
    q = "SELECT *, (SELECT max(id_orden_pago) FROM ordenes_pago_facturas inner join ordenes_pago on ordenes_pago_facturas.id_orden_pago=ordenes_pago.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND ordenes_pago.estado<>2  limit 1) as nro_orden"

    If widhCompensatorios Then
        q = q & ", (SELECT SUM(IF (tipo=1, c.importe,-c.importe)) FROM ordenes_pago_compensatorios c JOIN ordenes_pago op ON c.id_orden_pago=op.id AND op.estado=1 WHERE c.id_comprobante = AdminComprasFacturasProveedores.id AND c.cancelado=0 ) as total_compensado"

    Else

        q = q & ",0 as total_compensado "
    End If
    
    q = q & ",IFNULL((SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id WHERE op1.estado=1 AND opf.id_factura_proveedor=AdminComprasFacturasProveedores.id),0) AS total_abonado"
    q = q & ",IFNULL((SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id WHERE op1.estado=1 AND opf.id_factura_proveedor=AdminComprasFacturasProveedores.id),0) AS neto_gravado_abonado "
    q = q & ",IFNULL((SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id WHERE op1.estado=1 AND opf.id_factura_proveedor=AdminComprasFacturasProveedores.id),0) AS otros_abonado "
    q = q & " ,  CONVERT((SELECT IFNULL(GROUP_CONCAT(id_orden_pago),'-') FROM ordenes_pago_facturas INNER JOIN ordenes_pago ON ordenes_pago_facturas.id_orden_pago=ordenes_pago.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND ordenes_pago.estado<>2 ),NCHAR) AS ordenes_pago "
    q = q & ",   CONVERT((SELECT IFNULL(GROUP_CONCAT(numero_liq),'-') From liquidaciones_caja_facturas INNER JOIN liquidaciones_caja ON liquidaciones_caja_facturas.id_liquidacion_caja=liquidaciones_caja.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND liquidaciones_caja.estado<>2),NCHAR) AS num_liquidaciones_caja "
    q = q & " ,  CONVERT((SELECT IFNULL(GROUP_CONCAT(estado),'-') FROM ordenes_pago_facturas INNER JOIN ordenes_pago ON ordenes_pago_facturas.id_orden_pago=ordenes_pago.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND ordenes_pago.estado<>2 ),NCHAR) AS ordenes_pago_estado "
    q = q & ",   CONVERT((SELECT IFNULL(GROUP_CONCAT(estado),'-') From liquidaciones_caja_facturas INNER JOIN liquidaciones_caja ON liquidaciones_caja_facturas.id_liquidacion_caja=liquidaciones_caja.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND liquidaciones_caja.estado<>2),NCHAR) AS liquidaciones_caja_estado "
    
    q = q & " From" _
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
        & " LEFT JOIN usuarios ON AdminComprasFacturasProveedores.id_usuario_creador=usuarios.id " _
        & " LEFT JOIN liquidaciones_caja_facturas ON (AdminComprasFacturasProveedores.id = liquidaciones_caja_facturas.id_factura_proveedor) " _
        & " WHERE 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    If soloPropias Then
        q = q & " and AdminComprasFacturasProveedores.id_usuario_creador=" & funciones.GetUserObj.Id

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
    Dim op As OrdenPago

    While Not rs.EOF
        Set F = Map(rs, indice, "AdminComprasFacturasProveedores", "proveedores", "AdminConfigFacturasProveedor", "AdminConfigIVAProveedor", "AdminConfigMonedas")

        Dim neto_gravado_liquidado As Variant
        neto_gravado_liquidado = rs!neto_gravado_liquidado
        If Not IsNull(neto_gravado_liquidado) Then
            F.NetoGravadoAbonadoGlobal = rs!neto_gravado_abonado + neto_gravado_liquidado
        Else
            F.NetoGravadoAbonadoGlobal = rs!neto_gravado_abonado
        End If

        Dim otros_liquidado As Variant
        otros_liquidado = rs!otros_liquidado
        If Not IsNull(otros_liquidado) Then
            F.OtrosAbonadoGlobal = rs!otros_abonado + otros_liquidado
        Else
            F.OtrosAbonadoGlobal = rs!otros_abonado
        End If

        F.OrdenesPagoId = rs!ordenes_pago
        'F.LiquidacionesCajaId = rs!liquidaciones_caja
        F.LiquidacionesCajaId = rs!num_liquidaciones_caja
        ''        Debug.Print (num_liquidaciones_caja)

        F.OrdenesPagoEstado = rs!ordenes_pago_estado
        F.LiquidacionesCajaEstado = rs!liquidaciones_caja_estado


        If funciones.BuscarEnColeccion(col, CStr(F.Id)) Then
            Set F = col.item(CStr(F.Id))
        Else
            If withHistorial Then
                F.Historial = DaoFacturaProveedorHistorial.getAllByIdFactura(F.Id)
            End If
        End If

        If IsSomething(F.configFactura) Then
            F.configFactura.alicuotas.Add DAOAlicuotas.Map(rs, indice, "AdminConfigIvaAlicuotas")
        End If

        Set per = DAOPercepcionesAplicadas.Map(rs, indice, "AdminComprasFacturasProveedoresPercepciones", "AdminConfigPercepciones")
        If IsSomething(per) Then
            If Not funciones.BuscarEnColeccion(F.percepciones, CStr(per.Id)) Then
                If per.Id <> 0 Then F.percepciones.Add per, CStr(per.Id)
            End If
        End If

       
        Set cta = DAOCuentasFacturas.Map(rs, indice, "AdminComprasCuentasFacturas", "AdminComprasCuentasContables")
        If IsSomething(cta) Then
            If Not funciones.BuscarEnColeccion(F.cuentasContables, CStr(cta.Id)) Then
                F.cuentasContables.Add cta, CStr(cta.Id)
            End If
        End If

        Set Iva = DAOIvaAplicado.Map(rs, indice, "AdminComprasFacturasProveedoresIva", "a1")
        If IsSomething(Iva) Then
            If Not funciones.BuscarEnColeccion(F.IvaAplicado, CStr(Iva.Id)) Then
                F.IvaAplicado.Add Iva, CStr(Iva.Id)
            End If
        End If

        If Not funciones.BuscarEnColeccion(col, CStr(F.Id)) Then col.Add F, CStr(F.Id)
        rs.MoveNext

    Wend


    Set FindAll = col
    
    Exit Function

err1:
    Set FindAll = Nothing
    MsgBox "Error: " & Err.Description, vbExclamation
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaProveedor As String = vbNullString, _
                    Optional tablaAdminConfigFacturasProveedor As String = vbNullString, _
                    Optional tablaAdminConfigIVAProveedor As String = vbNullString, _
                    Optional tablaMoneda As String = vbNullString) As clsFacturaProveedor

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim fc As clsFacturaProveedor

    If Id > 0 Then
        Set fc = New clsFacturaProveedor
        fc.Id = Id
        fc.tipoDocumentoContable = GetValue(rs, indice, tabla, "tipo_doc_contable")
        fc.estado = GetValue(rs, indice, tabla, "estado")
        fc.FEcha = GetValue(rs, indice, tabla, "fecha")
        fc.ImpuestoInterno = GetValue(rs, indice, tabla, "impuesto_interno")
        fc.Monto = GetValue(rs, indice, tabla, "monto_neto")
        fc.numero = GetValue(rs, indice, tabla, "numero_factura")
        fc.redondeo = GetValue(rs, indice, tabla, "redondeo_iva")
        'fc.ConceptoNoGravado = GetValue(rs, indice, tabla, "no_gravado")
        fc.FormaPagoCuentaCorriente = GetValue(rs, indice, tabla, "forma_de_pago_cta_cte")
        fc.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")
        fc.TipoCambioPago = GetValue(rs, indice, tabla, "tipo_cambio_pago")

        fc.totalAbonado = GetValue(rs, indice, tabla, "total_abonado")
        fc.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")

        If indice.Exists(".neto_gravado_liquidado") Then fc.NetoGravadoAbonadoGlobal = fc.NetoGravadoAbonadoGlobal + GetValue(rs, indice, vbNullString, "neto_gravado_liquidado")
        If indice.Exists(".otros_liquidado") Then fc.OtrosAbonadoGlobal = fc.OtrosAbonadoGlobal + GetValue(rs, indice, vbNullString, "otros_liquidado")
        
        If indice.Exists(".total_liquidado") Then
            fc.TotalLiquidado = GetValue(rs, indice, vbNullString, "total_liquidado")
        Else
            fc.TotalLiquidado = 0
        End If

        fc.UltimaActualizacion = GetValue(rs, indice, tabla, "ultima_actualizacion")
        
        If indice.Exists(tabla & ".ARCA") Then
            fc.EsArca = CBool(GetValue(rs, indice, tabla, "ARCA"))
        Else
            fc.EsArca = False
        End If
        
        Set fc.UsuarioCarga = DAOUsuarios.Map(rs, indice, "usuarios")

        If LenB(tablaMoneda) > 0 Then Set fc.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)

        fc.Proveedor = DAOProveedor.Map2(rs, indice, tablaProveedor)
        If LenB(tablaAdminConfigFacturasProveedor) > 0 Then fc.configFactura = DAOConfigFacturaProveedor.Map(rs, indice, tablaAdminConfigFacturasProveedor, tablaAdminConfigIVAProveedor)

        If indice.Exists(".nro_orden") Then fc.OrdenPagoID = GetValue(rs, indice, vbNullString, "nro_orden")

        If indice.Exists(".total_compensado") Then fc.TotalCompensado = GetValue(rs, indice, vbNullString, "total_compensado")

        If indice.Exists(".num_liquidaciones_caja") Then fc.LiquidacionesCajaId = GetValue(rs, indice, vbNullString, "num_liquidaciones_caja")
        
        If indice.Exists(".ordenes_pago_estado") Then fc.OrdenesPagoEstado = GetValue(rs, indice, vbNullString, "ordenes_pago_estado")
        
        If indice.Exists(".liquidaciones_caja_estado") Then fc.LiquidacionesCajaEstado = GetValue(rs, indice, vbNullString, "liquidaciones_caja_estado")
        
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



    '#209
    If DAOSubdiarios.ComprobanteComprasLiquidado(facid) Then
        MsgBox "El comprobante se encuentra liquidado, no se puede eliminar o modificar.", vbCritical
        Exit Function
    End If



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
                If F.Id <> facid Then
                    facsOrphan.Add F
                End If
            Next F

            If Not DAOOrdenPago.Delete(op.Id, False) Then GoTo E

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
        MsgBox "El comprobante " & facProv.NumeroFormateado & " fue eliminado así como también la Orden de Pago Nº " & op.Id & " a la que pertenecia." & vbNewLine & "Los siguientes comprobantes formaban parte de esa orden de pago y ahora se encuentran en estado [Pendiente] para ser incluidos en alguna otra orden de pago:" & vbNewLine & funciones.JoinCollectionValues(facsOrphan, vbNewLine, "NumeroFormateado"), vbInformation + vbOKOnly
    Else
        If IsSomething(op) Then
            MsgBox "El comprobante " & facProv.NumeroFormateado & " fue eliminado así como también la Orden de Pago Nº " & op.Id & " a la que pertenecia.", vbInformation + vbOKOnly
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


'''Public Function FindAllByLiquidacionCaja(ByVal lcid As Long) As Collection
'''    Set FindAllByLiquidacionCaja = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id IN (SELECT id_factura_proveedor FROM liquidaciones_caja_facturas WHERE id_liquidacion_caja = " & lcid & ")")
'''End Function


Public Function FindAllByLiquidacionCaja(ByVal lcid As Long) As Collection
    On Error GoTo err1
    
    Dim indice As New Dictionary
    Dim q As String
    Dim rs As ADODB.Recordset
    Dim col As New Collection
    
    Dim F As clsFacturaProveedor
    Dim per As clsPercepcionesAplicadas
    Dim Iva As clsAlicuotaAplicada
    Dim cta As clsCuentaFactura
'''    Dim alic As clsAlicuota
    
    q = ""
    
    q = q & "SELECT "
    q = q & " AdminComprasFacturasProveedores.*, "
    q = q & " AdminConfigFacturasProveedor.*, "
    q = q & " proveedores.*, "
    q = q & " AdminConfigMonedas.*, "
    q = q & " AdminConfigIVAProveedor.*, "
    q = q & " AdminConfigIvaAlicuotas.*, "
    q = q & " AdminComprasFacturasProveedoresIva.*, "
    q = q & " AdminComprasFacturasProveedoresPercepciones.*, "
    q = q & " AdminConfigPercepciones.*, "
    q = q & " a1.*, "
    q = q & " AdminComprasCuentasFacturas.*, "
    q = q & " AdminComprasCuentasContables.*, "
    q = q & " usuarios.*, "
    q = q & " IFNULL(lcf.total_liquidado,0) AS total_liquidado, "
    q = q & " IFNULL(lcf.neto_gravado_liquidado,0) AS neto_gravado_liquidado, "
    q = q & " IFNULL(lcf.otros_liquidado,0) AS otros_liquidado, "
    q = q & " (SELECT MAX(id_orden_pago) "
    q = q & "    FROM ordenes_pago_facturas "
    q = q & "    INNER JOIN ordenes_pago ON ordenes_pago_facturas.id_orden_pago = ordenes_pago.id "
    q = q & "   WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & "     AND ordenes_pago.estado <> 2 "
    q = q & "   LIMIT 1) AS nro_orden, "
    q = q & " 0 AS total_compensado, "
    q = q & " IFNULL((SELECT SUM(total_abonado) "
    q = q & "           FROM ordenes_pago_facturas opf "
    q = q & "           JOIN ordenes_pago op1 ON opf.id_orden_pago = op1.id "
    q = q & "          WHERE op1.estado = 1 "
    q = q & "            AND opf.id_factura_proveedor = AdminComprasFacturasProveedores.id),0) AS total_abonado, "
    q = q & " IFNULL((SELECT SUM(neto_gravado_abonado) "
    q = q & "           FROM ordenes_pago_facturas opf "
    q = q & "           JOIN ordenes_pago op1 ON opf.id_orden_pago = op1.id "
    q = q & "          WHERE op1.estado = 1 "
    q = q & "            AND opf.id_factura_proveedor = AdminComprasFacturasProveedores.id),0) AS neto_gravado_abonado, "
    q = q & " IFNULL((SELECT SUM(otros_abonado) "
    q = q & "           FROM ordenes_pago_facturas opf "
    q = q & "           JOIN ordenes_pago op1 ON opf.id_orden_pago = op1.id "
    q = q & "          WHERE op1.estado = 1 "
    q = q & "            AND opf.id_factura_proveedor = AdminComprasFacturasProveedores.id),0) AS otros_abonado, "
    q = q & " CONVERT((SELECT IFNULL(GROUP_CONCAT(id_orden_pago),'-') "
    q = q & "            FROM ordenes_pago_facturas "
    q = q & "            INNER JOIN ordenes_pago ON ordenes_pago_facturas.id_orden_pago = ordenes_pago.id "
    q = q & "           WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & "             AND ordenes_pago.estado <> 2), NCHAR) AS ordenes_pago, "
    q = q & " CONVERT((SELECT IFNULL(GROUP_CONCAT(numero_liq),'-') "
    q = q & "            FROM liquidaciones_caja_facturas "
    q = q & "            INNER JOIN liquidaciones_caja ON liquidaciones_caja_facturas.id_liquidacion_caja = liquidaciones_caja.id "
    q = q & "           WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & "             AND liquidaciones_caja.estado <> 2), NCHAR) AS num_liquidaciones_caja, "
    q = q & " CONVERT((SELECT IFNULL(GROUP_CONCAT(estado),'-') "
    q = q & "            FROM ordenes_pago_facturas "
    q = q & "            INNER JOIN ordenes_pago ON ordenes_pago_facturas.id_orden_pago = ordenes_pago.id "
    q = q & "           WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & "             AND ordenes_pago.estado <> 2), NCHAR) AS ordenes_pago_estado, "
    q = q & " CONVERT((SELECT IFNULL(GROUP_CONCAT(estado),'-') "
    q = q & "            FROM liquidaciones_caja_facturas "
    q = q & "            INNER JOIN liquidaciones_caja ON liquidaciones_caja_facturas.id_liquidacion_caja = liquidaciones_caja.id "
    q = q & "           WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & "             AND liquidaciones_caja.estado <> 2), NCHAR) AS liquidaciones_caja_estado "
    
    q = q & " FROM AdminComprasFacturasProveedores "
    q = q & " LEFT JOIN AdminConfigFacturasProveedor "
    q = q & "   ON AdminComprasFacturasProveedores.id_config_factura = AdminConfigFacturasProveedor.id "
    q = q & " LEFT JOIN proveedores "
    q = q & "   ON AdminComprasFacturasProveedores.id_proveedor = proveedores.id "
    q = q & " LEFT JOIN AdminConfigMonedas "
    q = q & "   ON AdminComprasFacturasProveedores.id_moneda = AdminConfigMonedas.id "
    q = q & " LEFT JOIN AdminConfigIVAProveedor "
    q = q & "   ON AdminConfigFacturasProveedor.id_iva = AdminConfigIVAProveedor.id "
    q = q & " LEFT JOIN AdminConfigIvaAlicuotas "
    q = q & "   ON AdminConfigFacturasProveedor.id = AdminConfigIvaAlicuotas.id_config_factura "
    q = q & " LEFT JOIN AdminComprasFacturasProveedoresIva "
    q = q & "   ON AdminComprasFacturasProveedoresIva.id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & " LEFT JOIN AdminComprasFacturasProveedoresPercepciones "
    q = q & "   ON AdminComprasFacturasProveedoresPercepciones.id_factura_proveedor = AdminComprasFacturasProveedores.id "
    q = q & " LEFT JOIN AdminConfigPercepciones "
    q = q & "   ON AdminComprasFacturasProveedoresPercepciones.id_percepcion = AdminConfigPercepciones.id "
    q = q & " LEFT JOIN AdminConfigIvaAlicuotas AS a1 "
    q = q & "   ON AdminComprasFacturasProveedoresIva.id_iva = a1.id "
    q = q & " LEFT JOIN AdminComprasCuentasFacturas "
    q = q & "   ON AdminComprasFacturasProveedores.id = AdminComprasCuentasFacturas.id_factura "
    q = q & " LEFT JOIN AdminComprasCuentasContables "
    q = q & "   ON AdminComprasCuentasFacturas.id_cuenta = AdminComprasCuentasContables.id "
    q = q & " LEFT JOIN usuarios "
    q = q & "   ON AdminComprasFacturasProveedores.id_usuario_creador = usuarios.id "
    q = q & " INNER JOIN liquidaciones_caja_facturas lcf "
    q = q & "   ON AdminComprasFacturasProveedores.id = lcf.id_factura_proveedor "
    q = q & " WHERE lcf.id_liquidacion_caja = " & lcid
    
    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, indice
    
    While Not rs.EOF
        Set F = Map(rs, indice, "AdminComprasFacturasProveedores", "proveedores", "AdminConfigFacturasProveedor", "AdminConfigIVAProveedor", "AdminConfigMonedas")
        
        If Not F Is Nothing Then
            
            If funciones.BuscarEnColeccion(col, CStr(F.Id)) Then
                Set F = col.item(CStr(F.Id))
            Else
                F.TotalLiquidado = GetValue(rs, indice, vbNullString, "total_liquidado")
                F.NetoGravadoAbonadoGlobal = GetValue(rs, indice, vbNullString, "neto_gravado_abonado") + GetValue(rs, indice, vbNullString, "neto_gravado_liquidado")
                F.OtrosAbonadoGlobal = GetValue(rs, indice, vbNullString, "otros_abonado") + GetValue(rs, indice, vbNullString, "otros_liquidado")
                F.OrdenesPagoId = GetValue(rs, indice, vbNullString, "ordenes_pago")
                F.LiquidacionesCajaId = GetValue(rs, indice, vbNullString, "num_liquidaciones_caja")
                F.OrdenesPagoEstado = GetValue(rs, indice, vbNullString, "ordenes_pago_estado")
                F.LiquidacionesCajaEstado = GetValue(rs, indice, vbNullString, "liquidaciones_caja_estado")
                
                col.Add F, CStr(F.Id)
            End If
            
'''            If IsSomething(F.configFactura) Then
'''                Set alic = DAOAlicuotas.Map(rs, indice, "AdminConfigIvaAlicuotas")
'''                If IsSomething(alic) Then
'''                    If Not funciones.BuscarEnColeccion(F.configFactura.alicuotas, CStr(alic.Id)) Then
'''                        F.configFactura.alicuotas.Add alic, CStr(alic.Id)
'''                    End If
'''                End If
'''            End If
            
            Set per = DAOPercepcionesAplicadas.Map(rs, indice, "AdminComprasFacturasProveedoresPercepciones", "AdminConfigPercepciones")
            If IsSomething(per) Then
                If Not funciones.BuscarEnColeccion(F.percepciones, CStr(per.Id)) Then
                    If per.Id <> 0 Then F.percepciones.Add per, CStr(per.Id)
                End If
            End If
            
            Set cta = DAOCuentasFacturas.Map(rs, indice, "AdminComprasCuentasFacturas", "AdminComprasCuentasContables")
            If IsSomething(cta) Then
                If Not funciones.BuscarEnColeccion(F.cuentasContables, CStr(cta.Id)) Then
                    F.cuentasContables.Add cta, CStr(cta.Id)
                End If
            End If
            
            Set Iva = DAOIvaAplicado.Map(rs, indice, "AdminComprasFacturasProveedoresIva", "a1")
            If IsSomething(Iva) Then
                If Not funciones.BuscarEnColeccion(F.IvaAplicado, CStr(Iva.Id)) Then
                    F.IvaAplicado.Add Iva, CStr(Iva.Id)
                End If
            End If
        End If
        
        rs.MoveNext
    Wend
    
    Set FindAllByLiquidacionCaja = col
    Exit Function

err1:
    Set FindAllByLiquidacionCaja = New Collection
    MsgBox "Error en FindAllByLiquidacionCaja: " & Err.Description, vbExclamation
End Function


Public Function PagarEnEfectivo(fac As clsFacturaProveedor, fechaPago As Date, insideTransaction As Boolean) As Boolean

    On Error GoTo eh

    If insideTransaction Then conectar.BeginTransaction

    Dim op As New OrdenPago
    op.FacturasProveedor.Add fac
    fac.totalAbonado = fac.total
    fac.TipoCambio = 1
    fac.NetoGravadoAbonado = fac.NetoGravado
    fac.OtrosAbonado = fac.total - fac.NetoGravado

    op.FEcha = fechaPago
    op.estado = EstadoOrdenPago_pendiente
    Set op.moneda = fac.moneda


    Dim opeCaja As New operacion
    opeCaja.Pertenencia = OrigenOperacion.caja
    opeCaja.Monto = fac.total
    If fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then
        opeCaja.Monto = fac.total * -1
    End If

    Set opeCaja.moneda = fac.moneda
    opeCaja.FechaOperacion = op.FEcha


    opeCaja.FechaCarga = Now
    Set opeCaja.caja = DAOCaja.FindById(1)
    opeCaja.EntradaSalida = OPSalida
    op.OperacionesCaja.Add opeCaja



    Dim d As New clsDTOPadronIIBB
    Set d = DTOPadronIIBB.FindByCUIT(fac.Proveedor.cuit, TipoPadronRetencion)
    op.alicuota = d.alicuota

    op.StaticTotalFacturas = funciones.RedondearDecimales(MonedaConverter.Convertir(IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.total * -1, fac.total), fac.moneda.Id, op.moneda.Id))
    op.StaticTotalFacturasNG = funciones.RedondearDecimales(MonedaConverter.Convertir(IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.NetoGravado * -1, fac.NetoGravado), fac.moneda.Id, op.moneda.Id))
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


Public Function ExportarColeccion(col As Collection, Optional ProgressBar As Object) As Boolean
    On Error GoTo err1

    ExportarColeccion = True

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")

    'Dim xlApplication As New Excel.Application
    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate
    
    Dim titulo As String
        titulo = "Reporte de Comprobantes de compra"

    With xlWorksheet.Range("A1:T1")
        .Merge
        .value = titulo
        .Font.Bold = True
        .HorizontalAlignment = -4108 ' xlCenter
    End With

    'fila, columna

    Dim offset As Long
    offset = 3
    xlWorksheet.Cells(offset, 1).value = "Cuit"
    xlWorksheet.Cells(offset, 2).value = "Razon Social"
    xlWorksheet.Cells(offset, 3).value = "Comprobante"
    xlWorksheet.Cells(offset, 4).value = "Fecha"
    xlWorksheet.Cells(offset, 5).value = "Moneda"
    xlWorksheet.Cells(offset, 6).value = "Neto Gravado"
    xlWorksheet.Cells(offset, 7).value = "IVA"
    xlWorksheet.Cells(offset, 8).value = "No Gravado"
    xlWorksheet.Cells(offset, 9).value = "Percepciones"
    xlWorksheet.Cells(offset, 10).value = "Imp. Total"
    xlWorksheet.Cells(offset, 11).value = "Total"
    xlWorksheet.Cells(offset, 12).value = "Saldo a Pagar"
    xlWorksheet.Cells(offset, 13).value = "Cta. Contable"
    xlWorksheet.Cells(offset, 14).value = "Estado"
    xlWorksheet.Cells(offset, 15).value = "Forma de Pago"
    xlWorksheet.Cells(offset, 16).value = "Orden de Pago"
    xlWorksheet.Cells(offset, 17).value = "Liquidacion C"
    xlWorksheet.Cells(offset, 18).value = "Tipo de Cambio"
    xlWorksheet.Cells(offset, 19).value = "Usuario"
    xlWorksheet.Cells(offset, 20).value = "ID"

    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 20)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 20)).Interior.Color = &HC0C0C0


    '.Borders.LineStyle = xlContinuous

    Dim fac As clsFacturaProveedor
    Dim initoffset As Long
    initoffset = offset

    Dim c As Integer
    Dim total As Double
    Dim totalneto As Double
    Dim totalno As Double
    Dim totIva As Double
    'Agregar DNEMER 03/02/2021
    Dim totalpercep As Double
    Dim TotalPendiente As Double

    ProgressBar.min = 0
    ProgressBar.max = col.count


    Dim d As Long
    d = 0

    For Each fac In col
        If fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        total = total + MonedaConverter.Convertir(fac.total * c, fac.moneda.Id, MonedaConverter.Patron.Id)
        totalneto = totalneto + MonedaConverter.Convertir(fac.Monto * c - fac.TotalNetoGravadoDiscriminado(0) * c, fac.moneda.Id, MonedaConverter.Patron.Id)
        totalno = totalno + MonedaConverter.Convertir(fac.TotalNetoGravadoDiscriminado(0) * c, fac.moneda.Id, MonedaConverter.Patron.Id)
        totIva = totIva + MonedaConverter.Convertir(fac.TotalIVA * c, fac.moneda.Id, MonedaConverter.Patron.Id)

        'Agrega DNEMER 03/02/2021
        totalpercep = totalpercep + fac.totalPercepciones * c
        'Agrega DNEMER 24/04/2023
        TotalPendiente = TotalPendiente + ((fac.total - (fac.NetoGravadoAbonadoGlobal + fac.OtrosAbonadoGlobal)) * c)



        If fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

        d = d + 1
        ProgressBar.value = d

        offset = offset + 1
        xlWorksheet.Cells(offset, 1).value = fac.Proveedor.cuit
        xlWorksheet.Cells(offset, 2).value = fac.Proveedor.RazonSocial
        xlWorksheet.Cells(offset, 3).value = fac.NumeroFormateado
        xlWorksheet.Cells(offset, 4).value = fac.FEcha
        xlWorksheet.Cells(offset, 5).value = fac.moneda.NombreCorto
        xlWorksheet.Cells(offset, 6).value = funciones.FormatearDecimales(fac.total - (fac.TotalIVA + fac.totalPercepciones + fac.ImpuestoInterno + fac.TotalNetoGravadoDiscriminado(0))) * i
        xlWorksheet.Cells(offset, 7).value = funciones.FormatearDecimales(fac.TotalIVA * i)
        xlWorksheet.Cells(offset, 8).value = funciones.FormatearDecimales(fac.TotalNetoGravadoDiscriminado(0) * i)
        xlWorksheet.Cells(offset, 9).value = funciones.FormatearDecimales(fac.totalPercepciones * i)
        xlWorksheet.Cells(offset, 10).value = funciones.FormatearDecimales(fac.ImpuestoInterno * i)
        xlWorksheet.Cells(offset, 11).value = funciones.FormatearDecimales(fac.total * i)

        xlWorksheet.Cells(offset, 12).value = funciones.FormatearDecimales(fac.total - (fac.NetoGravadoAbonadoGlobal + fac.OtrosAbonadoGlobal)) * i

        If fac.cuentasContables.count > 0 Then xlWorksheet.Cells(offset, 13).value = fac.cuentasContables.item(1).cuentas.codigo
        xlWorksheet.Cells(offset, 14).value = enums.enumEstadoFacturaProveedor(fac.estado)
        If fac.FormaPagoCuentaCorriente Then xlWorksheet.Cells(offset, 15).value = "Cta. Cte." Else xlWorksheet.Cells(offset, 15).value = "Contado"

        'Definir primero las celdas como texto para que Excel
        'no convierta las listas de n�meros a notaci�n cient�fica.
        xlWorksheet.Cells(offset, 16).NumberFormat = "@"
        xlWorksheet.Cells(offset, 17).NumberFormat = "@"
        
        xlWorksheet.Cells(offset, 16).value = CStr(fac.OrdenesPagoId)
        xlWorksheet.Cells(offset, 17).value = CStr(fac.LiquidacionesCajaId)

        xlWorksheet.Cells(offset, 18).value = fac.TipoCambio
        xlWorksheet.Cells(offset, 19).value = fac.UsuarioCarga.Usuario
        xlWorksheet.Cells(offset, 20).value = fac.Id


        xlWorksheet.Range(xlWorksheet.Cells(initoffset, 1), xlWorksheet.Cells(offset, 20)).Borders.LineStyle = xlContinuous
        
    Next


    xlWorksheet.Cells(offset + 3, 2).value = "Total NG"
    xlWorksheet.Cells(offset + 4, 2).value = "Total NNG"
    xlWorksheet.Cells(offset + 5, 2).value = "Total Neto"
    xlWorksheet.Cells(offset + 6, 2).value = "Tota IVA"
    xlWorksheet.Cells(offset + 7, 2).value = "Tota Percepciones"
    xlWorksheet.Cells(offset + 8, 2).value = "Tota Pendiente"
    xlWorksheet.Cells(offset + 9, 2).value = "Total Filtrado"

    xlWorksheet.Cells(offset + 3, 3).value = totalneto
    xlWorksheet.Cells(offset + 4, 3).value = totalno
    xlWorksheet.Cells(offset + 5, 3).value = totalneto + totalno
    xlWorksheet.Cells(offset + 6, 3).value = totIva
    xlWorksheet.Cells(offset + 7, 3).value = totalpercep
    xlWorksheet.Cells(offset + 8, 3).value = TotalPendiente
    xlWorksheet.Cells(offset + 9, 3).value = total

    xlWorksheet.Range(xlWorksheet.Cells(offset + 3, 2), xlWorksheet.Cells(offset + 9, 3)).Borders.LineStyle = xlContinuous
    xlWorksheet.Range(xlWorksheet.Cells(offset + 3, 2), xlWorksheet.Cells(offset + 9, 2)).Interior.Color = &HC0C0C0

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

    ProgressBar.value = 0

    Exit Function
err1:
    ExportarColeccion = False
End Function


Public Function CrearTablaTempComprobantes(facturas) As Boolean
    On Error GoTo err1

    CrearTablaTempComprobantes = True

    Dim fac As clsFacturaProveedor
    Dim strsql As String

    Set cn = conectar.obternerConexion
    cn.BeginTrans
    cn.execute "TRUNCATE sp_temporal.ComprobantesCargadosSP"
    cn.CommitTrans
    cn.BeginTrans

    For Each fac In facturas
        Dim partes() As String
        partes = Split(fac.numero, "-") ' Dividir el valor en dos partes

        Dim puntoDeVenta As String
        Dim NumeroDesde As String
        puntoDeVenta = CStr(CLng(partes(0))) ' Convertir a n�mero sin procesar
        NumeroDesde = CStr(CLng(partes(1))) ' Convertir a n�mero sin procesar
        
        If (fac.Id = 83569) Then
        Debug.Print ("aca")
        End If
        

        strsql = "INSERT INTO sp_temporal.ComprobantesCargadosSP (idcomprobante, numero, cuit, clave, puntodeventa, numerodesde)" _
                 & " VALUES (" & fac.Id & ", '" & fac.numero & "', " & fac.Proveedor.cuit & ", '" & puntoDeVenta + "-" + NumeroDesde + fac.Proveedor.cuit & "', '" & puntoDeVenta & "', '" & NumeroDesde & "')"

        cn.execute strsql
        
'''        Debug.Print (fac.Id)
    Next fac

    cn.CommitTrans
    Exit Function
err1:
    CrearTablaTempComprobantes = False
    MsgBox ("Error en el cbte | ID: " & fac.Id & " | N�: " & fac.numero)
    cn.RollbackTrans
End Function


'=========================================================
' ANTICIPOS PENDIENTES A UNA FECHA DETERMINADA
'
' Incluye pagos a cuenta realizados hasta la fecha
' de corte que todavia no estaban aplicados mediante
' una OP aprobada a esa fecha.
'
' No depende solamente del estado actual del anticipo.
'=========================================================
Public Function TotalAnticiposPendientesAl( _
    ByVal idProveedor As Long, _
    ByVal FechaCorte As Date, _
    ByVal IdMoneda As Long _
) As Double

    On Error GoTo err1

    Dim q As String
    Dim rs As Recordset

    Dim FechaHasta As Date

    Dim numeroError As Long
    Dim descripcionError As String

    TotalAnticiposPendientesAl = 0

    If idProveedor <= 0 Then
        Err.Raise 5, _
                  "TotalAnticiposPendientesAl", _
                  "El proveedor no es valido."
    End If

    'Trabajar solamente con la fecha, sin horario.
    FechaHasta = DateSerial( _
        Year(FechaCorte), _
        Month(FechaCorte), _
        Day(FechaCorte))

    '------------------------------------------------------
    ' BUSCAR ANTICIPOS EXISTENTES AL CIERRE
    '------------------------------------------------------
    q = "SELECT " _
      & "IFNULL(SUM(p.static_total_origen), 0) " _
      & "AS total_anticipos " _
      & "FROM pagos_a_cuenta p " _
      & "WHERE p.id_proveedor = " & idProveedor & " " _
      & "AND p.id_moneda = " & IdMoneda & " " _
      & "AND p.fecha <= " & conectar.Escape(FechaHasta) & " " _
      & "AND p.estado IN (0, 1) "

    '------------------------------------------------------
    ' EXCLUIR ANTICIPOS QUE YA ESTABAN APLICADOS
    ' EN UNA OP APROBADA A LA FECHA DE CORTE
    '------------------------------------------------------
    q = q & "AND NOT EXISTS ( " _
          & " SELECT 1 " _
          & " FROM ordenes_pago_pagos_a_cuenta vinc " _
          & " INNER JOIN ordenes_pago op " _
          & " ON op.id = vinc.id_orden_pago " _
          & " WHERE vinc.id_pago_a_cuenta = p.id " _
          & " AND op.estado = 1 " _
          & " AND op.fecha <= " _
          & conectar.Escape(FechaHasta) & " " _
          & ")"

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF Then

        TotalAnticiposPendientesAl = _
            CDbl(rs!total_anticipos)

    End If

    Set rs = Nothing

    Exit Function

err1:

    numeroError = Err.Number
    descripcionError = Err.Description

    TotalAnticiposPendientesAl = 0

    Err.Raise numeroError, _
              "DAOFacturaProveedor.TotalAnticiposPendientesAl", _
              descripcionError

End Function



Public Function FindAllTotalizadores(Optional filtro As String = vbNullString, Optional FechaFin As String = vbNullString, Optional withHistorial As Boolean = False, Optional orderBy As String = vbNullString, Optional soloPropias As Boolean = False, Optional widhCompensatorios As Boolean = False) As Collection
    On Error Resume Next
    On Error GoTo err1
    Dim indice As New Dictionary
    Dim q As String
    

    Dim rs As Recordset
    Dim col As New Collection
    q = "SELECT *, (SELECT max(id_orden_pago) FROM ordenes_pago_facturas inner join ordenes_pago on ordenes_pago_facturas.id_orden_pago=ordenes_pago.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND ordenes_pago.estado<>2  limit 1) as nro_orden"
    
    
    If widhCompensatorios Then
        q = q & ", (SELECT SUM(IF (tipo=1, c.importe,-c.importe)) FROM ordenes_pago_compensatorios c JOIN ordenes_pago op ON c.id_orden_pago=op.id AND op.estado=1 WHERE c.id_comprobante = AdminComprasFacturasProveedores.id AND c.cancelado=0 ) as total_compensado"
    

    
    Else
        q = q & ",0 as total_compensado "
    End If

    q = q & ",IFNULL((SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id WHERE op1.estado=1 AND op1.fecha <= " & FechaFin & " AND opf.id_factura_proveedor=AdminComprasFacturasProveedores.id),0) AS total_abonado"
    q = q & ",IFNULL((SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id WHERE op1.estado=1 AND op1.fecha <= " & FechaFin & " AND opf.id_factura_proveedor=AdminComprasFacturasProveedores.id),0) AS neto_gravado_abonado "
    q = q & ",IFNULL((SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id WHERE op1.estado=1 AND op1.fecha <=" & FechaFin & " AND opf.id_factura_proveedor=AdminComprasFacturasProveedores.id),0) AS otros_abonado "
    
    '------------------------------------------------------
    ' LIQUIDACIONES APROBADAS HASTA LA FECHA DE CORTE
    '------------------------------------------------------
    
    q = q & ",IFNULL((" _
          & "SELECT SUM(lcf.neto_gravado_liquidado) " _
          & "FROM liquidaciones_caja_facturas lcf " _
          & "INNER JOIN liquidaciones_caja lc " _
          & "ON lc.id = lcf.id_liquidacion_caja " _
          & "WHERE lcf.id_factura_proveedor = " _
          & "AdminComprasFacturasProveedores.id " _
          & "AND lc.estado = 1 " _
          & "AND lc.fecha <= " & FechaFin _
          & "),0) AS ng_liquidado_al_corte "
    
    q = q & ",IFNULL((" _
          & "SELECT SUM(lcf.otros_liquidado) " _
          & "FROM liquidaciones_caja_facturas lcf " _
          & "INNER JOIN liquidaciones_caja lc " _
          & "ON lc.id = lcf.id_liquidacion_caja " _
          & "WHERE lcf.id_factura_proveedor = " _
          & "AdminComprasFacturasProveedores.id " _
          & "AND lc.estado = 1 " _
          & "AND lc.fecha <= " & FechaFin _
          & "),0) AS otros_liquidado_al_corte "
          
    
    q = q & " ,  CONVERT((SELECT IFNULL(GROUP_CONCAT(id_orden_pago),'-') FROM ordenes_pago_facturas INNER JOIN ordenes_pago ON ordenes_pago_facturas.id_orden_pago=ordenes_pago.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND ordenes_pago.estado<>2  AND ordenes_pago.fecha <=" & FechaFin & "),NCHAR) AS ordenes_pago "
    q = q & ",   CONVERT((SELECT IFNULL(GROUP_CONCAT(numero_liq),'-') From liquidaciones_caja_facturas INNER JOIN liquidaciones_caja ON liquidaciones_caja_facturas.id_liquidacion_caja=liquidaciones_caja.id WHERE id_factura_proveedor = AdminComprasFacturasProveedores.id AND liquidaciones_caja.estado<>2  AND liquidaciones_caja.fecha <=" & FechaFin & "),NCHAR) AS num_liquidaciones_caja "
    q = q & " From" _
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
        & " LEFT JOIN usuarios ON AdminComprasFacturasProveedores.id_usuario_creador=usuarios.id " _
        & " WHERE 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    
    If soloPropias Then
        q = q & " and AdminComprasFacturasProveedores.id_usuario_creador=" & funciones.GetUserObj.Id

    End If

    If LenB(orderBy) > 0 Then
        q = q & " ORDER BY " & orderBy
    End If


    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, indice
    


    Dim F As clsFacturaProveedor
    Dim per As clsPercepcionesAplicadas
    Dim Iva As clsAlicuotaAplicada

    While Not rs.EOF
        Set F = Map(rs, indice, "AdminComprasFacturasProveedores", "proveedores", "AdminConfigFacturasProveedor", "AdminConfigIVAProveedor", "AdminConfigMonedas")

        '------------------------------------------------------
        ' IMPORTES ABONADOS HASTA LA FECHA DE CORTE
        '------------------------------------------------------
        
        F.NetoGravadoAbonadoGlobal = _
            CDbl(rs!neto_gravado_abonado) + _
            CDbl(rs!ng_liquidado_al_corte)
        
        F.OtrosAbonadoGlobal = _
            CDbl(rs!otros_abonado) + _
            CDbl(rs!otros_liquidado_al_corte)

        F.OrdenesPagoId = rs!ordenes_pago
        F.LiquidacionesCajaId = rs!num_liquidaciones_caja


        If funciones.BuscarEnColeccion(col, CStr(F.Id)) Then
            Set F = col.item(CStr(F.Id))
        Else
            If withHistorial Then
                F.Historial = DaoFacturaProveedorHistorial.getAllByIdFactura(F.Id)
            End If
        End If

        If IsSomething(F.configFactura) Then
            F.configFactura.alicuotas.Add DAOAlicuotas.Map(rs, indice, "AdminConfigIvaAlicuotas")
        End If

        Set per = DAOPercepcionesAplicadas.Map(rs, indice, "AdminComprasFacturasProveedoresPercepciones", "AdminConfigPercepciones")
        If IsSomething(per) Then
            If Not funciones.BuscarEnColeccion(F.percepciones, CStr(per.Id)) Then
                If per.Id <> 0 Then F.percepciones.Add per, CStr(per.Id)
            End If
        End If
        
        Set Iva = DAOIvaAplicado.Map(rs, indice, "AdminComprasFacturasProveedoresIva", "a1")
        If IsSomething(Iva) Then
            If Not funciones.BuscarEnColeccion(F.IvaAplicado, CStr(Iva.Id)) Then
                F.IvaAplicado.Add Iva, CStr(Iva.Id)
            End If
        End If
        
       If Not funciones.BuscarEnColeccion(col, CStr(F.Id)) Then col.Add F, CStr(F.Id)
       rs.MoveNext
       
    Wend

    Set FindAllTotalizadores = col

    Exit Function

err1:
    MsgBox Err.Description
End Function


'=========================================================
' SITUACION DE UNA FACTURA A LA FECHA DE CORTE
'
' Utiliza los importes historicos que ya fueron
' cargados mediante FindAllTotalizadores.
'=========================================================
Public Function SituacionComprobanteAlCorte( _
    ByVal fac As clsFacturaProveedor _
) As String

    Dim importeFactura As Double
    Dim saldo As Double

    importeFactura = _
        fac.Monto + _
        fac.TotalIVA + _
        fac.totalPercepciones + _
        fac.ImpuestoInterno + _
        fac.redondeo

    saldo = funciones.FormatearDecimales( _
        importeFactura - fac.TotalAbonadoGlobal)

    If fac.tipoDocumentoContable = _
            tipoDocumentoContable.notaCredito Then

        SituacionComprobanteAlCorte = _
            "Nota de credito - ver saldo"

    ElseIf Abs(saldo) < 0.01 Then

        SituacionComprobanteAlCorte = _
            "Saldada al corte"

    ElseIf saldo < 0 Then

        SituacionComprobanteAlCorte = _
            "Saldo acreedor - revisar"

    ElseIf fac.TotalAbonadoGlobal > 0 Then

        SituacionComprobanteAlCorte = _
            "Pago parcial al corte"

    Else

        SituacionComprobanteAlCorte = _
            "Pendiente al corte"

    End If

End Function


Public Function ExportarColeccionTotalizadores( _
    col As Collection, _
    Optional ProgressBar As Object, _
    Optional FechaFin As String, _
    Optional ByVal IdProveedorResumen As Long = 0, _
    Optional ByVal SaldoFacturasResumen As Double = 0, _
    Optional ByVal AnticiposResumen As Double = 0, _
    Optional ByVal AplicacionesPosteriores As Object = Nothing _
) As Boolean

    On Error GoTo err1

    ExportarColeccionTotalizadores = False

    Dim xlApplication As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object

    Dim fac As clsFacturaProveedor

    Dim fila As Long
    Dim filaTotal As Long
    Dim filaResumen As Long

    Dim cantidad As Long
    Dim signo As Integer

    Dim importeFactura As Double
    Dim importePagado As Double
    Dim importeSaldo As Double

    Dim situacion As String
    Dim aplicacion As String

    Dim claveFactura As String
    Dim claveMoneda As String

    Dim totalesPorMoneda As New Dictionary

    Dim datosMoneda As Variant
    Dim monedaClave As Variant

    Dim fechaReporte As Date

    Dim ruta As String
    Dim nombreHoja As String

    '==================================================
    ' VALIDACIONES
    '==================================================

    If col Is Nothing Then
        Err.Raise 5, , _
            "No se recibio la coleccion de comprobantes."
    End If

    If LenB(Trim$(FechaFin)) = 0 Then
        Err.Raise 5, , _
            "No se recibio la fecha de corte."
    End If

    fechaReporte = DateValue(CDate(FechaFin))

    '==================================================
    ' INICIAR EXCEL
    '==================================================

    Set xlApplication = CreateObject("Excel.Application")

    Set xlWorkbook = xlApplication.Workbooks.Add

    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Range("A1:M1").Merge
    xlWorksheet.Range("A2:M2").Merge

    xlWorksheet.Range("A1:M2").HorizontalAlignment = _
        xlHAlignCenter

    xlWorksheet.Range("A1:M2").Font.Bold = True

    xlWorksheet.Cells(1, 1).value = _
        "REPORTE DE COMPROBANTES DE COMPRA ADEUDADOS AL " & _
        Format$(fechaReporte, "dd/mm/yyyy")

    xlWorksheet.Cells(2, 1).value = _
        "Las OP posteriores son informativas. " & _
        "Los anticipos asociados no se imputan " & _
        "individualmente a cada factura."

    '==================================================
    ' ENCABEZADOS - 13 COLUMNAS
    '==================================================

    fila = 3

    xlWorksheet.Cells(fila, 1).value = "ID"
    xlWorksheet.Cells(fila, 2).value = "Razon Social"
    xlWorksheet.Cells(fila, 3).value = "CUIT"
    xlWorksheet.Cells(fila, 4).value = "Documento"
    xlWorksheet.Cells(fila, 5).value = "Letra"
    xlWorksheet.Cells(fila, 6).value = "Numero"
    xlWorksheet.Cells(fila, 7).value = "Fecha"
    xlWorksheet.Cells(fila, 8).value = "Moneda"
    xlWorksheet.Cells(fila, 9).value = "Total"
    xlWorksheet.Cells(fila, 10).value = "Pagado"
    xlWorksheet.Cells(fila, 11).value = "Saldo"

    xlWorksheet.Cells(fila, 12).value = _
        "Situacion al corte"

    xlWorksheet.Cells(fila, 13).value = _
        "Aplicacion posterior / Anticipo asociado"

    With xlWorksheet.Range("A3:M3")

        .Font.Bold = True
        .Interior.Color = &HC0C0C0
        .Borders.LineStyle = xlContinuous

    End With

    'Conservar numeros de comprobante como texto.
    xlWorksheet.Columns("F").NumberFormat = "@"

    '==================================================
    ' BARRA DE PROGRESO
    '==================================================

    If Not ProgressBar Is Nothing Then

        ProgressBar(0).min = 0

        If col.count > 0 Then

            ProgressBar(0).max = col.count

        Else

            ProgressBar(0).max = 1

        End If

        ProgressBar(0).value = 0

    End If

    '==================================================
    ' RECORRER EXACTAMENTE LA COLECCION DE LA GRILLA
    '==================================================

    cantidad = 0

    For Each fac In col

        cantidad = cantidad + 1

        fila = fila + 1

        signo = 1

        If fac.tipoDocumentoContable = _
                tipoDocumentoContable.notaCredito Then

            signo = -1

        End If

        '----------------------------------------------
        ' MISMOS IMPORTES QUE grilla_UnboundReadData
        '----------------------------------------------

        importeFactura = _
            fac.Monto + _
            fac.TotalIVA + _
            fac.totalPercepciones + _
            fac.ImpuestoInterno + _
            fac.redondeo

        importePagado = fac.TotalAbonadoGlobal

        importeSaldo = _
            importeFactura - importePagado

        importeFactura = _
            funciones.FormatearDecimales(importeFactura) * signo

        importePagado = _
            funciones.FormatearDecimales(importePagado) * signo

        importeSaldo = _
            funciones.FormatearDecimales(importeSaldo) * signo

        '----------------------------------------------
        ' SITUACION HISTORICA
        '----------------------------------------------

        situacion = _
            DAOFacturaProveedor.SituacionComprobanteAlCorte(fac)

        '----------------------------------------------
        ' OP Y PAGOS A CUENTA POSTERIORES
        '----------------------------------------------

        aplicacion = ""

        claveFactura = CStr(fac.Id)

        If Not AplicacionesPosteriores Is Nothing Then

            If AplicacionesPosteriores.Exists( _
                    claveFactura) Then

                aplicacion = _
                    CStr(AplicacionesPosteriores( _
                        claveFactura))

            End If

        End If

        '----------------------------------------------
        ' ESCRIBIR LA FILA
        '----------------------------------------------

        xlWorksheet.Cells(fila, 1).value = fac.Id

        xlWorksheet.Cells(fila, 2).value = _
            UCase(funciones.RazonSocialFormateada( _
                fac.Proveedor.RazonSocial))

        xlWorksheet.Cells(fila, 3).value = _
            fac.Proveedor.cuit

        xlWorksheet.Cells(fila, 4).value = _
            enums.EnumTipoDocumentoContableShort( _
                fac.tipoDocumentoContable)

        xlWorksheet.Cells(fila, 5).value = _
            fac.configFactura.TipoFactura

        xlWorksheet.Cells(fila, 6).value = _
            CStr(fac.numero)

        xlWorksheet.Cells(fila, 7).value = _
            fac.FEcha

        xlWorksheet.Cells(fila, 8).value = _
            fac.moneda.NombreCorto

        xlWorksheet.Cells(fila, 9).value = _
            importeFactura

        xlWorksheet.Cells(fila, 10).value = _
            importePagado

        xlWorksheet.Cells(fila, 11).value = _
            importeSaldo

        xlWorksheet.Cells(fila, 12).value = _
            situacion

        xlWorksheet.Cells(fila, 13).value = _
            aplicacion

        '----------------------------------------------
        ' ACUMULAR TOTALES POR MONEDA
        '----------------------------------------------

        claveMoneda = CStr(fac.moneda.Id)

        If Not totalesPorMoneda.Exists(claveMoneda) Then

            datosMoneda = Array( _
                0#, _
                0#, _
                0#, _
                fac.moneda.NombreCorto)

            totalesPorMoneda.Add _
                claveMoneda, datosMoneda

        End If

        datosMoneda = totalesPorMoneda(claveMoneda)

        datosMoneda(0) = _
            CDbl(datosMoneda(0)) + importeFactura

        datosMoneda(1) = _
            CDbl(datosMoneda(1)) + importePagado

        datosMoneda(2) = _
            CDbl(datosMoneda(2)) + importeSaldo

        totalesPorMoneda(claveMoneda) = datosMoneda

        '----------------------------------------------
        ' ACTUALIZAR PROGRESO
        '----------------------------------------------

        If Not ProgressBar Is Nothing Then

            ProgressBar(0).value = cantidad

        End If

    Next fac

    '==================================================
    ' FORMATO DEL DETALLE
    '==================================================

    If cantidad > 0 Then

        xlWorksheet.Range( _
            xlWorksheet.Cells(4, 7), _
            xlWorksheet.Cells(fila, 7)).NumberFormat = _
            "dd/mm/yyyy"

        xlWorksheet.Range( _
            xlWorksheet.Cells(4, 9), _
            xlWorksheet.Cells(fila, 11)).NumberFormat = _
            "#,##0.00"

        xlWorksheet.Range( _
            xlWorksheet.Cells(3, 1), _
            xlWorksheet.Cells(fila, 13)).Borders.LineStyle = _
            xlContinuous

        xlWorksheet.Range( _
            xlWorksheet.Cells(3, 1), _
            xlWorksheet.Cells(fila, 13)).AutoFilter

    Else

        fila = fila + 1

        xlWorksheet.Cells(fila, 1).value = _
            "No se encontraron comprobantes."

    End If

    '==================================================
    ' TOTALES SEPARADOS POR MONEDA
    '==================================================

    filaTotal = fila

    For Each monedaClave In totalesPorMoneda.Keys

        datosMoneda = totalesPorMoneda(monedaClave)

        filaTotal = filaTotal + 1

        xlWorksheet.Cells(filaTotal, 8).value = _
            "TOTAL " & CStr(datosMoneda(3))

        xlWorksheet.Cells(filaTotal, 9).value = _
            CDbl(datosMoneda(0))

        xlWorksheet.Cells(filaTotal, 10).value = _
            CDbl(datosMoneda(1))

        xlWorksheet.Cells(filaTotal, 11).value = _
            CDbl(datosMoneda(2))

        With xlWorksheet.Range( _
            xlWorksheet.Cells(filaTotal, 1), _
            xlWorksheet.Cells(filaTotal, 13))

            .Font.Bold = True
            .Interior.Color = &HC0C0C0

        End With

        xlWorksheet.Range( _
            xlWorksheet.Cells(filaTotal, 9), _
            xlWorksheet.Cells(filaTotal, 11)).NumberFormat = _
            "#,##0.00"

    Next monedaClave

    '==================================================
    ' RESUMEN GLOBAL DEL PROVEEDOR EN PESOS
    '
    ' Este resumen NO depende de los filtros parciales
    ' de la grilla.
    '==================================================

    If IdProveedorResumen > 0 Then

        filaResumen = filaTotal + 3

        xlWorksheet.Cells(filaResumen, 8).value = _
            "PROVEEDOR ID: " & IdProveedorResumen & _
            " - RESUMEN GLOBAL AR$"

        xlWorksheet.Cells(filaResumen + 1, 8).value = _
            "Deuda por facturas al corte"

        xlWorksheet.Cells(filaResumen + 1, 11).value = _
            SaldoFacturasResumen

        xlWorksheet.Cells(filaResumen + 2, 8).value = _
            "Anticipos sin aplicar al corte"

        xlWorksheet.Cells(filaResumen + 2, 11).value = _
            -AnticiposResumen

        xlWorksheet.Cells(filaResumen + 3, 8).value = _
            "DEUDA NETA AR$ DEL PROVEEDOR"

        xlWorksheet.Cells(filaResumen + 3, 11).value = _
            SaldoFacturasResumen - AnticiposResumen

        With xlWorksheet.Range( _
            xlWorksheet.Cells(filaResumen, 8), _
            xlWorksheet.Cells(filaResumen + 3, 11))

            .Font.Bold = True

        End With

        xlWorksheet.Range( _
            xlWorksheet.Cells(filaResumen + 1, 11), _
            xlWorksheet.Cells(filaResumen + 3, 11)).NumberFormat = _
            "#,##0.00"

        xlWorksheet.Cells(filaResumen + 5, 8).value = _
            "Resumen informativo: validar con cuenta corriente."

    End If

    '==================================================
    ' FORMATO FINAL
    '==================================================

    xlWorksheet.Cells.EntireColumn.AutoFit

    xlWorksheet.Columns("M").ColumnWidth = 60

    xlWorksheet.Columns("M").WrapText = True

    xlWorksheet.Columns("L").ColumnWidth = 26

    xlApplication.ScreenUpdating = False

    nombreHoja = xlWorksheet.Name

    xlWorkbook.Sheets(nombreHoja).Select

    xlApplication.ScreenUpdating = True

    '==================================================
    ' GUARDAR
    '==================================================

    ruta = Environ$("TEMP")

    If LenB(ruta) = 0 Then
        ruta = Environ$("TMP")
    End If

    If LenB(ruta) = 0 Then
        ruta = App.path
    End If

    ruta = ruta & "\" & _
           funciones.CreateGUID() & ".xls"

    xlWorkbook.SaveAs ruta

    xlWorkbook.Saved = True

    xlWorkbook.Close False

    xlApplication.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    If Not ProgressBar Is Nothing Then
        ProgressBar(0).value = 0
    End If

    'Abrir el archivo generado.
    ShellExecute -1, "open", ruta, "", "", 4

    ExportarColeccionTotalizadores = True

    Exit Function

err1:

    ExportarColeccionTotalizadores = False

    Dim descripcionError As String

    descripcionError = _
        Err.Number & " - " & Err.Description

    On Error Resume Next

    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
    End If

    If Not xlApplication Is Nothing Then
        xlApplication.Quit
    End If

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    If Not ProgressBar Is Nothing Then
        ProgressBar(0).value = 0
    End If

    MsgBox _
        "Error al exportar los comprobantes." & _
        vbCrLf & vbCrLf & descripcionError, _
        vbCritical, _
        "Exportacion a Excel"

End Function


'ESTA FUNCION DEJA TODO EN 0 PARA QUE PUEDA SER EDITADO
Public Function ColeccionAEditable(col As Collection) As Boolean
    On Error GoTo err1

    ColeccionAEditable = True

    Dim fac As clsFacturaProveedor
    
    Set cn = conectar.obternerConexion
    
    For Each fac In col
        cn.execute "UPDATE liquidaciones_caja_facturas SET total_liquidado=0, neto_gravado_liquidado=0, otros_liquidado=0 WHERE id_factura_proveedor=" & fac.Id
    Next

    Exit Function
err1:
    ColeccionAEditable = False
    
End Function


Public Function MarcarFacturaFisicaRecibida(ByVal idFactura As Long) As Boolean
    On Error GoTo err1

    Dim q As String

    q = "UPDATE AdminComprasFacturasProveedores " & _
        "SET ARCA = 0, ultima_actualizacion = " & Escape(Now) & " " & _
        "WHERE id = " & idFactura

    MarcarFacturaFisicaRecibida = conectar.execute(q)

    Exit Function

err1:
    MarcarFacturaFisicaRecibida = False
End Function
