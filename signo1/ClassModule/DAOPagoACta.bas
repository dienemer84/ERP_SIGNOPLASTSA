Attribute VB_Name = "DAOPagoACta"
Option Explicit


Public Function FindAbonadoPendienteEnEstaOP(facid As Long, ocid As Long) As Collection

    Dim q As String

    q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & "),0 ) AS total_pendiente, " _
      & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & "),0 ) AS netogravado_pendiente, " _
      & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & "),0 ) AS otros_pendiente "

    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim tot As Double, ng As Double, Otros As Double
    tot = rs!total_pendiente
    ng = rs!netogravado_pendiente
    Otros = rs!otros_pendiente

    Dim C As New Collection
    C.Add tot
    C.Add ng
    C.Add Otros
    Set FindAbonadoPendienteEnEstaOP = C
End Function




Public Function FindAbonadoPendiente(facid As Long, ocid As Long) As Collection

    Dim q As String

    q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago <> " & ocid & "),0 ) AS total_pendiente, " _
      & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago <> " & ocid & "),0 ) AS netogravado_pendiente, " _
      & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago <> " & ocid & "),0 ) AS otros_pendiente "
    
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim tot As Double, ng As Double, Otros As Double
    tot = rs!total_pendiente
    ng = rs!netogravado_pendiente
    Otros = rs!otros_pendiente

    Dim C As New Collection
    C.Add tot
    C.Add ng
    C.Add Otros
    Set FindAbonadoPendiente = C
    
End Function


Public Function FindAbonadoFactura(facid As Long, ocid As Long) As Collection

    Dim q As String

    q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & " and op1.estado=1),0 ) AS total_pendiente, " _
      & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & " and op1.estado=1),0 ) AS netogravado_pendiente, " _
      & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
      & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & " and op1.estado=1),0 ) AS otros_pendiente "

    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim tot As Double, ng As Double, Otros As Double
    tot = rs!total_pendiente
    ng = rs!netogravado_pendiente
    Otros = rs!otros_pendiente

    Dim C As New Collection
    C.Add tot
    C.Add ng
    C.Add Otros
    Set FindAbonadoFactura = C
End Function


Public Function FindLast() As OrdenPago
    Set FindLast = FindAll("ordenes_pago.id = (SELECT MAX(id) FROM ordenes_pago)")(1)
End Function


Public Function FindByFacturaId(facid As Long) As OrdenPago
    Dim col As Collection
    Set col = FindAll("ordenes_pago.id = (SELECT DISTINCT id_orden_pago from ordenes_pago_facturas opf inner join ordenes_pago op on opf.id_orden_pago=op.id WHERE id_factura_proveedor = " & facid & " AND op.estado=1)")
    If col.count > 0 Then
        Set FindByFacturaId = col(1)
    Else
        Set FindByFacturaId = Nothing
    End If
End Function


Public Function FindById(Id As Long) As clsPagoACta
    Set FindById = FindAll("pagos_a_cuenta.id=" & Id)(1)
End Function


Public Function FindAllByProveedor(provid As Long, Optional cond As String, Optional soloOp As Boolean = False) As Collection
    Dim q As String
    q = "pagos_a_cuenta.id IN (SELECT DISTINCT opf.id from pagos_a_cuenta opf WHERE opf.id_proveedor = " & provid & " )"

    If LenB(cond) > 0 Then
        q = q & "  " & cond
        'ver aca
    End If

    If soloOp Then
        Set FindAllByProveedor = FindAllSoloOP(q)
        '  Debug.Print FindAllByProveedor.count

    Else
        Set FindAllByProveedor = FindAll(q)
        '  Debug.Print FindAllByProveedor.count
    End If
End Function


Public Function FindAllSoloOP(Optional filter As String = "1 = 1", Optional orderBy As String = "1") As Collection
    Dim q As String
    q = "SELECT * " _
      & " From pagos_a_cuenta" _
      & " LEFT JOIN AdminConfigMonedas ON (AdminConfigMonedas.id = pagos_a_cuenta.id_moneda)"

    q = q & " WHERE " & filter
    q = q & " ORDER BY " & orderBy
    Dim col As New Collection
    Dim op As OrdenPago
    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx
    While Not rs.EOF
        Set op = Map(rs, idx, "pagos_a_cuenta", "AdminConfigMonedas")    ', "certificados_retencion")
        If funciones.BuscarEnColeccion(col, CStr(op.Id)) Then
            Set op = col.item(CStr(op.Id))
        Else
            col.Add op, CStr(op.Id)
        End If
        rs.MoveNext
    Wend

    Set FindAllSoloOP = col

End Function

Public Function FindAll(Optional filter As String = "1 = 1", Optional orderBy As String = "1") As Collection
    Dim q As String
    q = "SELECT *, (operaciones.pertenencia + 0) as pertenencia2" _
      & " From pagos_a_cuenta" _
      & " LEFT JOIN pagos_a_cuenta_cheques ON (pagos_a_cuenta.id = pagos_a_cuenta_cheques.id_pago_a_cuenta)" _
      & " LEFT JOIN pagos_a_cuenta_operaciones ON (pagos_a_cuenta.id = pagos_a_cuenta_operaciones.id_pago_a_cuenta)" _
      & " LEFT JOIN operaciones ON (operaciones.id = pagos_a_cuenta_operaciones.id_operacion)" _
      & " LEFT JOIN Cheques ON (Cheques.id = pagos_a_cuenta_cheques.id_cheque)" _
      & " LEFT JOIN Chequeras ON (Chequeras.id = Cheques.id_chequera)" _
      & " LEFT JOIN AdminConfigBancos monbanco ON (monbanco.id = Chequeras.id_banco)" _
      & " LEFT JOIN AdminConfigMonedas monchequera ON (monchequera.id = Chequeras.id_moneda)" _
      & " LEFT JOIN AdminConfigMonedas ON (AdminConfigMonedas.id = pagos_a_cuenta.id_moneda)" _
      & " LEFT JOIN AdminConfigMonedas monedaoperacion ON (monedaoperacion.id = operaciones.moneda_id)" _
      & " LEFT JOIN AdminComprasCuentasContables ON (AdminComprasCuentasContables.id = operaciones.cuenta_contable_id)" _
      & " LEFT JOIN cajas ON (cajas.id = operaciones.cuentabanc_o_caja_id)" _
      & " LEFT JOIN AdminConfigCuentas ON (AdminConfigCuentas.id = operaciones.cuentabanc_o_caja_id)" _
      & " LEFT JOIN AdminConfigMonedas moncuentabancaria ON (moncuentabancaria.id = AdminConfigCuentas.moneda_id)" _
      & " LEFT JOIN AdminConfigMonedas moncheque ON (moncheque.id = Cheques.id_moneda)" _
      & " LEFT JOIN AdminRecibosCheques rec ON rec.idCheque= Cheques.id" _
      & " LEFT JOIN AdminConfigBancos ON (AdminConfigBancos.id = AdminConfigCuentas.idBanco)" _
      & " LEFT JOIN AdminConfigBancos bancocheque ON (bancocheque.id = Cheques.id_banco)" _
      & " LEFT JOIN proveedores proveedores ON (pagos_a_cuenta.id_proveedor = proveedores.id)"

    q = q & " WHERE " & filter
    q = q & " ORDER BY " & orderBy

    Dim col As New Collection
    Dim op As clsPagoACta
    Dim fac As clsFacturaProveedor
    Dim che As cheque
    Dim oper As operacion

    Dim idx As Dictionary
    Dim rs As Recordset
    Dim ra As DTORetencionAlicuota
    
    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set op = Map(rs, idx, "pagos_a_cuenta", "AdminConfigMonedas", "cuentacontableordenpago", "retenciones", "proveedores")   ', "certificados_retencion")

        If funciones.BuscarEnColeccion(col, CStr(op.Id)) Then
            Set op = col.item(CStr(op.Id))
        Else
            col.Add op, CStr(op.Id)
        End If

 
       Set che = DAOCheques.Map(rs, idx, "Cheques", "bancocheque", "moncheque", "Chequeras", "monchequera", "monbanco", "rec")
        

If IsSomething(che) Then
            If che.Propio Then
                If Not funciones.BuscarEnColeccion(op.ChequesPropios, CStr(che.Id)) Then
                    op.ChequesPropios.Add che, CStr(che.Id)
                End If
            Else
                If Not funciones.BuscarEnColeccion(op.ChequesTerceros, CStr(che.Id)) Then
                    op.ChequesTerceros.Add che, CStr(che.Id)
                End If
            End If
        End If

        Set oper = DAOOperacion.Map(rs, idx, "operaciones", "AdminComprasCuentasContables", "monedaoperacion", "AdminConfigCuentas", "cajas")
        If IsSomething(oper) Then
            If oper.Pertenencia = Banco Then
                If Not funciones.BuscarEnColeccion(op.operacionesBanco, CStr(oper.Id)) Then

                    op.operacionesBanco.Add oper, CStr(oper.Id)
                End If
            ElseIf oper.Pertenencia = caja Then
                If Not funciones.BuscarEnColeccion(op.operacionesCaja, CStr(oper.Id)) Then
                    op.operacionesCaja.Add oper, CStr(oper.Id)
                End If
            End If
        End If


        rs.MoveNext
        
    Wend

    Set FindAll = col
    
End Function

Public Function Map(rs As Recordset, indice As Dictionary, _
                    tabla As String, _
                    Optional ByVal tablaMoneda As String = vbNullString, _
                    Optional ByVal tablaCuentaContable As String = vbNullString, _
                    Optional ByVal TablaRetenciones As String = vbNullString, _
                    Optional ByVal tablaProveedor As String = vbNullString, _
                    Optional tablaAdminConfigFacturasProveedor As String = vbNullString, _
                    Optional tablaAdminConfigIVAProveedor As String = vbNullString _
                  ) As clsPagoACta

'Optional ByVal tablaCertRetencion As String = vbNullString _

  Dim op As clsPagoACta


    
'id_certificado_retencion
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set op = New clsPagoACta
        op.Id = Id

        op.FEcha = GetValue(rs, indice, tabla, "fecha")
        op.estado = GetValue(rs, indice, tabla, "estado")
        op.StaticTotalFacturas = GetValue(rs, indice, tabla, "static_total_facturas")
        op.StaticTotalFacturasNG = GetValue(rs, indice, tabla, "static_total_factura_ng")
        op.StaticTotalRetenido = GetValue(rs, indice, tabla, "static_total_a_retener")
        op.StaticTotalOrigenes = GetValue(rs, indice, tabla, "static_total_origen")

        If LenB(tablaProveedor) > 0 Then Set op.Proveedor = DAOProveedor.Map2(rs, indice, tablaProveedor)

        op.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")
        op.DiferenciaCambioEnNG = GetValue(rs, indice, tabla, "dif_cambio_ng")
        op.DiferenciaCambioEnTOTAL = GetValue(rs, indice, tabla, "dif_cambio_total")
        
        op.Creada = GetValue(rs, indice, tabla, "creada")
        
        If LenB(tablaMoneda) > 0 Then Set op.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        'If LenB(tablaCertRetencion) > 0 Then Set op.CertificadoRetencion = DAOCertificadoRetencion.Map(rs, indice, tablaCertRetencion)
        
            Debug.Print (op.Id)
    End If

    Set Map = op
End Function



Public Function MapAlicuotaRetencion(rs As Recordset, indice As Dictionary, _
                                     tabla As String, _
                                     ByVal TablaRetenciones As String) As DTORetencionAlicuota

'Optional ByVal tablaCertRetencion As String = vbNullString _

  Dim ra As DTORetencionAlicuota

'id_certificado_retencion
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, "id_retencion")

    If Id > 0 Then
        Set ra = New DTORetencionAlicuota
        ra.alicuotaRetencion = GetValue(rs, indice, tabla, "alicuota")
        Set ra.Retencion = DAORetenciones.Map(rs, indice, TablaRetenciones)
        ra.importe = GetValue(rs, indice, tabla, "total")
        ra.certificados = GetValue(rs, indice, tabla, "certificados")

        'If LenB(tablaCertRetencion) > 0 Then Set op.CertificadoRetencion = DAOCertificadoRetencion.Map(rs, indice, tablaCertRetencion)
    End If

    Set MapAlicuotaRetencion = ra
End Function


Public Function Save(pcta As clsPagoACta, Optional cascada As Boolean = False) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    Save = Guardar(pcta, cascada)
    conectar.CommitTransaction
    Exit Function
err1:
    Save = False
    conectar.RollBackTransaction
End Function


Public Function aprobar(op_mem As OrdenPago, insideTransaction As Boolean) As Boolean

    On Error GoTo err1
    If insideTransaction Then conectar.BeginTransaction

    '3-10-2020 recargo la OP para que se actualicen los estados de las facturas y se validen bien
    
    Dim op As OrdenPago
    
    Set op = DAOOrdenPago.FindById(op_mem.Id)

    If Not IsSomething(op) Then
        GoTo err1
    End If

    'VALIDAR BIEN LOS TOTALES ANTES DE PODER APROBAR
    'verificar que las facturas esten todas aprobadsa...
    Dim F As clsFacturaProveedor
    Dim nopago As Double
    Dim nopago1 As Double

    Dim otrosvalores As Double

    Dim esf As EstadoFacturaProveedor
    For Each F In op.FacturasProveedor

        Dim fac As clsFacturaProveedor
        Set fac = DAOFacturaProveedor.FindById(F.Id)
        '            'debug.print (F.Id & "- " & F.NumeroFormateado)

        If fac.estado = EstadoFacturaProveedor.EnProceso Then
            Err.Raise 44, "aprobar op", "La factura " & fac.NumeroFormateado & " no está aprobada. No se pudo aprobar la OP"
        End If

        Dim x

        Set x = DAOOrdenPago.FindAbonadoPendienteEnEstaOP(fac.Id, op.Id)

        nopago1 = fac.total - fac.TotalAbonadoGlobal    '- (funciones.RedondearDecimales(funciones.RedondearDecimales(CDbl(x(1))) + funciones.RedondearDecimales(CDbl(x(2))) + funciones.RedondearDecimales(CDbl(x(3)))))

        'nopago = fac.Total - fac.TotalAbonadoGlobal - funciones.RedondearDecimales(funciones.RedondearDecimales(CDbl(x(1))) + funciones.RedondearDecimales(CDbl(x(2))) + funciones.RedondearDecimales(CDbl(x(3))))

        otrosvalores = funciones.RedondearDecimales(funciones.RedondearDecimales(CDbl(x(1))) + funciones.RedondearDecimales(CDbl(x(2))) + funciones.RedondearDecimales(CDbl(x(3))))

        nopago = funciones.RedondearDecimales(nopago1) - otrosvalores

        esf = EstadoFacturaProveedor.Aprobada

        If nopago < 0 Then
            Err.Raise 44, "aprobar op", "La factura " & fac.NumeroFormateado & " tiene un error y no se pudo aprobar la OP"
        End If
        If nopago > 0 Then
            esf = EstadoFacturaProveedor.pagoParcial
        Else
            esf = EstadoFacturaProveedor.Saldada
        End If
        conectar.execute "UPDATE AdminComprasFacturasProveedores SET estado = " & esf & " WHERE id = " & fac.Id
    Next F


    If op.estado = EstadoOrdenPago_pendiente Then
        Dim es As EstadoOrdenPago
        es = op.estado
        op.estado = EstadoOrdenPago_Aprobada

        If op.EsParaFacturaProveedor Then
           
           If op.FacturasProveedor.count > 0 Then
                If op.FacturasProveedor(1).Proveedor.estado <> 2 Then
                    Dim d As New clsDTOPadronIIBB
                    'todo: cambiar validacion
                    Set d = DTOPadronIIBB.FindByCUIT(op.FacturasProveedor(1).Proveedor.Cuit, TipoPadronRetencion)
                    Dim ret As Double

                    If IsSomething(d) Then
                        ret = d.alicuota

                        If ret <> op.alicuota Then
                            If MsgBox("La alicuota de retención actual del proveedor en el padrón difiere de la especificada en la orden de pago." & vbNewLine & "¿Quiere editar la orden de pago con la nueva alicuota de retención? o ¿Usar la especificada de todas maneras?" & vbNewLine & "[SI] - Continuar usando la especificada." & "[NO] - Cancelar y editar la orden de pago.", vbQuestion + vbYesNo) = vbNo Then
                                GoTo err1
                            End If
                        End If

                    End If
                Else
                    MsgBox "El proveedor es de tipo contado! " & vbNewLine & "No se le realizará ninguna retención!", vbInformation, "Información"
                End If
            End If
        End If

        'analizo las facturas de proveedores

        'TODO: debo verificar que los deudas por compensatorio no esten utilizadas en otra OP aprobada ni que esten ya canceladas en otro proceso

        If Guardar(op) Then

            Dim fac1 As clsFacturaProveedor
            For Each fac1 In op.FacturasProveedor
                If fac1.estado = EstadoFacturaProveedor.Saldada Then
                    If Not DaoFacturaProveedorHistorial.agregar(fac1, "SALDADA") Then GoTo err1
                End If
                If fac1.estado = EstadoFacturaProveedor.pagoParcial Then
                    If Not DaoFacturaProveedorHistorial.agregar(fac1, "PAGO PARCIAL") Then GoTo err1
                End If
            Next

            If op.StaticTotalRetenido > 0 Then
                Dim ra As DTORetencionAlicuota
                For Each ra In op.RetencionesAlicuota

                    If IsSomething(DAOCertificadoRetencion.Create(op, ra.Retencion, ra.alicuotaRetencion, True)) Then
                        MsgBox "Se creo un certificado de Retenciones para la Orden de Pago. ", vbInformation
                    Else
                        GoTo err1
                    End If
                Next

            End If
        Else
            GoTo err1
        End If
        
    End If
    
'        MsgBox (op.Id)
        
    DaoHistorico.Save "orden_pago_historial", "OP Aprobada", op.Id
    aprobar = True
    
'        MsgBox (op.Id)

    If insideTransaction Then conectar.CommitTransaction
    Exit Function
err1:

'        MsgBox (op.Id)

    op.estado = es
    If insideTransaction Then conectar.RollBackTransaction
    aprobar = False
End Function


Public Function Guardar(pcta As clsPagoACta, Optional cascada As Boolean = False) As Boolean

'TODO: tengo que revisar que las facturas no esten en otra op aprobada antes de continuar

    Dim q As String
    Dim rs As Recordset
    On Error GoTo E
    Dim Nueva As Boolean: Nueva = False
    If pcta.Id = 0 Then
        Nueva = True
        q = "INSERT INTO pagos_a_cuenta (id_moneda, fecha, id_proveedor, estado, static_total_facturas, static_total_factura_ng, static_total_a_retener, static_total_origen, dif_cambio_ng,dif_cambio_total)" _
          & " VALUES ('id_moneda', 'fecha', 'id_proveedor', '0', 'static_total_facturas', 'static_total_factura_ng', 'static_total_a_retener', 'static_total_origen', 'dif_cambio_ng','dif_cambio_total')"
    Else
        q = "UPDATE pagos_a_cuenta" _
          & " SET id_moneda = 'id_moneda'," _
          & " fecha = 'fecha'," _
          & " id_proveedor = 'id_proveedor'," _
          & " estado = 'estado'," _
          & " static_total_facturas = 'static_total_facturas'," _
          & " static_total_factura_ng = 'static_total_factura_ng'," _
          & " static_total_a_retener = 'static_total_a_retener'," _
          & " static_total_origen = 'static_total_origen'," _
          & " tipo_cambio = 'tipo_cambio'," _
          & " dif_cambio_ng = 'dif_cambio_ng'," _
          & " dif_cambio_total = 'dif_cambio_total'" _
          & " WHERE id = 'id'"
        q = Replace(q, "'id'", GetEntityId(pcta))
    End If

    q = Replace(q, "'id_moneda'", GetEntityId(pcta.moneda))
    q = Replace(q, "'id_proveedor'", Escape(pcta.Proveedor.Id))
    q = Replace(q, "'fecha'", Escape(pcta.FEcha))
    q = Replace(q, "'estado'", Escape(pcta.estado))
    q = Replace(q, "'static_total_facturas'", Escape(pcta.StaticTotalFacturas))
    q = Replace(q, "'static_total_factura_ng'", Escape(pcta.StaticTotalFacturasNG))
    q = Replace(q, "'static_total_a_retener'", Escape(pcta.StaticTotalRetenido))
    q = Replace(q, "'static_total_origen'", Escape(pcta.StaticTotalOrigenes))
    q = Replace(q, "'dif_cambio'", Escape(pcta.DiferenciaCambio))
    q = Replace(q, "'id_moneda_pago'", Escape(pcta.IdMonedaPago))
    q = Replace(q, "'tipo_cambio'", Escape(pcta.TipoCambio))
    q = Replace(q, "'dif_cambio_ng'", Escape(pcta.DiferenciaCambioEnNG))
    q = Replace(q, "'dif_cambio_total'", Escape(pcta.DiferenciaCambioEnTOTAL))


    If Not conectar.execute(q) Then GoTo E

    If Nueva Then pcta.Id = conectar.UltimoId2()
    If pcta.Id = 0 Then GoTo E

    '------------------------------------------------------
    
    If cascada Then

        q = "SELECT id_cheque FROM pagos_a_cuenta_cheques WHERE id_pago_a_cuenta = " & pcta.Id
        q = q & " AND id_cheque NOT IN (-1"
        If pcta.ChequesTerceros.count > 0 Then
            q = q & ", " & funciones.JoinCollectionValues(pcta.ChequesTerceros, ", ", "id")
        End If
        If pcta.ChequesPropios.count > 0 Then
            q = q & ", " & funciones.JoinCollectionValues(pcta.ChequesPropios, ", ", "id")
        End If
        q = q & ")"
        Set rs = conectar.RSFactory(q)
        While Not rs.EOF
            q = "UPDATE Cheques SET en_cartera = 1, observaciones = NULL, origen= NULL WHERE id = " & rs!id_cheque
            If Not conectar.execute(q) Then GoTo E
            rs.MoveNext
        Wend


        q = "DELETE FROM pagos_a_cuenta_cheques WHERE id_pago_a_cuenta = " & pcta.Id
        If Not conectar.execute(q) Then GoTo E

        Dim che As cheque
        For Each che In pcta.ChequesTerceros
            che.EnCartera = False
            che.IdOrdenPagoOrigen = pcta.Id
            che.FechaEmision = pcta.FEcha
            
            If Not DAOCheques.Guardar(che) Then GoTo E

            q = "INSERT INTO pagos_a_cuenta_cheques VALUES (" & pcta.Id & ", " & che.Id & ")"
            If Not conectar.execute(q) Then GoTo E
            
        Next che

        For Each che In pcta.ChequesPropios
            che.EnCartera = False
            che.IdOrdenPagoOrigen = pcta.Id
            che.FechaEmision = pcta.FEcha
        
        
        If Not DAOCheques.Guardar(che) Then GoTo E
            q = "INSERT INTO pagos_a_cuenta_cheques VALUES (" & pcta.Id & ", " & che.Id & ")"
            If Not conectar.execute(q) Then GoTo E
        Next che
 
        '------------------------------------------------------

        q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM pagos_a_cuenta_operaciones WHERE id_pago_a_cuenta = " & pcta.Id & ")"
        If Not conectar.execute(q) Then GoTo E
        q = "DELETE FROM pagos_a_cuenta_operaciones WHERE id_pago_a_cuenta = " & pcta.Id
        If Not conectar.execute(q) Then GoTo E

        Dim oper As operacion
        For Each oper In pcta.operacionesBanco
            oper.FechaCarga = Now
            If DAOOperacion.Save(oper) Then
                oper.Id = conectar.UltimoId2
                If oper.Id = 0 Then GoTo E
                q = "INSERT INTO pagos_a_cuenta_operaciones VALUES (" & pcta.Id & ", " & oper.Id & ")"
                If Not conectar.execute(q) Then GoTo E
            Else
                GoTo E
            End If
        Next oper

        For Each oper In pcta.operacionesCaja
            oper.FechaCarga = Now
            If DAOOperacion.Save(oper) Then
                oper.Id = conectar.UltimoId2
                If oper.Id = 0 Then GoTo E
                q = "INSERT INTO pagos_a_cuenta_operaciones VALUES (" & pcta.Id & ", " & oper.Id & ")"
                If Not conectar.execute(q) Then GoTo E
            Else
                GoTo E
            End If
        Next oper

    End If

    Guardar = True

    Exit Function
E:
    Guardar = False
    If Nueva Then pcta.Id = 0

End Function


Public Function RemoveFactura(opid As Long, facid As Long) As Boolean
    RemoveFactura = False

    Dim op As OrdenPago
    Set op = DAOOrdenPago.FindById(opid)
    If IsSomething(op) Then
        If op.estado = EstadoOrdenPago_pendiente Then    'si esta aprobada no se puede eliminar una factura
            Dim q As String

            q = "UPDATE AdminComprasFacturasProveedores SET estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id = " & opid
            RemoveFactura = conectar.execute(q)

            If RemoveFactura Then
                q = "DELETE FROM ordenes_pago_facturas WHERE id_factura_proveedor = " & facid
                RemoveFactura = conectar.execute(q)
            Else
                'vuelvo la factura al estado anterior, para no hacer una transaccion
                q = "UPDATE AdminComprasFacturasProveedores SET estado = " & EstadoFacturaProveedor.Saldada & " WHERE id = " & opid
                RemoveFactura = conectar.execute(q)
            End If
            'DaoHistorico.Save "orden_pago_historial", "Factura Id " & facid & " removida, opid"

        End If
    End If

End Function


Public Function Delete(opid As Long, useInternalTransaction As Boolean) As Boolean
    On Error GoTo E

    Dim op As OrdenPago
    Set op = DAOOrdenPago.FindById(opid)

    If useInternalTransaction Then conectar.BeginTransaction

    Dim q As String

    q = "UPDATE AdminComprasFacturasProveedores SET estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id IN (SELECT id_factura_proveedor FROM ordenes_pago_facturas WHERE id_orden_pago = " & opid & ")"
    If Not conectar.execute(q) Then GoTo E
    ' q = "DELETE FROM ordenes_pago_facturas WHERE id_orden_pago = " & opid
    '     If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM ordenes_pago_operaciones WHERE id_orden_pago = " & opid & ")"
    If Not conectar.execute(q) Then GoTo E
    q = "DELETE FROM ordenes_pago_operaciones WHERE id_orden_pago = " & opid
    If Not conectar.execute(q) Then GoTo E


    'se deben borrar los cheques creados para esta orden de pago (solo los propios)
    'fix 14-10-2020
    'q = "UPDATE Cheques SET orden_pago_origen=0, fecha_emision=NULL, monto=0, en_cartera = 0, fecha_vencimiento=NULL, observaciones = NULL, origen= NULL WHERE id IN (SELECT id_cheque FROM ordenes_pago_cheques WHERE id_orden_pago = " & opid & ")"
    q = "UPDATE Cheques SET orden_pago_origen=0, fecha_emision=NULL, monto=0, en_cartera = 0, fecha_vencimiento=NULL, observaciones = NULL, origen= NULL WHERE id IN (SELECT id_cheque FROM ordenes_pago_cheques WHERE id_orden_pago = " & opid & ") and propio=1"
    If Not conectar.execute(q) Then GoTo E

    q = "UPDATE Cheques SET orden_pago_origen=0,en_cartera = 1  WHERE id IN (SELECT id_cheque FROM ordenes_pago_cheques WHERE id_orden_pago = " & opid & ") and propio=0"

    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM ordenes_pago_cheques WHERE id_orden_pago = " & opid
    If Not conectar.execute(q) Then GoTo E

    q = "DELETE FROM ordenes_pago_compensatorios WHERE id_orden_pago = " & opid
    If Not conectar.execute(q) Then GoTo E

    '    q = "DELETE FROM ordenes_pago WHERE id = " & opid
    If Not conectar.execute(q) Then GoTo E
    Dim estado_anterior As EstadoOrdenPago
    estado_anterior = op.estado
    op.estado = EstadoOrdenPago_Anulada
    If Not DAOOrdenPago.Guardar(op, False) Then GoTo E


    DaoHistorico.Save "orden_pago_historial", "OP Anulada", op.Id

    If useInternalTransaction Then conectar.CommitTransaction

    Delete = True
    Exit Function
E:
    op.estado = estado_anterior
    If useInternalTransaction Then conectar.RollBackTransaction
    Delete = False
End Function


Public Function ResumenPagos(ByRef cheques As Collection, ByRef caja As Collection, ByRef bancos As Collection, ByRef comp As Collection, ByRef retenciones As Collection, ByRef cheques3 As Collection, Optional filtro As String, Optional idProveedor As Long = -1) As Boolean
    On Error GoTo err1
    ResumenPagos = True
    Dim q As String
    Dim rs As Recordset

    '#'CHEQUES'
    q = "SELECT b.Nombre,SUM(monto * acm.cambio) as monto FROM ordenes_pago  op " _
      & " INNER JOIN ordenes_pago_cheques opc ON opc.id_orden_pago=op.id " _
      & " LEFT JOIN Cheques c ON opc.id_cheque=c.id " _
      & " LEFT JOIN AdminConfigBancos b ON c.id_banco=b.id " _
      & " LEFT JOIN AdminConfigMonedas acm ON c.id_moneda=acm.id WHERE c.propio=1 and 1=1 " _

If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    q = q & " GROUP BY b.id "

    Dim d As DTONombreMonto

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        cheques.Add d
        rs.MoveNext
    Wend


    '#OPERACIONES CAJA
    q = " SELECT ca.nombre,SUM(monto * acm.cambio ) as monto FROM ordenes_pago op " _
      & " INNER JOIN ordenes_pago_operaciones opo ON opo.id_orden_pago=op.id " _
      & " LEFT JOIN operaciones o ON opo.id_operacion=o.id " _
      & " LEFT JOIN cajas ca ON ca.id=o.cuentabanc_o_caja_id " _
      & " LEFT JOIN AdminConfigMonedas acm ON o.moneda_id=acm.id " _
      & " WHERE o.pertenencia='caja' AND entrada_salida=-1 AND 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & " GROUP BY o.cuentabanc_o_caja_id"

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        caja.Add d
        rs.MoveNext
    Wend

    '#OPERACIONES BANCO
    q = "SELECT  ba.nombre,  SUM(monto * acm.cambio ) AS monto " _
      & " FROM ordenes_pago op   INNER JOIN ordenes_pago_operaciones opo     ON opo.id_orden_pago = op.id " _
      & " LEFT JOIN operaciones o     ON opo.id_operacion = o.id " _
      & " LEFT JOIN AdminConfigCuentas cba     ON cba.id = o.cuentabanc_o_caja_id " _
      & " INNER JOIN AdminConfigBancos ba ON cba.idBanco=ba.id    LEFT JOIN AdminConfigMonedas acm     ON o.moneda_id = acm.id " _
      & " WHERE o.pertenencia='banco' AND entrada_salida=-1 AND 1=1 "
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & "GROUP BY o.cuentabanc_o_caja_id"
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto

        If Not IsNull(rs!nombre) Then d.nombre = rs!nombre Else d.nombre = vbNullString

        bancos.Add d
        rs.MoveNext
    Wend


    '#compensatorios
    q = "SELECT fp.numero_factura, (IF (com.tipo=1,(com.importe * acm.cambio),(com.importe * acm.cambio*-1))) AS monto  FROM ordenes_pago op " _
      & " INNER JOIN ordenes_pago_compensatorios com ON com.id_orden_pago=op.id " _
      & " INNER JOIN AdminComprasFacturasProveedores fp ON com.id_comprobante=fp.id " _
      & " INNER JOIN AdminConfigMonedas acm ON fp.id_moneda=acm.id  where 1=1 "

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If


    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        If Not IsNull(rs!numero_factura) Then d.nombre = rs!numero_factura Else d.nombre = vbNullString

        comp.Add d
        rs.MoveNext
    Wend

    q = "SELECT 'IIBB' AS nombre, SUM(static_total_a_retener) AS monto FROM ordenes_pago op WHERE 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    'q = q & " GROUP BY id "

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        retenciones.Add d
        rs.MoveNext
    Wend


    '#'CHEQUES 3ROS'
    q = "SELECT b.Nombre,SUM(monto * acm.cambio) as monto FROM ordenes_pago  op " _
      & " INNER JOIN ordenes_pago_cheques opc ON opc.id_orden_pago=op.id " _
      & " LEFT JOIN Cheques c ON opc.id_cheque=c.id " _
      & " LEFT JOIN AdminConfigBancos b ON c.id_banco=b.id " _
      & " LEFT JOIN AdminConfigMonedas acm ON c.id_moneda=acm.id WHERE c.propio=0 and 1=1 " _

If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    q = q & " GROUP BY b.id "


    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        If Not IsNull(rs!nombre) Then d.nombre = rs!nombre Else d.nombre = vbNullString
        cheques3.Add d
        rs.MoveNext
    Wend

    Exit Function
err1:
    ResumenPagos = False
End Function


Public Function ExportarColeccion(col As Collection, Optional ProgressBar As Object) As Boolean
    On Error GoTo err1

    ExportarColeccion = True

    Dim xlApplication As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object

    ' Crear una instancia de Excel
    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    ' Fila inicial
    Dim offset As Long
    offset = 3

    ' Escribir encabezados
    xlWorksheet.Cells(offset, 1).value = "Número Pago a Cuenta"
    xlWorksheet.Cells(offset, 2).value = "Proveedor"
    xlWorksheet.Cells(offset, 3).value = "Fecha"
    xlWorksheet.Cells(offset, 4).value = "Moneda"
    xlWorksheet.Cells(offset, 5).value = "Valor"
    xlWorksheet.Cells(offset, 6).value = "Estado"
    xlWorksheet.Cells(offset, 7).value = "Creada"

    ' Formatear encabezados
    With xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 7))
        .Font.Bold = True
        .Interior.Color = &HC0C0C0
    End With

    ' Escribir datos
    Dim PagoACta As clsPagoACta
    Dim initoffset As Long
    initoffset = offset

    If Not ProgressBar Is Nothing Then
        ProgressBar.min = 0
        ProgressBar.max = col.count
    End If

    Dim d As Long
    d = 0

    For Each PagoACta In col
        d = d + 1
        If Not ProgressBar Is Nothing Then ProgressBar.value = d

        offset = offset + 1

        xlWorksheet.Cells(offset, 1).value = PagoACta.Id
        xlWorksheet.Cells(offset, 2).value = PagoACta.Proveedor.RazonSocial
        xlWorksheet.Cells(offset, 3).value = PagoACta.FEcha
        xlWorksheet.Cells(offset, 4).value = PagoACta.moneda.NombreCorto
        xlWorksheet.Cells(offset, 5).value = PagoACta.StaticTotalOrigenes
        xlWorksheet.Cells(offset, 6).value = enums.enumEstadoPagoACuenta(PagoACta.estado)
        xlWorksheet.Cells(offset, 7).value = PagoACta.Creada
    Next

    ' Centrar los datos de la primera columna
    With xlWorksheet.Range(xlWorksheet.Cells(initoffset + 1, 1), xlWorksheet.Cells(offset, 1))
        .HorizontalAlignment = -4108 ' xlCenter
    End With

    ' Aplicar bordes a los datos
    xlWorksheet.Range(xlWorksheet.Cells(initoffset, 1), xlWorksheet.Cells(offset, 7)).Borders.LineStyle = 1 ' xlContinuous

    ' Ajustar el ancho de las columnas
    xlApplication.ScreenUpdating = False
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlApplication.ScreenUpdating = True

    ' Guardar el archivo
    Dim ruta As String
    ruta = Environ$("TEMP")
    If LenB(ruta) = 0 Then ruta = Environ$("TMP")
    If LenB(ruta) = 0 Then ruta = App.path
    ruta = ruta & "\" & funciones.CreateGUID() & ".xls"

    xlWorkbook.SaveAs ruta
    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    ' Abrir el archivo
    ShellExecute -1, "open", ruta, "", "", 4

    ' Limpiar objetos
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    If Not ProgressBar Is Nothing Then ProgressBar.value = 0

    Exit Function

err1:
    ExportarColeccion = False
    If Not xlApplication Is Nothing Then xlApplication.Quit
    If Not ProgressBar Is Nothing Then ProgressBar.value = 0
End Function
