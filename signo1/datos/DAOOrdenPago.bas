Attribute VB_Name = "DAOOrdenPago"
Option Explicit


Public Function FindAbonadoPendienteEnEstaOP(facid As Long, ocid As Long) As Collection

Dim q As String

q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & "),0 ) AS total_pendiente, " _
    & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & "),0 ) AS netogravado_pendiente, " _
     & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & "),0 ) AS otros_pendiente "

'q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS total_pendiente, " _
'    & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS netogravado_pendiente, " _
'     & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS otros_pendiente "



Dim rs As Recordset
 Set rs = conectar.RSFactory(q)

Dim tot As Double, ng As Double, Otros As Double
tot = rs!total_pendiente
ng = rs!netogravado_pendiente
Otros = rs!otros_pendiente

Dim c As New Collection
c.Add tot
c.Add ng
c.Add Otros
Set FindAbonadoPendienteEnEstaOP = c
End Function




Public Function FindAbonadoPendiente(facid As Long, ocid As Long) As Collection

Dim q As String

q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago <> " & ocid & "),0 ) AS total_pendiente, " _
    & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago <> " & ocid & "),0 ) AS netogravado_pendiente, " _
     & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago <> " & ocid & "),0 ) AS otros_pendiente "

'q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS total_pendiente, " _
'    & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS netogravado_pendiente, " _
'     & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS otros_pendiente "



Dim rs As Recordset
 Set rs = conectar.RSFactory(q)

Dim tot As Double, ng As Double, Otros As Double
tot = rs!total_pendiente
ng = rs!netogravado_pendiente
Otros = rs!otros_pendiente

Dim c As New Collection
c.Add tot
c.Add ng
c.Add Otros
Set FindAbonadoPendiente = c
End Function

Public Function FindAbonadoFactura(facid As Long, ocid As Long) As Collection

Dim q As String

q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & " and op1.estado=1),0 ) AS total_pendiente, " _
    & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & " and op1.estado=1),0 ) AS netogravado_pendiente, " _
     & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " and opf.id_orden_pago = " & ocid & " and op1.estado=1),0 ) AS otros_pendiente "

'q = "SELECT IFNULL( (SELECT SUM(total_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS total_pendiente, " _
'    & " IFNULL( (SELECT SUM(neto_gravado_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS netogravado_pendiente, " _
'     & " IFNULL( (SELECT SUM(otros_abonado) FROM ordenes_pago_facturas opf JOIN ordenes_pago op1 ON opf.id_orden_pago=op1.id " _
'    & " WHERE op1.estado=0 AND opf.id_factura_proveedor=" & facid & " AND opf.id_orden_pago=" & ocid & "),0 ) AS otros_pendiente "



Dim rs As Recordset
 Set rs = conectar.RSFactory(q)

Dim tot As Double, ng As Double, Otros As Double
tot = rs!total_pendiente
ng = rs!netogravado_pendiente
Otros = rs!otros_pendiente

Dim c As New Collection
c.Add tot
c.Add ng
c.Add Otros
Set FindAbonadoFactura = c
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

Public Function FindAllByProveedor(provid As Long, Optional cond As String, Optional soloOp As Boolean = False) As Collection
    Dim q As String
    q = "ordenes_pago.id IN (SELECT DISTINCT opf.id_orden_pago from ordenes_pago_facturas opf INNER JOIN AdminComprasFacturasProveedores cfp ON  cfp.id = opf.id_factura_proveedor WHERE cfp.id_proveedor = " & provid & " )"

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

Public Function FindById(id As Long) As OrdenPago
    Set FindById = FindAll("ordenes_pago.id=" & id)(1)
End Function



Public Function FindAllSoloOP(Optional filter As String = "1 = 1", Optional orderBy As String = "1") As Collection
    Dim q As String
    q = "SELECT * " _
        & " From ordenes_pago" _
        & " LEFT JOIN AdminConfigMonedas ON (AdminConfigMonedas.id = ordenes_pago.id_moneda)"

    q = q & " WHERE " & filter
    q = q & " ORDER BY " & orderBy
    Dim col As New Collection
    Dim op As OrdenPago
    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx
    While Not rs.EOF
        Set op = Map(rs, idx, "ordenes_pago", "AdminConfigMonedas")    ', "certificados_retencion")
        If funciones.BuscarEnColeccion(col, CStr(op.id)) Then
            Set op = col.item(CStr(op.id))
        Else
            col.Add op, CStr(op.id)
        End If
        rs.MoveNext
    Wend

    Set FindAllSoloOP = col

End Function
Public Function FindAll(Optional filter As String = "1 = 1", Optional orderBy As String = "1") As Collection
    Dim q As String
    q = "SELECT *, (operaciones.pertenencia + 0) as pertenencia2" _
        & " From ordenes_pago" _
        & " LEFT JOIN ordenes_pago_cheques ON (ordenes_pago.id = ordenes_pago_cheques.id_orden_pago)" _
        & " LEFT JOIN ordenes_pago_operaciones ON (ordenes_pago.id = ordenes_pago_operaciones.id_orden_pago)" _
        & " LEFT JOIN ordenes_pago_facturas ON (ordenes_pago.id = ordenes_pago_facturas.id_orden_pago)" _
        & " LEFT JOIN AdminComprasCuentasContables cuentacontableordenpago ON (ordenes_pago.id_cuenta_contable = cuentacontableordenpago.id)" _
        & " LEFT JOIN operaciones ON (operaciones.id = ordenes_pago_operaciones.id_operacion)" _
        & " LEFT JOIN Cheques ON (Cheques.id = ordenes_pago_cheques.id_cheque)" _
        & " LEFT JOIN Chequeras ON (Chequeras.id = Cheques.id_chequera)" _
        & " LEFT JOIN AdminConfigBancos monbanco ON (monbanco.id = Chequeras.id_banco)" _
        & " LEFT JOIN AdminConfigMonedas monchequera ON (monchequera.id = Chequeras.id_moneda)" _
        & " LEFT JOIN AdminComprasFacturasProveedores ON (AdminComprasFacturasProveedores.id = ordenes_pago_facturas.id_factura_proveedor)" _
        & " LEFT JOIN AdminConfigMonedas ON (AdminConfigMonedas.id = ordenes_pago.id_moneda)" _
        & " LEFT JOIN AdminConfigMonedas monFacProv ON (monFacProv.id = AdminComprasFacturasProveedores.id_moneda)" _
        & " LEFT JOIN AdminConfigFacturasProveedor ON (AdminComprasFacturasProveedores.id_config_factura = AdminConfigFacturasProveedor.id)" _
        & " LEFT JOIN AdminConfigMonedas monedaoperacion ON (monedaoperacion.id = operaciones.moneda_id)" _
        & " LEFT JOIN AdminComprasCuentasContables ON (AdminComprasCuentasContables.id = operaciones.cuenta_contable_id)" _
        & " LEFT JOIN cajas ON (cajas.id = operaciones.cuentabanc_o_caja_id)" _
        & " LEFT JOIN AdminConfigCuentas ON (AdminConfigCuentas.id = operaciones.cuentabanc_o_caja_id)" _
        & " LEFT JOIN AdminConfigMonedas moncuentabancaria ON (moncuentabancaria.id = AdminConfigCuentas.moneda_id)" _
        & " LEFT JOIN AdminConfigMonedas moncheque ON (moncheque.id = Cheques.id_moneda)" _
        & " LEFT JOIN usuarios ON AdminComprasFacturasProveedores.id_usuario_creador=usuarios.id " _
      & " LEFT JOIN AdminConfigBancos ON (AdminConfigBancos.id = AdminConfigCuentas.idBanco)"
    q = q & " LEFT JOIN AdminConfigBancos bancocheque  ON (bancocheque.id = Cheques.id_banco)" _
        & " LEFT JOIN proveedores ON (proveedores.id = AdminComprasFacturasProveedores.id_proveedor)" _
       ' & " LEFT JOIN certificados_retencion ON (certificados_retencion.id_orden_pago = ordenes_pago.id)" _
       ' & " LEFT JOIN retenciones ON (certificados_retencion.id_retencion = retenciones.id)"
       
       
q = q & " LEFT JOIN ordenes_pago_retenciones opr ON opr.id_pago = ordenes_pago.id " _
      & " LEFT JOIN retenciones r ON r.id = opr.id_retencion "
       q = q & " WHERE " & filter
    q = q & " ORDER BY " & orderBy

    Dim col As New Collection
    Dim op As OrdenPago
    Dim fac As clsFacturaProveedor
    Dim che As cheque
    Dim oper As operacion

    Dim idx As Dictionary
    Dim rs As Recordset
    Dim ra As DTORetencionAlicuota
    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set op = Map(rs, idx, "ordenes_pago", "AdminConfigMonedas", "cuentacontableordenpago", "retenciones")    ', "certificados_retencion")

        If funciones.BuscarEnColeccion(col, CStr(op.id)) Then
            Set op = col.item(CStr(op.id))
        Else
            col.Add op, CStr(op.id)
        End If

        Set fac = DAOFacturaProveedor.Map(rs, idx, "AdminComprasFacturasProveedores", "proveedores", "AdminConfigFacturasProveedor", , "monFacProv")

        If IsSomething(fac) Then
            If Not funciones.BuscarEnColeccion(op.FacturasProveedor, CStr(fac.id)) Then
                op.FacturasProveedor.Add fac, CStr(fac.id)
            End If
        End If

        Set che = DAOCheques.Map(rs, idx, "Cheques", "bancocheque", "moncheque", "Chequeras", "monchequera", "monbanco")
        If IsSomething(che) Then
            If che.Propio Then
                If Not funciones.BuscarEnColeccion(op.ChequesPropios, CStr(che.id)) Then
                    op.ChequesPropios.Add che, CStr(che.id)
                End If
            Else
                If Not funciones.BuscarEnColeccion(op.ChequesTerceros, CStr(che.id)) Then
                    op.ChequesTerceros.Add che, CStr(che.id)
                End If
            End If
        End If

        Set oper = DAOOperacion.Map(rs, idx, "operaciones", "AdminComprasCuentasContables", "monedaoperacion", "AdminConfigCuentas", "cajas")
        If IsSomething(oper) Then
            If oper.Pertenencia = Banco Then
                If Not funciones.BuscarEnColeccion(op.OperacionesBanco, CStr(oper.id)) Then
                
                    op.OperacionesBanco.Add oper, CStr(oper.id)
                End If
            ElseIf oper.Pertenencia = caja Then
                If Not funciones.BuscarEnColeccion(op.OperacionesCaja, CStr(oper.id)) Then
                    op.OperacionesCaja.Add oper, CStr(oper.id)
                End If
            End If
        End If

        Set ra = MapAlicuotaRetencion(rs, idx, "opr", "r")
          If IsSomething(ra) Then
            If Not funciones.BuscarEnColeccion(op.RetencionesAlicuota, CStr(ra.Retencion.id)) Then
                    op.RetencionesAlicuota.Add ra, CStr(ra.Retencion.id)
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
                    Optional ByVal TablaRetenciones As String = vbNullString _
                    ) As OrdenPago

    'Optional ByVal tablaCertRetencion As String = vbNullString _

     Dim op As OrdenPago

    'id_certificado_retencion
    Dim id As Long
    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set op = New OrdenPago
        op.id = id

        op.FEcha = GetValue(rs, indice, tabla, "fecha")
        op.CuentaContableDescripcion = GetValue(rs, indice, tabla, "cuenta_contable_desc")
        op.estado = GetValue(rs, indice, tabla, "estado")
        op.alicuota = GetValue(rs, indice, tabla, "alicuota")

        op.StaticTotalFacturas = GetValue(rs, indice, tabla, "static_total_facturas")
        op.StaticTotalFacturasNG = GetValue(rs, indice, tabla, "static_total_factura_ng")
        op.StaticTotalRetenido = GetValue(rs, indice, tabla, "static_total_a_retener")
        op.StaticTotalOrigenes = GetValue(rs, indice, tabla, "static_total_origen")

        op.TipoCambio = GetValue(rs, indice, tabla, "tipo_cambio")
        op.DiferenciaCambio = GetValue(rs, indice, tabla, "dif_cambio")
        op.OtrosDescuentos = GetValue(rs, indice, tabla, "otros_descuentos")
        op.DiferenciaCambioEnNG = GetValue(rs, indice, tabla, "dif_cambio_ng")
        op.DiferenciaCambioEnTOTAL = GetValue(rs, indice, tabla, "dif_cambio_total")
        If LenB(tablaCuentaContable) > 0 Then Set op.CuentaContable = DAOCuentaContable.Map(rs, indice, tablaCuentaContable)
        If LenB(tablaMoneda) > 0 Then Set op.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        'If LenB(tablaCertRetencion) > 0 Then Set op.CertificadoRetencion = DAOCertificadoRetencion.Map(rs, indice, tablaCertRetencion)
    End If

    Set Map = op
End Function



Public Function MapAlicuotaRetencion(rs As Recordset, indice As Dictionary, _
                    tabla As String, _
                  ByVal TablaRetenciones As String) As DTORetencionAlicuota

    'Optional ByVal tablaCertRetencion As String = vbNullString _

     Dim ra As DTORetencionAlicuota

    'id_certificado_retencion
    Dim id As Long
    id = GetValue(rs, indice, tabla, "id_retencion")

    If id > 0 Then
        Set ra = New DTORetencionAlicuota
        ra.alicuotaRetencion = GetValue(rs, indice, tabla, "alicuota")
       Set ra.Retencion = DAORetenciones.Map(rs, indice, TablaRetenciones)
      ra.importe = GetValue(rs, indice, tabla, "total")

            'If LenB(tablaCertRetencion) > 0 Then Set op.CertificadoRetencion = DAOCertificadoRetencion.Map(rs, indice, tablaCertRetencion)
    End If

    Set MapAlicuotaRetencion = ra
End Function




Public Function Save(op As OrdenPago, Optional cascada As Boolean = False) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    Save = Guardar(op, cascada)
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
    Set op = DAOOrdenPago.FindById(op_mem.id)
    
    If Not IsSomething(op) Then
        GoTo err1
    End If

    'VALIDAR BIEN LOS TOTALES ANTES DE PODER APROBAR
    'verificar que las facturas esten todas aprobadsa...
    Dim f As clsFacturaProveedor
    Dim nopago As Double
    Dim esf As EstadoFacturaProveedor
    For Each f In op.FacturasProveedor
        
            Dim fac As clsFacturaProveedor
            Set fac = DAOFacturaProveedor.FindById(f.id)
            
            If fac.estado = EstadoFacturaProveedor.EnProceso Then
                Err.Raise 44, "aprobar op", "La factura " & fac.NumeroFormateado & " no está aprobada. No se pudo aprobar la OP"
            End If
            
            Dim x
            
           Set x = DAOOrdenPago.FindAbonadoPendienteEnEstaOP(fac.id, op.id)
            
             nopago = fac.Total - fac.TotalAbonadoGlobal - (x(1) + x(2) + x(3))
            esf = EstadoFacturaProveedor.Aprobada
            
            If nopago < 0 Then
                Err.Raise 44, "aprobar op", "La factura " & fac.NumeroFormateado & " tiene un error y no se pudo aprobar la OP"
            End If
             If nopago > 0 Then
                esf = EstadoFacturaProveedor.pagoParcial
            Else
                esf = EstadoFacturaProveedor.Saldada
            End If
               conectar.execute "UPDATE AdminComprasFacturasProveedores SET estado = " & esf & " WHERE id = " & fac.id
    Next f





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
DaoHistorico.Save "orden_pago_historial", "OP Aprobada", op.id
    aprobar = True
    If insideTransaction Then conectar.CommitTransaction
    Exit Function
err1:
    op.estado = es
    If insideTransaction Then conectar.RollBackTransaction
    aprobar = False
End Function


Public Function Guardar(op As OrdenPago, Optional cascada As Boolean = False) As Boolean
    
    
'TODO: tengo que revisar que las facturas no esten en otra op aprobada antes de continuar

    Dim q As String
    Dim rs As Recordset
    On Error GoTo E
    Dim Nueva As Boolean: Nueva = False
    If op.id = 0 Then
        Nueva = True
        q = "INSERT INTO ordenes_pago (id_moneda_pago,tipo_cambio,id_moneda, fecha, id_cuenta_contable,cuenta_contable_desc,estado,alicuota,static_total_facturas, static_total_factura_ng, static_total_a_retener, static_total_origen,dif_cambio, otros_descuentos,dif_cambio_ng,dif_cambio_total)" _
            & " VALUES ('id_moneda_pago','tipo_cambio','id_moneda', 'fecha', 'id_cuenta_contable', 'cuenta_contable_desc','0','alicuota','static_total_facturas', 'static_total_factura_ng', 'static_total_a_retener', 'static_total_origen', 'dif_cambio', 'otros_descuentos','dif_cambio_ng','dif_cambio_total')"
    Else
        q = "UPDATE ordenes_pago" _
            & " SET id_moneda = 'id_moneda'," _
            & " fecha = 'fecha'," _
            & " id_cuenta_contable = 'id_cuenta_contable'," _
            & " alicuota = 'alicuota'," _
            & " cuenta_contable_desc = 'cuenta_contable_desc'," _
            & " estado = 'estado'," _
            & " static_total_facturas = 'static_total_facturas'," _
            & " static_total_factura_ng = 'static_total_factura_ng'," _
            & " static_total_a_retener = 'static_total_a_retener'," _
            & " static_total_origen = 'static_total_origen'," _
            & " dif_cambio = 'dif_cambio'," _
            & " otros_descuentos = 'otros_descuentos'," _
            & " tipo_cambio = 'tipo_cambio'," _
            & " id_moneda_pago = 'id_moneda_pago'," _
            & " dif_cambio_ng = 'dif_cambio_ng'," _
            & " dif_cambio_total = 'dif_cambio_total'" _
            & " WHERE id = 'id'"
        q = Replace(q, "'id'", GetEntityId(op))
    End If

    q = Replace(q, "'id_moneda'", GetEntityId(op.moneda))
    q = Replace(q, "'alicuota'", Escape(op.alicuota))
    q = Replace(q, "'fecha'", Escape(op.FEcha))
    q = Replace(q, "'id_cuenta_contable'", GetEntityId(op.CuentaContable))
    q = Replace(q, "'cuenta_contable_desc'", Escape(op.CuentaContableDescripcion))
    q = Replace(q, "'estado'", Escape(op.estado))
    q = Replace(q, "'static_total_facturas'", Escape(op.StaticTotalFacturas))
    q = Replace(q, "'static_total_factura_ng'", Escape(op.StaticTotalFacturasNG))
    q = Replace(q, "'static_total_a_retener'", Escape(op.StaticTotalRetenido))
    q = Replace(q, "'static_total_origen'", Escape(op.StaticTotalOrigenes))
    q = Replace(q, "'dif_cambio'", Escape(op.DiferenciaCambio))
    q = Replace(q, "'otros_descuentos'", Escape(op.OtrosDescuentos))
    q = Replace(q, "'id_moneda_pago'", Escape(op.IdMonedaPago))
    q = Replace(q, "'tipo_cambio'", Escape(op.TipoCambio))
    q = Replace(q, "'dif_cambio_ng'", Escape(op.DiferenciaCambioEnNG))
    q = Replace(q, "'dif_cambio_total'", Escape(op.DiferenciaCambioEnTOTAL))


    If Not conectar.execute(q) Then GoTo E

    If Nueva Then op.id = conectar.UltimoId2()
    If op.id = 0 Then GoTo E

    If cascada Then

        q = "SELECT id_cheque FROM ordenes_pago_cheques WHERE id_orden_pago = " & op.id
        q = q & " AND id_cheque NOT IN (-1"
        If op.ChequesTerceros.count > 0 Then
            q = q & ", " & funciones.JoinCollectionValues(op.ChequesTerceros, ", ", "id")
        End If
        If op.ChequesPropios.count > 0 Then
            q = q & ", " & funciones.JoinCollectionValues(op.ChequesPropios, ", ", "id")
        End If
        q = q & ")"
        Set rs = conectar.RSFactory(q)
        While Not rs.EOF
            q = "UPDATE Cheques SET  en_cartera = 1, observaciones = NULL, origen= NULL WHERE id = " & rs!id_cheque
            If Not conectar.execute(q) Then GoTo E
            rs.MoveNext
        Wend


        q = "DELETE FROM ordenes_pago_cheques WHERE id_orden_pago = " & op.id
        If Not conectar.execute(q) Then GoTo E

        Dim che As cheque
        For Each che In op.ChequesTerceros
            che.EnCartera = False
            che.IdOrdenPagoOrigen = op.id
            che.FechaEmision = op.FEcha
            'che.Observaciones = "Utilizado en Orden de Pago Nº " & op.Id
            If Not DAOCheques.Guardar(che) Then GoTo E

            q = "INSERT INTO ordenes_pago_cheques VALUES (" & op.id & ", " & che.id & ")"
            If Not conectar.execute(q) Then GoTo E
        Next che

        For Each che In op.ChequesPropios
            che.EnCartera = False
            che.IdOrdenPagoOrigen = op.id
            che.FechaEmision = op.FEcha
            'che.Observaciones = "Utilizado en Orden de Pago Nº " & op.Id
            If op.EsParaFacturaProveedor And op.FacturasProveedor.count > 0 Then che.OrigenDestino = op.FacturasProveedor(1).Proveedor.RazonSocial
            If Not DAOCheques.Guardar(che) Then GoTo E

            q = "INSERT INTO ordenes_pago_cheques VALUES (" & op.id & ", " & che.id & ")"
            If Not conectar.execute(q) Then GoTo E
        Next che
        '------------------------------------------------------

        '------------------------------------------------------

        Dim fcp As clsFacturaProveedor
        For Each fcp In op.FacturasProveedor
            q = "UPDATE AdminComprasFacturasProveedores SET tipo_cambio_pago= " & fcp.TipoCambioPago & ", estado = " & EstadoFacturaProveedor.Aprobada & " WHERE id = " & fcp.id
            If Not conectar.execute(q) Then GoTo E
        Next





        q = "DELETE FROM ordenes_pago_facturas WHERE id_orden_pago = " & op.id
        If Not conectar.execute(q) Then GoTo E


        Dim es As EstadoFacturaProveedor
        Dim nopago As Double
        Dim compe As Compensatorio
        Dim cp As Compensatorio
        Dim fac As clsFacturaProveedor
        For Each fac In op.FacturasProveedor
            q = "INSERT INTO ordenes_pago_facturas VALUES (" & op.id & ", " & fac.id & "," & fac.ImporteTotalAbonado & "," & fac.NetoGravadoAbonado & "," & fac.OtrosAbonado & ")"

             If Not conectar.execute(q) Then GoTo E

'            If BuscarEnColeccion(op.Compensatorios, CStr(fac.id)) Then
'
'                Set compe = op.Compensatorios(CStr(fac.id))
'                nopago = compe.Monto
'            Else
'                nopago = 0
'            End If
            nopago = 0
         'validar si se pagan facturas o compensatorios
           
           For Each cp In op.Compensatorios
                nopago = nopago + cp.Monto
                
            Next cp
            
            
            nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado
            
             q = "DELETE FROM orden_pago_deuda_compensatorios WHERE id_orden_pago = " & op.id
        If Not conectar.execute(q) Then GoTo E

        Dim c As Compensatorio
            
        For Each c In op.DeudaCompensatorios
        If c.Cancelado Then Err.Raise "El compensatorio " & c.id & " ya está cancelado!"
         q = "INSERT INTO  orden_pago_deuda_compensatorios (id_orden_pago, id_compensatorio) values (" & op.id & "," & c.id & ")"
            
            
            If Not conectar.execute(q) Then GoTo E

        q = "UPDATE  ordenes_pago_compensatorios   SET cancelado=1 where id_orden_pago=" & op.id & " and id=" & c.id
            
            
            If Not conectar.execute(q) Then GoTo E


        Next c
            
            
            
            
            'If op.estado = EstadoOrdenPago_Aprobada Then
            
             'nopago = fac.Total - fac.TotalAbonadoGlobal - fac.TotalAbonado
            es = EstadoFacturaProveedor.Aprobada
             If nopago > 0 Then
                es = EstadoFacturaProveedor.pagoParcial
            Else
                es = EstadoFacturaProveedor.Saldada
            End If

            ' Else

            ' End If
            q = "UPDATE AdminComprasFacturasProveedores SET estado = " & es & " WHERE id = " & fac.id
            'q = "UPDATE AdminComprasFacturasProveedores SET estado = " & fac.AnalizarEstado & " WHERE id = " & fac.Id
            If Not conectar.execute(q) Then GoTo E


        Next fac


        '------------------------------------------------------


        q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.id & ")"
        If Not conectar.execute(q) Then GoTo E
        q = "DELETE FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.id
        If Not conectar.execute(q) Then GoTo E

        Dim oper As operacion
        For Each oper In op.OperacionesBanco
            'oper.IdPertenencia = op.Id no se usa creo
            oper.FechaCarga = Now
            If DAOOperacion.Save(oper) Then
                oper.id = conectar.UltimoId2
                If oper.id = 0 Then GoTo E
                q = "INSERT INTO ordenes_pago_operaciones VALUES (" & op.id & ", " & oper.id & ")"
                If Not conectar.execute(q) Then GoTo E
            Else
                GoTo E
            End If
        Next oper

        For Each oper In op.OperacionesCaja
            'oper.IdPertenencia = op.Id no se usa creo
            oper.FechaCarga = Now
            If DAOOperacion.Save(oper) Then
                oper.id = conectar.UltimoId2
                If oper.id = 0 Then GoTo E
                q = "INSERT INTO ordenes_pago_operaciones VALUES (" & op.id & ", " & oper.id & ")"
                If Not conectar.execute(q) Then GoTo E
            Else
                GoTo E
            End If
        Next oper
        
        
        
        'guardo las retenciones aplicadas
'   q = "DELETE FROM ordenes_pago_cheques WHERE id_orden_pago = " & op.id
'        If Not conectar.execute(q) Then GoTo E
'
'        Dim r As DTORetencionAlicuota
'        For Each r In op.RetencionesAlicuota
'
'
'
'          ' q = "INSERT INTO retenciones_alicuotas VALUES (" & op.id & ", " & che.id & ")"
'          '  If Not conectar.execute(q) Then GoTo E
'        Next r


        'guardo los compensatorios
        q = "DELETE FROM ordenes_pago_compensatorios WHERE id_orden_pago = " & op.id
        If Not conectar.execute(q) Then GoTo E

        Dim c1 As Compensatorio

        For Each c1 In op.Compensatorios
            c1.IdOrdenPago = op.id
            If Not DAOCompensatorios.Guardar(c1) Then GoTo E

        Next c1
        
        
           'guardo las retenciones
        q = "DELETE FROM ordenes_pago_retenciones WHERE id_pago = " & op.id
        If Not conectar.execute(q) Then GoTo E

        Dim ra As DTORetencionAlicuota
        
         For Each ra In op.RetencionesAlicuota
            
           
           q = " INSERT INTO ordenes_pago_retenciones (id_pago,id_retencion,fecha,alicuota,total) values('id_pago','id_retencion','fecha','alicuota','total')"
           
             q = Replace(q, "'id_pago'", GetEntityId(op))
             q = Replace(q, "'id_retencion'", GetEntityId(ra.Retencion))
             q = Replace(q, "'fecha'", Escape(op.FEcha))
             q = Replace(q, "'alicuota'", Escape(ra.alicuotaRetencion))
             q = Replace(q, "'total'", Escape(ra.importe))
              If Not conectar.execute(q) Then GoTo E
        Next ra

    End If
    
    Dim msg As String
    msg = "OP Creada"
    
    If Not Nueva Then msg = "OP Actualizada"
    DaoHistorico.Save "orden_pago_historial", msg, op.id
    
    Guardar = True

    Exit Function
E:
    Guardar = False
    If Nueva Then op.id = 0

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


DaoHistorico.Save "orden_pago_historial", "OP Anulada", op.id

    If useInternalTransaction Then conectar.CommitTransaction

    Delete = True
    Exit Function
E:
    op.estado = estado_anterior
    If useInternalTransaction Then conectar.RollBackTransaction
    Delete = False
End Function
Public Function ResumenPagos(ByRef Cheques As Collection, ByRef caja As Collection, ByRef bancos As Collection, ByRef comp As Collection, ByRef retenciones As Collection, ByRef cheques3 As Collection, Optional filtro As String, Optional idProveedor As Long = -1) As Boolean
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
        Cheques.Add d
        rs.MoveNext
    Wend



    '
    '
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


    '
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



Public Function PrintOP(Orden As OrdenPago, pic As PictureBox) As Boolean
    Dim TAB1 As Integer
    Dim TAB2 As Integer
    Dim maxw As Single
    Dim c As Long
    Dim mtxt As String
    Dim textw As Single
    Dim lmargin As Integer



    pic.Picture = LoadResPicture(101, vbResBitmap)

    Dim A As Single
    lmargin = 720



    TAB1 = 300
    TAB2 = 300
    Printer.CurrentY = lmargin
    maxw = Printer.Width - lmargin * 2
    A = lmargin + (maxw - 3200) / 2
    Printer.PaintPicture pic.Picture, A, 100, 3200, 600
    Printer.FontBold = True

    Printer.FontSize = 12
    mtxt = "Orden de Pago Nº " & Orden.id
    textw = Printer.TextWidth(mtxt)
    Printer.CurrentX = lmargin + (maxw - textw) / 2
    Printer.Print mtxt
    Printer.FontSize = 10

    Printer.CurrentX = lmargin
    Printer.Print "Fecha: ";
    Printer.FontBold = False
    Printer.Print Orden.FEcha



    If Orden.FacturasProveedor.count > 0 Then

        Printer.FontBold = True
        Printer.CurrentX = lmargin
        Printer.Print "Proveedor: ";
        Printer.FontBold = False
        Printer.Print Orden.FacturasProveedor(1).Proveedor.RazonSocial
    End If
  
    Dim existeIIBB As Boolean
    existeIIBB = False
    
    Dim ra As DTORetencionAlicuota
    For Each ra In Orden.RetencionesAlicuota
   
            Printer.FontBold = True
            Printer.CurrentX = lmargin
            Printer.Print "Alícuota " & ra.Retencion.nombre & ": ";
            Printer.FontBold = False
            Printer.Print ra.alicuotaRetencion & "%"
            
            If ra.Retencion.id = 5 Then existeIIBB = True
                
      Next
      
      If Not existeIIBB Then
              Printer.FontBold = True
            Printer.CurrentX = lmargin
            Printer.Print "Alícuota IIBB BS AS: ";
            Printer.FontBold = False
            Printer.Print Orden.alicuota & "%"
    End If




    Dim cert As CertificadoRetencion
    Set cert = DAOCertificadoRetencion.FindByOrdenPago(Orden.id)


    Dim allcert As Collection
    Set allcert = DAOCertificadoRetencion.FindAllByOrdenPago(Orden.id)
    
   ' For Each cert In cert.CertificadoRetencion
   ' Printer.Print cert.id & "%"
   ' Next
    
   If allcert Is Nothing Then
   Set allcert = New Collection
   End If
   
    If allcert Is Nothing Or Not IsSomething(allcert) Or allcert.count = 0 Then
        Printer.FontBold = True
        Printer.CurrentX = lmargin
        Printer.Print "Certificado IIBB Nº: ";
        Printer.FontBold = False
        Printer.Print " NO POSEE"
    Else
        If allcert.count = 1 Then
            Printer.FontBold = True
            Printer.CurrentX = lmargin
            Printer.Print "Certificado IIBB Nº: ";
            Printer.FontBold = False
            Printer.Print allcert(1).id
        Else
            
             Printer.FontBold = True
            Printer.CurrentX = lmargin
            Printer.Print "Certificados IIBB Nº: ";
            Printer.FontBold = False
            
            Dim c1 As CertificadoRetencion
            Dim t1 As String
            For Each c1 In allcert
               t1 = t1 & c1.id & " "
            Next
          
                Printer.Print t1
            
        End If
    End If
    
      
    
    
    
    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Moneda: ";
    Printer.FontBold = False
    Printer.Print Orden.moneda.NombreCorto & " " & Orden.moneda.NombreLargo

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Otros Descuentos: ";
    Printer.FontBold = False
    Printer.Print Orden.moneda.NombreCorto & " " & Orden.OtrosDescuentos

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Dif. Por Tipo de Cambio: ";
    Printer.FontBold = False
    Printer.Print Orden.moneda.NombreCorto & " " & Orden.DiferenciaCambio

    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)



    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Facturas: "
    Printer.FontBold = False
    Printer.FontSize = 8
    Set Orden.FacturasProveedor = DAOFacturaProveedor.FindAllByOrdenPago(Orden.id)
    Dim f As clsFacturaProveedor
    Dim facs As New Collection
    c = 0
    For Each f In Orden.FacturasProveedor
        c = c + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print f.NumeroFormateado & String$(8, " del ") & f.FEcha & String$(8, " por ") & f.moneda.NombreCorto & " " & f.Total
    Next f
    If c = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE FACTURAS ASOCIADAS"
    End If
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.FontSize = 10
    Printer.CurrentX = lmargin
    Printer.FontBold = True
    Printer.Print "Valores: "
    Printer.FontSize = 8
    Printer.CurrentX = lmargin
    Printer.Print "Cheques Propios: "
    Printer.FontBold = False
    Dim cheq As cheque
    Dim tmpCol As New Collection
    c = 0
    For Each cheq In Orden.ChequesPropios
        c = c + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print cheq.numero & String$(8, " ") & cheq.Banco.nombre & String$(24, " ") & cheq.FechaVencimiento & String$(8, " ") & cheq.moneda.NombreCorto & " " & cheq.Monto
    Next cheq
    If c = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE CHEQUES PROPIOS"
    End If
    Printer.Print
    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Cheques de Terceros: "
    Printer.FontBold = False
    Set tmpCol = New Collection
    c = 0
    For Each cheq In Orden.ChequesTerceros
        c = c + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print cheq.numero & String$(8, " ") & cheq.Banco.nombre & String$(16, " ") & cheq.FechaVencimiento & String$(8, " ") & cheq.moneda.NombreCorto & " " & cheq.Monto
    Next cheq
    If c = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE CHEQUES DE TERCEROS"
    End If
    Printer.Print
    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Transferencias: "
    Printer.FontBold = False

    Dim op As operacion
    Set tmpCol = New Collection
    c = 0
    For Each op In Orden.OperacionesBanco
        c = c + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print op.FechaOperacion & String$(8, " ") & op.moneda.NombreCorto & " " & op.Monto
    Next op
    If c = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE TRANSFERENCIAS"
    End If
    Printer.Print
    Printer.FontBold = True
    Printer.CurrentX = lmargin + TAB1
    Printer.Print "Efectivo: "
    Printer.FontBold = False


    Set tmpCol = New Collection
    c = 0
    For Each op In Orden.OperacionesCaja
        c = c + 1
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print op.FechaOperacion & String$(8, " ") & op.moneda.NombreCorto & " " & op.Monto
    Next op
    If c = 0 Then
        Printer.CurrentX = lmargin + TAB1 + TAB2
        Printer.Print "NO POSEE OPERACIONES EN EFECTIVO"
    End If

    Printer.FontSize = 11
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Total Facturas: ";
    Printer.FontBold = False
    Printer.Print Orden.moneda.NombreCorto & " " & Orden.StaticTotalFacturas

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Total Retenido: ";
    Printer.FontBold = False
    Printer.Print Orden.moneda.NombreCorto & " " & Orden.StaticTotalRetenido

    Printer.FontBold = True
    Printer.CurrentX = lmargin
    Printer.Print "Total Abonado: ";
    Printer.FontBold = False
    Printer.Print Orden.moneda.NombreCorto & " " & Orden.StaticTotalOrigenes
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.EndDoc
    
        DaoHistorico.Save "orden_pago_historial", "OP Impresa", Orden.id
End Function

