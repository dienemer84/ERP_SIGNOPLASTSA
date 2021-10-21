Attribute VB_Name = "DAOFactura"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                     (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                      ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function PagosRealizados(factura_id As Long) As Double
    Dim q As String
    q = "SELECT IFNULL(SUM(monto_pagado),0) as suma FROM AdminRecibosDetalleFacturas WHERE idFactura = " & factura_id
    Dim r As Recordset

    Set r = conectar.RSFactory(q)

    If r.EOF Then
        PagosRealizados = 0
    Else
        PagosRealizados = r!suma
    End If

End Function


Public Function FindAllByEstadoSaldoAndCliente(estado As TipoSaldadoFactura, Estado1 As EstadoFacturaCliente, Optional cliente_id As Long = -1) As Collection


    Dim filtro As String

    filtro = "saldada=" & estado & " and AdminFacturas.estado =" & Estado1

    If (cliente_id > 0) Then filtro = filtro & " and idCliente=" & cliente_id

    Set FindAllByEstadoSaldoAndCliente = DAOFactura.FindAll(filtro)


End Function



Public Function FindAllNoSaldadasNiVencidas(desde As Date, hasta As Date, Optional cliente_id As Long = 0) As Collection
    Dim F As String
    F = " ((AdminFacturas.saldada IN (0,2) AND NOW() <=  ADDDATE(AdminFacturas.FechaEmision, AdminFacturas.FormaPago)) " _
        & " OR (AdminFacturas.saldada IN (0,2) AND NOW() >  ADDDATE(AdminFacturas.FechaEmision, AdminFacturas.FormaPago) AND AdminFacturas.propuesta >= NOW())) " _
        & " AND (IF(IFNULL(propuesta, ADDDATE(FechaEmision,FormaPago)), ADDDATE(FechaEmision,FormaPago) >= " & conectar.Escape(desde) & ", propuesta >= " & conectar.Escape(desde) & " ))" _
        & " AND (IF(IFNULL(propuesta, ADDDATE(FechaEmision,FormaPago)), ADDDATE(FechaEmision,FormaPago) <= " & conectar.Escape(hasta) & ", propuesta <= " & conectar.Escape(hasta) & "))" _

If cliente_id <> 0 Then
        F = F & " AND AdminFacturas.idCliente = " & cliente_id
    End If

    Set FindAllNoSaldadasNiVencidas = FindAll(F)
End Function

Public Function FindAllNoSaldadasTotalByCliente(cliente_id As Long, Optional includeDetalles As Boolean = False, Optional includeEntregasWithDetalles As Boolean = False) As Collection
    Set FindAllNoSaldadasTotalByCliente = FindAll("AdminFacturas.idCliente = " & cliente_id & " AND AdminFacturas.estado <> " & EstadoFacturaCliente.Anulada & " AND AdminFacturas.saldada IN (" & TipoSaldadoFactura.NoSaldada & ", " & TipoSaldadoFactura.SaldadoParcial & "," & TipoSaldadoFactura.notaCreditoParcial & ")", includeDetalles, includeEntregasWithDetalles)
End Function
Public Function FindAll(Optional ByVal filter As String = "1 = 1", Optional includeDetalles As Boolean = False, Optional includeEntregasWithDetalles As Boolean = False, Optional Orden As String = vbNullString) As Collection

    On Error GoTo err1
    Dim q As String
    q = "SELECT *, ADDDATE(AdminFacturas.FechaEmision, AdminFacturas.FormaPago) AS FechaVencimiento " _
        
        If includeDetalles Then
            
            q = q & ",CAST((SELECT   GROUP_CONCAT(DISTINCT r.numero SEPARATOR ',') AS lista_remitos FROM AdminFacturasDetalleAplicacionRemitos a " _
                       & "INNER JOIN entregas e ON e.id=a.idRemitoDetalle INNER JOIN remitos r ON e.Remito = r.id  WHERE a.idFacturaDetalle= AdminFacturasDetalleNueva.id) AS CHAR) as lista_remitos_aplicados "
        
       
            q = q & ",CAST((SELECT   COUNT(DISTINCT r.numero) AS cantidad_remitos FROM AdminFacturasDetalleAplicacionRemitos a " _
                       & "INNER JOIN entregas e ON e.id=a.idRemitoDetalle INNER JOIN remitos r ON e.Remito = r.id  WHERE a.idFacturaDetalle= AdminFacturasDetalleNueva.id) AS CHAR) as cantidad_remitos_aplicados "
        
        End If
        q = q & " From AdminFacturas" _
        & " LEFT JOIN AdminConfigFacturasTiposDiscriminado acftd      ON (       acftd.id = AdminFacturas.id_tipo_discriminado    ) " _
        & " LEFT JOIN AdminConfigFacturasTipos acft     ON (acftd.id_tipo_factura = acft.id)  " _
        & " LEFT JOIN AdminConfigFacturasTiposDiscriminadoIva acftdi      ON (       acftd.`id` = acftdi.`id_tipo_factura_discriminado`   ) " _
        & " LEFT JOIN AdminConfigIVA ivaFac      ON (ivaFac.idIVA = acftdi.id_iva) " _
        & " LEFT JOIN AdminConfigFacturaPuntoVenta pv      ON (acftd.id_punto_venta = pv.id) " _
        & " LEFT JOIN clientes ON (AdminFacturas.idCliente = clientes.id)" _
        & " LEFT JOIN Localidades ON (clientes.id_localidad = Localidades.ID)" _
        & " LEFT JOIN Provincia ON (clientes.id_provincia = Provincia.ID)" _
        & " LEFT JOIN AdminConfigIVA iva ON (iva.idIVA = clientes.iva)" _
        & " LEFT JOIN AdminConfigMonedas ON (AdminFacturas.idMoneda = AdminConfigMonedas.id)" _
        & " LEFT JOIN usuarios ON AdminFacturas.idUsuarioEmision=usuarios.id " _
        & " LEFT JOIN usuarios as usuarios2 ON AdminFacturas.idUsuarioAprobacion=usuarios2.id "


    If includeDetalles Then
        q = q & " LEFT JOIN  AdminFacturasDetalleNueva ON AdminFacturasDetalleNueva.idFactura = AdminFacturas.id "
        If includeEntregasWithDetalles Then
            q = q & " LEFT JOIN  entregas ON entregas.id = AdminFacturasDetalleNueva.idEntrega "
        End If
    End If

    q = q & " WHERE " & filter

    Dim col As New Collection
    Dim F As Factura
    Dim idx As Dictionary
    Dim deta As FacturaDetalle
    Dim rs As Recordset


' Modificar el orden en el que se muestran los comprobantes en el grid de Facturas Emitidas
' Nemer 29/08/2020

    If LenB(Orden) > 0 Then
        q = q & " ORDER BY " & Orden
    Else
        q = q & " ORDER BY AdminFacturas.FechaEmision DESC"
    End If

    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, idx
    While Not rs.EOF
        Set F = Map(rs, idx, "AdminFacturas", "clientes", "AdminConfigMonedas", "iva", "acftd", "ivaFac", "acft", "pv")
Debug.Print F.id
        If funciones.BuscarEnColeccion(col, CStr(F.id)) Then
            Set F = col.item(CStr(F.id))
        Else
            F.Detalles = New Collection
            col.Add F, CStr(F.id)
        End If

        If includeDetalles Then
            Set deta = DAOFacturaDetalles.Map(rs, idx, "AdminFacturasDetalleNueva")
            

            
            If IsSomething(deta) Then
                  If rs!cantidad_remitos_aplicados > 0 Then
                    deta.ListaRemitosAplicados = rs!lista_remitos_aplicados
                  End If
                    deta.CantidadRemitosAplicados = rs!cantidad_remitos_aplicados
                If Not funciones.BuscarEnColeccion(F.Detalles, CStr(deta.id)) Then
                    Set deta.Factura = F
                    F.Detalles.Add deta, CStr(deta.id)
                End If

                If includeEntregasWithDetalles Then
                    Set deta.detalleRemito = DAORemitoSDetalle.Map(rs, idx, "entregas")
                End If
            End If


        End If

        rs.MoveNext
    Wend

    Set FindAll = col

    Exit Function
    'Debug.Print "DAOFactura.FindAll()", GetTickCount - duracion
err1:
    Set FindAll = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FindById(id As Long, Optional includeDetalles As Boolean = False, Optional includeEntregasWithDetalles As Boolean = False) As Factura
    Dim col As Collection: Set col = FindAll("AdminFacturas.id = " & id, includeDetalles, includeEntregasWithDetalles)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function
Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaCliente As String = vbNullString, _
                    Optional tablaMoneda As String = vbNullString, Optional tablaClienteIva As String = vbNullString, Optional tablaFactTipoFacturaDiscriminado As String = vbNullString, Optional tablaIVATipo As String = vbNullString, Optional tablaTipoFactura As String = vbNullString, Optional tablaPuntoVenta As String = vbNullString) As Factura
    Dim F As Factura
    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    If id > 0 Then
        Set F = New Factura
        F.id = id
        F.numero = GetValue(rs, indice, tabla, "NroFactura")
        'f.Descuento = GetValue(rs, indice, tabla, "descuento")
        F.EstaImpresa = GetValue(rs, indice, tabla, "impresa")
        F.EstaDiscriminada = GetValue(rs, indice, tabla, "discriminada")
        F.FechaEmision = GetValue(rs, indice, tabla, "FechaEmision")
        F.FechaEntrega = GetValue(rs, indice, tabla, "fecha_entrega")
        F.CBU = GetValue(rs, indice, tabla, "CBU")
        F.AnulacionAFIP = GetValue(rs, indice, tabla, "anulacion_afip")
        F.MotivosAnulacionAFIP = GetValue(rs, indice, tabla, "motivo_anulacion_afip")
        F.fechaPago = GetValue(rs, indice, tabla, "fecha_pago")
        F.esCredito = GetValue(rs, indice, tabla, "EsCredito")
        F.AprobadaAFIP = GetValue(rs, indice, tabla, "aprobacion_afip")
        'fce_nemer_29052020
        F.FechaVtoDesde = GetValue(rs, indice, tabla, "fecha_vto_desde")
        F.FechaVtoHasta = GetValue(rs, indice, tabla, "fecha_vto_hasta")
        F.TextoAdicional = GetValue(rs, indice, tabla, "texto_adicional")
        F.CAE = GetValue(rs, indice, tabla, "cae")
        F.CAEVto = GetValue(rs, indice, tabla, "cae_vto")
        F.FechaVencimientoSQL = GetValue(rs, indice, vbNullString, "FechaVencimiento")
        F.ConceptoIncluir = GetValue(rs, indice, tabla, "id_concepto_incluir")
        F.OrdenCompra = GetValue(rs, indice, tabla, "OrdenCompra")
        F.Saldado = GetValue(rs, indice, tabla, "saldada")
        F.observaciones = GetValue(rs, indice, tabla, "observaciones")
                F.Opcional27 = GetValue(rs, indice, tabla, "opcional27")
        F.observaciones_cancela = GetValue(rs, indice, tabla, "observaciones_cancela")
        If Trim(F.observaciones) = "-" Or (F.observaciones) = "." Then F.observaciones = vbNullString
        'If F.id = 6415 Then Stop
        If LenB(tablaFactTipoFacturaDiscriminado) Then Set F.Tipo = DAOTipoFacturaDiscriminado.Map(rs, indice, tablaFactTipoFacturaDiscriminado, tablaTipoFactura, tablaPuntoVenta)
        If LenB(tablaIVATipo) Then Set F.TipoIVA = DAOTipoIva.Map(rs, indice, tablaIVATipo)
        F.AlicuotaAplicada = GetValue(rs, indice, tabla, "alicuotaAplicada")
        'ver bien que onda
        Dim porcenPercep As Double
        porcenPercep = GetValue(rs, indice, tabla, "AliPercIB")
        F.AlicuotaPercepcionesIIBB = porcenPercep
        F.estado = GetValue(rs, indice, tabla, "estado")
        F.CantDiasPago = GetValue(rs, indice, tabla, "FormaPago")
        F.FechaAprobacion = GetValue(rs, indice, tabla, "FechaAprobacion")
        F.origenFacturado = GetValue(rs, indice, tabla, "origenFacturado")
        F.FechaPropuestaPago = GetValue(rs, indice, tabla, "propuesta")
        F.Cancelada = GetValue(rs, indice, tabla, "cancelada")
        F.MotivoNC = GetValue(rs, indice, tabla, "nc_motivo")
        F.CambioAPatron = GetValue(rs, indice, tabla, "cambio_a_patron")
        If LenB(tablaMoneda) > 0 Then Set F.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaCliente) > 0 Then Set F.cliente = DAOCliente.Map(rs, indice, tablaCliente, tablaClienteIva, "Localidades", "", "Provincia")
        
        
        'MAP DE USUARIOS PARA FACTURAS VENTAS
        Set F.usuarioCreador = DAOUsuarios.Map(rs, indice, "usuarios")
        Set F.UsuarioAprobacion = DAOUsuarios.Map(rs, indice, "usuarios2")


        F.TasaAjusteMensual = GetValue(rs, indice, tabla, "tasa_ajuste_mensual")
        Set F.TotalEstatico = New FacturaTotalEstatico
        F.TotalEstatico.Total = GetValue(rs, indice, tabla, "total_estatico")
        F.TotalEstatico.TotalExento = GetValue(rs, indice, tabla, "total_exento_estatico")
        F.TotalEstatico.TotalIVA = GetValue(rs, indice, tabla, "total_iva_estatico")
        F.TotalEstatico.TotalIVADiscrimandoONo = GetValue(rs, indice, tabla, "total_iva_discono_estatico")
        F.TotalEstatico.TotalNetoGravado = GetValue(rs, indice, tabla, "total_neto_estatico")
        F.TotalEstatico.TotalPercepcionesIB = GetValue(rs, indice, tabla, "total_perIB_estatico")
        F.IdMonedaAjuste = GetValue(rs, indice, tabla, "id_moneda_ajuste")
        F.TipoCambioAjuste = GetValue(rs, indice, tabla, "tipo_cambio_ajuste")


    End If

    Set Map = F
End Function
Public Function Save(Factura As Factura, Optional Cascade As Boolean = False) As Boolean
  On Error GoTo err1
    
    conectar.BeginTransaction
    Save = True
    If Not Guardar(Factura, Cascade) Then
        GoTo err1
    End If
    conectar.CommitTransaction
    Exit Function
err1:
    conectar.RollBackTransaction
    Save = False
End Function

Public Function hacktipofactura(idIVA, Tipo) As Long
    '''''''''''''''''''''''''''''''''''''' HACK
    Dim qTemp As String
    Dim TipoFactura As Long: TipoFactura = -1
    Dim tmpRS As Recordset
    qTemp = "SELECT id FROM AdminConfigFacturas WHERE idIVA = " & idIVA & " AND TipoFactura = " & Tipo

    Set tmpRS = conectar.RSFactory(qTemp)
    If Not tmpRS.EOF Then TipoFactura = tmpRS!id
    hacktipofactura = TipoFactura
End Function


Public Function RechazoAfip(F As Factura)
On Error GoTo err1

    If Not F.esCredito Then
        Err.Raise 1235, "recha", "La anulación corresponde a facturas de crédito"
    End If

    
     
     Dim q As String
  q = "Update sp.AdminFacturas  SET motivo_anulacion_afip='motivo_anulacion_afip',anulacion_afip='anulacion_afip' where id='id'"
  
    q = Replace$(q, "'id'", conectar.Escape(F.id))
    q = Replace$(q, "'motivo_anulacion_afip'", conectar.Escape(F.MotivosAnulacionAFIP))
    q = Replace$(q, "'anulacion_afip'", conectar.Escape(F.AnulacionAFIP))
   
    If Not conectar.execute(q) Then
           Err.Raise 112233, "No se pudo actualizar el estado de rechazo del comprobante"
    End If
Exit Function
err1:
Err.Raise Err.Number, Err.Description
End Function

Public Function ActualizarCAE(F As Factura)
On Error GoTo err1

    If Not F.Tipo.PuntoVenta.CaeManual Then
        Err.Raise 1235, "cae", "no se puede actualizar datos de CAE en comprobantes con PV que no acepten cae manual"
    End If

    If F.estado = EstadoFacturaCliente.Anulada Or F.estado = EstadoFacturaCliente.EnProceso Then
            Err.Raise 1444, "cae", "Solo puede modificar CAE en comprobantes aprobados. "
     End If
     
     
     Dim q As String
  q = "Update sp.AdminFacturas  SET cae='cae',cae_vto='cae_vto', aprobacion_afip='aprobacion_afip' where id='id'"
  
    q = Replace$(q, "'id'", conectar.Escape(F.id))
    q = Replace$(q, "'cae'", conectar.Escape(F.CAE))
    q = Replace$(q, "'cae_vto'", conectar.Escape(F.CAEVto))
    q = Replace$(q, "'aprobacion_afip'", conectar.Escape(F.AprobadaAFIP))
    If Not conectar.execute(q) Then
           Err.Raise 112233, "No se pudieron actualizar los datos del CAE"
    End If
Exit Function
err1:
Err.Raise Err.Number, Err.Description
End Function
Public Function Guardar(F As Factura, Optional Cascade As Boolean = False) As Boolean
    Dim q As String
    '    conectar.BeginTransaction
    Dim esNueva As Boolean

    If F.id > 0 Then

        '& " descuento = 'descuento' , "
        q = "Update sp.AdminFacturas  SET " _
            & " NroFactura = 'NroFactura' , idCliente = 'idCliente' ,  tipoFactura_borrar = 'tipoFactura_borrar' ," _
            & " idMoneda = 'idMoneda' ,cae='cae',cae_vto='cae_vto',aprobacion_afip='aprobacion_afip', " _
            & " FechaEmision = 'FechaEmision' , EsCredito = 'EsCredito'," _
            & " idUsuarioEmision = 'idUsuarioEmision' ," _
            & " FechaAprobacion = 'FechaAprobacion' , " _
            & " idUsuarioAprobacion = 'idUsuarioAprobacion' ," _
            & " OrdenCompra = 'OrdenCompra' , " _
            & " origenFacturado = 'origenFacturado' , " _
            & " estado = 'estado' ,id_concepto_incluir='id_concepto_incluir' , " _
            & " alicuotaAplicada = 'alicuotaAplicada' , " _
            & " discriminada = 'discriminada' , " _
            & " impresa = 'impresa' , anulacion_afip='anulacion_afip'," _
            & " tipo_borrar= 'tipo_borrar' , " _
            & " saldada = 'saldada' , id_tipo_discriminado= 'id_tipo_discriminado', " _
            & " observaciones = 'observaciones', texto_adicional = 'texto_adicional'," _
            & " AliPercIB = 'AliPercIB', Opcional27='Opcional27', " _
            & " cambio_a_patron = 'cambio_a_patron' ," _
            & " FormaPago = 'FormaPago' , fecha_entrega = 'fecha_entrega' , " _
            & " propuesta = 'propuesta', fecha_serv_desde = 'fecha_serv_desde', fecha_serv_hasta = 'fecha_serv_hasta' , " _
            & " cancelada = 'cancelada' , CBU = 'CBU' , fecha_pago = 'fecha_pago' , fecha_vto_desde = 'fecha_vto_desde' , fecha_vto_hasta = 'fecha_vto_hasta' , " _
            & " nc_motivo = 'nc_motivo', id_moneda_ajuste='id_moneda_ajuste', tipo_cambio_ajuste='tipo_cambio_ajuste' ," _
            & " total_estatico = 'total_estatico' ,  total_iva_estatico = 'total_iva_estatico' , total_perIB_estatico = 'total_perIB_estatico' ," _
            & " total_neto_estatico = 'total_neto_estatico' , total_exento_estatico = 'total_exento_estatico' ," _
            & " total_iva_discono_estatico = 'total_iva_discono_estatico', tasa_ajuste_mensual = 'tasa_ajuste_mensual'  Where id = 'id'"

        q = Replace$(q, "'id'", conectar.Escape(F.id))
        q = Replace$(q, "'FechaAprobacion'", conectar.Escape(F.FechaAprobacion))
    Else
        esNueva = True

        q = "INSERT INTO sp.AdminFacturas " _
            & " (NroFactura, " _
            & " idCliente, " _
            & " tipoFactura_borrar, " _
            & " idMoneda, " _
            & " FechaEmision, EsCredito," _
            & " idUsuarioEmision, " _
            & " FechaAprobacion, " _
            & " idUsuarioAprobacion, " _
            & " OrdenCompra, " _
            & " origenFacturado, " _
            & " estado, id_concepto_incluir, " _
            & " alicuotaAplicada, " _
            & " discriminada, " _
            & " impresa, " _
            & " tipo_borrar, " _
            & " saldada, " _
            & " observaciones, texto_adicional," _
            & " AliPercIB, " _
            & " cambio_a_patron, " _
            & " FormaPago, " _
            & " fecha_entrega, Opcional27," _
            & " propuesta, id_tipo_discriminado, fecha_serv_desde, fecha_serv_hasta, " _
            & " cancelada, id_moneda_ajuste, tipo_cambio_ajuste, CBU, fecha_pago, fecha_vto_desde, fecha_vto_hasta, " _
    & " nc_motivo, tasa_ajuste_mensual) Values "
        q = q & "('NroFactura', " _
            & " 'idCliente', " _
            & " 'tipoFactura_borrar', " _
            & " 'idMoneda', " _
            & " 'FechaEmision', 'EsCredito'," _
            & " 'idUsuarioEmision', " _
            & " 'FechaAprobacion', " _
            & " 'idUsuarioAprobacion', " _
            & " 'OrdenCompra', " _
            & " 'origenFacturado', " _
            & " 'estado','id_concepto_incluir', " _
            & " 'alicuotaAplicada', " _
            & " 'discriminada', " _
            & " 'impresa', " _
            & " 'tipo_borrar', " _
            & " 'saldada', " _
            & " 'observaciones', 'texto_adicional'," _
            & " 'AliPercIB', " _
            & " 'cambio_a_patron', " _
            & " 'FormaPago', " _
            & " 'fecha_entrega', 'Opcional27'," _
            & " 'propuesta', 'id_tipo_discriminado', 'fecha_serv_desde', 'fecha_serv_hasta', " _
            & " 'cancelada', 'id_moneda_ajuste','tipo_cambio_ajuste', 'CBU', 'fecha_pago', 'fecha_vto_desde','fecha_vto_hasta', " _
            & " 'nc_motivo','tasa_ajuste_mensual' " _
            & ")"

        Set F.usuarioCreador = funciones.GetUserObj()
        q = Replace$(q, "'FechaAprobacion'", "'0000-00-00 00:00:00'")
    End If
    
    '    '''''''''''''''''''''''''''''''''''''' HACK
    '    Dim qTemp As String
    '    Dim TipoFactura As Long: TipoFactura = -1
    '    Dim tmpRS As Recordset
    '    qTemp = "SELECT id FROM AdminConfigFacturas WHERE idIVA = " & F.TipoIVA.idIva & " AND TipoFactura = " & F.tipo.Id
    '
    '    Set tmpRS = conectar.RSFactory(qTemp)
    '    If Not tmpRS.EOF Then TipoFactura = tmpRS!Id
    '
    '    'apunta a AdminConfigFacturas
    '    q = Replace$(q, "'tipoFactura'", TipoFactura)
    '    ''''''''''''''''''''''''''''''''''''''
    
    
    q = Replace$(q, "'NroFactura'", conectar.Escape(F.numero))
    q = Replace$(q, "'idCliente'", conectar.GetEntityId(F.cliente))
    q = Replace$(q, "'fecha_entrega'", conectar.Escape(F.FechaEntrega))
    q = Replace$(q, "'cae'", conectar.Escape(F.CAE))
    q = Replace$(q, "'cae_vto'", conectar.Escape(F.CAEVto))
    q = Replace$(q, "'aprobacion_afip'", conectar.Escape(F.AprobadaAFIP))
    q = Replace$(q, "'Opcional27'", conectar.Escape(F.Opcional27))
    q = Replace$(q, "'anulacion_afip'", conectar.Escape(F.AnulacionAFIP))
    q = Replace$(q, "'id_concepto_incluir'", conectar.Escape(F.ConceptoIncluir))
    q = Replace$(q, "'id_tipo_discriminado'", F.Tipo.id)
    q = Replace$(q, "'tipoFactura_borrar'", F.Tipo.TipoFactura.id)
    q = Replace$(q, "'idMoneda'", conectar.GetEntityId(F.moneda))
    q = Replace$(q, "'FechaEmision'", conectar.Escape(F.FechaEmision))
    q = Replace$(q, "'EsCredito'", conectar.Escape(F.esCredito))
    q = Replace$(q, "'idUsuarioEmision'", conectar.GetEntityId(F.usuarioCreador))
    q = Replace$(q, "'OrdenCompra'", conectar.Escape(F.OrdenCompra))
    q = Replace$(q, "'origenFacturado'", conectar.Escape(F.origenFacturado))
    q = Replace$(q, "'estado'", conectar.Escape(F.estado))
    q = Replace$(q, "'alicuotaAplicada'", conectar.Escape(F.AlicuotaAplicada))
    q = Replace$(q, "'discriminada'", conectar.Escape(F.EstaDiscriminada))
    q = Replace$(q, "'impresa'", conectar.Escape(F.EstaImpresa))
    q = Replace$(q, "'tipo_borrar'", conectar.Escape(F.TipoDocumento))
    q = Replace$(q, "'saldada'", conectar.Escape(F.Saldado))
    q = Replace$(q, "'observaciones'", conectar.Escape(F.observaciones))
    q = Replace$(q, "'texto_adicional'", conectar.Escape(F.TextoAdicional))
    'q = Replace$(q, "'AliPercIB'", conectar.Escape(1 + (f.AlicuotaPercepcionesIIBB / 100)))
    q = Replace$(q, "'AliPercIB'", conectar.Escape(F.AlicuotaPercepcionesIIBB))
    q = Replace$(q, "'cambio_a_patron'", conectar.Escape(F.CambioAPatron))
    q = Replace$(q, "'FormaPago'", conectar.Escape(F.CantDiasPago))
    q = Replace$(q, "'propuesta'", conectar.Escape(F.FechaPropuestaPago))
    q = Replace$(q, "'cancelada'", conectar.Escape(F.Cancelada))
    q = Replace$(q, "'nc_motivo'", conectar.Escape(F.MotivoNC))
    q = Replace$(q, "'idUsuarioAprobacion'", conectar.GetEntityId(F.UsuarioAprobacion))

    q = Replace$(q, "'id_moneda_ajuste'", conectar.Escape(F.IdMonedaAjuste))
    q = Replace$(q, "'tipo_cambio_ajuste'", conectar.Escape(F.TipoCambioAjuste))

    q = Replace$(q, "'total_estatico'", conectar.Escape(F.TotalEstatico.Total))
    q = Replace$(q, "'total_iva_estatico'", conectar.Escape(F.TotalEstatico.TotalIVA))
    q = Replace$(q, "'total_perIB_estatico'", conectar.Escape(F.TotalEstatico.TotalPercepcionesIB))
    q = Replace$(q, "'total_neto_estatico'", conectar.Escape(F.TotalEstatico.TotalNetoGravado))
    q = Replace$(q, "'total_exento_estatico'", conectar.Escape(F.TotalEstatico.TotalExento))
    q = Replace$(q, "'total_iva_discono_estatico'", conectar.Escape(F.TotalEstatico.TotalIVADiscrimandoONo))
    'q = Replace$(q, "'id_ot_anticipo'", conectar.Escape(F.IdOTAnticipo))
    q = Replace$(q, "'tasa_ajuste_mensual'", conectar.Escape(F.TasaAjusteMensual))
    q = Replace$(q, "'CBU'", conectar.Escape(F.CBU))
    q = Replace$(q, "'fecha_pago'", conectar.Escape(F.fechaPago))
    
    'fce_nemer_28052020
    q = Replace$(q, "'fecha_vto_desde'", conectar.Escape(F.FechaVtoDesde))
    q = Replace$(q, "'fecha_vto_hasta'", conectar.Escape(F.FechaVtoHasta))
    
    'fce_nemer_02062020_#113
    'q = Replace$(q, "'fecha_serv_desde'", conectar.Escape(F.FechaServDesde))
    'q = Replace$(q, "'fecha_serv_hasta'", conectar.Escape(F.FechaServHasta))
    
    

    If conectar.execute(q) Then
        If esNueva Then F.id = conectar.UltimoId2
        If F.id = 0 Then GoTo err1

        'me fijo que el nro y el tipo no se repita
        'solo en las q no son electronicas



        If DAOFactura.FindAll("AdminFacturas.id_tipo_discriminado=" & F.Tipo.id & "  and  AdminFacturas.NroFactura = " & F.numero & " And AdminFacturas.Id <> " & F.id).count > 0 Then
            
            If Not F.Tipo.PuntoVenta.EsElectronico And Not F.Tipo.PuntoVenta.CaeManual Then GoTo err1
    

            'el nro de factura y tipo se repite
            'valida tambien q no se repita el tipo de comprobante nc-nd-fc   29/7/13

        Else
            Dim A As New classAdministracion

            If Not F.Tipo.PuntoVenta.EsElectronico Then
                If F.numero >= DAOFactura.proximaFactura(F) Then  'F.TipoDocumento, F.Tipo.TipoFactura.id) Then


                    If Not conectar.execute("UPDATE   AdminConfigFacturasTiposDiscriminado SET   numeracion = " & F.numero & " Where id_tipo_factura = " & F.Tipo.TipoFactura.id & "   AND tipo_documento = " & F.TipoDocumento & " and id_punto_venta = " & F.Tipo.PuntoVenta.id) Then
                        GoTo err1
                    End If
                End If
        
            Else
            If F.AprobadaAFIP Then  'F.TipoDocumento, F.Tipo.TipoFactura.id) Then


                    If Not conectar.execute("UPDATE   AdminConfigFacturasTiposDiscriminado SET   numeracion = " & F.numero & " Where id_tipo_factura = " & F.Tipo.TipoFactura.id & "   AND tipo_documento = " & F.TipoDocumento & " and id_punto_venta = " & F.Tipo.PuntoVenta.id) Then
                        GoTo err1
                    End If
                End If

           End If
        End If

    Else
        GoTo err1
    End If


    Dim Ot As OrdenTrabajo


    If F.id <> 0 Then
        q = "UPDATE pedidos SET id_anticipo_factura = 0 WHERE id_anticipo_factura = " & F.id
        If F.OTsFacturadasAnticipo.count > 0 Then
            q = q & " AND id NOT IN (" & funciones.JoinCollectionValues(F.OTsFacturadasAnticipo, ", ", "Id") & ")"
        End If

        If Not conectar.execute(q) Then GoTo err1
    End If

    For Each Ot In F.OTsFacturadasAnticipo
        If Not conectar.execute("UPDATE pedidos SET id_anticipo_factura = " & F.id & " WHERE id = " & Ot.id) Then GoTo err1
    Next Ot


    If Cascade Then
        Dim det As FacturaDetalle

        'hacer delete de los detalles
        'If Not DAOFacturaDetalles.Delete("idFactura = " & F.id) Then
        '    GoTo err1
        'End If

        DAOFacturaDetalles.Delete "idFactura=" & F.id

        For Each det In F.Detalles
            det.id = 0    'asi hace insert
            det.idFactura = F.id
            If Not DAOFacturaDetalles.Guardar(det) Then
                Err.Raise 8773, "Guardando el historial", "Se produjo un error al guardar el historial"
            End If
        Next det

    End If


    Dim hist As Boolean
    hist = True
    If esNueva Then
        hist = DAOFacturaHistorial.agregar(F, "Factura Creada")
    Else
        hist = DAOFacturaHistorial.agregar(F, "Factura Modificada")
    End If
    If Not hist Then Err.Raise 8773, "Guardando el historial", "Se produjo un error al guardar el historial"


    'conectar.CommitTransaction


    Dim EVENTO As New clsEventoObserver
    Set EVENTO.Elemento = F

    If esNueva Then
        EVENTO.EVENTO = agregar_
    Else
        EVENTO.EVENTO = modificar_
    End If


    Set EVENTO.Originador = Nothing
    EVENTO.Tipo = FacturaCliente_

    Channel.Notificar EVENTO, FacturaCliente_



    Guardar = True
    Exit Function
err1:
    If esNueva Then F.id = 0
    Guardar = False
    Err.Raise Err.Number, Err.Source, Err.Description
    'conectar.RollBackTransaction
End Function




Public Function Anular(Factura As Factura) As Boolean

    
    On Error GoTo err5
    Factura.Detalles = DAOFacturaDetalles.FindByFactura(Factura.id)
    Anular = True



    If Factura.Tipo.PuntoVenta.EsElectronico Then
        MsgBox "Imposible anular un comprobante electronico", vbOKOnly + vbExclamation
        Anular = False
        Exit Function

    End If

    If DAOFactura.EnLiquidacionSubdiarioVentas(Factura.id) Then
        MsgBox "La factura se encuentra  liquidada", vbOKOnly + vbExclamation
        Anular = False
    End If




    Dim estadoAnterior As EstadoFacturaCliente
    Dim deta As FacturaDetalle
    Dim Remito As Remito
    Dim remito_detalle As remitoDetalle

    Dim totAnt As Double
    Dim TotEx As Double
    Dim totIV As Double
    Dim TotIVDisc As Double
    Dim TotNG As Double
    Dim TotIB As Double


    totAnt = Factura.TotalEstatico.Total
    TotEx = Factura.TotalEstatico.TotalExento
    totIV = Factura.TotalEstatico.TotalIVA
    TotIVDisc = Factura.TotalEstatico.TotalIVADiscrimandoONo
    TotNG = Factura.TotalEstatico.TotalNetoGravado
    TotIB = Factura.TotalEstatico.TotalPercepcionesIB

    conectar.BeginTransaction
    estadoAnterior = Factura.estado
    Factura.estado = EstadoFacturaCliente.Anulada

    Factura.TotalEstatico.Total = 0
    Factura.TotalEstatico.TotalExento = 0
    Factura.TotalEstatico.TotalIVA = 0
    Factura.TotalEstatico.TotalIVADiscrimandoONo = 0
    Factura.TotalEstatico.TotalNetoGravado = 0
    Factura.TotalEstatico.TotalPercepcionesIB = 0


    For Each deta In Factura.Detalles
        'luego sacar
        If IsSomething(deta.detalleRemito) Then
            conectar.execute "update detalles_pedidos set cantidad_facturada=cantidad_facturada-" & deta.detalleRemito.Cantidad & "  where id=" & deta.detalleRemito.idDetallePedido
            'vuelvo la cantidad facturada por anulacion de factura
            If Not DAODetalleOrdenTrabajo.SaveCantidad(deta.detalleRemito.idDetallePedido, -deta.detalleRemito.Cantidad, CantidadFacturada_, deta.Bruto, Factura.id, Factura.moneda.id, Factura.CambioAPatron, Factura.TipoCambioAjuste) Then GoTo err5
            'marco el remito como no facturado
            Set Remito = DAORemitoS.FindById(deta.detalleRemito.Remito)
            Remito.EstadoFacturado = RemitoNoFacturado
            If Not DAORemitoS.CambiarEstadoFacturado(Remito.id, RemitoNoFacturado) Then GoTo err5
            Set remito_detalle = DAORemitoSDetalle.FindById(deta.DetalleRemitoId)

            remito_detalle.Facturado = False
            If Not DAORemitoSDetalle.Guardar(remito_detalle) Then GoTo err5
        End If
        If Not DAOFacturaDetalles.Guardar(deta) Then GoTo err5


    Next

    If Factura.TipoDocumento = tipoDocumentoContable.notaCredito Then

        If Factura.Cancelada > 0 Then
            Dim ftmp As Factura
            Set ftmp = DAOFactura.FindById(Factura.Cancelada)
            Dim tmpestado As EstadoFacturaCliente
            tmpestado = ftmp.estado
            ftmp.estado = EstadoFacturaCliente.Aprobada


            If Not conectar.execute("update AdminFacturas set cancelada=" & 0 & " where id=" & ftmp.id) Then GoTo err5

            ' If Not conectar.execute("INSERT INTO AdminFacturas_NC (idFactura, idNC) VALUES (" & idFactura & "," & idnc & ")") Then GoTo er12


            If Not conectar.execute("Delete FROM `sp`.`AdminFacturas_NC` WHERE `idFactura` = " & ftmp.id & " AND idNC=" & Factura.id) Then GoTo err5


            'If Not conectar.execute("update AdminFacturas set cancelada=" & 9 & " where id=" & idnc) Then GoTo er12
            If Not DAOFactura.Guardar(ftmp, False) Then GoTo err5

            Dim evento2 As New clsEventoObserver

            Set evento2.Elemento = Remito
            evento2.EVENTO = agregar_
            Set evento2.Originador = Nothing
            evento2.Tipo = RemitosDetalle_
            Channel.Notificar evento2, RemitosDetalle_




        End If
    End If


    If Factura.OTsFacturadasAnticipo.count > 0 And Factura.origenFacturado = OrigenFacturadoAnticipoOT Then   'si la factura es de Anticipo
        Dim Ot As New OrdenTrabajo

        For Each Ot In Factura.OTsFacturadasAnticipo
            Ot.AnticipoFacturado = False    'True
            Ot.AnticipoFacturadoIdFactura = 0    'Factura.Id
            If Not DAOOrdenTrabajo.Guardar(Ot, False) Then GoTo err5
        Next Ot

    End If
    If Not DAOFactura.Guardar(Factura, False) Then GoTo err5
    DAOEvento.Publish Factura.id, TipoEventoBroadcast.TEB_FacturaAnulada

    conectar.CommitTransaction

    Exit Function
err5:
    conectar.RollBackTransaction
    Factura.estado = estadoAnterior
    'ftmp.estado = tmpestado
    Factura.TotalEstatico.Total = totAnt
    Factura.TotalEstatico.TotalExento = TotEx

    Factura.TotalEstatico.TotalIVA = totIV
    Factura.TotalEstatico.TotalIVADiscrimandoONo = TotIVDisc
    Factura.TotalEstatico.TotalNetoGravado = TotNG
    Factura.TotalEstatico.TotalPercepcionesIB = TotIB
    Anular = False
End Function


Public Function desAnular(Factura As Factura) As Boolean
Exit Function
    
    On Error GoTo err5

    desAnular = True


    Dim estadoAnterior As EstadoFacturaCliente
    Dim deta As FacturaDetalle
    Dim Remito As Remito
    conectar.BeginTransaction
    estadoAnterior = Factura.estado
    Factura.estado = EstadoFacturaCliente.Aprobada

    If Not DAOFacturaHistorial.agregar(Factura, "COMPROBANTE DESANULADO") Then GoTo err5


    If Not DAOFactura.Guardar(Factura, False) Then GoTo err5

    conectar.CommitTransaction

    Exit Function
err5:
    conectar.RollBackTransaction
    Factura.estado = estadoAnterior
    desAnular = False

End Function

'Public Function aprobarAFIP(factura As factura) As Boolean
' On Error GoTo err5
' Set factura = DAOFactura.FindById(factura.id)
'
'
'
'   If (factura.estado = EstadoFacturaCliente.EnProceso) Then
'         Err.Raise 110013, "Factura", "La factura debe aprobarse localmente primero"
'    End If
'
'    If factura.AprobadaAFIP Then
'         Err.Raise 110011, "Factura", "Factura ya aprobada en otra sesión"
'    End If
'
' If Not factura.Tipo.PuntoVenta.EsElectronico Then
'    Err.Raise 110012, "Factura", "La factura a aprobar debe ser electrónica"
' End If
'
'  'si es electronica, debo pedir el cae
'
'
'        'pido el proximo comprobante
'
'
'      conectar.BeginTransaction
'        Dim response As New CAESolicitar
'        'validar la aprobacion o desplegar los errores
'        Set response = ERPHelper.CreateFECaeSolicitarRequest(factura)
'
'        'vlidar  cae y poner nro de factura, cae y fecha de vencimiento y volver a guarar
'
'        If IsSomething(response) Then
'
'            If response.Resultado = "APROBADO" Then
'                factura.numero = response.Comprobante
'                factura.AprobadaAFIP = True
'                factura.FechaEmision = response.getFechaFromString(response.FechaEmision)
'                factura.CAE = response.CAE
'                factura.CAEVto = response.getFechaFromString(response.CAEVencimiento)
'                factura.CAEFechaProceso = response.FechaProceso
'            Else
'                Err.Raise 1000, "Afip", "Comprobante no autorizado " & response.Errores
'
'
'            End If
'
'
'            If Not DAOFactura.Guardar(factura) Then GoTo err5
'
'
'aprobarAFIP = True
'
'        End If
' conectar.CommitTransaction
'  Exit Function
'err5:
'If Err.Number = 110011 Or Err.Number = 110012 Or Err.Number = 110013 Then
'    Err.Raise Err.Number, , Err.Description
'Else
'    conectar.RollBackTransaction
'    factura.AprobadaAFIP = False
'    End If
'
'
'End Function


'esto es para aprobar locamente la factura.
Public Function aprobarV2(Factura As Factura, aprobarLocal As Boolean, enviarAfip As Boolean) As Boolean
    On Error GoTo err5
    
    Set Factura = DAOFactura.FindById(Factura.id)
    
    conectar.BeginTransaction
  
            aprobarV2 = True
    If aprobarLocal Then
            If (Factura.estado = EstadoFacturaCliente.Aprobada) Then
          
                Err.Raise 110011, "Factura", "Factura aprobada en otra sesión"
            End If
            
            Dim idf As Long
        
            If Factura.moneda.Cambio > 1 Then
                If MsgBox("¿Desea asumir el valor para  " & Factura.moneda.NombreCorto & " cómo " & Factura.moneda.Cambio & "?", vbYesNo, "Confirmación") = vbNo Then GoTo err5
            End If
        
            Dim CambioAnterior As Double
            Dim estadoAnterior
            CambioAnterior = Factura.CambioAPatron
            estadoAnterior = Factura.estado
        
        
        
            Factura.Detalles = DAOFacturaDetalles.FindByFactura(Factura.id)    'DAOFactura.FindById(Factura.id, True)
            Dim d As FacturaDetalle
            For Each d In Factura.Detalles
                Set d.Factura = Factura
       
            Next
            Factura.CambioAPatron = Factura.moneda.Cambio
            Factura.FechaAprobacion = Now
            'Factura.FechaEntrega = Date
            Factura.TotalEstatico.Total = Factura.Total
            Factura.TotalEstatico.TotalExento = Factura.TotalExento
            Factura.TotalEstatico.TotalIVA = Factura.TotalIVA
            Factura.TotalEstatico.TotalIVADiscrimandoONo = Factura.TotalIVADiscrimandoONo
            Factura.TotalEstatico.TotalNetoGravado = Factura.TotalNetoGravado
            Factura.TotalEstatico.TotalPercepcionesIB = Factura.totalPercepciones
            Factura.estado = EstadoFacturaCliente.Aprobada
        
            
            Set Factura.UsuarioAprobacion = funciones.GetUserObj
        
            Dim T As Factura
            Set T = Factura
            If Not DAOFactura.Guardar(Factura) Then GoTo err5
            idf = Factura.id
            ' Set Factura = DAOFactura.FindById(i)
            If Not DAOFacturaHistorial.agregar(Factura, "FACTURA APROBADA!") Then GoTo err5
            Dim col As New Collection
            Dim deta As FacturaDetalle
            Dim q As String
            Set Factura = T
            For Each deta In Factura.Detalles
        
        
                If IsSomething(deta.detalleRemito) Then    'si tiene detalleremito es porq se facturo un remito, sino se facturo one concept
        
                    q = "INSERT INTO AdminFacturasDetalleAplicacionRemitos (idFacturaDetalle, idRemitoDetalle, cantidadAplicada) VALUES (" & deta.id & ", " & deta.detalleRemito.id & "  ,  " & deta.detalleRemito.Cantidad & ")"
                    If Not conectar.execute(q) Then GoTo err5
        
                    If deta.detalleRemito.Facturado Then
                        Err.Raise 100000, , "Detalle de remito ya facturado!"
                    End If
        
                    If Not deta.detalleRemito.Facturado Then
        
                        deta.detalleRemito.Facturado = True
                        If Not DAORemitoSDetalle.Guardar(deta.detalleRemito) Then Err.Raise 200, "Detalle de remito", "Imposible guardar el detalle de remito"
        
                        If deta.detalleRemito.Origen = OrigenRemitoOt Then
                            'luego quitar
                            conectar.execute "update detalles_pedidos set cantidad_facturada=cantidad_facturada+" & deta.Cantidad & " where id=" & deta.detalleRemito.idDetallePedido
                            DAODetalleOrdenTrabajo.SaveCantidad deta.detalleRemito.idDetallePedido, deta.Cantidad, CantidadFacturada_, deta.Bruto, Factura.id, Factura.moneda.id, Factura.CambioAPatron, Factura.TipoCambioAjuste
        
                        ElseIf deta.detalleRemito.Origen = OrigenRemitooe Then
                            conectar.execute "update detallesPedidosEntregas set cantidad_facturada=cantidad_facturada+" & deta.Cantidad & " where id=" & deta.detalleRemito.idDetallePedido
                        End If
                    End If
        
                    If Factura.EsAnticipo And Factura.DetallesMismaOT Then
                        Dim Ot As OrdenTrabajo
                        Set Ot = DAOOrdenTrabajo.FindById(deta.detalleRemito.idpedido)
                        If Ot.Anticipo = 100 Then DAODetalleOrdenTrabajo.SaveCantidad deta.detalleRemito.idDetallePedido, deta.detalleRemito.DetallePedido.CantidadPedida, CantidadFacturada_, deta.detalleRemito.Valor, Factura.id, Factura.moneda.id, Factura.CambioAPatron, Factura.TipoCambioAjuste
                    End If
        
        
        
        
                    If Not BuscarEnColeccion(col, CStr(deta.detalleRemito.Remito)) Then
                        col.Add deta.detalleRemito.Remito, CStr(deta.detalleRemito.Remito)
                    End If
                    Dim x As Long
                    Dim rto As Long
                    Dim Remito As Remito
                    For x = 1 To col.count
                        rto = col.item(x)
                        Set Remito = DAORemitoS.FindById(rto)
                        Remito.EstadoFacturado = DAORemitoS.AnalizarEstadoFacturado(Remito.id)
        If Remito.estado = RemitoPendiente Then Err.Raise 206, "Remito " & Remito.numero, "El Remito no esta aprobado"
                        If Remito.estado = RemitoAnulado Then Err.Raise 205, "Remito " & Remito.numero, "El Remito fue anulado en otra sesion"
                        If Not DAORemitoS.Guardar(Remito) Then Err.Raise 201, "Remito", "Imposible guardar el remito " & Remito.numero
                    Next
        
                End If
            Next
        
            If Factura.OTsFacturadasAnticipo.count > 0 And Factura.origenFacturado = OrigenFacturadoAnticipoOT Then   'si la factura es de Anticipo
                If Not EnlazarFacturaAnticipoConOT(Factura) Then GoTo err5
            End If
       
    End If
    
    If enviarAfip Then

              If (Factura.estado = EstadoFacturaCliente.EnProceso) Then
                Err.Raise 110013, "Factura", "La factura debe aprobarse localmente primero"
           End If
           
           If Factura.AprobadaAFIP Then
                Err.Raise 110011, "Factura", "Factura ya aprobada en otra sesión"
           End If
        
        If Not Factura.Tipo.PuntoVenta.EsElectronico Then
           Err.Raise 110012, "Factura", "La factura a aprobar debe ser electrónica"
    End If
 
        Dim response As New CAESolicitar
        'validar la aprobacion o desplegar los errores
        Set response = ERPHelper.CreateFECaeSolicitarRequest(Factura)

        'vlidar  cae y poner nro de factura, cae y fecha de vencimiento y volver a guarar

        If IsSomething(response) Then

            If response.Resultado = "APROBADO" Then
                Factura.numero = response.Comprobante
                Factura.AprobadaAFIP = True
                Factura.FechaEmision = response.getFechaFromString(response.FechaEmision)
                Factura.CAE = response.CAE
                Factura.CAEVto = response.getFechaFromString(response.CAEVencimiento)
                Factura.CAEFechaProceso = response.FechaProceso
            Else
                Err.Raise 1000, "Afip", "Comprobante no autorizado " & response.Errores

            End If

            If Not DAOFactura.Guardar(Factura) Then GoTo err5
            
            'actualizo campo observaciones_cancela (si corresponde)
            Dim tmp As Factura
            Set tmp = DAOFactura.FindById(Factura.Cancelada)
                     
            If IsSomething(tmp) And Factura.Cancelada > 0 Then
              Dim msg1 As String
               Dim MSG2 As String

            If Factura.TipoDocumento = tipoDocumentoContable.Factura Then
            
              
                 msg1 = conectar.Escape("CANCELADA POR " & tmp.GetShortDescription(False, True))
                 MSG2 = conectar.Escape("CANCELA A " & Factura.GetShortDescription(False, True))
                     
        Else
                   msg1 = conectar.Escape("CANCELA A " & tmp.GetShortDescription(False, True))
                 MSG2 = conectar.Escape("CANCELADA POR " & Factura.GetShortDescription(False, True))
        End If
        
        
        
                If Not conectar.execute("update AdminFacturas set observaciones_cancela=" & msg1 & " where id=" & Factura.id) Then GoTo err5
                If Not conectar.execute("update AdminFacturas set  observaciones_cancela=" & MSG2 & " where id=" & tmp.id) Then GoTo err5
            End If
    
    End If

   
        If idf > 0 Then
            Set Factura = DAOFactura.FindById(idf)
            DAOEvento.Publish Factura.id, TipoEventoBroadcast.TEB_FacturaAprobada
        End If
    End If
            conectar.CommitTransaction
    Exit Function
err5:

conectar.RollBackTransaction

If Err.Number = 110011 Or Err.Number = 110012 Or Err.Number = 110013 Then
    Err.Raise Err.Number, , Err.Description
Else
    'conectar.RollBackTransaction
    aprobarV2 = False
    Factura.CambioAPatron = CambioAnterior
    Factura.FechaAprobacion = 0
    Factura.estado = estadoAnterior
    Factura.AprobadaAFIP = False
    Set Factura.UsuarioAprobacion = Nothing
    Err.Raise Err.Number, , Err.Description
 End If
End Function

Public Function desaprobar(Factura As Factura) As Boolean



    conectar.BeginTransaction
    Factura.Detalles = DAOFacturaDetalles.FindByFactura(Factura.id)    'DAOFactura.FindById(Factura.id, True)
    Dim d As FacturaDetalle
    For Each d In Factura.Detalles
        Set d.Factura = Factura
    Next

    On Error GoTo err5
    desaprobar = True

    Dim CambioAnterior As Double
    Dim estadoAnterior
    Dim usuAnterior As clsUsuario
    Set usuAnterior = Factura.UsuarioAprobacion
    CambioAnterior = Factura.CambioAPatron
    estadoAnterior = Factura.estado
    Factura.CambioAPatron = Factura.moneda.Cambio
    Factura.FechaAprobacion = Null

    'Factura.FechaEntrega = Date
    Factura.TotalEstatico.Total = Factura.Total
    Factura.TotalEstatico.TotalExento = Factura.TotalExento
    Factura.TotalEstatico.TotalIVA = Factura.TotalIVA
    Factura.TotalEstatico.TotalIVADiscrimandoONo = Factura.TotalIVADiscrimandoONo
    Factura.TotalEstatico.TotalNetoGravado = Factura.TotalNetoGravado
    Factura.TotalEstatico.TotalPercepcionesIB = Factura.totalPercepciones
    Factura.estado = EstadoFacturaCliente.EnProceso
    Set Factura.UsuarioAprobacion = Nothing
    If Not DAOFactura.Guardar(Factura) Then GoTo err5
    If Not DAOFacturaHistorial.agregar(Factura, "FACTURA DESAPROBADA!") Then GoTo err5
    Dim col As New Collection
    Dim deta As FacturaDetalle
    Dim q As String
    For Each deta In Factura.Detalles


        If IsSomething(deta.detalleRemito) Then    'si tiene detalleremito es porq se facturo un remito, sino se facturo one concept

            q = "INSERT INTO AdminFacturasDetalleAplicacionRemitos (idFacturaDetalle, idRemitoDetalle, cantidadAplicada) VALUES (" & deta.id & ", " & deta.detalleRemito.id & "  ,  " & deta.detalleRemito.Cantidad & ")"
            If Not conectar.execute(q) Then GoTo err5

            If deta.detalleRemito.Facturado Then
                'Err.Clear
                'Err.Raise 100000, , "Detalle de remito ya facturado!"
            End If





            If Factura.EsAnticipo And Factura.DetallesMismaOT Then
                Dim Ot As OrdenTrabajo
                Set Ot = DAOOrdenTrabajo.FindById(deta.detalleRemito.idpedido)
                If Ot.Anticipo = 100 Then DAODetalleOrdenTrabajo.SaveCantidad deta.detalleRemito.idDetallePedido, deta.detalleRemito.DetallePedido.CantidadPedida, CantidadFacturada_, deta.detalleRemito.Valor, Factura.id, Factura.moneda.id, Factura.CambioAPatron, Factura.TipoCambioAjuste
            End If




            If Not BuscarEnColeccion(col, CStr(deta.detalleRemito.Remito)) Then
                col.Add deta.detalleRemito.Remito, CStr(deta.detalleRemito.Remito)
            End If
            Dim x As Long
            Dim rto As Long
            Dim Remito As Remito
            For x = 1 To col.count
                rto = col.item(x)
                Set Remito = DAORemitoS.FindById(rto)
                Remito.EstadoFacturado = DAORemitoS.AnalizarEstadoFacturado(Remito.id)
                If Not DAORemitoS.Guardar(Remito) Then GoTo err5
            Next

        End If
    Next

    If Factura.OTsFacturadasAnticipo.count > 0 And Factura.origenFacturado = OrigenFacturadoAnticipoOT Then   'si la factura es de Anticipo
        If Not EnlazarFacturaAnticipoConOT(Factura) Then GoTo err5
    End If

    conectar.CommitTransaction
    DAOEvento.Publish Factura.id, TipoEventoBroadcast.TEB_FacturaAprobada

    Exit Function
err5:
    conectar.RollBackTransaction
    desaprobar = False
    Factura.CambioAPatron = CambioAnterior
    Factura.FechaAprobacion = 0
    Factura.estado = estadoAnterior
    Set Factura.UsuarioAprobacion = Nothing

    'Err.Raise Err.Number, , Err.Description
End Function

Public Function EnlazarFacturaAnticipoConOT(Factura As Factura, Optional implicitTransaction As Boolean = False) As Boolean
    EnlazarFacturaAnticipoConOT = True

    If implicitTransaction Then conectar.BeginTransaction

    Dim Ot As OrdenTrabajo
    Dim sumaOt As Double

Dim Cambio As Double
Cambio = Factura.CambioAPatron

    For Each Ot In Factura.OTsFacturadasAnticipo
        Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id)
        Ot.AnticipoFacturado = True
        Ot.AnticipoFacturadoIdFactura = Factura.id
        EnlazarFacturaAnticipoConOT = DAOOrdenTrabajo.Guardar(Ot, False)

        If Not EnlazarFacturaAnticipoConOT Then Exit For
        sumaOt = sumaOt + funciones.RedondearDecimales(((MonedaConverter.Convertir(Ot.Total, Ot.moneda.id, Factura.moneda.id)) * Ot.Anticipo) / 100)
    Next Ot

    If EnlazarFacturaAnticipoConOT Then
        EnlazarFacturaAnticipoConOT = (funciones.RedondearDecimales(sumaOt) = funciones.RedondearDecimales(Factura.TotalNetoGravado - DAOFactura.MontoTotalAplicadoNCFC(Factura.id, True)))
    End If

    If implicitTransaction Then
        If EnlazarFacturaAnticipoConOT Then

            If Factura.id <> 0 Then
                Dim q As String
                q = "UPDATE pedidos SET id_anticipo_factura = 0 WHERE id_anticipo_factura = " & Factura.id
                If Factura.OTsFacturadasAnticipo.count > 0 Then
                    q = q & " AND id NOT IN (" & funciones.JoinCollectionValues(Factura.OTsFacturadasAnticipo, ", ", "Id") & ")"
                End If

                EnlazarFacturaAnticipoConOT = conectar.execute(q)
                If Not EnlazarFacturaAnticipoConOT Then
                    conectar.RollBackTransaction
                End If
            End If

            If DAOFactura.Guardar(Factura) Then
                conectar.CommitTransaction
            Else
                conectar.RollBackTransaction
                EnlazarFacturaAnticipoConOT = False
            End If
        Else
            conectar.RollBackTransaction
        End If
    End If
End Function
Public Function EnLiquidacionSubdiarioVentas(factura_id As Long) As Boolean
    Dim q As String
    q = "SELECT 1 FROM liquidacion_subdiario_detalles WHERE id_factura = " & factura_id
    Dim r As Recordset
    Set r = conectar.RSFactory(q)
    EnLiquidacionSubdiarioVentas = Not r.EOF
End Function


Public Function proximaFactura(F As Factura) As Long 'TipoDocumento As tipoDocumentoContable, Optional TipoFactura As Long = -1, Optional idFactura As Long = -1) As Long
  On Error GoTo err1
    
    If F.Tipo.PuntoVenta.EsElectronico And Not F.Tipo.PuntoVenta.CaeManual Then
    
      If IsSomething(F) And IsSomething(F.Tipo) And IsSomething(F.Tipo.PuntoVenta) Then
        
        
          If F.Tipo.PuntoVenta.EsElectronico Then
            Dim ultimoautorizado As Integer
            ultimoautorizado = ERPHelper.ObtenerUltimoActual(F)
    proximaFactura = ultimoautorizado + 1
            
          
          End If
          End If

            
    Else
    
    
    Dim idTipoFacturaDiscriminado  As Long
    idTipoFacturaDiscriminado = F.Tipo.id
    Dim rs As Recordset
    ' 'If TipoFactura = -1 And idFactura = -1 Then Exit Function

   'If TipoFactura > 0 Then
        ' Set rs = conectar.RSFactory("select ft.TipoFactura,ft.numeracion+1 as ult from AdminConfigFacturas f inner join AdminConfigFacturasTipos ft on f.TipoFactura=ft.id where f.TipoFactura=" & TipoFactura)


        '   Set rs = conectar.RSFactory("SELECT DISTINCT  ftd.tipo_documento,   ftd.numeracion +1 AS ult From " _
            & "  AdminConfigFacturas f  INNER JOIN AdminConfigFacturasTipos ft  ON f.TipoFactura = ft.id  " _
            & "  LEFT JOIN AdminConfigFacturasTiposDiscriminado ftd " _
            & " ON ftd.id_tipo_factura=ft.id WHERE ftd.id_tipo_factura = " & TipoFactura & " AND ftd.tipo_documento=" & TipoDocumento)

        Dim q As String
        q = "SELECT DISTINCT    acftd.tipo_documento,   acftd.numeracion + 1 AS ult From    AdminConfigFacturasTiposDiscriminado acftd " _
            & "  LEFT JOIN AdminConfigFacturasTipos acft   ON acftd.id_tipo_factura= acft.id " _
            & "  LEFT JOIN AdminConfigFacturaPuntoVenta pv ON acftd.id_punto_venta=pv.id " _
            & " Where acftd.id=" & idTipoFacturaDiscriminado
           ' & " Where acftd.id_tipo_factura = " & TipoFactura & "   AND acftd.tipo_documento = " & TipoDocumento


        Set rs = conectar.RSFactory(q)

        'Me.ejecutarConsulta ("select numeracion+1 as ult from AdminConfigFacturas where id=" & tipoFactura)
        If Not rs.EOF And Not rs.BOF Then
            proximaFactura = Format(rs!Ult, "0000")
        End If
        Exit Function

End If

Exit Function
err1:
proximaFactura = -1
End Function



Public Function Imprimir(idFactura As Long) As Boolean
    On Error GoTo err91

    Imprimir = False

    Dim objFac As Factura
    Set objFac = DAOFactura.FindById(idFactura, True, True)


    Dim objDeta As FacturaDetalle
    Printer.CurrentY = 300
    Printer.CurrentX = 6800
    Printer.Font.Size = 12
    Printer.Print objFac.TipoDocumentoDescription    'comp

    Printer.CurrentY = 1150
    Printer.CurrentX = 6800
    Printer.Font.Size = 6
    Printer.Print "Control " & objFac.numero  'rs!nroFactura



    Dim x
    Dim xval
    Dim A
    Dim b
    Dim d
    Dim ss

    Printer.Font.Size = 14
    Printer.Line (8800, 1400)-(10100, 1400)

    Printer.Line (8800, 1900)-(10100, 1900)


    Printer.CurrentY = 1500
    Printer.CurrentX = 8900
    'Printer.Print Format(diaa, "00") & "/" & Format(mess, "00") & "/" & Format(anioo - 2000, "00")
    Printer.Print Format(Day(objFac.FechaEmision), "00") & "/" & Format(Month(objFac.FechaEmision), "00") & "/" & Format(Year(objFac.FechaEmision) - 2000, "00")


    Printer.Font = "arial"
    'posiciono los datos del cliente
    'Cliente = Format(nroCli, "0000") & " - " & Cliente
    Printer.CurrentY = 3700
    Printer.Font.Size = 9
    Printer.Print Tab(4);
    Printer.Print "Señor/es: ";
    Printer.FontBold = True
    Printer.Font.Size = 9
    'Printer.Print truncar(Cliente, 64)
    Printer.Print truncar(Format(objFac.cliente.id, "0000") & " - " & objFac.cliente.razon, 100)
    Printer.Font.Size = 9
    Printer.Print Tab(4);
    Printer.FontBold = False
    Printer.Print "I.V.A.: ";
    Printer.FontBold = True
    'Printer.Print truncar(Ivva, 50);
    Printer.Print truncar(objFac.cliente.TipoIVA.detalle, 50);
    Printer.Print Tab(65);
    Printer.FontBold = False
    Printer.Print "C.U.I.T.: ";
    Printer.FontBold = True
    'Printer.Print truncar(Cuit, 50)
    Printer.Print truncar(objFac.cliente.Cuit, 50)
    Printer.Print Tab(4);
    Printer.FontBold = False
    Printer.Print "Domicilio: ";
    Printer.FontBold = True
    Printer.Print truncar(UCase(objFac.cliente.Domicilio), 90);
    Printer.Print Tab(4);
    Printer.FontBold = False
    Printer.Print "Ref: ";
    Printer.FontBold = True
    Printer.Print truncar(objFac.OrdenCompra, 50);


    Printer.Print Tab(65);
    Printer.FontBold = False
    Printer.Print "Localidad: ";
    Printer.FontBold = True
    Printer.Print truncar(UCase(objFac.cliente.localidad.nombre), 30);
    Printer.FontBold = False

    Printer.Print Tab(70);
    Printer.FontBold = False
    Printer.Print "Provincia: ";
    Printer.FontBold = True

    Printer.Print truncar(UCase(objFac.cliente.provincia.nombre), 30)
    Printer.FontBold = False
    Printer.Print Tab(4);
    Printer.Print "Condición: ";
    Printer.FontBold = True
    'Printer.Print truncar(condicion, 40) & " días FF";
    Printer.Print truncar(objFac.CantDiasPago, 40) & " días FF";
    Printer.Print Tab(65);
    Printer.FontBold = False
    '    Printer.Print "C.P. ";
    '    Printer.FontBold = True
    '    'Printer.Print truncar(oc, 50)
    '    Printer.Print truncar(UCase$(objFac.Cliente.CP), 50);
    '    Printer.FontBold = False
    Printer.Print Tab(4);

    Printer.Print "";
    Printer.Print Tab(4);
    'Printer.Font.Size = 8
    'Printer.Print Observaciones
    Printer.Print objFac.observaciones

    Printer.Font.Size = 7
    'detalle y encabezado de detalle de la factura
    Printer.CurrentY = 6700
    Printer.Print Tab(7);
    Printer.Print "Cant";
    Printer.Print Tab(14);
    Printer.Print "Rto      Pos";
    Printer.Print Tab(27);
    Printer.Print "Detalle";
    Printer.Print Tab(110);
    Printer.Print "% Desc";
    'strsql = "select ib,iva,idEntrega,detalle, valor, cantidad from AdminFacturasDetalleNueva where idFactura=" & idFactura

    'Me.ejecutarConsulta strsql
    'tot = 0
    'ali = 1 + (Alicuota / 100)
    Printer.CurrentY = 7000

    'While Not rs.EOF
    For Each objDeta In objFac.Detalles

        Printer.Print Tab(12);
        'ss = funciones.formatearDecimales(rs!Cantidad, 2)
        ss = funciones.FormatearDecimales(objDeta.Cantidad, 2)
        x = Printer.CurrentX
        xval = x - Printer.TextWidth(ss)
        Printer.CurrentX = xval
        Printer.Print ss;




        Printer.Print Tab(14);
        'Printer.Print Trim(remito);
        If IsSomething(objDeta.detalleRemito) Then
            If objDeta.detalleRemito.idDetallePedido > 0 Then
                Set objDeta.detalleRemito = DAORemitoSDetalle.FindById(objDeta.DetalleRemitoId)
                Dim rto As Remito
                Set rto = DAORemitoS.FindById(objDeta.detalleRemito.Remito)
                If IsSomething(objDeta.detalleRemito.DetallePedido) Then
                    Printer.Print Trim(rto.numero & " | " & objDeta.detalleRemito.DetallePedido.item);
                Else
                    Printer.Print Trim(rto.numero & " | " & objDeta.detalleRemito.VerOrigen);

                End If
            Else
                Set rto = DAORemitoS.FindById(objDeta.detalleRemito.Remito)
                Printer.Print Trim(rto.numero);
            End If
        End If


        Dim kk As Long
        Dim jj As Long
        Dim detalle As String
        detalle = UCase(AjustarLineas(objDeta.detalle))

        kk = funciones.InstrCount(detalle, vbNewLine)

        Printer.Print Tab(26);
        If kk = 0 Then
            Printer.Print detalle;
        Else
            Printer.Print Split(detalle, vbNewLine)(0);
        End If

        For jj = 1 To kk
            Printer.Print Tab(26);
            Printer.Print Split(detalle, vbNewLine)(jj);
        Next jj


        Printer.Print Tab(110);
        Printer.Print funciones.FormatearDecimales(objDeta.PorcentajeDescuento);

        'alineo a la izquierda
        Printer.Print Tab(135);
        x = Printer.CurrentX
        xval = x - Printer.TextWidth(funciones.FormatearDecimales(objDeta.SubTotal))

        Printer.CurrentX = xval
        Printer.Print funciones.FormatearDecimales(objDeta.SubTotal);

        'alineo a la izquierda
        Printer.Print Tab(165);
        x = Printer.CurrentX
        xval = x - Printer.TextWidth(funciones.FormatearDecimales(objDeta.Total))
        Printer.CurrentX = xval


        'Printer.Print montoformateado
        Printer.Print funciones.FormatearDecimales(objDeta.Total)

    Next objDeta
    'Next x
    'totalles

    Printer.FontBold = True
    Printer.Font.Size = 11

    Printer.CurrentY = 14900


    'imprimo el primer subtotal alineado a la derecha
    Printer.Print Tab(18);
    x = Printer.CurrentX
    xval = x - Printer.TextWidth(vbNullString)
    Printer.CurrentX = xval
    Printer.Print vbNullString;

    'imprimo el descuento alineado a la derecha
    Printer.Print Tab(35);
    'dtoFormateado = funciones.formatearDecimales(dtoAplicado, 2)
    x = Printer.CurrentX
    'xval = x - Printer.TextWidth(dtoFormateado)
    xval = x - Printer.TextWidth(vbNullString)
    Printer.CurrentX = xval
    'Printer.Print dtoFormateado;
    Printer.Print vbNullString;

    'imprimo el segundo subtotal alineado a la derecha
    Printer.Print Tab(53);

    x = Printer.CurrentX
    xval = x - Printer.TextWidth(funciones.FormatearDecimales(objFac.TotalSubTotal))
    Printer.CurrentX = xval
    Printer.Print funciones.FormatearDecimales(objFac.TotalSubTotal);


    If objFac.EstaDiscriminada Then
        Printer.Print Tab(70);
        x = Printer.CurrentX
        xval = x - Printer.TextWidth(funciones.FormatearDecimales(objFac.TotalIVA))
        Printer.CurrentX = xval
        Printer.Print funciones.FormatearDecimales(objFac.TotalIVA);
    End If



    Dim per As Double
    Printer.Print Tab(84);
    x = Printer.CurrentX
    xval = x - Printer.TextWidth(funciones.FormatearDecimales(objFac.totalPercepciones))
    Printer.CurrentX = xval

    Dim i As Integer
    i = Printer.Font.Size

    Dim cy As Integer
    cy = Printer.CurrentY
    Dim cx As Integer
    cx = Printer.CurrentX
    Printer.CurrentY = Printer.CurrentY - 150
    Printer.Font.Size = 6
    Printer.Print "IIBB  Pcia.Bs.As."
    Printer.CurrentY = Printer.CurrentY + Printer.TextHeight(vbNullString) - 100
    Printer.Font.Size = i
    Printer.CurrentX = cx
    Printer.Print funciones.FormatearDecimales(objFac.totalPercepciones);


    'imprimo el total
    Printer.Print Tab(105);
    Printer.CurrentY = cy
    x = Printer.CurrentX

    xval = x - Printer.TextWidth(funciones.FormatearDecimales(objFac.Total))
    Printer.CurrentX = xval

    Printer.Print funciones.FormatearDecimales(objFac.Total);


    Printer.FontBold = False


    'imprimo el total en letras
    Printer.CurrentY = 12900
    Printer.Print Tab(3);
    Dim c As New classNumericas
    Printer.FontBold = False
    Printer.Font.Size = 8
    Dim queMon As String
    Dim Largo As String

    If objFac.TasaAjusteMensual > 0 Then
        Printer.Print "Esta factura devengará un interés mensual de " & objFac.TasaAjusteMensual & "%"
    End If

    A = "     La cancelación del monto indicado en la presente factura, se efectuará convirtiendo este importe a " & UCase(MonedaConverter.Patron.NombreLargo)    'UCase(nombre_largo_patron)
    b = "     De acuerdo  con la cotización de la moneda extranjera Vigente al día anterior del efectivo pago"
    d = "     Tipo de cambio de referencia en la presente factura :" & MonedaConverter.Patron.NombreCorto & " " & objFac.CambioAPatron & " " & MonedaConverter.Patron.NombreLargo

    If MonedaConverter.Patron.id <> objFac.moneda.id Then
        Printer.Print A
        Printer.Print b
        Printer.Print d
    End If
    Dim mon As clsMoneda
    Set mon = DAOMoneda.GetById(objFac.IdMonedaAjuste)
    If IsSomething(mon) Then
        Dim tot_cbio As Double
        tot_cbio = funciones.RedondearDecimales(objFac.Total / objFac.TipoCambioAjuste)
        A = "      El importe del presente documento equivale a " & mon.NombreCorto & " " & tot_cbio
        b = "      al tipo de cambio de " & DAOMoneda.FindFirstByPatronOrDefault.NombreCorto & " " & objFac.TipoCambioAjuste
        d = "      la presente debera ser abonada al TC BNA tipo comprador del dia anterior al efectivo pago"
        If objFac.moneda.id <> objFac.IdMonedaAjuste And objFac.moneda.id <> 1 Then
            Printer.Print A
            Printer.Print b
            Printer.Print d
        End If
    End If

    Printer.Print "    SON: " & objFac.moneda.NombreLargo & " " & objFac.moneda.NombreCorto  ' & strnum;
    Printer.Print vbTab & c.ValorEnLetras(objFac.Total, objFac.moneda.NombreLargo);


    Printer.EndDoc
    Imprimir = True
    conectar.BeginTransaction

    conectar.execute "update AdminFacturas set impresa=impresa+1 where id=" & idFactura
    conectar.execute "insert into AdminFacturasHistorial (idFactura,Nota,Fecha,idusuario) values (" & idFactura & ",'Factura impresa','" & funciones.datetimeFormateada(Now) & "'," & getUser & " )"
    conectar.CommitTransaction


    Exit Function
err91:

    Imprimir = False
    conectar.RollBackTransaction
    MsgBox Err.Description
End Function

Public Function FindAllByRemitos(remitosNumeros As Collection) As Dictionary

    Dim recordsetConItems As Boolean
    Dim recordsetConItems2 As Boolean
    Dim q As String
    q = "SELECT DISTINCT r.numero, fd.idFactura " _
        & " FROM AdminFacturasDetalleNueva fd" _
        & " INNER JOIN entregas e" _
        & " ON e.id = fd.idEntrega" _
        & " INNER JOIN remitos r" _
        & " ON r.id = e.Remito" _
        & " WHERE r.id IN (" & funciones.JoinCollectionValues(remitosNumeros, ", ") & ")"

    Dim rs As Recordset
    Dim facturas_id As New Collection

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        If Not funciones.BuscarEnColeccion(facturas_id, CStr(rs.Fields("idFactura").value)) Then facturas_id.Add rs.Fields("idFactura").value, CStr(rs.Fields("idFactura").value)
        rs.MoveNext
        recordsetConItems = True
    Wend


    Dim rs2 As Recordset
    q = "SELECT DISTINCT r.numero, fd.idFactura" _
        & " FROM AdminFacturasDetalleAplicacionRemitos ar" _
        & " INNER JOIN entregas e ON e.id = ar.idRemitoDetalle" _
        & " INNER JOIN remitos r ON r.id = e.Remito" _
        & " INNER JOIN AdminFacturasDetalleNueva fd ON fd.id = ar.idFacturaDetalle" _
        & " WHERE r.id IN (" & funciones.JoinCollectionValues(remitosNumeros, ", ") & ")"    'ID O NUMERO
    Set rs2 = conectar.RSFactory(q)
    While Not rs2.EOF
        If Not funciones.BuscarEnColeccion(facturas_id, CStr(rs2.Fields("idFactura").value)) Then facturas_id.Add rs2.Fields("idFactura").value, CStr(rs2.Fields("idFactura").value)
        rs2.MoveNext
        recordsetConItems2 = True
    Wend




    Dim remitosFacturas As New Dictionary

    Dim facturas As Collection
    If facturas_id.count > 0 Then
        Set facturas = DAOFactura.FindAll("AdminFacturas.id IN (" & funciones.JoinCollectionValues(facturas_id, ", ") & ")")
    End If
    Dim Factura As Factura


    If recordsetConItems Then rs.MoveFirst
    While Not rs.EOF
        If funciones.BuscarEnColeccion(facturas, CStr(rs.Fields("idFactura").value)) Then
            If Not remitosFacturas.Exists(CStr(rs.Fields("numero").value)) Then
                remitosFacturas.Add CStr(rs.Fields("numero").value), vbNullString
            End If

            Set Factura = facturas.item(CStr(rs.Fields("idFactura").value))
            remitosFacturas.item(CStr(rs.Fields("numero").value)) = remitosFacturas.item(CStr(rs.Fields("numero").value)) & Factura.GetShortDescription(False, True) & ", "
        End If
        rs.MoveNext
    Wend

    If recordsetConItems2 Then rs2.MoveFirst
    While Not rs2.EOF
        If funciones.BuscarEnColeccion(facturas, CStr(rs2.Fields("idFactura").value)) Then
            If Not remitosFacturas.Exists(CStr(rs2.Fields("numero").value)) Then
                remitosFacturas.Add CStr(rs2.Fields("numero").value), vbNullString
            End If

            Set Factura = facturas.item(CStr(rs2.Fields("idFactura").value))
            If InStr(1, remitosFacturas.item(CStr(rs2.Fields("numero").value)), Factura.GetShortDescription(False, True)) = 0 Then
                remitosFacturas.item(CStr(rs2.Fields("numero").value)) = remitosFacturas.item(CStr(rs2.Fields("numero").value)) & Factura.GetShortDescription(False, True) & ", "
            End If
        End If
        rs2.MoveNext
    Wend


    Set FindAllByRemitos = remitosFacturas
End Function



Public Function aplicarANC(idOrigen As Long, idNCDestino As Long)
  Dim esreto As EstadoRemitoFacturado
    Dim rs As Recordset
    Dim rs_rto As Recordset
    On Error GoTo er12
    aplicarANC = True
    conectar.BeginTransaction


    Dim nc As Factura
    Dim fc As Factura

    Set nc = DAOFactura.FindById(idNCDestino)
    nc.Detalles = DAOFacturaDetalles.FindByFactura(nc.id)
    Set fc = DAOFactura.FindById(idOrigen)
    fc.Detalles = DAOFacturaDetalles.FindByFactura(fc.id)


'    '23-8  Si quiero aplicar una FC a una NC, ambas deben estar aprobadas localmente y no informadas a la afip?
'    If Not fc.Modificable Then
'             Err.Raise 821, "bb", "La FC ya fué enviada a la AFIP, no se puede realizar la asociación"
'    End If
'  If Not nc.Modificable Then
'             Err.Raise 822, "bb", "La NC ya fué enviada a la AFIP, no se puede realizar la asociación"
'    End If

' 23-8 si quiero asociar una FC a una NC, laNC no debe estar informada y deberá controlar que este aprobada
'localmente antes de informar NC a afip
  If Not nc.Modificable Then
         Err.Raise 821, "bb", "La NC no debe estar informada para poder hacer la asociación"
   End If

    Dim ok As Boolean
Dim saldadoTotal As Boolean
saldadoTotal = False
    If MonedaConverter.Convertir(fc.TotalEstatico.Total, fc.moneda.id, nc.moneda.id) <> (nc.TotalEstatico.Total + DAOFactura.MontoTotalAplicadoNCFC(idFactura)) Then
        If MsgBox("La NC a aplicar debe ser del mismo monto que la FC!" & vbNewLine & "¿Desea aplicar de todas maneras?", vbQuestion + vbYesNo) = vbYes Then
     
    saldadoTotal = False
            ok = True
        End If    '
    Else
        ok = True
    End If

    
    
    If ok Then

If saldadoTotal Then
        nc.estado = CanceladaNC
        nc.Saldado = saldadoTotal
        Else
        nc.Saldado = notaCreditoParcial
        nc.estado = CanceladaNCParcial
    End If
    
        
        If Not conectar.execute("update AdminFacturas set cancelada=" & idNCDestino & " where id=" & idFactura) Then GoTo er12

        If Not conectar.execute("INSERT INTO AdminFacturas_NC (idFactura, idNC) VALUES (" & idOrigen & "," & idNCDestino & ")") Then GoTo er12

        If Not conectar.execute("update AdminFacturas set cancelada=" & idOrigen & " where id=" & idNCDestino) Then GoTo er12
        If Not conectar.execute("update AdminFacturas set cancelada=" & idNCDestino & " where id=" & idFactura) Then GoTo er12

            'fix #197
            Dim msg1 As String
            'msg1 = conectar.Escape(fc.observaciones & " / CANCELADA POR " & nc.GetShortDescription(False, True))
            ''If LenB(fc.observaciones) = 0 Then msg1 = conectar.Escape(" / CANCELADA POR " & nc.GetShortDescription(False, True))
            msg1 = conectar.Escape("CANCELADA POR " & nc.GetShortDescription(False, True))
            
            Dim MSG2 As String
            'MSG2 = conectar.Escape(nc.observaciones & " / CANCELA A " & fc.GetShortDescription(False, True))
            'If LenB(fc.observaciones) = 0 Then MSG2 = conectar.Escape(" / CANCELA A " & fc.GetShortDescription(False, True))
            MSG2 = conectar.Escape("CANCELA A " & fc.GetShortDescription(False, True))


       ' If Not conectar.execute("update AdminFacturas set saldada=" & TipoSaldadoFactura.notaCredito & ", estado=" & EstadoFacturaCliente.CanceladaNC & ", observaciones=" & msg1 & " where id=" & fc.id) Then GoTo er12
         '   If Not conectar.execute("update AdminFacturas set saldada=" & TipoSaldadoFactura.notaCredito & ", estado=" & EstadoFacturaCliente.CanceladaNC & ", observaciones=" & MSG2 & " where id=" & nc.id) Then GoTo er12
 If Not conectar.execute("update AdminFacturas set saldada=" & TipoSaldadoFactura.notaCredito & ", estado=" & EstadoFacturaCliente.CanceladaNC & ", observaciones_cancela=" & msg1 & " where id=" & fc.id) Then GoTo er12
            If Not conectar.execute("update AdminFacturas set saldada=" & TipoSaldadoFactura.notaCredito & ", estado=" & EstadoFacturaCliente.CanceladaNC & ", observaciones_cancela=" & MSG2 & " where id=" & nc.id) Then GoTo er12



    Else
        GoTo er12
    End If



    conectar.CommitTransaction

    Exit Function
er12:
    aplicarANC = False
    conectar.RollBackTransaction
End Function

Public Function aplicarNCaFC(idFactura As Long, idnc As Long) As Boolean
    Dim esreto As EstadoRemitoFacturado
    Dim rs As Recordset
    Dim rs_rto As Recordset
    On Error GoTo er12
    aplicarNCaFC = True
    conectar.BeginTransaction


    Dim nc As Factura
    Dim fc As Factura

    Set nc = DAOFactura.FindById(idnc)
    nc.Detalles = DAOFacturaDetalles.FindByFactura(nc.id)
    Set fc = DAOFactura.FindById(idFactura)
    fc.Detalles = DAOFacturaDetalles.FindByFactura(fc.id)


' 23-8 si quiero asociar una FC a una NC, laNC no debe estar informada y deberá controlar que este aprobada
'localmente antes de informar NC a afip

'02.09.20 DNEMER
'Desactivo este mensaje de ERROR porque finalmente las aplicaciones se hacen para los comprobantes electronicos si están informados tambien.
' En el caso de que sean Mi Pymes no van a llegar hasta esta comprobación porque no va a estar disponible la aplicacion en el menu

'  If Not nc.Modificable Then
'         Err.Raise 821, "bb", "La NC no debe estar informada para poder hacer la asociación"
'   End If

' FIN

    Dim ok As Boolean
Dim saldadoTotal As Boolean
saldadoTotal = False
    If MonedaConverter.Convertir(fc.TotalEstatico.Total, fc.moneda.id, nc.moneda.id) <> (nc.TotalEstatico.Total + DAOFactura.MontoTotalAplicadoNCFC(idFactura)) Then
        If MsgBox("La NC a aplicar no es del mismo que la FC!" & vbNewLine & "¿Desea aplicar de todas maneras?", vbQuestion + vbYesNo) = vbYes Then
            '     If Not conectar.execute("INSERT INTO AdminFacturas_NC (idFactura, idNC) VALUES (" & idFactura & "," & idnc & ")") Then GoTo er12
            '     If Not conectar.execute("update AdminFacturas set estado=" & EstadoFacturaCliente.CanceladaNC & ", observaciones=" & conectar.Escape("CANCELADA POR " & nc.NumeroFormateado) & " where id=" & fc.Id) Then GoTo er12
    saldadoTotal = False
            ok = True
        End If    '
    Else
        ok = True
    End If

    If ok Then
        ' If MsgBox("La NC a aplicar debe ser del mismo monto que la FC!" & vbNewLine & "¿Desea aplicar de todas maneras?", vbQuestion + vbYesNo) = vbNo Then
        '           GoTo er12
        '    End If

If saldadoTotal Then
        nc.estado = CanceladaNC
        nc.Saldado = TipoSaldadoFactura.notaCredito
        Else
        nc.Saldado = notaCreditoParcial
        nc.estado = CanceladaNCParcial
    End If
    
    
        
        If Not conectar.execute("update AdminFacturas set cancelada=" & idnc & " where id=" & idFactura) Then GoTo er12

        If Not conectar.execute("INSERT INTO AdminFacturas_NC (idFactura, idNC) VALUES (" & idFactura & "," & idnc & ")") Then GoTo er12

        If Not conectar.execute("update AdminFacturas set cancelada=" & idFactura & " where id=" & idnc) Then GoTo er12
        If Not conectar.execute("update AdminFacturas set cancelada=" & idnc & " where id=" & idFactura) Then GoTo er12


        Dim deta As FacturaDetalle

        For Each deta In fc.Detalles
            If IsSomething(deta.detalleRemito) Then
                Set deta.detalleRemito.DetallePedido = DAODetalleOrdenTrabajo.FindById(deta.detalleRemito.idDetallePedido)
                If IsSomething(deta.detalleRemito.DetallePedido) Then
                    ' chequear si descuenta la cantidad facturada
                    If Not DAODetalleOrdenTrabajo.SaveCantidad(deta.detalleRemito.idDetallePedido, deta.Cantidad, CantidadFacturada_, deta.Cantidad, deta.id, nc.moneda.id, nc.CambioAPatron, nc.TipoCambioAjuste) Then GoTo er12
                End If
            End If
            '   Next deta
            'libero el remito de la FC aplicada

            ' Set rs = conectar.RSFactory("select idEntrega from AdminFacturasDetalleNueva where idFactura=" & fc.Id)
            'comentado 13-11-12 para poder restablecer detalles de remitos cuando se aplican varios a un solo item de factura
            Set rs = conectar.RSFactory("select idRemitoDetalle as idEntrega from AdminFacturasDetalleAplicacionRemitos where idFacturaDetalle=" & deta.id)


            Dim ide As Long
            Dim reto As Long
            While Not rs.EOF
                ide = rs!idEntrega



                If ide > 0 Then
                    Set rs_rto = conectar.RSFactory("select remito from entregas where id=" & ide)

                    If Not rs_rto.EOF And Not rs_rto.BOF Then
                        reto = rs_rto!Remito
                        'si el origen es remito entonces pongo el item como no facturado
                        If Not conectar.execute("update entregas set facturado=0 where id=" & ide) Then GoTo er12

                        esreto = DAORemitoS.AnalizarEstadoFacturado(reto)

                        If Not conectar.execute("update remitos set estadoFacturado=" & esreto & " where id=" & reto) Then GoTo er12

                    Else
                        GoTo er12
                    End If


                End If
                rs.MoveNext
            Wend
            '#197
            Dim msg1 As String
           ' msg1 = conectar.Escape(fc.observaciones & " / CANCELADA POR " & nc.GetShortDescription(False, True))
            ' LenB(fc.observaciones) = 0 Then msg1 = conectar.Escape(" / CANCELADA POR " & nc.GetShortDescription(False, True))
            msg1 = conectar.Escape("APLICADA DE " & nc.GetShortDescription(False, True))
            
            Dim MSG2 As String
            'MSG2 = conectar.Escape(nc.observaciones & " / CANCELA A " & fc.GetShortDescription(False, True))
           ' If LenB(fc.observaciones) = 0 Then MSG2 = conectar.Escape(" / CANCELA A " & fc.GetShortDescription(False, True))
            MSG2 = conectar.Escape("APLICADA A " & fc.GetShortDescription(False, True))

            If Not conectar.execute("update AdminFacturas set saldada=" & nc.Saldado & ", estado=" & nc.estado & ", observaciones=" & msg1 & " where id=" & fc.id) Then GoTo er12
            If Not conectar.execute("update AdminFacturas set saldada=" & nc.Saldado & ", estado=" & nc.estado & ", observaciones=" & MSG2 & " where id=" & nc.id) Then GoTo er12

        Next deta



    Else
        GoTo er12
    End If



    conectar.CommitTransaction

    Exit Function
er12:
    aplicarNCaFC = False
    conectar.RollBackTransaction
   MsgBox Err.Description, vbCritical, "Error"
End Function

Public Function CrearCopiaFiel(F As Factura, Tipo As tipoDocumentoContable) As Factura
    Dim nuevaF As New Factura
    nuevaF.Cancelada = F.Cancelada
    nuevaF.origenFacturado = F.origenFacturado

    Set nuevaF.cliente = F.cliente
    Set nuevaF.moneda = F.moneda
    nuevaF.CambioAPatron = F.CambioAPatron
    Set nuevaF.Tipo = DAOTipoFacturaDiscriminado.FindByTipoDocumentoAndPuntoVentaAndTipoFactura(F.Tipo.TipoFactura.id, Tipo, F.Tipo.PuntoVenta.id, F.TipoIVA.idIVA)


    If F.Tipo.PuntoVenta.EsElectronico Then
        nuevaF.numero = 0
    Else
        nuevaF.numero = CStr(DAOFactura.proximaFactura(nuevaF))   'Tipo, nuevaF.Tipo.TipoFactura.id))
    End If
    Set nuevaF.TipoIVA = F.TipoIVA
    nuevaF.FechaEmision = Date
    nuevaF.observaciones = F.observaciones
    nuevaF.EstaDiscriminada = F.EstaDiscriminada
    nuevaF.OrdenCompra = F.OrdenCompra
    nuevaF.ConceptoIncluir = F.ConceptoIncluir
    nuevaF.TipoCambioAjuste = F.TipoCambioAjuste

    nuevaF.Saldado = NoSaldada
    nuevaF.AlicuotaAplicada = F.AlicuotaAplicada
    nuevaF.AlicuotaPercepcionesIIBB = F.AlicuotaPercepcionesIIBB
    nuevaF.estado = EstadoFacturaCliente.EnProceso
    Set nuevaF.usuarioCreador = funciones.GetUserObj
    nuevaF.CantDiasPago = F.CantDiasPago
    Set nuevaF.UsuarioAprobacion = Nothing
    Set nuevaF.TotalEstatico = F.TotalEstatico
    nuevaF.TotalEstatico.TotalPercepcionesIB = F.TotalEstatico.TotalPercepcionesIB


    Dim deta As FacturaDetalle
    Dim detaNew As FacturaDetalle

    F.Detalles = DAOFacturaDetalles.FindByFactura(F.id)
    nuevaF.Detalles = New Collection
    For Each deta In F.Detalles
        Set detaNew = New FacturaDetalle
        Set detaNew.detalleRemito = Nothing
        detaNew.Bruto = deta.Bruto
        detaNew.detalle = deta.detalle
        detaNew.Cantidad = deta.Cantidad
        detaNew.IvaAplicado = deta.IvaAplicado
        detaNew.IBAplicado = deta.IBAplicado
        detaNew.PorcentajeDescuento = deta.PorcentajeDescuento
        detaNew.Observacion = deta.Observacion
        Set detaNew.Factura = nuevaF

        nuevaF.Detalles.Add detaNew

    Next deta
    nuevaF.TotalEstatico.TotalPercepcionesIB = F.TotalEstatico.TotalPercepcionesIB
    If DAOFactura.Save(nuevaF, True) Then
        Set CrearCopiaFiel = nuevaF
    Else
        Set CrearCopiaFiel = Nothing
    End If

End Function

Public Function CrearFacturaDesdeRemito(rto As Remito) As Boolean
    Dim F As New Factura



End Function


Public Function MontoTotalAplicadoNCFC(idFac As Long, Optional porNetoGravado As Boolean = False) As Double
    Dim facturas As Collection
    Dim tot As Double: tot = 0
    Set facturas = DAOFactura.FindAll("AdminFacturas.id IN (SELECT idNC from AdminFacturas_NC where idFactura = " & idFac & ")", True)
    Dim fac As Factura
    For Each fac In facturas
        If porNetoGravado Then
            tot = tot + fac.TotalEstatico.TotalNetoGravado
        Else
            tot = tot + fac.TotalEstatico.Total
        End If
    Next fac

    MontoTotalAplicadoNCFC = tot
End Function

Public Function VerFacturaElectronicaParaImpresion(idFactura As Long)
    On Error GoTo err1
    'Printer.PaperSize = 9
    Dim F As Factura
    Set F = DAOFactura.FindById(idFactura, True, False)
    Dim seccion As Section
    Dim c As Object
    
    rptFacturaElectronica.LeftMargin = 250
    
    If IsSomething(F) Then
    
    Dim Largo As Double
    
    
        Set seccion = rptFacturaElectronica.Sections("header")
        
        
        Set c = seccion.Controls.item("lblTipoDocumento")
        c.caption = F.Tipo.TipoFactura.Tipo
        
        Set c = seccion.Controls.item("lblFce")
        c.Visible = F.esCredito
        c.caption = F.DescripcionCreditoAdicional
             
        Set c = seccion.Controls.item("lblCbuEmisorFce")
        c.Visible = F.esCredito And F.TipoDocumento = tipoDocumentoContable.Factura
        c.caption = "CBU del Emisor: " & F.CBU

        Set c = seccion.Controls.item("lblCodigoDocumento")
        c.caption = "Código Nº" & Format(F.GetCodigoDocumentoAfip, "00")

        Set c = seccion.Controls.item("LBLDescripcionCodigoDocumento")
        c.caption = F.GetDescripciopnDocumentoAfip

        Set c = seccion.Controls.item("lblFecha")
        c.caption = "Fecha de Emisión: " & Format(F.FechaEmision, "dd/mm/yyyy")
   
        'fce_nemer_2905/2020
        Set c = seccion.Controls.item("lblNumeroDocumento")
        c.caption = "Punto de Venta: " & Format(F.Tipo.PuntoVenta.PuntoVenta, "0000")
   
        Set c = seccion.Controls.item("lblNumeroDocumentoComp")
        c.caption = "Compr. Nro: " & Format(F.numero, "00000000")

        Set seccion = rptFacturaElectronica.Sections("detailsHead")
        
        'fce_nemer_10062020
        
        
        Set c = seccion.Controls.item("lblFechaPagoFce")
        c.Visible = F.TipoDocumento = tipoDocumentoContable.Factura
        
        Set c = seccion.Controls.item("lblFechaPagoFceDato")
                If F.fechaPago = "30/12/1899" Then
                    c.Visible = F.TipoDocumento = tipoDocumentoContable.Factura
                    c.caption = "S/D"
                Else
                    c.Visible = F.TipoDocumento = tipoDocumentoContable.Factura
                    c.caption = Format(F.fechaPago, "dd/mm/yyyy")
                End If
        
        'fce_nemer_09062020
        Set c = seccion.Controls.item("lblDias")
        If F.CantDiasPago = 1 Then
                c.caption = "/ " & F.CantDiasPago & " día"
        Else
                c.caption = "/ " & F.CantDiasPago & " días"
        End If
           
           
        Set c = seccion.Controls.item("lblFechaPagoFceDesde")
        c.Visible = F.esCredito
        
        Set c = seccion.Controls.item("FechaPagoFceDesdeDato")
        c.Visible = F.esCredito
        c.caption = Format(F.FechaVtoDesde, "dd/mm/yyyy")
        
        
        Set c = seccion.Controls.item("lblFechaPagoFceHasta")
        c.Visible = F.esCredito
              
        Set c = seccion.Controls.item("FechaPagoFceHastaDato")
        c.Visible = F.esCredito
        c.caption = Format(F.FechaVtoHasta, "dd/mm/yyyy")
        
                
        Set c = seccion.Controls.item("lblConceptoTexto")
        'fce_nemer_09062020
        c.caption = F.MostrarConcepto
        
        
        
                
        'fce_nemer_10062020_#113
        'Set c = seccion.Controls.item("lblFechaServFceDesde")
        'If F.MostrarConcepto = "Productos" Then
        '    c.caption = ""
        '   Else
        '    c.caption = "Fecha del Servicio Desde:"
        'End If
        
        'fce_nemer_10062020_#113
        'Set c = seccion.Controls.item("FechaServFceDesdeDato")
       'If F.MostrarConcepto = "Productos" Then
        '    c.caption = ""
        '    Else
        '    c.caption = Format(F.FechaServDesde, "dd/mm/yyyy")
        'End If

        
        'fce_nemer_10062020_#113
        'Set c = seccion.Controls.item("lblFechaServFceHasta")
        'If F.MostrarConcepto = "Productos" Then
        '    c.caption = ""
         '   Else
        '    c.caption = "Hasta:"
       ' End If
        
        'fce_nemer_10062020_#113
        'Set c = seccion.Controls.item("FechaServFceHastaDato")
       ' If F.MostrarConcepto = "Productos" Then
        '    c.caption = ""
        '    Else
        '    c.caption = Format(F.FechaServHasta, "dd/mm/yyyy")
        'End If
        
        
    
        seccion.Controls.item("lblCliente").caption = Format(F.cliente.id, "0000") & " - " & F.cliente.razon
        seccion.Controls.item("lblCuit").caption = F.cliente.Cuit
        seccion.Controls.item("lblIva").caption = F.cliente.TipoIVA.detalle
        
        'fce_nemer_29052020
        seccion.Controls.item("lblCondicionPagoFCE").caption = F.observaciones
        
        seccion.Controls.item("lblDireccion").caption = F.cliente.Domicilio & ", " & F.cliente.localidad.nombre & ", " & F.cliente.provincia.nombre
        seccion.Controls.item("lblReferencia").caption = F.OrdenCompra


        Set seccion = rptFacturaElectronica.Sections("footer")

        Set c = seccion.Controls.item("lblBarcode")
        c.caption = F.CodigoBarrasAfip
        Set c = seccion.Controls.item("lblBarcodeCode")
        c.caption = F.CodigoBarrasAfip
        
        Set c = seccion.Controls.item("lblTextoAdicional")
        c.caption = F.TextoAdicional

        Set c = seccion.Controls.item("lblCae")
        c.caption = "CAE: " & F.CAE
        Set c = seccion.Controls.item("lblCaeVencimiento")
        
        c.caption = "VTO CAE: " & F.CAEVto
      
Dim tip As String
tip = vbNullString
If F.TasaAjusteMensual > 0 Then
        
        tip = "Esta factura devengará un interés mensual de " & F.TasaAjusteMensual & "%"

End If
seccion.Controls.item("lblIntereses").caption = tip


 If F.TipoCambioAjuste > 0 Then
    Dim mon As clsMoneda
    'FIX #001
    'cambio de tipo de cambio comprador a vendedor
    
    If F.moneda.id = DAOMoneda.FindFirstByPatronOrDefault.id Then
       'si esta factura en moneda patron
    Set mon = DAOMoneda.GetById(F.IdMonedaAjuste)
       tip = "***  El total de la presente factura, equvale a " & mon.NombreCorto & " " & funciones.RedondearDecimales(F.Total / F.TipoCambioAjuste) & " al tipo de cambio " & mon.NombreCorto & " " & F.TipoCambioAjuste & ".  La presente deberá ser abonada al tipo de cambio BNA tipo vendedor del dia anterior al efectivo pago.  ***"
    Else
        'si esta facturada en otra moneda
    Set mon = DAOMoneda.FindFirstByPatronOrDefault '  DAOMoneda.GetById(F.IdMonedaAjuste)
    'tip = "***  El total de la presente factura, equvale a " & mon.NombreCorto & " " & funciones.RedondearDecimales(F.Total * F.CambioAPatron) & " al tipo de cambio " & mon.NombreCorto & " " & F.CambioAPatron & ".  La presente deberá ser abonada al tipo de cambio BNA tipo comprador del dia anterior al efectivo pago.  ***"
    'FIX 001 - MT
    
    Dim idPatron As Long
    idPatron = DAOMoneda.FindFirstByPatronOrDefault.id
    If F.IdMonedaAjuste <> idPatron And F.moneda.id = idPatron Then
    'factura en pesos, pero  convertida de dolares
        tip = "***  El total de la presente factura, equivale a " & mon.NombreCorto & " " & funciones.RedondearDecimales(F.Total * F.CambioAPatron) & " al tipo de cambio " & mon.NombreCorto & " " & F.CambioAPatron & ".  La presente deberá ser abonada al tipo de cambio BNA tipo vendedor del dia anterior al efectivo pago.  ***"
   Else
   'fix 000
   'factura en dolares
        tip = "***  El total de la presente factura, equivale a " & F.moneda.NombreCorto & " " & funciones.RedondearDecimales(F.Total) & " al tipo de cambio " & mon.NombreCorto & " " & F.CambioAPatron & " ***"
   End If
    
    End If
    
    
    'FIX #001
    'If Not F.moneda.Patron Then
    seccion.Controls.item("lblCambio").caption = tip
   ' MsgBox tip
    
'   End If
    
End If
seccion.Controls.item("lblCambio").Visible = F.IdMonedaAjuste <> idPatron Or F.moneda.id <> idPatron 'F.TipoCambioAjuste > 0 'fix #003 es este comentario And F.IdMonedaAjuste <> DAOMoneda.FindFirstByPatronOrDefault.id



Dim n As New classNumericas

seccion.Controls.item("lblTotalLetras").caption = "Son " & F.moneda.NombreLargo & " " & F.moneda.NombreCorto & " " & LCase(n.ValorEnLetras(F.Total))
seccion.Controls.item("lblSubTotal").caption = funciones.FormatearDecimales(F.TotalSubTotal)
seccion.Controls.item("lblTotalIva").caption = funciones.FormatearDecimales(F.TotalIVA)
seccion.Controls.item("lblTotalTributos").caption = funciones.FormatearDecimales(F.totalPercepciones)
seccion.Controls.item("lblTotal").caption = funciones.FormatearDecimales(F.Total)

QRHelper.generar F
Set seccion.Controls.item("qrcode").Picture = LoadPicture(App.path & "\" & F.id & ".bmp")


'rptFacturaElectronica.ReportWidth = Largo

        Dim r_tmp As New Recordset
        With r_tmp
            .Fields.Append "cantidad", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "remito", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "item", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "descripcion", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "descuento", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "unitario", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "importe", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        End With


    Dim deta As FacturaDetalle
        r_tmp.Open
        For Each deta In F.Detalles
            r_tmp.AddNew
            r_tmp!Cantidad = deta.Cantidad
            
            r_tmp!Remito = vbNullString
            r_tmp!item = vbNullString
            
            If deta.DetalleRemitoId > 0 Then
            Set deta.detalleRemito = DAORemitoSDetalle.FindById(deta.DetalleRemitoId)
        End If


        If deta.CantidadRemitosAplicados > 1 Then
             r_tmp!Remito = "Varios"
        Else
            If IsSomething(deta.detalleRemito) Then
                r_tmp!Remito = deta.detalleRemito.RemitoAlQuePertenece.numero
                
                If IsSomething(deta.detalleRemito.DetallePedido) Then
                    r_tmp!item = deta.detalleRemito.DetallePedido.item
                End If
            End If
       End If
        
            r_tmp!descripcion = deta.detalle
            
            If deta.CantidadRemitosAplicados > 1 Then
            r_tmp!descripcion = r_tmp!descripcion & " (Remitos: " & deta.ListaRemitosAplicados & ")"
            End If
            r_tmp!unitario = funciones.FormatearDecimales(deta.SubTotal)
            r_tmp!Descuento = deta.PorcentajeDescuento
            r_tmp!importe = funciones.FormatearDecimales(deta.Total)
            
            r_tmp.Update

        Next deta

        rptFacturaElectronica.Title = F.GetShortDescription(True, False) & F.Tipo.TipoFactura.Tipo & "-" & Format(F.Tipo.PuntoVenta.PuntoVenta, "000") & "-" & Format(F.numero, "00000000") & " - " & F.cliente.razonFixed
        rptFacturaElectronica.caption = rptFacturaElectronica.Title
        
        Set rptFacturaElectronica.DataSource = r_tmp
        
        rptFacturaElectronica.PrintReport True
        
        conectar.BeginTransaction

        conectar.execute "update AdminFacturas set impresa=impresa+1 where id=" & idFactura
        conectar.execute "insert into AdminFacturasHistorial (idFactura,Nota,Fecha,idusuario) values (" & idFactura & ",'Factura impresa','" & funciones.datetimeFormateada(Now) & "'," & getUser & " )"
        conectar.CommitTransaction


    Else
        MsgBox "Factura no disponible!", vbCritical, "Error"
    End If
    Exit Function
err1:
    MsgBox Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'Public Function GenerarPdf(idFactura As Long) As String
'    GenerarPdf = vbNullString
'    Dim scaleMode As Integer
'    scaleMode = Printer.scaleMode
'    Const COLOR_GRIS = &HC0C0C0
'    Printer.scaleMode = vbPoints
'    On Error GoTo err1
'
'
'    Dim F As Factura
'    Set F = DAOFactura.FindById(idFactura)
'    Dim anchoPagina As Single: anchoPagina = Mm2PT(210)
'    Dim largoPagina As Single: largoPagina = Mm2PT(297)
'    Dim margen As Double: margen = Mm2PT(5)
'    ' Set the PDF title and filename
'    o.PDFTitle = F.GetShortDescription(False, False)
'    o.PDFFileName = funciones.CreateGUID & "_" & F.GetShortDescription(True, True) & F.Tipo.TipoFactura.Tipo & "-" & F.Tipo.PuntoVenta.PuntoVenta & "-" & F.numero & ".pdf"
'
'    ' We must tell the class where the PDF fonts are located
'    o.PDFLoadAfm = App.path & "\"
'    o.PDFSetLayoutMode = LAYOUT_DEFAULT
'    o.PDFFormatPage = FORMAT_A4
'    o.PDFOrientation = ORIENT_PORTRAIT
'    o.PDFSetUnit = UNIT_PT
'
'
'    ' View the PDF file after we create it
'    o.PDFView = True
'
'    ' Begin our PDF document
'    o.PDFBeginDoc
'
'    ' Set the font name, size, and style
'    o.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
'
'    'o.PDFPageWidth = anchoPagina
'    ' o.PDFPageHeight = largoPagina
'
'    '
'    Dim mitadx As Double
'    mitadx = o.PDFGetPageWidth / 2
'
'
'    'dibujo perimetro
'
'
'
'    o.PDFSetDrawColor = vbWhite
'    o.PDFSetLineColor = vbBlack
'    o.PDFSetLineStyle = PDFStyleLgn.pPDF_SOLID
'    o.PDFSetLineWidth = 1.25
'    o.PDFSetDrawMode = DRAW_DRAWBORDER
'    o.PDFDrawRectangle 0 + margen, 0 + margen, o.PDFGetPageWidth - (margen * 2), o.PDFGetPageHeight - (margen * 2)
'    o.PDFSetLineWidth = 0.75
'    o.PDFDrawLine mitadx, margen, mitadx, 150
'    o.PDFDrawLine margen, 150, o.PDFGetPageWidth - (margen), 150
'
'
'    o.PDFSetLineWidth = 1.75
'    o.PDFDrawLine margen, 240, o.PDFGetPageWidth - (margen), 240
'
'
'    'dibujo tipo factura
'
'    o.PDFSetDrawColor = vbBlack
'    o.PDFSetTextColor = vbWhite
'    'o.PDFSetAlignement = ALIGN_CENTER
'    o.PDFSetBorder = BORDER_ALL
'    o.PDFSetFill = True
'    o.PDFDrawRectangle mitadx - 20, margen + 1, 40, 40
'
'
'
'
'    o.PDFSetFont FONT_ARIAL, 30, FONT_BOLD
'    o.PDFSetTextColor = vbWhite
'
'    Dim tam_witdh As Double
'    Dim tam_height As Double
'
'    Dim tip As String
'
'    tip = F.Tipo.TipoFactura.Tipo
'    tam_witdh = Printer.TextWidth(tip)
'    tam_height = Printer.TextHeight(tip)
'    o.PDFTextOut tip, (mitadx - tam_witdh * 2), (margen * 2) + tam_height + 5
'    o.PDFSetFont FONT_ARIAL, 5, FONT_NORMAL
'    tip = "Codigo N " & Format(F.GetCodigoDocumentoAfip, "00")
'    tam_witdh = Printer.TextWidth(tip)
'    tam_height = Printer.TextHeight(tip)
'    o.PDFTextOut tip, (mitadx - 15), (margen * 2) + tam_height + 15
'
'    'coloco logo
'    ' Set the text color
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetFont FONT_ARIAL, 7, FONT_NORMAL
'
'
'
'    tip = "Administracion: Arieta 4720 - Planta: Almafuerte 4670"
'    o.PDFTextOut tip, (mitadx / 2) - 80, 90
'    tip = "B1766DSD Tablada - Pcia. Bs. As. - Argentina"
'    o.PDFTextOut tip, (mitadx / 2) - 70, 100
'    tip = "Tel: (5411) 4651-0051. Fax: 4651-0050"
'    o.PDFTextOut tip, (mitadx / 2) - 60, 110
'    tip = "Email: sp@signoplast.com.ar"
'    o.PDFTextOut tip, (mitadx / 2) - 40, 120
'
'
'    o.PDFImage App.path & "\logo.jpg", 55, 35, 500 / 2.5, 110 / 2.5, "http://www.signoplast.com.ar"
'
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetFont FONT_ARIAL, 20, FONT_BOLD
'    tip = F.GetDescripciopnDocumentoAfip
'    o.PDFTextOut tip, 360, 50
'    o.PDFSetFont FONT_ARIAL, 18, FONT_NORMAL
'    tip = "Nº " & Format(F.Tipo.PuntoVenta.PuntoVenta, "0000") & "-" & Format(F.numero, "00000000")
'    o.PDFTextOut tip, 360, 65
'
'    o.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
'    tip = F.FechaEmision
'    o.PDFTextOut tip, 390, 85
'
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetFont FONT_ARIAL, 7, FONT_NORMAL
'
'
'
'    tip = "CUIT: 30-65760497-2"
'    o.PDFTextOut tip, 360, 100
'    tip = "IIBB: 901-988021-1"
'    o.PDFTextOut tip, 360, 110
'    o.PDFSetFont FONT_ARIAL, 7, FONT_BOLD
'    tip = "CONVENIO MULTILATERAL"
'    o.PDFTextOut tip, 360, 120
'    tip = "INICIO DE ACTIVIDADES: 07-1992"
'    o.PDFTextOut tip, 360, 130
'
'    'fin encabezado
'
'
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'
'    tip = "CUIT:"
'    o.PDFTextOut tip, 35, 185
'
'
'    tip = "Cliente:"
'    o.PDFTextOut tip, 35, 170
'    tip = "Domicilio:"
'    o.PDFTextOut tip, 35, 200
'
'    tip = "Condicion de venta:"
'    o.PDFTextOut tip, 35, 215
'
'    tip = "IVA:"
'    o.PDFTextOut tip, 160, 185
'    tip = "Referencia:"
'    o.PDFTextOut tip, 35, 230
'
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'
'
'    tip = F.Cliente.razon
'    o.PDFTextOut tip, 85, 170
'
'    tip = F.getDescripcionCondicion
'    o.PDFTextOut tip, 157, 215
'
'    tip = F.OrdenCompra
'    o.PDFTextOut tip, 108, 230
'
'    tip = F.Cliente.Domicilio & " - " & F.Cliente.localidad.nombre & " - " & F.Cliente.provincia.nombre
'    o.PDFTextOut tip, 100, 200
'
'    tip = F.Cliente.Cuit
'    o.PDFTextOut tip, 73, 185
'    tip = F.Cliente.TipoIVA.detalle
'    o.PDFTextOut tip, 190, 185
'
'
'    'head detalle
'
'    o.PDFSetDrawColor = COLOR_GRIS
'    o.PDFSetLineWidth = 0.5
'    o.PDFSetTextColor = vbWhite
'    'o.PDFSetAlignement = ALIGN_CENTER
'    o.PDFSetBorder = BORDER_NONE
'    o.PDFSetFill = False
'
'    o.PDFCell " ", margen + 4, 250, o.PDFGetPageWidth - margen - margen - 2, 15
'
'
'
'
'    Dim ccol1 As Double
'    Dim ccol2 As Double
'    Dim ccol3 As Double
'    Dim ccol4 As Double
'    Dim ccol5 As Double
'    Dim ccol6 As Double
'    Dim ccol7 As Double
'    ccol1 = margen + 5
'    ccol2 = ccol1 + 25 + 2
'    ccol3 = ccol2 + 25 + 2
'    ccol4 = ccol3 + 20 + 2
'    ccol5 = ccol4 + 390 + 2
'    ccol6 = ccol5 + 30 + 2
'    ccol7 = ccol6 + 30 + 2
'
'    Dim yhead As Double
'    yhead = 250
'    Dim hcell As Double
'    hcell = 15
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetFont FONT_ARIAL, 8, FONT_BOLD
'
'    o.PDFSetDrawColor = COLOR_GRIS
'    o.PDFSetAlignement = ALIGN_CENTER
'    o.PDFSetBorder = BORDER_NONE
'    o.PDFSetFill = True
'
'    o.PDFCell "Cant", ccol1, yhead, 25, hcell
'    o.PDFCell "Rto", ccol2, yhead, 25, hcell
'    o.PDFCell "Pos", ccol3, yhead, 20, hcell
'    o.PDFCell "Detalle", ccol4, yhead, 390, hcell
'    o.PDFCell "% Desc", ccol5, yhead, 30, hcell
'    o.PDFCell "Precio", ccol6, yhead, 30, hcell
'    o.PDFCell "Importe", ccol7, yhead, 30, hcell
'
'    'detalle
'
'    Dim ydetalle As Double
'    ydetalle = 268
'    Dim x As Integer
'    hcell = 12
'    o.PDFSetFont FONT_ARIAL, 5, FONT_NORMAL
'    Dim deta As FacturaDetalle
'    F.Detalles = DAOFacturaDetalles.FindByFactura(F.id)
'
'    For Each deta In F.Detalles
'
'        Set deta.Factura = F
'
'        o.PDFSetDrawColor = vbRed
'        o.PDFSetTextColor = vbBlack
'        o.PDFSetAlignement = ALIGN_CENTER
'        o.PDFSetBorder = BORDER_NONE
'        o.PDFSetFill = False
'
'        o.PDFCell funciones.FormatearDecimales(deta.Cantidad), ccol1, ydetalle, 25, hcell
'
'        If deta.DetalleRemitoId > 0 Then
'            Set deta.detalleRemito = DAORemitoSDetalle.FindById(deta.DetalleRemitoId)
'        End If
'
'        If IsSomething(deta.detalleRemito) Then
'            o.PDFCell deta.detalleRemito.Remito, ccol2, ydetalle, 25, hcell
'            If IsSomething(deta.detalleRemito.DetallePedido) Then
'                o.PDFCell deta.detalleRemito.DetallePedido.item, ccol3, ydetalle, 20, hcell
'            End If
'        End If
'
'
'        o.PDFSetAlignement = ALIGN_LEFT
'        o.PDFCell deta.detalle, ccol4, ydetalle, 390, hcell
'        o.PDFSetAlignement = ALIGN_CENTER
'        o.PDFCell deta.PorcentajeDescuento, ccol5, ydetalle, 30, hcell
'        o.PDFSetAlignement = ALIGN_RIGHT
'        o.PDFCell funciones.FormatearDecimales(deta.SubTotal), ccol6, ydetalle, 30, hcell
'        o.PDFCell funciones.FormatearDecimales(deta.Total), ccol7, ydetalle, 30, hcell
'        o.PDFSetLineWidth = 0.2
'        o.PDFDrawLine margen, ydetalle + hcell + 2, o.PDFGetPageWidth - (margen), ydetalle + hcell + 2
'        ydetalle = ydetalle + hcell + 4
'
'    Next
'
'
'
'
'    'footer
'    Dim footery As Double: footery = 760
'
'    o.PDFSetLineColor = COLOR_GRIS
'    o.PDFSetLineWidth = 1.75
'    o.PDFDrawLine margen, footery, o.PDFGetPageWidth - margen, footery
'    o.PDFSetLineColor = vbBlack
'    o.PDFDrawLine margen, footery + 36, o.PDFGetPageWidth - margen, footery + 36
'
'
'
'
'    o.PDFSetDrawColor = COLOR_GRIS
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetAlignement = ALIGN_CENTER
'    o.PDFSetBorder = BORDER_NONE
'    o.PDFSetFill = True
'
'    Dim cellw As Double: cellw = 90
'    Dim cells As Double: cells = 4
'    Dim pos_sub1 As Double: pos_sub1 = margen + 7
'    Dim pos_dto As Double: pos_dto = pos_sub1 + pos_dto + cellw + cells
'    Dim pos_sub2 As Double: pos_sub2 = pos_dto + pos_sub2 + cellw + cells
'    Dim pos_iva As Double: pos_iva = pos_sub2 + pos_iva + cellw + cells
'    Dim pos_perc As Double: pos_perc = pos_iva + pos_perc + cellw + cells
'    Dim pos_total As Double: pos_total = pos_perc + pos_total + cellw + cells
'
'    Dim posxx As Double: posxx = footery + 8
'
'    o.PDFSetFont FONT_ARIAL, 7, FONT_NORMAL
'    ' o'.PDFCell funciones.FormatearDecimales(0), pos_sub1, posxx, cellw, 25
'    '  o.PDFCell funciones.FormatearDecimales(0), pos_dto, posxx, cellw, 25
'    o.PDFCell funciones.FormatearDecimales(F.TotalSubTotal), pos_sub2, posxx, cellw, 25
'
'    Dim totIva As Double
'    If F.EstaDiscriminada Then
'        totIva = F.TotalIVA
'    Else
'        totIva = 0
'    End If
'
'    o.PDFCell funciones.FormatearDecimales(totIva), pos_iva, posxx, cellw, 25
'    o.PDFCell funciones.FormatearDecimales(F.totalPercepciones), pos_perc, posxx, cellw, 25
'    o.PDFCell funciones.FormatearDecimales(F.Total), pos_total, posxx, cellw, 25
'
'
'
'    o.PDFSetDrawColor = COLOR_GRIS
'    o.PDFSetTextColor = vbBlack
'    o.PDFSetAlignement = ALIGN_CENTER
'    o.PDFSetBorder = BORDER_NONE
'    o.PDFSetFill = False
'    o.PDFSetFont FONT_ARIAL, 5, FONT_NORMAL
'
'    posxx = footery - 8
'    ' o.PDFCell "Subtotal", pos_sub1, posxx, cellw, 25
'    ' o.PDFCell "Descuento", pos_dto, posxx, cellw, 25
'    o.PDFCell "Subtotal", pos_sub2, posxx, cellw, 25
'    o.PDFCell "IVA", pos_iva, posxx, cellw, 25
'    o.PDFCell "Percepciones", pos_perc, posxx, cellw, 25
'    o.PDFCell "Total", pos_total, posxx, cellw, 25
'
'    o.PDFTextOut "IIBB Pcia. Bs. As.", pos_perc + 20, posxx + 21
'
'    o.PDFSetFont FONT_ARIAL, 7, FONT_NORMAL
'    If F.TasaAjusteMensual > 0 Then
'        tip = "Esta factura devengará un interés mensual de " & F.TasaAjusteMensual & "%"
'        o.PDFTextOut tip, margen + 10, footery + 10
'
'    End If
'
'    Dim c As New classNumericas
'
'
'
'    tip = "SON: " & F.Moneda.NombreLargo & " " & F.Moneda.NombreCorto & " " & c.ValorEnLetras(F.Total, F.Moneda.NombreLargo)
'    o.PDFTextOut tip, margen + 10, footery - 5
'
'
'    o.PDFSetFont FONT_ARIAL, 7, FONT_ITALIC
'    o.PDFSetAlignement = ALIGN_RIGHT
'    If LenB(F.CAE) > 0 Then
'
'        o.PDFCell "CAE.: " & F.CAE, 420, footery + 40, 160, 15
'
'    End If
'
'    If LenB(F.CAEVto) > 0 Then
'        o.PDFCell "Vencimiento CAE.: " & F.CAEVto, 420, footery + 50, 160, 15
'    End If
'    o.PDFSetFont FONT_3OF9, 30, FONT_NORMAL
'
'    o.PDFTextOut "aca va el codigo de barras ", 120, footery + 50
'    ' End our PDF document (this will save it to the filename)
'    o.PDFEndDoc
'    GenerarPdf = o.PDFGetFileName
'    conectar.execute "update AdminFacturas set impresa=impresa+1 where id=" & F.id
'    conectar.execute "insert into AdminFacturasHistorial (idFactura,Nota,Fecha,idusuario) values (" & F.id & ",'Factura electronica generada en un PDF','" & funciones.datetimeFormateada(Now) & "'," & getUser & " )"
'    Printer.scaleMode = scaleMode
'
'    Exit Function
'err1:
'    Printer.scaleMode = scaleMode
'    GenerarPdf = vbNullString
'End Function

Private Function Mm2PT(valueInMm As Double) As Double
    Mm2PT = valueInMm * 2.835016835017
    '1 pt = (INCHES * 72)
End Function


Private Function Mm2Twips(valueInMm As Double) As Double
    Mm2Twips = valueInMm * 56.692913386
    '1 pt = (INCHES * 72)
End Function


