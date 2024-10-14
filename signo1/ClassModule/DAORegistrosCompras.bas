Attribute VB_Name = "DAORegistrosCompras"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset


Public Function FindAllComprobantes(Optional ByRef filter As String = vbNullString, Optional ByRef order As String = vbNullString) As Collection
    
    Dim rs As ADODB.Recordset
    Dim q As String

    q = "TRUNCATE TABLE Z_AlicuotasCompras"
    
    Set rs = conectar.RSFactory(q)
       
    q = "INSERT INTO Z_AlicuotasCompras (id, valor) " _
        & " SELECT " _
        & " cp.id AS id_factura_proveedor," _
        & " ROUND(SUM(" _
        & " CASE " _
        & " WHEN b.id_iva IN (10, 2, 4) THEN ROUND((21.00 * b.valor)/100 + 0.0000000001, 2)" _
        & " WHEN b.id_iva = 5 THEN ROUND((10.50 * b.valor)/100 + 0.0000000001, 2)" _
        & " WHEN b.id_iva = 6 THEN ROUND((27.00 * b.valor)/100 + 0.0000000001, 2)" _
        & " WHEN b.id_iva = 19 THEN ROUND((05.00 * b.valor)/100 + 0.0000000001, 2)" _
        & " ELSE 0 END ), 2) AS SUMA" _
        & " FROM sp.AdminComprasFacturasProveedores cp" _
        & " LEFT JOIN " _
        & " sp.AdminComprasFacturasProveedoresIva b ON cp.id = b.id_factura_proveedor" _
        & " WHERE b.id_factura_proveedor IS NOT NULL" _
        & " GROUP BY cp.id;" _

    Set rs = conectar.RSFactory(q)
    
    q = "TRUNCATE TABLE Z_PercepcionesComprasSinIva"
    
    Set rs = conectar.RSFactory(q)
    
    q = "TRUNCATE TABLE Z_PercepcionesComprasSoloIva"
    
    Set rs = conectar.RSFactory(q)
    
    q = "INSERT INTO Z_PercepcionesComprasSinIva (id, valor_percepcionSINIVA)" _
        & " SELECT pp.id_factura_proveedor," _
        & " ROUND(SUM(DISTINCT pp.valor),2)" _
        & " FROM sp.AdminComprasFacturasProveedoresPercepciones pp" _
        & " Where pp.id_percepcion <> 2" _
        & " GROUP BY pp.id_factura_proveedor"
    
    Set rs = conectar.RSFactory(q)
    
    q = "INSERT INTO Z_PercepcionesComprasSoloIva (id, valor_percepcionSOLOIVA)" _
       & " SELECT pp.id_factura_proveedor," _
       & " ROUND(SUM(pp.valor),2)" _
       & " FROM sp.AdminComprasFacturasProveedoresPercepciones pp" _
       & " Where pp.id_percepcion = 2" _
       & " GROUP BY pp.id_factura_proveedor"
    
    Set rs = conectar.RSFactory(q)


    Dim registros As New Collection

    q = "SELECT" _
      & " fecha, numero_factura, tipo_doc_contable, id_config_factura, monto_neto, redondeo_iva, impuesto_interno," _
      & " pv.cuit, pv.razon," _
      & " COALESCE(SUM(DISTINCT pp.valor), 00,00) AS percepciones_valor," _
      & " (iva.valor) AS iva_valor," _
      & " (po.valor_percepcionSOLOIVA) AS percepcion_soloiva," _
      & " (ps.valor_percepcionSINIVA) AS percepcion_siniva," _
      & " COUNT(b.id_iva) AS cantidadAlicuotas," _
      & " COUNT(b.id) AS cantidadidIVA," _
      & " b.id_iva," _
      & " COUNT(DISTINCT b.id) AS cantidadidIVADistintas, b.valor AS valor_alicuota" _
      & " FROM AdminComprasFacturasProveedores cp" _
      & " LEFT JOIN sp.proveedores pv" _
      & " ON cp.id_proveedor = pv.id" _
      & " LEFT JOIN sp.AdminComprasFacturasProveedoresIva b" _
      & " ON cp.id=b.id_factura_proveedor" _
      & " LEFT JOIN sp.AdminComprasFacturasProveedoresPercepciones pp" _
      & " ON cp.id=pp.id_factura_proveedor" _
      & " LEFT JOIN sp.Z_AlicuotasCompras iva" _
      & " ON cp.id=iva.id" _
      & " LEFT JOIN sp.Z_PercepcionesComprasSoloIva po" _
      & " ON cp.id = po.id" _
      & " LEFT JOIN sp.Z_PercepcionesComprasSinIva ps" _
      & " ON cp.id = ps.id" _
      & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

        q = q & " GROUP BY cp.id ORDER BY LEFT (cp.numero_factura,6) ASC, RIGHT(cp.numero_factura,8) ASC, RIGHT(cp.id_proveedor,20)"


    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAllComprobantes = New Collection

    While Not rs.EOF
        registros.Add MapComprobantes(rs, fieldsIndex, "cp", "pv", "b", "pp", "iva", "po", "ps")
        rs.MoveNext
    Wend

    Set FindAllComprobantes = registros
    
End Function


Public Function FindAllAlicuotas(Optional ByRef filter As String = vbNullString, Optional ByRef order As String = vbNullString) As Collection
    
    Dim rs As ADODB.Recordset
    Dim q As String
    
    q = "TRUNCATE TABLE Z_AlicuotasCompras"
    
    Set rs = conectar.RSFactory(q)
       
    q = "INSERT INTO Z_AlicuotasCompras (id, valor) " _
        & " SELECT " _
        & " cp.id AS id_factura_proveedor," _
        & " ROUND(SUM(" _
        & " CASE " _
        & " WHEN b.id_iva IN (10, 2, 4) THEN ROUND((21.00 * b.valor)/100 + 0.0000000001, 2)" _
        & " WHEN b.id_iva = 5 THEN ROUND((10.50 * b.valor)/100 + 0.0000000001, 2)" _
        & " WHEN b.id_iva = 6 THEN ROUND((27.00 * b.valor)/100 + 0.0000000001, 2)" _
        & " WHEN b.id_iva = 19 THEN ROUND((05.00 * b.valor)/100 + 0.0000000001, 2)" _
        & " ELSE 0 END ), 2) AS SUMA" _
        & " FROM sp.AdminComprasFacturasProveedores cp" _
        & " LEFT JOIN " _
        & " sp.AdminComprasFacturasProveedoresIva b ON cp.id = b.id_factura_proveedor" _
        & " WHERE b.id_factura_proveedor IS NOT NULL" _
        & " GROUP BY cp.id;" _

    Set rs = conectar.RSFactory(q)
    
    q = "TRUNCATE TABLE Z_PercepcionesComprasSinIva"
    
    Set rs = conectar.RSFactory(q)
    
    q = "TRUNCATE TABLE Z_PercepcionesComprasSoloIva"
    
    Set rs = conectar.RSFactory(q)
    
    q = "INSERT INTO Z_PercepcionesComprasSinIva (id, valor_percepcionSINIVA)" _
        & " SELECT pp.id_factura_proveedor," _
        & " ROUND(SUM(DISTINCT pp.valor),2)" _
        & " FROM sp.AdminComprasFacturasProveedoresPercepciones pp" _
        & " Where pp.id_percepcion <> 2" _
        & " GROUP BY pp.id_factura_proveedor"
    
    Set rs = conectar.RSFactory(q)
    
    q = "INSERT INTO Z_PercepcionesComprasSoloIva (id, valor_percepcionSOLOIVA)" _
       & " SELECT pp.id_factura_proveedor," _
       & " ROUND(SUM(pp.valor),2)" _
       & " FROM sp.AdminComprasFacturasProveedoresPercepciones pp" _
       & " Where pp.id_percepcion = 2" _
       & " GROUP BY pp.id_factura_proveedor"
    
    Set rs = conectar.RSFactory(q)
    
    Dim registros As New Collection

    q = "SELECT" _
      & " fecha, numero_factura, tipo_doc_contable, id_config_factura, monto_neto, redondeo_iva, impuesto_interno," _
      & " pv.cuit, pv.razon," _
      & " (iva.valor) AS iva_valor," _
      & " (po.valor_percepcionSOLOIVA) AS percepcion_soloiva," _
      & " (ps.valor_percepcionSINIVA) AS percepcion_siniva," _
      & " b.id_iva," _
      & " b.valor AS valor_alicuota" _
      & " FROM sp.AdminComprasFacturasProveedoresIva b" _
      & " LEFT JOIN sp.AdminComprasFacturasProveedores cp" _
      & " ON cp.id=b.id_factura_proveedor" _
      & " LEFT JOIN sp.proveedores pv" _
      & " ON cp.id_proveedor = pv.id" _
      & " LEFT JOIN sp.Z_AlicuotasCompras iva" _
      & " ON cp.id=iva.id" _
      & " LEFT JOIN sp.Z_PercepcionesComprasSoloIva po" _
      & " ON cp.id = po.id" _
      & " LEFT JOIN sp.Z_PercepcionesComprasSinIva ps" _
      & " ON cp.id = ps.id" _
      & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If
        
        q = q & " AND NOT (cp.tipo_doc_contable=0 AND cp.id_config_factura = 2)" _
                & " AND NOT (cp.tipo_doc_contable=0 AND cp.id_config_factura = 3)" _
                & " AND NOT (cp.tipo_doc_contable=0 AND cp.id_config_factura = 7)" _
                & " AND NOT (cp.tipo_doc_contable=0 AND cp.id_config_factura =10)" _
                & " AND NOT (cp.tipo_doc_contable=1 AND cp.id_config_factura = 3)" _
                & " AND NOT (cp.tipo_doc_contable=1 AND cp.id_config_factura = 7)" _
                & " AND NOT (cp.tipo_doc_contable=1 AND cp.id_config_factura =10)" _
                & " AND NOT (cp.tipo_doc_contable=2 AND cp.id_config_factura =10)"
        
        q = q & "ORDER BY LEFT (cp.numero_factura,6) ASC, RIGHT(cp.numero_factura,8) ASC, RIGHT(cp.id_proveedor,20)"


    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAllAlicuotas = New Collection

    While Not rs.EOF
        registros.Add MapAlicuotas(rs, fieldsIndex, "cp", "pv", "b", "pp", "iva", "po", "ps")
        rs.MoveNext
    Wend

    Set FindAllAlicuotas = registros
    
End Function

Public Function MapComprobantes(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef tablaProveedores As String = vbNullString, _
                    Optional ByRef tablaProveedoresIva As String = vbNullString, _
                    Optional ByRef tablaPercepciones As String = vbNullString, _
                    Optional ByRef tablaIVA As String = vbNullString, _
                    Optional ByRef Z_PercepcionesComprasSoloIva As String = vbNullString, _
                    Optional ByRef Z_PercepcionesComprasSinIva As String = vbNullString) As clsRegistroIVACompras
                    
   
    Dim ric As clsRegistroIVACompras

        Set ric = New clsRegistroIVACompras
        ric.FEcha = GetValue(rs, fieldsIndex, tableNameOrAlias, "fecha")
        ric.numerodecomprobante = GetValue(rs, fieldsIndex, tableNameOrAlias, "numero_factura")
        ric.tipodoccontable = GetValue(rs, fieldsIndex, tableNameOrAlias, "tipo_doc_contable")
        ric.idconfigfactura = GetValue(rs, fieldsIndex, tableNameOrAlias, "id_config_factura")
        ric.Cuit = GetValue(rs, fieldsIndex, tablaProveedores, "cuit")
        ric.denominacionvendedor = GetValue(rs, fieldsIndex, tablaProveedores, "razon")
        ric.montoneto = GetValue(rs, fieldsIndex, tableNameOrAlias, "monto_neto")
        ric.redondeoiva = GetValue(rs, fieldsIndex, tableNameOrAlias, "redondeo_iva")
        ric.impuestosinternos = GetValue(rs, fieldsIndex, tableNameOrAlias, "impuesto_interno")
        ric.ivavalor = GetValue(rs, fieldsIndex, tablaIVA, "iva_valor")
        
        ric.percepcionesSoloIva = GetValue(rs, fieldsIndex, Z_PercepcionesComprasSoloIva, "percepcion_soloiva")
        ric.percepcionessSinIva = GetValue(rs, fieldsIndex, Z_PercepcionesComprasSinIva, "percepcion_siniva")
        
        ric.cantidaddealicuotas = GetValue(rs, fieldsIndex, vbNullString, "cantidadAlicuotas")
        ric.contadoridIVA = GetValue(rs, fieldsIndex, vbNullString, "cantidadidIVA")
        ric.idIVA = GetValue(rs, fieldsIndex, tablaProveedoresIva, "id_iva")
        ric.cantidadidIVADistintas = GetValue(rs, fieldsIndex, vbNullString, "cantidadidIVADistintas")
        
        ric.valorAlicuota = GetValue(rs, fieldsIndex, tablaProveedoresIva, "valor_alicuota")

    Set MapComprobantes = ric

End Function


Public Function MapAlicuotas(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef tablaProveedores As String = vbNullString, _
                    Optional ByRef tablaProveedoresIva As String = vbNullString, _
                    Optional ByRef tablaPercepciones As String = vbNullString, _
                    Optional ByRef tablaIVA As String = vbNullString, _
                    Optional ByRef Z_PercepcionesComprasSoloIva As String = vbNullString, _
                    Optional ByRef Z_PercepcionesComprasSinIva As String = vbNullString) As clsRegistroIVACompras
                    
   
    Dim ric As clsRegistroIVACompras

        Set ric = New clsRegistroIVACompras
        ric.FEcha = GetValue(rs, fieldsIndex, tableNameOrAlias, "fecha")
        ric.numerodecomprobante = GetValue(rs, fieldsIndex, tableNameOrAlias, "numero_factura")
        ric.tipodoccontable = GetValue(rs, fieldsIndex, tableNameOrAlias, "tipo_doc_contable")
        ric.idconfigfactura = GetValue(rs, fieldsIndex, tableNameOrAlias, "id_config_factura")
        ric.Cuit = GetValue(rs, fieldsIndex, tablaProveedores, "cuit")
        ric.denominacionvendedor = GetValue(rs, fieldsIndex, tablaProveedores, "razon")
        ric.montoneto = GetValue(rs, fieldsIndex, tableNameOrAlias, "monto_neto")
        ric.redondeoiva = GetValue(rs, fieldsIndex, tableNameOrAlias, "redondeo_iva")
        ric.impuestosinternos = GetValue(rs, fieldsIndex, tableNameOrAlias, "impuesto_interno")
        ric.ivavalor = GetValue(rs, fieldsIndex, tablaIVA, "iva_valor")
        
        ric.idIVA = GetValue(rs, fieldsIndex, tablaProveedoresIva, "id_iva")
        
        ric.valorAlicuota = GetValue(rs, fieldsIndex, tablaProveedoresIva, "valor_alicuota")

    Set MapAlicuotas = ric

End Function

