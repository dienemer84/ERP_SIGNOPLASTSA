Attribute VB_Name = "DAOIvaAplicado"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection


Public Function listByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim b As clsAlicuotaAplicada
    Dim a As clsTipoIVA
    Set rs = conectar.RSFactory("select * from AdminComprasFacturasProveedoresIva where id_factura_proveedor=" & id_factura)
    While Not rs.EOF
        Set b = New clsAlicuotaAplicada
        Set a = New clsTipoIVA
        b.Monto = rs!Valor
        b.Alicuota = DAOAlicuotas.GetById(rs!id_iva)
        col.Add b
        rs.MoveNext
    Wend

    Set listByIdFactura = col
End Function
Public Function Save(fc As clsFacturaProveedor) As Boolean
    Save = True
    Set cn = conectar.obternerConexion
    On Error GoTo er1:
    cn.execute "delete from AdminComprasFacturasProveedoresIva where id_factura_proveedor=" & fc.id


    For K = 1 To fc.IvaAplicado.count
        cn.execute "insert into AdminComprasFacturasProveedoresIva (id_iva, valor, id_factura_proveedor) values (" & fc.IvaAplicado(K).Alicuota.id & "," & fc.IvaAplicado(K).Monto & "," & fc.id & ")"
    Next K

    Exit Function
er1:
    Save = False
End Function



Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, tablaAlicuota As String) As clsAlicuotaAplicada

    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim a As clsAlicuotaAplicada

    If id > 0 Then
        Set a = New clsAlicuotaAplicada
        a.id = id
        a.Alicuota = DAOAlicuotas.Map(rs, indice, tablaAlicuota)
        a.Monto = GetValue(rs, indice, tabla, "valor")
    End If
    Set Map = a
End Function
