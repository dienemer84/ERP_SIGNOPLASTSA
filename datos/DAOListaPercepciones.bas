Attribute VB_Name = "DAOPercepcionesAplicadas"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection


Public Function listByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim a As clsPercepcionesAplicadas

    Set rs = conectar.RSFactory("select * from AdminComprasFacturasProveedoresPercepciones where id_factura_proveedor=" & id_factura)

    While Not rs.EOF


        Set a = New clsPercepcionesAplicadas
        a.Monto = rs!Valor
        a.Percepcion = DAOPercepciones.GetById(rs!id_percepcion)
        col.Add a

        rs.MoveNext
    Wend

    Set listByIdFactura = col
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, tablaPercepcion As String) As clsPercepcionesAplicadas

    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim P As clsPercepcionesAplicadas

    If id > 0 Then
        Set P = New clsPercepcionesAplicadas
        P.id = id
        P.Monto = GetValue(rs, indice, tabla, "valor")
        P.Percepcion = DAOPercepciones.Map(rs, indice, tablaPercepcion)
    End If
    Set Map = P
End Function

Public Function Save(fc As clsFacturaProveedor) As Boolean
    Set cn = conectar.obternerConexion
    Save = True
    On Error GoTo er1:
    cn.execute "delete from AdminComprasFacturasProveedoresPercepciones where id_factura_proveedor=" & fc.id

    For K = 1 To fc.percepciones.count
        cn.execute "insert into AdminComprasFacturasProveedoresPercepciones (id_percepcion, valor, id_factura_proveedor) values (" & fc.percepciones(K).Percepcion.id & "," & fc.percepciones(K).Monto & "," & fc.id & ")"
    Next K
    Exit Function
er1:
    Save = False
End Function
