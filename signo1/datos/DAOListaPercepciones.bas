Attribute VB_Name = "DAOPercepcionesAplicadas"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection


Public Function listByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim A As clsPercepcionesAplicadas

    Set rs = conectar.RSFactory("select * from AdminComprasFacturasProveedoresPercepciones where id_factura_proveedor=" & id_factura)

    While Not rs.EOF


        Set A = New clsPercepcionesAplicadas
        A.Monto = rs!Valor
        A.Percepcion = DAOPercepciones.GetById(rs!id_percepcion)
        col.Add A

        rs.MoveNext
    Wend

    Set listByIdFactura = col
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, tablaPercepcion As String) As clsPercepcionesAplicadas

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim P As clsPercepcionesAplicadas

    If Id > 0 Then
        Set P = New clsPercepcionesAplicadas
        P.Id = Id
        P.Monto = GetValue(rs, indice, tabla, "valor")
        P.Percepcion = DAOPercepciones.Map(rs, indice, tablaPercepcion)
    End If
    Set Map = P
End Function

Public Function Save(fc As clsFacturaProveedor) As Boolean
    Set cn = conectar.obternerConexion
    Save = True
    On Error GoTo er1:
    cn.execute "delete from AdminComprasFacturasProveedoresPercepciones where id_factura_proveedor=" & fc.Id

    For K = 1 To fc.percepciones.count
        cn.execute "insert into AdminComprasFacturasProveedoresPercepciones (id_percepcion, valor, id_factura_proveedor) values (" & fc.percepciones(K).Percepcion.Id & "," & fc.percepciones(K).Monto & "," & fc.Id & ")"
    Next K
    Exit Function
er1:
    Save = False
End Function
