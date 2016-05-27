Attribute VB_Name = "DAOCuentasFacturas"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset


Public Function GetByFactura(id_factura As Long) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim rs As Recordset
    Dim cta As clsCuentaFactura



    Set rs = conectar.RSFactory("select * from AdminComprasCuentasFacturas where id_factura=" & id_factura)



    While Not rs.EOF

        Set cta = New clsCuentaFactura
        cta.cuentas = DAOCuentaContable.GetById(rs!id_cuenta)
        cta.Monto = rs!Monto

        col.Add cta
        rs.MoveNext
    Wend

    Set GetByFactura = col
    Exit Function
err1:
    Set GetByFactura = Nothing

End Function
Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, tablaCtaContable As String) As clsCuentaFactura
    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim a As clsCuentaFactura

    If id > 0 Then
        Set a = New clsCuentaFactura
        a.id = id
        a.cuentas = DAOCuentaContable.Map(rs, indice, tablaCtaContable)
        a.Monto = GetValue(rs, indice, tabla, "monto")
    End If
    Set Map = a
End Function
Public Function Save(Factura As clsFacturaProveedor) As Boolean
    On Error GoTo err1
    Save = True
    Set cn = conectar.obternerConexion
    cn.execute "delete from AdminComprasCuentasFacturas where id_factura=" & Factura.id
    Dim ctatmp As clsCuentaFactura
    For P = 1 To Factura.cuentasContables.count
        Set ctatmp = Factura.cuentasContables(P)

        cn.execute "insert into AdminComprasCuentasFacturas   (id_factura, id_cuenta,monto)   values  (" & Factura.id & "," & ctatmp.cuentas.id & "," & ctatmp.Monto & ")"
    Next P
    Exit Function
err1:
    Save = False

End Function

