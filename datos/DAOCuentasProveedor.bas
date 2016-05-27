Attribute VB_Name = "DAOCuentasProveedor"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset


Public Function GetByProvedor(id_proveedor As Long) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim rs As Recordset
    Dim cta As clsCuentaContable
    Set rs = conectar.RSFactory("select * from AdminComprasCuentasProveedores where id_proveedor=" & id_proveedor)

    While Not rs.EOF
        Set cta = DAOCuentaContable.GetById(rs!id_cuenta)

        col.Add cta
        rs.MoveNext
    Wend

    Set GetByProvedor = col
    Exit Function
err1:
    Set GetByProvedor = Nothing

End Function





Public Function Save(Proveedor As clsProveedor) As Boolean
    On Error GoTo err1
    Save = True
    Set cn = conectar.obternerConexion
    cn.BeginTrans
    cn.execute "delete from AdminComprasCuentasProveedores where id_proveedor=" & Proveedor.id

    For P = 1 To Proveedor.cuentasContables.count

        cn.execute "insert into AdminComprasCuentasProveedores   (id_proveedor, id_cuenta)   values  (" & Proveedor.id & "," & Proveedor.cuentasContables(P).id & ")"
    Next P


    cn.CommitTrans
    Exit Function
err1:
    Save = False
    cn.RollbackTrans
End Function





