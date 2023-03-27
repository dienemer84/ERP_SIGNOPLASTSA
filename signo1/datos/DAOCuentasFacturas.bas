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
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim A As clsCuentaFactura

    If Id > 0 Then
        Set A = New clsCuentaFactura
        A.Id = Id
        A.cuentas = DAOCuentaContable.Map(rs, indice, tablaCtaContable)
        A.Monto = GetValue(rs, indice, tabla, "monto")
    End If
    Set Map = A
End Function
Public Function Save(Factura As clsFacturaProveedor) As Boolean
    On Error GoTo err1
    Save = True
    Set cn = conectar.obternerConexion
    cn.execute "delete from AdminComprasCuentasFacturas where id_factura=" & Factura.Id
    Dim ctatmp As clsCuentaFactura
    For P = 1 To Factura.cuentasContables.count
        Set ctatmp = Factura.cuentasContables(P)

        cn.execute "insert into AdminComprasCuentasFacturas   (id_factura, id_cuenta,monto)   values  (" & Factura.Id & "," & ctatmp.cuentas.Id & "," & ctatmp.Monto & ")"
    Next P
    Exit Function
err1:
    Save = False

End Function

