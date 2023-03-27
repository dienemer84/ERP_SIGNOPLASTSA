Attribute VB_Name = "DaoFacturaProveedorHistorial"
Dim rs As ADODB.Recordset

Public Function getAllByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim A As clsHistorial
    Set rs = conectar.RSFactory("select * from AdminComprasFacturasProveedoresHistorial where id_factura=" & id_factura)
    While Not rs.EOF
        Set A = New clsHistorial
        A.FEcha = rs!FEcha
        A.mensaje = rs!mensaje
        A.usuario = DAOUsuarios.GetById(rs!id_usuario)
        col.Add A

        rs.MoveNext
    Wend
    Set A = Nothing
    Set getAllByIdFactura = col
End Function

Public Function agregar(Factura As clsFacturaProveedor, mensaje As String) As Boolean
    On Error GoTo err1
    Dim usus As New clsUsuario
    Set cn = conectar.obternerConexion
    Set usus = DAOUsuarios.GetById(funciones.getUser)
    fech = funciones.datetimeFormateada(Now)
    If Not conectar.execute("insert into AdminComprasFacturasProveedoresHistorial (id_factura,fecha,mensaje,id_usuario) values (" & Factura.Id & ",'" & fech & "','" & UCase(mensaje) & "'," & usus.Id & ")") Then GoTo err1
    agregar = True
    Exit Function
err1:
    agregar = False
End Function
