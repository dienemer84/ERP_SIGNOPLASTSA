Attribute VB_Name = "DaoFacturaProveedorHistorial"
Dim rs As ADODB.Recordset

Public Function getAllByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim a As clsHistorial
    Set rs = conectar.RSFactory("select * from AdminComprasFacturasProveedoresHistorial where id_factura=" & id_factura)
    While Not rs.EOF
        Set a = New clsHistorial
        a.FEcha = rs!FEcha
        a.mensaje = rs!mensaje
        a.usuario = DAOUsuarios.GetById(rs!id_usuario)
        col.Add a

        rs.MoveNext
    Wend
    Set a = Nothing
    Set getAllByIdFactura = col
End Function

Public Function agregar(Factura As clsFacturaProveedor, mensaje As String) As Boolean
    On Error GoTo err1
    Dim usus As New clsUsuario
    Set cn = conectar.obternerConexion
    Set usus = DAOUsuarios.GetById(funciones.getUser)
    fech = funciones.datetimeFormateada(Now)
    If Not conectar.execute("insert into AdminComprasFacturasProveedoresHistorial (id_factura,fecha,mensaje,id_usuario) values (" & Factura.id & ",'" & fech & "','" & UCase(mensaje) & "'," & usus.id & ")") Then GoTo err1
    agregar = True
    Exit Function
err1:
    agregar = False
End Function
