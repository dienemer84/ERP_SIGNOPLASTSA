Attribute VB_Name = "DAOFacturaHistorial"
Option Explicit

Dim rs As ADODB.Recordset

Public Function getAllByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim A As clsHistorial
    Set rs = conectar.RSFactory("select * from AdminFacturasHistorial where idFactura=" & id_factura)
    While Not rs.EOF
        Set A = New clsHistorial
        A.FEcha = rs!FEcha
        A.mensaje = rs!Nota
        A.usuario = DAOUsuarios.GetById(rs!idUsuario)
        col.Add A

        rs.MoveNext
    Wend
    Set A = Nothing
    Set getAllByIdFactura = col
End Function

Public Function agregar(Factura As Factura, mensaje As String) As Boolean
    agregar = conectar.execute("insert into AdminFacturasHistorial (idFactura,fecha,nota,idUsuario) values (" & Factura.Id & "," & conectar.Escape(Now) & ",'" & UCase(mensaje) & "'," & conectar.GetEntityId(funciones.GetUserObj) & ")")
End Function

