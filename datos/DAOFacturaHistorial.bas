Attribute VB_Name = "DAOFacturaHistorial"
Option Explicit

Dim rs As ADODB.Recordset

Public Function getAllByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim a As clsHistorial
    Set rs = conectar.RSFactory("select * from AdminFacturasHistorial where idFactura=" & id_factura)
    While Not rs.EOF
        Set a = New clsHistorial
        a.FEcha = rs!FEcha
        a.mensaje = rs!Nota
        a.usuario = DAOUsuarios.GetById(rs!idUsuario)
        col.Add a

        rs.MoveNext
    Wend
    Set a = Nothing
    Set getAllByIdFactura = col
End Function

Public Function agregar(Factura As Factura, mensaje As String) As Boolean
    agregar = conectar.execute("insert into AdminFacturasHistorial (idFactura,fecha,nota,idUsuario) values (" & Factura.id & "," & conectar.Escape(Now) & ",'" & UCase(mensaje) & "'," & conectar.GetEntityId(funciones.GetUserObj) & ")")
End Function

