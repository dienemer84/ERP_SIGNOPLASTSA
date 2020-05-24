Attribute VB_Name = "DAOOrdenTrabajoHistorial"

Option Explicit
Dim rs As Recordset
Dim cn As ADODB.Connection
Public Function getAllByOrdenTrabajo(id_ot As Long, Optional ShowIt As Boolean = False) As Collection
    Dim campos As New Dictionary
    Dim col As New Collection
    Dim a As clsHistorial

    Set rs = conectar.RSFactory("SELECT h.*, u.* FROM historial_pedido h LEFT JOIN usuarios u ON h.autor = u.id WHERE idpedido=" & id_ot)

    conectar.BuildFieldsIndex rs, campos

    While Not rs.EOF
        Set a = New clsHistorial
        a.FEcha = GetValue(rs, campos, "h", "fecha")
        a.mensaje = GetValue(rs, campos, "h", "Nota")
        a.usuario = DAOUsuarios.Map(rs, campos, "u")
        col.Add a

        rs.MoveNext
    Wend
    Set a = Nothing
    Set getAllByOrdenTrabajo = col

    If ShowIt Then
        frmHistoriales.lista = col
        frmHistoriales.Show
    End If
End Function

Public Function agregar(Ot As OrdenTrabajo, mensaje As String)
    Dim fech As Date
    fech = funciones.datetimeFormateada(Now)
    Dim usuario As clsUsuario
    Set usuario = funciones.GetUserObj
    agregar = conectar.execute("insert into historial_pedido (idPedido,fecha,nota,autor) values (" & Ot.id & ",'" & funciones.datetimeFormateada(fech) & "','" & UCase(mensaje) & "'," & usuario.id & " )")
End Function

