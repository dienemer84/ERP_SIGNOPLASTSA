Attribute VB_Name = "DAOPresupuestoHistorial"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function getAllByPresu(id_presu As Long, Optional ShowIt As Boolean = False) As Collection
    Dim campos As New Dictionary
    Dim col As New Collection
    Dim a As clsHistorial
    Set rs = conectar.RSFactory("SELECT h.*, u.* FROM historico_presupuesto h LEFT JOIN usuarios u ON h.idUsuario = u.id WHERE idPresupuesto=" & id_presu)

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
    Set getAllByIdPresu = col

    If ShowIt Then
        frmHistoriales.lista = col
        frmHistoriales.Show
    End If

End Function

Public Function agregar(presu As clsPresupuesto, mensaje As String)
    Set cn = conectar.obternerConexion
    fech = funciones.datetimeFormateada(Now)
    Dim usuario As clsUsuario
    Set usuario = funciones.GetUserObj
    conectar.execute "insert into historico_presupuesto (idPresupuesto,fecha,nota,idUsuario) values (" & presu.id & ",'" & fech & "','" & UCase(mensaje) & "'," & usuario.id & " )"
End Function


