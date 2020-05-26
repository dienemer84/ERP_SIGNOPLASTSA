Attribute VB_Name = "DAORequeHistorial"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function getAllByIdReque(id_reque As Long) As Collection
    Dim col As New Collection
    Dim a As clsHistorial
    Set rs = conectar.RSFactory("select * from ComprasRequerimientosHistorial where idReque=" & id_reque)

    While Not rs.EOF
        Set a = New clsHistorial
        a.FEcha = rs!FEcha
        a.mensaje = rs!Nota
        a.usuario = DAOUsuarios.GetById(rs!idUsuario)
        col.Add a

        rs.MoveNext
    Wend
    Set a = Nothing
    Set getAllByIdReque = col
End Function

Public Function agregar(reque As clsRequerimiento, mensaje As String)
    Set cn = conectar.obternerConexion
    fech = funciones.datetimeFormateada(Now)
    Dim usuario As clsUsuario
    Set usuario = DAOUsuarios.GetById(funciones.getUser)
    conectar.execute "insert into ComprasRequerimientosHistorial (idReque,fecha,nota,idUsuario) values (" & reque.id & ",'" & fech & "','" & UCase(mensaje) & "'," & usuario.id & " )"
End Function

