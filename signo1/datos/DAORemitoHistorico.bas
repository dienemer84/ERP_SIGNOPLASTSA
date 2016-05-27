Attribute VB_Name = "DAORemitoHistorico"
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function getAllByIdRemito(id As Long) As Collection
    On Error Resume Next
    Dim col As New Collection
    Dim a As clsHistorial
    Set rs = conectar.RSFactory("select * from remito_historico where id_remito=" & id)

    While Not rs.EOF
        Set a = New clsHistorial
        a.FEcha = rs!FEcha
        a.mensaje = rs!Nota
        a.usuario = DAOUsuarios.GetById(rs!usuario)
        col.Add a

        rs.MoveNext
    Wend
    Set a = Nothing
    Set getAllByIdRemito = col
End Function

Public Function agregar(rto As Remito, mensaje As String) As Boolean
    On Error GoTo err1
    Set cn = conectar.obternerConexion
    Dim usuario As clsUsuario
    Set usuario = DAOUsuarios.GetById(funciones.getUser)
    conectar.execute "insert into remito_historico (id_remito,fecha,nota,usuario) values (" & rto.id & "," & Escape(Now) & ",'" & UCase(mensaje) & "'," & usuario.id & " )"
    agregar = True

    Exit Function
err1:
    agregar = False
End Function
