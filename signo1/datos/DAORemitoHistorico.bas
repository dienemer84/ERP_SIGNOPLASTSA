Attribute VB_Name = "DAORemitoHistorico"
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function getAllByIdRemito(Id As Long) As Collection
    On Error Resume Next
    Dim col As New Collection
    Dim A As clsHistorial
    Set rs = conectar.RSFactory("select * from remito_historico where id_remito=" & Id)

    While Not rs.EOF
        Set A = New clsHistorial
        A.FEcha = rs!FEcha
        A.mensaje = rs!Nota
        A.usuario = DAOUsuarios.GetById(rs!usuario)
        col.Add A

        rs.MoveNext
    Wend
    Set A = Nothing
    Set getAllByIdRemito = col
End Function

Public Function agregar(rto As Remito, mensaje As String) As Boolean
    On Error GoTo err1
    Set cn = conectar.obternerConexion
    Dim usuario As clsUsuario
    Set usuario = DAOUsuarios.GetById(funciones.getUser)
    conectar.execute "insert into remito_historico (id_remito,fecha,nota,usuario) values (" & rto.Id & "," & Escape(Now) & ",'" & UCase(mensaje) & "'," & usuario.Id & " )"
    agregar = True

    Exit Function
err1:
    agregar = False
End Function
