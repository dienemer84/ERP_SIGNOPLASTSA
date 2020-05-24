Attribute VB_Name = "DaoHistorico"
Option Explicit



Public Function Save(tabla As String, mensaje As String, idSource As Long) As Boolean
On Error GoTo err1
Dim q As String
q = "insert into " & tabla & " ( mensaje, usuario, id_source) values ('" & mensaje & "', '" & funciones.GetUserObj.usuario & "','" & idSource & "')"
conectar.execute (q)

Exit Function

err1:
Err.Raise Err.Number
End Function

Public Function GetAll(tabla As String, id As Long) As Collection
Dim rs As Recordset
Dim col As New Collection

Dim h As Historial
Set rs = RSFactory("select * from " & tabla & " where id_source = " & id)


While Not rs.EOF And Not rs.BOF
Set h = New Historial
h.Autor = rs!usuario
h.FEcha = rs!FEcha
h.mensaje = rs!mensaje

col.Add h

rs.MoveNext

Wend

Set GetAll = col
End Function

