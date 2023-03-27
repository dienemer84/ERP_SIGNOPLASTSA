Attribute VB_Name = "DAOMaterialHistorico"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Function getAllByMaterial(id_material As Long) As Collection
    Dim col As New Collection
    Dim A As clsMaterialHistorico

    Set rs = conectar.RSFactory("select * from historico where id_material=" & id_material)
    While Not rs.EOF
        Set A = New clsMaterialHistorico
        A.FEcha = rs!FEcha_actualizacion
        A.Valor = rs!Valor
        A.moneda = DAOMoneda.GetById(rs!id_moneda)
        col.Add A
        rs.MoveNext
    Wend
    Set A = Nothing
    Set getAllByMaterial = col
End Function




Public Function crear(Material As clsMaterial) As Boolean
    On Error GoTo er1
    Set cn = conectar.obternerConexion
    crear = True

    cn.execute "insert into historico (id_material,valor,fecha_actualizacion,id_moneda) VALUES (" & Material.Id & "," & Material.Valor & " ,'" & Format(Material.FechaValor, "yyyy-mm-dd") & "'," & Material.moneda.Id & ")"

    Exit Function
er1:
    crear = False
End Function
