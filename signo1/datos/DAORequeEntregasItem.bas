Attribute VB_Name = "DAORequeEntregas"
Dim rs As Recordset
Public Function GetEntregaById(id_detalle_reque As Long, Tipo As tipoEntrega) As Collection
    On Error GoTo er1
    Dim a As clsRequeEntregas
    Dim col As New Collection
    Set rs = conectar.RSFactory("select * from ComprasRequerimientosDetallesEntregas where tipo=" & Tipo & " and id_detalle_material=" & id_detalle_reque)
    While Not rs.EOF
        Set a = New clsRequeEntregas
        a.Cantidad = rs!Cantidad
        a.FEcha = rs!FEcha
        a.id = rs!id
        a.Tipo = rs!Tipo
        col.Add a
        rs.MoveNext
    Wend
    Set GetEntregaById = col
    Exit Function
er1:
    Set GetEntregaById = Nothing
End Function
Public Function saveAll(T As Collection, id_detalle As Long) As Boolean
    On Error GoTo err1
    Dim tmp As clsRequeEntregas
    saveAll = True
    'saveAll = conectar.execute("delete from ComprasRequerimientosDetallesEntregas where id_detalle_material=" & id_detalle)
    For i = 1 To T.count
        Set tmp = T.item(i)
        saveAll = conectar.execute("insert into ComprasRequerimientosDetallesEntregas (id_detalle_material, cantidad, fecha, tipo) values (" & id_detalle & "," & tmp.Cantidad & ",'" & funciones.dateFormateada(tmp.FEcha) & "'," & tmp.Tipo & ")")
    Next i

    Exit Function
err1:
    saveAll = False
End Function




