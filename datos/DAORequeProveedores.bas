Attribute VB_Name = "DAORequeProveedores"
Dim rs As ADODB.Recordset

Public Function GetAllByReque(T As clsRequerimiento) As Collection
    On Error GoTo err1
    Dim prov As clsProveedor
    Dim col As New Collection
    Dim rs As Recordset
    Dim strsql As String

    strsql = "select idProveedor from ComprasRequerimientosProveedores p inner join ComprasRequerimientosDetalleMaterial m on p.idDetalleReque=m.id where m.idReque=" & T.id & " group by idProveedor"
    Set rs = conectar.RSFactory(strsql)

    While Not rs.EOF And Not rs.BOF

        Set prov = DAOProveedor.FindById(rs!idProveedor)
        col.Add prov
        rs.MoveNext
    Wend

    Set GetAllByReque = col
    Exit Function

err1:
    Set GetAllByReque = Nothing


End Function

Public Function GetByDetalleReque(id_detalle_reque As Long, b As tipoEntrega) As Collection
    On Error GoTo err1
    Dim rs As Recordset
    Dim col As New Collection
    Dim prov As clsProveedor
    strsql = "select * from ComprasRequerimientosProveedores where idDetalleReque=" & id_detalle_reque & " and tipoDetalleReque=" & b
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF And Not rs.BOF
        Set prov = DAOProveedor.FindById(rs!idProveedor)
        col.Add prov
        rs.MoveNext
    Wend
    Set GetByDetalleReque = col
    Exit Function
err1:
    Set GetByDetalleReque = Nothing
End Function
Public Function Save(T As Collection, id_reque_detalle As Long, Tipo As tipoEntrega) As Boolean
    On Error GoTo err1
    Save = True
    Dim prove As clsProveedor

    'conectar.execute "delete from ComprasRequerimientosProveedores where idDetalleReque=" & id_reque_detalle


    For i = 1 To T.count
        Set prove = T.item(i)

        conectar.execute "insert into ComprasRequerimientosProveedores (idDetalleReque,idProveedor,tipoDetalleReque) values (" & id_reque_detalle & "," & prove.id & "," & Tipo & ")"
    Next


    Exit Function
err1:
    Save = False
End Function
