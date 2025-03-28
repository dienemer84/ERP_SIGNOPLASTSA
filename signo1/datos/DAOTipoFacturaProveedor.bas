Attribute VB_Name = "DAOAlicuotas"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function getByIdConfigFactura(id_factura As Long)

    Dim col As New Collection
    Dim A As clsAlicuotas
    Set rs = conectar.RSFactory("select * from AdminConfigIvaAlicuotas where id_config_factura=" & id_factura)
    While Not rs.EOF
        Set A = New clsAlicuotas
        A.alicuota = rs!alicuota
        col.Add A
        rs.MoveNext
    Wend


    Set A = Nothing
    Set getByIdConfigFactura = col

End Function





Public Function GetById(Id As Long) As clsAlicuotas
    Dim A As clsAlicuotas
    Set rs = conectar.RSFactory("select * from AdminConfigIvaAlicuotas where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Set A = New clsAlicuotas
        A.alicuota = rs!alicuota
        A.Id = rs!Id

    End If
    Set GetById = A
    Set A = Nothing
End Function


Public Function getByTipoFactura(id_config_factura As Long) As Collection
    Dim col As New Collection
    Dim A As clsAlicuotas

    Set rs = conectar.RSFactory("select * from AdminConfigIvaAlicuotas where id_config_factura=" & id_config_factura)
    While Not rs.EOF
        Set A = New clsAlicuotas
        A.alicuota = rs!alicuota
        A.Id = rs!Id
        col.Add A
        rs.MoveNext
    Wend
    Set getByTipoFactura = col
    Set col = Nothing
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsAlicuotas

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim A As clsAlicuotas

    If Id > 0 Then
        Set A = New clsAlicuotas
        A.Id = Id
        A.alicuota = GetValue(rs, indice, tabla, "alicuota")
    End If

    Set Map = A
End Function
