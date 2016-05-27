Attribute VB_Name = "DAOAlicuotas"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function getByIdConfigFactura(id_factura As Long)

    Dim col As New Collection
    Dim a As clsAlicuotas
    Set rs = conectar.RSFactory("select * from AdminConfigIvaAlicuotas where id_config_factura=" & id_factura)
    While Not rs.EOF
        Set a = New clsAlicuotas
        a.Alicuota = rs!Alicuota
        col.Add a
        rs.MoveNext
    Wend


    Set a = Nothing
    Set getByIdConfigFactura = col

End Function





Public Function GetById(id As Long) As clsAlicuotas
    Dim a As clsAlicuotas
    Set rs = conectar.RSFactory("select * from AdminConfigIvaAlicuotas where id=" & id)
    If Not rs.EOF And Not rs.BOF Then
        Set a = New clsAlicuotas
        a.Alicuota = rs!Alicuota
        a.id = rs!id

    End If
    Set GetById = a
    Set a = Nothing
End Function


Public Function getByTipoFactura(id_config_factura As Long) As Collection
    Dim col As New Collection
    Dim a As clsAlicuotas

    Set rs = conectar.RSFactory("select * from AdminConfigIvaAlicuotas where id_config_factura=" & id_config_factura)
    While Not rs.EOF
        Set a = New clsAlicuotas
        a.Alicuota = rs!Alicuota
        a.id = rs!id
        col.Add a
        rs.MoveNext
    Wend
    Set getByTipoFactura = col
    Set col = Nothing
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsAlicuotas

    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim a As clsAlicuotas

    If id > 0 Then
        Set a = New clsAlicuotas
        a.id = id
        a.Alicuota = GetValue(rs, indice, tabla, "alicuota")
    End If

    Set Map = a
End Function
