Attribute VB_Name = "DAOConfigFacturaProveedor"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim configFactura As clsConfigFacturaProveedor
Public Function getByIdIVA(id_iva) As Collection
    Dim col As New Collection
    Set rs = conectar.RSFactory("select * from AdminConfigFacturasProveedor where id_iva=" & id_iva)


    While Not rs.EOF And Not rs.BOF
        Set configFactura = New clsConfigFacturaProveedor
        configFactura.Id = rs!Id
        configFactura.alicuotas = DAOAlicuotas.getByIdConfigFactura(rs!Id)
        configFactura.Discrimina = rs!Discrimina
        configFactura.TipoFactura = rs!TipoFactura
        configFactura.FormateaNumero = rs!Formatea_numero

        col.Add configFactura
        rs.MoveNext
    Wend

    Set getByIdIVA = col
End Function
Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaIVA As String) As clsConfigFacturaProveedor

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim c As clsConfigFacturaProveedor

    If Id >= 0 Then    'comienza con id= 0 la tabla
        Set c = New clsConfigFacturaProveedor
        c.Id = Id
        c.Discrimina = GetValue(rs, indice, tabla, "discrimina")
        c.TipoFactura = GetValue(rs, indice, tabla, "tipoFactura")
        c.FormateaNumero = GetValue(rs, indice, tabla, "formatea_numero")
        If LenB(tablaIVA) > 0 Then Set c.TipoIvaProveedor = DAOTipoIvaProveedor.Map(rs, indice, tablaIVA)
    End If

    Set Map = c
End Function

Public Function GetById(Id) As clsConfigFacturaProveedor
    Set rs = conectar.RSFactory("select * from AdminConfigFacturasProveedor where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Set configFactura = New clsConfigFacturaProveedor

        configFactura.Id = rs!Id
        configFactura.alicuotas = DAOAlicuotas.getByIdConfigFactura(rs!Id)
        configFactura.Discrimina = rs!Discrimina
        configFactura.TipoFactura = rs!TipoFactura
        configFactura.FormateaNumero = rs!Formatea_numero

    Else
        Set configFactura = Nothing
    End If

    Set GetById = configFactura
End Function

