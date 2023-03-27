Attribute VB_Name = "DAOTipoFactura"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_DISCRIMINA As String = "discrimina"
Public Const CAMPO_NUMERACION As String = "numeracion"
Public Const CAMPO_TIPO_FACTURA As String = "TipoFactura"


Public Function FindAllByFilter(filter As String) As Collection
    Dim c As New Collection
    Dim TipoFactura As New clsTipoFactura

    'Set rs = conectar.RSFactory("SELECT * FROM AdminConfigFacturasTipos acft LEFT JOIN AdminConfigFacturaPuntoVenta pv ON  acft.id_punto_venta=pv.id where " & filter)
    Dim q As String
    q = "SELECT   * From   AdminConfigFacturasTiposDiscriminado acftd " _
      & "LEFT JOIN AdminConfigFacturaPuntoVenta pv  ON acftd.id_punto_venta = pv.id " _
      & "LEFT JOIN AdminConfigFacturasTipos acft ON (acftd.`id_tipo_factura`=acft.`id`) " _
      & " LEFT JOIN AdminConfigFacturasTiposDiscriminadoIva acftdi ON (acftdi.id_tipo_factura_discriminado = acftd.id) " _
      & "WHERE " & filter    'acftdi.`id_iva`=3 AND tipo_documento=0"

    Set rs = conectar.RSFactory(q)
    Dim idx As New Dictionary
    conectar.BuildFieldsIndex rs, idx
    While Not rs.EOF

        c.Add Map(rs, idx, "acft")
        rs.MoveNext
    Wend

    Set FindAllByFilter = c
End Function

Public Function FindFirstByFilter(filter As String) As clsTipoFactura
    Dim TipoFactura As New clsTipoFactura
    Set rs = conectar.RSFactory("select * from AdminConfigFacturasTipos t where " & filter)
    If Not rs.EOF And Not rs.EOF Then
        TipoFactura.Tipo = rs!TipoFactura
        TipoFactura.Id = rs!Id
        'TipoFactura.numeracion = rs!numeracion
        TipoFactura.Discrimina = rs!Discrimina
        Set FindFirstByFilter = TipoFactura
    Else
        Set FindFirstByFilter = Nothing
    End If
End Function

Public Function GetById(Id) As clsTipoFactura
    Dim TipoFactura As New clsTipoFactura
    Set rs = conectar.RSFactory("SELECT * FROM AdminConfigFacturasTiposDiscriminado acftd INNER JOIN AdminConfigFacturasTipos acft ON  acftd.`id`=acft.id WHERE acftd.id=" & Id)
    If Not rs.EOF And Not rs.EOF Then
        TipoFactura.Tipo = rs!TipoFactura
        Set TipoFactura.PuntoVenta = DAOPuntoVenta.FindById(rs!id_punto_venta)

        TipoFactura.Id = rs!Id
        'TipoFactura.numeracion = rs!numeracion
        TipoFactura.Discrimina = rs!Discrimina
        Set GetById = TipoFactura

    Else
        GetById = Nothing
    End If
End Function

Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String) As clsTipoFactura
    Dim tfact As clsTipoFactura
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If Id > 0 Then
        Set tfact = New clsTipoFactura
        tfact.Id = Id
        tfact.Discrimina = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DISCRIMINA)
        tfact.Tipo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_TIPO_FACTURA)
        tfact.ExcentoIVA = GetValue(rs, fieldsIndex, tableNameOrAlias, "excento_iva")
    End If

    Set Map = tfact
End Function
