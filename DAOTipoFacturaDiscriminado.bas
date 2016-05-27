Attribute VB_Name = "DAOTipoFacturaDiscriminado"
Option Explicit

Public Function FindAllByFilter(filter As String) As Collection
    Dim c As New Collection
    Dim TipoFactura As New clsTipoFacturaDiscriminado
    Dim rs As New Recordset
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
        Set TipoFactura = New clsTipoFacturaDiscriminado

        c.Add Map(rs, idx, "acftd", "acft", "pv")
        rs.MoveNext
    Wend

    Set FindAllByFilter = c
End Function


Public Function FindByTipoDocumentoAndPuntoVentaAndTipoFactura(idTipoFActura As Long, tipoDocumentoContable As tipoDocumentoContable, IdPuntoVenta As Long, IdTipoIva As Long) As clsTipoFacturaDiscriminado
    Set FindByTipoDocumentoAndPuntoVentaAndTipoFactura = FindAllByFilter("acftd.id_tipo_factura=" & idTipoFActura & " and acftd.tipo_documento=" & tipoDocumentoContable & " and acftd.id_punto_venta=" & IdPuntoVenta & " and acftdi.id_iva=" & IdTipoIva)(1)
End Function

Public Function FindById(id As Long) As clsTipoFacturaDiscriminado
    Set FindById = FindAllByFilter("acftd.id=" & id)(1)
End Function


Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, ByRef tablaTipoFactura As String, ByRef tablaPuntoVenta As String) As clsTipoFacturaDiscriminado
    Dim tfact As clsTipoFacturaDiscriminado
    Dim id As Variant

    id = GetValue(rs, fieldsIndex, tableNameOrAlias, "id")
    
    If id > 0 Then
        Set tfact = New clsTipoFacturaDiscriminado
        tfact.id = id
        tfact.Numeracion = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NUMERACION)
        tfact.TipoDoc = GetValue(rs, fieldsIndex, tableNameOrAlias, "tipo_documento")
        Set tfact.TipoFactura = DAOTipoFactura.Map(rs, fieldsIndex, tablaTipoFactura)
        Set tfact.PuntoVenta = DAOPuntoVenta.Map(rs, fieldsIndex, tablaPuntoVenta)

    End If

    Set Map = tfact
End Function
