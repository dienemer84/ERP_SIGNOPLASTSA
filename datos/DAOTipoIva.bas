Attribute VB_Name = "DAOTipoIva"
Public Const CAMPO_ID As String = "idIVA"
Public Const CAMPO_DETALLE As String = "Detalle"
Public Const CAMPO_ALICUOTA As String = "Alicuota"
Public Const CAMPO_VALIDO As String = "valido"

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function GetAll() As Collection
    Dim col As New Collection
    Dim a As clsTipoIVA
    Set rs = conectar.RSFactory("select * from AdminConfigIVA")
    While Not rs.EOF
        Set a = New clsTipoIVA
        a.idIVA = rs!idIVA
        a.Alicuota = rs!Alicuota
        a.detalle = rs!detalle
        a.valido = rs!valido
        'Set a.TipoFactura = DAOTipoFactura.GetById(rs!tipo_factura)
        col.Add a
        rs.MoveNext
    Wend
    Set a = Nothing
    Set GetAll = col
End Function
Public Function GetById(id) As clsTipoIVA
    Dim TipoIVA As New clsTipoIVA
    Set rs = conectar.RSFactory("select * from AdminConfigIVA where idIVA=" & id)
    If Not rs.EOF And Not rs.EOF Then
        TipoIVA.Alicuota = rs!Alicuota
        TipoIVA.valido = rs!valido
        TipoIVA.detalle = rs!detalle
        TipoIVA.idIVA = rs!idIVA
        'Set TipoIVA.TipoFactura = DAOTipoFactura.GetById(rs!tipo_factura)
        Set GetById = TipoIVA
    End If
End Function
Public Function LlenarCombo(cbo As ComboBox)
    Dim col As New Collection
    Dim tmp As clsTipoIVA
    Set col = DAOTipoIva.GetAll
    For x = 1 To col.count
        Set tmp = col(x)
        cbo.AddItem tmp.detalle
        cbo.ItemData(cbo.NewIndex) = tmp.idIVA
    Next x
    If cbo.ListCount > 0 Then cbo.ListIndex = 0

End Function

Public Function llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As New Collection
    Dim tmp As clsTipoIVA
    Set col = DAOTipoIva.GetAll
    For x = 1 To col.count
        Set tmp = col(x)
        cbo.AddItem tmp.detalle
        cbo.ItemData(cbo.NewIndex) = tmp.idIVA
    Next x
    If cbo.ListCount > 0 Then cbo.ListIndex = 0

End Function



Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String) As clsTipoIVA
    Dim TIVA As clsTipoIVA
    Dim id As Variant

    id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If id >= 0 Then
        Set TIVA = New clsTipoIVA
        TIVA.idIVA = id
        TIVA.Alicuota = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ALICUOTA)
        TIVA.detalle = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DETALLE)
        TIVA.valido = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_VALIDO)
        'If LenB(tipoFacturaTableNameOrAlias) > 0 Then Set tiva.TipoFactura = DAOTipoFactura.Map(rs, fieldsIndex, tipoFacturaTableNameOrAlias)
    End If

    Set Map = TIVA
End Function
