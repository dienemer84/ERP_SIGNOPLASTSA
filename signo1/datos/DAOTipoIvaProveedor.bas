Attribute VB_Name = "DAOTipoIvaProveedor"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function GetById(id) As clsTipoIvaProveedor
    Dim TipoIVA As New clsTipoIvaProveedor
    Set rs = conectar.RSFactory("select * from AdminConfigIVAProveedor where id=" & id)
    If Not rs.EOF And Not rs.EOF Then
        TipoIVA.detalle = rs!detalle
        TipoIVA.id = rs!id
        TipoIVA.configFacturas = DAOConfigFacturaProveedor.getByIdIVA(rs!id)
        Set GetById = TipoIVA
    End If
End Function
Public Function GetAll() As Collection
    Dim col As New Collection
    Dim a As clsTipoIvaProveedor
    q = "SELECT * From  AdminConfigIVAProveedor  LEFT JOIN AdminConfigFacturasProveedor " _
        & " ON (AdminConfigIVAProveedor.id = AdminConfigFacturasProveedor.id) "


    Set rs = conectar.RSFactory(q)
    Dim indice As Dictionary
    conectar.BuildFieldsIndex rs, indice
    Dim configFc As clsConfigFacturaProveedor

    While Not rs.EOF
        Set a = Map(rs, indice, "AdminConfigIVAProveedor", "AdminConfigFacturasProveedor")

        If Not BuscarEnColeccion(col, CStr(a.id)) Then
            col.Add a, CStr(a.id)
        Else
            Set a = col.item(CStr(a.id))
        End If


        Set configFc = DAOConfigFacturaProveedor.Map(rs, indice, "AdminConfigFacturasProveedor")
        If IsSomething(configFc) Then
            If BuscarEnColeccion(a.configFacturas, CStr(configFc.id)) Then
                a.configFacturas.Add configFc, CStr(configFc.id)
                End
            End If

        End If


        rs.MoveNext
    Wend
    Set a = Nothing
    Set GetAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaConfigFactura As String = vbNullString) As clsTipoIvaProveedor

    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    Dim T As clsTipoIvaProveedor

    If id >= 0 Then    'comienza con id = 0 en la tabla
        Set T = New clsTipoIvaProveedor
        T.id = id
        T.detalle = GetValue(rs, indice, tabla, "Detalle")
    End If

    Set Map = T
End Function



Public Sub llenarComboXtremeSuite(cboIVA As Xtremesuitecontrols.ComboBox)
    cboIVA.Clear
    Dim col As Collection
    Set col = DAOTipoIvaProveedor.GetAll
    Dim d As clsTipoIvaProveedor
    For P = 1 To col.count
        cboIVA.AddItem col(P).detalle
        cboIVA.ItemData(cboIVA.NewIndex) = col(P).id
    Next P
    If cboIVA.ListCount > 0 Then
        cboIVA.ListIndex = 0
    End If

    cboIVA.ListIndex = 3
End Sub
