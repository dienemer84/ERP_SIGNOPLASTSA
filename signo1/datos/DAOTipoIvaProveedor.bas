Attribute VB_Name = "DAOTipoIvaProveedor"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Public Function GetById(Id) As clsTipoIvaProveedor
    Dim TipoIVA As New clsTipoIvaProveedor
    Set rs = conectar.RSFactory("select * from AdminConfigIVAProveedor where id=" & Id)
    If Not rs.EOF And Not rs.EOF Then
        TipoIVA.detalle = rs!detalle
        TipoIVA.Id = rs!Id
        TipoIVA.configFacturas = DAOConfigFacturaProveedor.getByIdIVA(rs!Id)
        Set GetById = TipoIVA
    End If
End Function
Public Function GetAll() As Collection
    Dim col As New Collection
    Dim A As clsTipoIvaProveedor
    Dim q As String
    q = "SELECT * From  AdminConfigIVAProveedor  LEFT JOIN AdminConfigFacturasProveedor " _
      & " ON (AdminConfigIVAProveedor.id = AdminConfigFacturasProveedor.id) "


    Set rs = conectar.RSFactory(q)
    Dim indice As Dictionary
    conectar.BuildFieldsIndex rs, indice
    Dim configFc As clsConfigFacturaProveedor

    While Not rs.EOF
        Set A = Map(rs, indice, "AdminConfigIVAProveedor", "AdminConfigFacturasProveedor")

        If Not BuscarEnColeccion(col, CStr(A.Id)) Then
            col.Add A, CStr(A.Id)
        Else
            Set A = col.item(CStr(A.Id))
        End If


        Set configFc = DAOConfigFacturaProveedor.Map(rs, indice, "AdminConfigFacturasProveedor")
        If IsSomething(configFc) Then
            If BuscarEnColeccion(A.configFacturas, CStr(configFc.Id)) Then
                A.configFacturas.Add configFc, CStr(configFc.Id)
                End
            End If

        End If


        rs.MoveNext
    Wend
    Set A = Nothing
    Set GetAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaConfigFactura As String = vbNullString) As clsTipoIvaProveedor

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim T As clsTipoIvaProveedor

    If Id >= 0 Then    'comienza con id = 0 en la tabla
        Set T = New clsTipoIvaProveedor
        T.Id = Id
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
        cboIVA.ItemData(cboIVA.NewIndex) = col(P).Id
    Next P
    If cboIVA.ListCount > 0 Then
        cboIVA.ListIndex = 0
    End If

    cboIVA.ListIndex = 3
End Sub
