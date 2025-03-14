Attribute VB_Name = "DAOCliente"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset



Public Const CAMPO_ID As String = "id"
Public Const CAMPO_RAZON_SOCIAL As String = "razon"
Public Const CAMPO_DOMICILIO As String = "domicilio"
Public Const CAMPO_LOCALIDAD As String = "localidad"
Public Const CAMPO_CODIGO_POSTAL As String = "CP"
Public Const CAMPO_TELEFONO As String = "telefono"
Public Const CAMPO_FAX As String = "Fax"
Public Const CAMPO_EMAIL As String = "email"
Public Const CAMPO_CUIT As String = "cuit"

Public Const CAMPO_PROVINCIA As String = "id_provincia"

Public Const CAMPO_ESTADO As String = "estado"
Public Const CAMPO_PASSWORD_SISTEMA As String = "passwordSistema"
Public Const CAMPO_FP As String = "FP"
Public Const CAMPO_FORMA_PAGO As String = "FP_detalle"
Public Const TABLA_CLIENTE As String = "c"
Public Const CAMPO_VALIDO_REMITO_FACTURA = "valido_remito_factura"


Public Function BuscarPorID(Id As Long) As clsCliente

    Dim col As Collection
    Set col = DAOCliente.FindAll("c.id = " & Id)
    If col.count > 0 Then
        Set BuscarPorID = col.item(1)
    Else
        Set BuscarPorID = Nothing
    End If

End Function


Public Function crear(Cliente As clsCliente) As Boolean
    Set cn = conectar.obternerConexion
    On Error GoTo err1
    crear = True
    With Cliente
        strsql = "INSERT INTO clientes (cuit_pais, id_impositivo, id_localidad,CP, id_moneda_default, razon,domicilio,telefono,Fax,email,cuit,iva,id_provincia,FP,FP_detalle,valido_remito_factura) VALUES " _
               & "('" & .CuitPais & "', '" & .IDImpositivo & "'," & .localidad.Id & ", " & .CodigoPostal & ", " & .idMonedaDefault & ",'" & .razon & "','" & .Domicilio & "','" & .telefono & "','" & .Fax & "','" & .Email & "','" & .Cuit & "'," & .TipoIVA.idIVA & "," & .provincia.Id & "," & .FP & ",'" & .FormaPago & "'," & conectar.Escape(.ValidoRemitoFactura) & ")"
        cn.execute strsql
    End With
    Exit Function
err1:
    crear = False

End Function


Public Function modificar(Cliente As clsCliente) As Boolean
    Set cn = conectar.obternerConexion
    On Error GoTo err11
    modificar = True
    With Cliente
        strsql = "UPDATE clientes SET id_localidad=" & .localidad.Id & ", CP=" & .CodigoPostal & ",id_moneda_default=" & .idMonedaDefault & ", razon='" & .razon & " ',domicilio='" & .Domicilio & "',telefono='" & .telefono & "',Fax='" & .Fax & "',email='" & .Email & "',cuit='" & .Cuit & "',iva=" & .TipoIVA.idIVA & ",id_provincia='" & .provincia.Id & "',FP=" & .FP & ", FP_detalle='" & .FormaPago & "',valido_remito_factura = " & conectar.Escape(.ValidoRemitoFactura) & "  where id=" & .Id
        cn.execute strsql
    End With
    Exit Function
err11:
    modificar = False
End Function


Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef ivaTableNameOrAlias As String = vbNullString, _
                    Optional ByRef tablaLocalidad As String = vbNullString, _
                    Optional ByRef tablaPais As String = vbNullString, _
                    Optional ByRef tablaProvincia As String = vbNullString) As clsCliente
    Dim C As clsCliente
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)


    If Id > 0 Then
        Set C = New clsCliente
        C.Id = Id
        C.razon = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_RAZON_SOCIAL)
        C.Domicilio = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DOMICILIO)
        C.exLocalidad = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_LOCALIDAD)
        C.CodigoPostal = GetValue(rs, fieldsIndex, tableNameOrAlias, "CP")
        C.telefono = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_TELEFONO)
        C.Fax = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FAX)
        C.TipoDocumento = GetValue(rs, fieldsIndex, tableNameOrAlias, "tipo_doc")

        C.Email = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_EMAIL)
        C.Cuit = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CUIT)
        If LenB(ivaTableNameOrAlias) > 0 Then Set C.TipoIVA = DAOTipoIva.Map(rs, fieldsIndex, ivaTableNameOrAlias)

        If LenB(tablaProvincia) > 0 Then Set C.provincia = DAOProvincias.Map(rs, fieldsIndex, tablaProvincia, tablaPais)
        If LenB(tablaLocalidad) > 0 Then Set C.localidad = DAOLocalidades.Map(rs, fieldsIndex, tablaLocalidad, tablaProvincia, tablaPais)

        C.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ESTADO)
        C.PasswordSistema = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PASSWORD_SISTEMA)
        C.FP = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FP)
        C.FormaPago = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FORMA_PAGO)
        C.TipoIvaID = GetValue(rs, fieldsIndex, tableNameOrAlias, "iva")
        C.ValidoRemitoFactura = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_VALIDO_REMITO_FACTURA)
        C.idMonedaDefault = GetValue(rs, fieldsIndex, tableNameOrAlias, "id_moneda_default")
        
        C.CuitPais = GetValue(rs, fieldsIndex, tableNameOrAlias, "cuit_pais")
        C.IDImpositivo = GetValue(rs, fieldsIndex, tableNameOrAlias, "id_impositivo")
    End If

    Set Map = C
End Function


Public Function LlenarCombo(cbo As ComboBox, Optional todos As Boolean = False, Optional verValidos As Boolean = True, Optional eliminados As Boolean = False)
    Dim col As Collection
    Dim cli As clsCliente

    If verValidos Then
        Set col = DAOCliente.FindAll(DAOCliente.TABLA_CLIENTE & "." & DAOCliente.CAMPO_ESTADO & "=" & EstadoCliente.activo & " ORDER BY " & DAOCliente.CAMPO_RAZON_SOCIAL)
    Else
        Set col = DAOCliente.FindAll(" 1=1  ORDER BY " & DAOCliente.CAMPO_RAZON_SOCIAL)
    End If

    cbo.Clear

    If todos Then
        cbo.AddItem "Todos"
        cbo.ItemData(cbo.NewIndex) = -1
    End If

    If eliminados Then
        cbo.AddItem "Eliminados"
        cbo.ItemData(cbo.NewIndex) = -2
    End If


    'While Not rs.EOF

    For Each cli In col
        cbo.AddItem cli.razon
        cbo.ItemData(cbo.NewIndex) = cli.Id
    Next
    'rs.MoveNext
    'Wend
    'rs.Close

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

End Function


Public Function llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox, Optional todos As Boolean = False, Optional verValidos As Boolean = True, Optional eliminados As Boolean = False)
    Dim col As Collection
    Dim cli As clsCliente
    If verValidos Then
        Set col = DAOCliente.FindAll(DAOCliente.TABLA_CLIENTE & "." & DAOCliente.CAMPO_ESTADO & "=" & EstadoCliente.activo & " ORDER BY " & DAOCliente.CAMPO_RAZON_SOCIAL)
    Else
        Set col = DAOCliente.FindAll(, DAOCliente.CAMPO_RAZON_SOCIAL)
    End If

    cbo.Clear

    If todos Then
        cbo.AddItem "Todos"
        cbo.ItemData(cbo.NewIndex) = -1
    End If

    If eliminados Then
        cbo.AddItem "Eliminados"
        cbo.ItemData(cbo.NewIndex) = -2
    End If

    For Each cli In col
        cbo.AddItem cli.razon
        cbo.ItemData(cbo.NewIndex) = cli.Id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function



Public Function FindAll(Optional ByRef filter As String = vbNullString, Optional ByRef order As String = vbNullString) As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim clientes As New Collection

    q = "SELECT * " _
      & " FROM clientes c" _
      & " LEFT JOIN AdminConfigIVA iva" _
      & " ON (iva.idIVA = c.iva)" _
      & " LEFT JOIN AdminConfigFacturasTipos tfact" _
      & " ON (tfact.id = iva.tipo_factura)" _
      & " LEFT JOIN Localidades l" _
      & " ON (c.id_localidad = l.ID)" _
      & " LEFT JOIN Provincia p" _
      & " ON (l.idProvincia = p.ID)" _
      & " LEFT JOIN Pais pa" _
      & " ON (p.idPais = pa.ID)" _
      & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    If LenB(order) > 0 Then
        q = q & " ORDER BY " & order
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAll = New Collection

    Const ivaTabla As String = "iva"
    Const tipoFacturaTabla As String = "tfact"

    While Not rs.EOF
        clientes.Add Map(rs, fieldsIndex, TABLA_CLIENTE, ivaTabla, "l", "pa", "p")
        rs.MoveNext
    Wend

    tickend = GetTickCount

    Set FindAll = clientes
End Function


