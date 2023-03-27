Attribute VB_Name = "DAOContacto"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Const TABLA_CONTACTOS As String = "contact"

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "Nombre"
Public Const CAMPO_TELEFONO As String = "Tel"
Public Const CAMPO_CELULAR As String = "celular"
Public Const CAMPO_CARGO As String = "cargo"
Public Const CAMPO_DIRECCION As String = "Direccion"
Public Const CAMPO_LOCALIDAD As String = "Localidad"
Public Const CAMPO_PROVINCIA As String = "Provincia"
Public Const CAMPO_PAIS As String = "País"
Public Const CAMPO_EMAIL As String = "email"
Public Const CAMPO_TIPO As String = "tipo"

Public Const CAMPO_ID_PERSONA As String = "idCliente"


Public Function GetAllByPersona(Id As Long, Tipo As TipoPersona) As Collection
    On Error GoTo err34
    Dim col As New Collection
    strsql = "select * from contactos where idCliente=" & Id & " and tipo=" & Tipo

    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set contacto = New clsContacto
        contacto.idPersona = rs!idCliente
        contacto.Cargo = rs!Cargo
        contacto.celular = rs!celular
        contacto.detalle = rs!detalle
        contacto.Domicilio = rs!direccion
        contacto.email = rs!email
        contacto.Id = rs!Id
        contacto.localidad = rs!localidad
        contacto.nombre = rs!nombre
        contacto.pais = rs!país
        contacto.provincia = rs!provincia
        contacto.telefono = rs!tel
        contacto.Tipo = rs!Tipo
        col.Add contacto
        rs.MoveNext
    Wend
    Set GetAllByPersona = col
    Exit Function
err34:
    Set GetAllByPersona = Nothing
End Function

Public Function LlenarComboByPersona(cbo As Xtremesuitecontrols.ComboBox, Id As Long, Tipo As TipoPersona)

    Dim contacto As clsContacto
    Set col = GetAllByPersona(Id, Tipo)

    For Each contacto In col
        If IsSomething(contacto) Then
            cbo.AddItem contacto.nombre
            cbo.ItemData(cbo.NewIndex) = contacto.Id
        End If
    Next


    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

End Function

Public Function agregar(contacto As clsContacto) As Boolean
    On Error GoTo err1
    Set cn = conectar.obternerConexion
    cn.BeginTrans
    agregar = True
    With contacto
        strsql = "insert into contactos (idCliente, Nombre, Tel, Direccion, Localidad, Provincia, País, Detalle,cargo,celular,email,tipo) values( " & .idPersona & ",'" & .nombre & "','" & .telefono & "','" & .Domicilio & "','" & .localidad & "','" & .provincia & "','" & .pais & "','" & .detalle & "','" & .Cargo & "','" & .celular & "','" & .email & "',1)"
        cn.execute strsql

        Set rs = conectar.RSFactory("select last_insert_id() AS Ultimo from contactos")
        If Not rs.EOF And Not rs.BOF Then
            contacto.Id = rs!ultimo
        Else
            GoTo err1
        End If
        cn.CommitTrans

    End With
    Exit Function
err1:
    cn.RollbackTrans
    agregar = False
End Function

Public Function modificar(contacto As clsContacto) As Boolean
    On Error GoTo err1
    Set cn = conectar.obternerConexion
    modificar = True
    With contacto
        strsql = "update contactos set idCliente= " & .idPersona & ",nombre='" & .nombre & "',tel='" & .telefono & "',direccion='" & .Domicilio & "',localidad='" & .localidad & "',provincia='" & .provincia & "', país='" & .pais & "',detalle='" & .detalle & "',cargo='" & .Cargo & "',celular='" & .celular & "',email='" & .email & "',tipo=" & .Tipo & " where id=" & .Id
        cn.execute strsql
    End With
    Exit Function
err1:
    modificar = False

End Function
Public Function FindById(Tipo As TipoPersona, Id As Long) As clsContacto
    Set FindById = FindAll(Tipo, "id=" & Id)(1)
End Function

Public Function FindAll(ByRef Tipo As TipoPersona, Optional ByRef filter As String = vbNullString) As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim contactos As New Collection

    q = "SELECT" _
      & " contact.*" _
      & " FROM contactos contact" _
      & " WHERE contact.tipo = " & Tipo

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAll = New Collection


    While Not rs.EOF
        contactos.Add Map(rs, fieldsIndex, TABLA_CONTACTOS)
        rs.MoveNext
    Wend
    tickend = GetTickCount
    Debug.Print tickend - tickStart, "ms elapsed"

    Set FindAll = contactos
End Function



Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String) As clsContacto
    Dim c As clsContacto
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If Id > 0 Then
        Set c = New clsContacto
        c.Id = Id
        c.Cargo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CARGO)
        c.celular = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CELULAR)
        'c.detalle = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOContacto.CAMPO_DETALLE)
        c.Domicilio = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DIRECCION)
        c.email = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_EMAIL)
        c.localidad = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_LOCALIDAD)
        c.nombre = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NOMBRE)
        c.pais = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PAIS)
        c.provincia = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PROVINCIA)
        c.telefono = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_TELEFONO)
        c.Tipo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_TIPO)

        c.idPersona = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID_PERSONA)
    End If

    Set Map = c
End Function
