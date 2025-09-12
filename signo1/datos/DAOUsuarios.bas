Attribute VB_Name = "DAOUsuarios"

Dim rs As ADODB.Recordset

Public Const CAMPO_ID As String = "id"
'Public Const CAMPO_ESTADO As String = "estado"
Public Const CAMPO_USUARIO As String = "usuario"
Public Const CAMPO_PASSWORD As String = "password"


Public Function FindAll() As Collection
    Dim usuarios As New Collection
    Dim Usuario As clsUsuario
    Dim q As String
    Dim rs As New Recordset
    q = "SELECT * FROM usuarios LEFT JOIN personal p ON p.id = usuarios.idEmpleado"
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    While Not rs.EOF
        Set Usuario = DAOUsuarios.Map(rs, fieldsIndex, "usuarios", "p")
        usuarios.Add Usuario, CStr(Usuario.Id)
        rs.MoveNext
    Wend

    Set FindAll = usuarios
End Function


Public Function GetAll(Optional filtro As String = Empty) As Collection

    On Error GoTo err1
    Dim col As New Collection
    Dim Usuario As clsUsuario
    Dim indice As Dictionary
    Dim q As String

    q = "select * from usuarios usu where 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Set rs = conectar.RSFactory(q)

    conectar.BuildFieldsIndex rs, indice
    While Not rs.EOF
        Set Usuario = New clsUsuario
        Set Usuario = Map(rs, indice, TABLA_USUARIO)
        col.Add Usuario, CStr(Usuario.Id)
        rs.MoveNext
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function


Public Function GetById(Optional Id As Long) As clsUsuario

    On Error GoTo err1
    Dim Usuario As clsUsuario
    Set rs = conectar.RSFactory("select * from usuarios where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Set Usuario = New clsUsuario
        If Not IsNull(rs!IdEmpleado) Then Usuario.Empleado = DAOEmpleados.GetById(rs!IdEmpleado)
        'Usuario.estado = rs!estado
        Usuario.Id = rs!Id
        Usuario.PassWord = rs!PassWord
        Usuario.Usuario = rs!Usuario
        'usuario.Memo = rs!Memo_interno
        Set GetById = Usuario
    Else
        Set GetById = Nothing
    End If

    Exit Function
err1:

End Function

Public Function GetEmpleado(Usuario As clsUsuario) As clsEmpleado
    GetEmpleado = DAOEmpleados.GetById(Usuario.Id)
End Function

Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, Optional ByVal tablaEmpleado As String = vbNullString) As clsUsuario
    Dim u As clsUsuario
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If Id > 0 Then
        Set u = New clsUsuario
        u.Id = Id
        'u.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ESTADO)
        u.Usuario = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_USUARIO)
        u.PassWord = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PASSWORD)
        u.Memo = GetValue(rs, fieldsIndex, tableNameOrAlias, "memo_interno")
        If LenB(tablaEmpleado) > 0 Then
            u.Empleado = DAOEmpleados.Map2(rs, fieldsIndex, tablaEmpleado)
        End If

    End If

    Set Map = u
End Function
