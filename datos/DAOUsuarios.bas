Attribute VB_Name = "DAOUsuarios"

Dim rs As ADODB.Recordset

Public Const CAMPO_ID As String = "id"
'Public Const CAMPO_ESTADO As String = "estado"
Public Const CAMPO_USUARIO As String = "usuario"
Public Const CAMPO_PASSWORD As String = "password"

Public Function FindAll() As Collection
    Dim usuarios As New Collection
    Dim usuario As clsUsuario
    Dim q As String
    Dim rs As New Recordset
    q = "SELECT * FROM usuarios LEFT JOIN personal p ON p.id = usuarios.idEmpleado"
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    While Not rs.EOF
        Set usuario = DAOUsuarios.Map(rs, fieldsIndex, "usuarios", "p")
        usuarios.Add usuario, CStr(usuario.id)
        rs.MoveNext
    Wend

    Set FindAll = usuarios
End Function


Public Function GetById(Optional id As Long) As clsUsuario


    On Error GoTo err1
    Dim usuario As clsUsuario
    Set rs = conectar.RSFactory("select * from usuarios where id=" & id)
    If Not rs.EOF And Not rs.BOF Then
        Set usuario = New clsUsuario
        If Not IsNull(rs!IdEmpleado) Then usuario.Empleado = DAOEmpleados.GetById(rs!IdEmpleado)
        'Usuario.estado = rs!estado
        usuario.id = rs!id
        usuario.PassWord = rs!PassWord
        usuario.usuario = rs!usuario
        'usuario.Memo = rs!Memo_interno
        Set GetById = usuario
    Else
        Set GetById = Nothing
    End If

    Exit Function
err1:

End Function

Public Function GetEmpleado(usuario As clsUsuario) As clsEmpleado
    GetEmpleado = DAOEmpleados.GetById(usuario.id)
End Function

Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, Optional ByVal tablaEmpleado As String = vbNullString) As clsUsuario
    Dim u As clsUsuario
    Dim id As Variant

    id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If id > 0 Then
        Set u = New clsUsuario
        u.id = id
        'u.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ESTADO)
        u.usuario = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_USUARIO)
        u.PassWord = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PASSWORD)
        u.Memo = GetValue(rs, fieldsIndex, tableNameOrAlias, "memo_interno")
        If LenB(tablaEmpleado) > 0 Then
            u.Empleado = DAOEmpleados.Map2(rs, fieldsIndex, tablaEmpleado)
        End If

    End If

    Set Map = u
End Function
