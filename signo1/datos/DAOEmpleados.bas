Attribute VB_Name = "DAOEmpleados"
Dim rs As ADODB.Recordset



Public Function GetAllByTareaId(tareaId As Long) As Collection
    Dim col As New Collection
    Dim a As clsEmpleado

    Set rs = conectar.RSFactory("SELECT DISTINCT * FROM personal p INNER JOIN empleado_tarea et ON et.personal_id = p.id WHERE et.tarea_id = " & tareaId)
    While Not rs.EOF
        col.Add Map(rs)
        rs.MoveNext
    Wend
    Set GetAllByTareaId = col
End Function


Public Function GetAll(Optional filter As String = vbNullString) As Collection
    Dim col As New Collection
    Dim a As clsEmpleado

    Set rs = conectar.RSFactory("select * from personal where 1 = 1 " & filter)
    While Not rs.EOF
        col.Add Map(rs)
        rs.MoveNext
    Wend
    Set GetAll = col
End Function
Public Function GetByLegajo(legajo As String) As clsEmpleado
    On Error Resume Next
    Set GetByLegajo = GetAll(" and legajo=" & legajo)(1)
End Function


Public Function GetById(id As Long) As clsEmpleado
    On Error GoTo err1
    Dim a As New clsEmpleado
    Set rs = conectar.RSFactory("select * from personal where id=" & id)
    If Not rs.EOF And Not rs.BOF Then


        Set GetById = Map(rs)
    Else
        GoTo err1
    End If
    Exit Function

err1:
    Set GetById = Nothing

End Function


Private Function Map(rs As Recordset) As clsEmpleado

    Set a = New clsEmpleado
    a.Apellido = rs!Apellido
    a.nombre = rs!nombre
    a.id = rs!id
    a.direccion = rs!direccion
    a.documento = rs!documento
    If Not IsNull(rs!email) Then a.email = rs!email
    a.estado = rs!estado
    a.legajo = rs!legajo
    a.localidad = rs!localidad
    a.nombre = rs!nombre
    a.Nombres = rs!Nombres
    Set a.sectores = DAOSectores.GetByIdEmpleado(rs!id)
    a.Telefono1 = rs!Telefono1
    a.Telefono2 = rs!Telefono2


    If Not IsNull(rs!fecha_ingreso) Then a.FechaIngreso = rs!fecha_ingreso
    If Not IsNull(rs!fecha_nacimiento) Then a.FechaNacimiento = rs!fecha_nacimiento



    If Not IsNull(rs!grupo_sanguineo) Then a.GrupoSanguineo = rs!grupo_sanguineo


    Set Map = a
End Function


Public Function GetEmpleadosByTareaId(tarea_id) As Collection
    Dim q As String: q = "SELECT DISTINCT emp.* FROM personal emp inner join empleado_tarea et on et.personal_id = emp.id WHERE et.tarea_id = " & tarea_id & " ORDER BY emp.legajo"
    Dim r As Recordset
    Set r = RSFactory(q)
    Dim idx As New Dictionary
    Dim col As New Collection
    BuildFieldsIndex r, idx
    Dim emp As clsEmpleado

    While Not r.EOF
        Set emp = Map2(r, idx, "emp")
        col.Add emp, CStr(emp.id)
        r.MoveNext
    Wend

    Set GetEmpleadosByTareaId = col
End Function
Public Function GetTareasIdAsignadasByPersonalId(personalId As Long) As Dictionary
    Dim q As String: q = "SELECT et.tarea_id FROM empleado_tarea et  WHERE et.personal_id = " & personalId
    Dim r As Recordset
    Set r = RSFactory(q)
    Dim col As New Dictionary
    While Not r.EOF
        col.Add r.Fields("tarea_id").value, 0
        r.MoveNext
    Wend
    Set GetTareasIdAsignadasByPersonalId = col
End Function

Public Function SetTareaAsignada(personal_id As Long, tarea_id As Long, Delete As Boolean) As Boolean
    Dim q As String
    If Delete Then
        q = "DELETE FROM empleado_tarea where personal_id = " & personal_id & " AND tarea_id = " & tarea_id
    Else
        q = "INSERT INTO empleado_tarea (personal_id, tarea_id) values (" & personal_id & ", " & tarea_id & ")"
    End If

    SetTareaAsignada = conectar.execute(q)
End Function


Public Function Map2(rs As Recordset, indice As Dictionary, tabla As String, Optional withFoto As Boolean = False) As clsEmpleado
    Dim E As clsEmpleado
    Dim id As Long

    id = GetValue(rs, indice, tabla, "id")

    If id <> 0 Then
        Set E = New clsEmpleado
        E.id = id
        E.Apellido = GetValue(rs, indice, tabla, "apellido")
        E.direccion = GetValue(rs, indice, tabla, "direccion")
        E.documento = GetValue(rs, indice, tabla, "documento")
        E.email = GetValue(rs, indice, tabla, "email")
        E.estado = GetValue(rs, indice, tabla, "estado")
        E.FechaIngreso = GetValue(rs, indice, tabla, "fecha_ingreso")
        E.FechaNacimiento = GetValue(rs, indice, tabla, "fecha_nacimiento")
        E.GrupoSanguineo = GetValue(rs, indice, tabla, "grupo_sanguineo")
        E.legajo = GetValue(rs, indice, tabla, "legajo")
        E.localidad = GetValue(rs, indice, tabla, "localidad")
        E.nombre = GetValue(rs, indice, tabla, "nombre")
        E.Nombres = GetValue(rs, indice, tabla, "nombres")
        'e.sectores
        E.Telefono1 = GetValue(rs, indice, tabla, "telefono1")
        E.Telefono2 = GetValue(rs, indice, tabla, "telefono2")

        If withFoto Then
            Set E.Foto = DAOArchivo.FindAll(OA_FotoEmpleado, "id=" & E.idFoto)(0)
        End If

    End If

    Set Map2 = E
End Function

Public Function Save(Empleado As clsEmpleado) As Boolean
    On Error GoTo err1
    Dim b As Boolean
    conectar.BeginTransaction
    Dim q As String
    If Empleado.id = 0 Then    'el insert se hace classPersonal
        q = "INSERT INTO personal " _
            & " (legajo," _
            & " documento," _
            & " apellido," _
            & " nombre," _
            & " direccion," _
            & " localidad," _
            & " telefono1," _
            & " telefono2," _
            & " estado," _
            & " nombres," _
            & " email," _
            & " grupo_sanguineo," _
            & " fecha_ingreso," _
          & " fecha_nacimiento)"
        q = q & " Values" _
            & " ('legajo'," _
            & " 'documento'," _
            & " 'apellido'," _
            & " 'nombre'," _
            & " 'direccion'," _
            & " 'localidad'," _
            & " 'telefono1'," _
            & " 'telefono2'," _
            & " 'estado'," _
            & " 'nombres'," _
            & " 'email'," _
            & " 'grupo_sanguineo'," _
            & " 'fecha_ingreso'," _
            & " 'fecha_nacimiento')"

        q = Replace$(q, "'legajo'", Escape(Empleado.legajo))
        q = Replace$(q, "'documento'", Escape(Empleado.documento))
        q = Replace$(q, "'apellido'", Escape(Empleado.Apellido))
        q = Replace$(q, "'nombre'", Escape(Empleado.nombre))
        q = Replace$(q, "'direccion'", Escape(Empleado.direccion))
        q = Replace$(q, "'localidad'", Escape(Empleado.localidad))
        q = Replace$(q, "'telefono1'", Escape(Empleado.Telefono1))
        q = Replace$(q, "'telefono2'", Escape(Empleado.Telefono2))
        q = Replace$(q, "'estado'", Escape(Empleado.estado))
        q = Replace$(q, "'nombres'", Escape(Empleado.Nombres))
        q = Replace$(q, "'email'", Escape(Empleado.email))
        q = Replace$(q, "'grupo_sanguineo'", Escape(Empleado.GrupoSanguineo))
        q = Replace$(q, "'fecha_ingreso'", Escape(Empleado.FechaIngreso))
        q = Replace$(q, "'fecha_nacimiento'", Escape(Empleado.FechaNacimiento))


        'duarante el alta creo el usuario del sistema

        b = True

    Else

        q = "Update personal" _
            & " SET" _
            & " legajo = " & Escape(Empleado.legajo) & " ," _
            & " documento = " & Escape(Empleado.documento) & " ," _
            & " apellido = " & Escape(Empleado.Apellido) & " ," _
            & " nombre = " & Escape(Empleado.nombre) & " ," _
            & " direccion = " & Escape(Empleado.direccion) & " ," _
            & " localidad = " & Escape(Empleado.localidad) & " ," _
            & " telefono1 = " & Escape(Empleado.Telefono1) & " ," _
            & " telefono2 = " & Escape(Empleado.Telefono2) & " ," _
            & " estado = " & Escape(Empleado.estado) & " ," _
            & " nombres = " & Escape(Empleado.Nombres) & " ," _
            & " email = " & Escape(Empleado.email) & " ," _
            & " grupo_sanguineo = " & Escape(Empleado.GrupoSanguineo) & " ," _
            & " fecha_ingreso = " & Escape(Empleado.FechaIngreso) & " ," _
            & " fecha_nacimiento = " & Escape(Empleado.FechaNacimiento) _
            & " Where" _
            & " id = " & Escape(Empleado.id)

    End If

    Save = conectar.execute(q)
    If Save Then If Empleado.id = 0 Then Empleado.id = conectar.UltimoId2




    If b Then
        Dim usu As New clsUsuario
        Dim md As New classMD5
        usu.Empleado = Empleado
        'usu.estado = Empleado.estado
        usu.usuario = Trim(LCase(crearUsuario(Empleado.nombre, Empleado.Apellido)))
        usu.PassWord = md.DigestStrToHexStr(usu.usuario)


        conectar.execute " INSERT INTO sp.usuarios  (usuario,  PASSWORD,  idEmpleado,  estado    )   Values    ('" & usu.usuario & "',  '" & usu.PassWord & "', '" & usu.Empleado.id & "','" & usu.Empleado.estado & "')"
        usu.id = conectar.UltimoId2


        conectar.execute "insert into sp_permisos.Config (idUsuario,Activo) values (" & usu.id & "," & Abs(True) & ")"
        conectar.execute "insert into sp_permisos.Plan (idUsuario) values (" & usu.id & ")"
        conectar.execute "insert into sp_permisos.Desarrollo (idUsuario) values (" & usu.id & ")"
        conectar.execute "insert into sp_permisos.Ventas (idUsuario) values (" & usu.id & ")"
        conectar.execute "insert into sp_permisos.Administracion (idUsuario) values (" & usu.id & ")"
        conectar.execute "insert into sp_permisos.Compras (idUsuario) values (" & usu.id & ")"

        conectar.execute "insert into sp_permisos.rrhh (idUsuario) values (" & usu.id & ")"



    End If


    conectar.CommitTransaction
    Exit Function
err1:
    conectar.RollBackTransaction

End Function