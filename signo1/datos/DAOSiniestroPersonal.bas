Attribute VB_Name = "DAOSiniestroPersonal"
Option Explicit


Public Function Save(sin As SiniestroPersonal) As Boolean
    On Error GoTo E

    Dim q As String

    conectar.BeginTransaction

    If sin.id = 0 Then
        q = "INSERT INTO siniestros_personal  (id_empleado_asegurado,id_empleado_supervisor, id_accidente, nro_siniestro, fecha_ocurrido, diagnostico, prestador_medico, tipo_accidente, tipo_tratamiento,tipo_gravedad,renauda_tareas,gestor,id_art, id_sector) VALUES (" _
            & "'id_empleado_asegurado','id_empleado_supervisor', 'id_accidente'," _
            & "'nro_siniestro'," _
            & "'fecha_ocurrido'," _
            & "'diagnostico'," _
            & "'prestador_medico'," _
            & "'tipo_accidente'," _
            & "'tipo_tratamiento'," _
            & "'tipo_gravedad'," _
            & "'renauda_tareas'," _
            & "'gestor', 'id_art', 'id_sector')"
    Else
        q = "Update siniestros_personal SET" _
            & " id = 'id' ," _
            & " id_empleado_asegurado = 'id_empleado_asegurado' ," _
            & " id_empleado_supervisor = 'id_empleado_supervisor'," _
            & " id_accidente  = 'id_accidente'," _
            & " nro_siniestro = 'nro_siniestro' ," _
            & " fecha_ocurrido = 'fecha_ocurrido' ," _
            & " diagnostico = 'diagnostico' ," _
            & " prestador_medico = 'prestador_medico' ," _
            & " tipo_accidente = 'tipo_accidente' ," _
            & " tipo_tratamiento = 'tipo_tratamiento' ," _
            & " tipo_gravedad = 'tipo_gravedad' ," _
            & " renauda_tareas = 'renauda_tareas' ," _
            & " gestor = 'gestor'," _
            & " id_art = 'id_art'," _
            & " id_sector = 'id_sector'" _
            & " Where id = 'id'"
    End If

    q = Replace$(q, "'id'", sin.id)
    q = Replace$(q, "'id_empleado_asegurado'", GetEntityId(sin.Asegurado))
    q = Replace$(q, "'fecha_ocurrido'", Escape(sin.FechaHoraOcurrido))
    q = Replace$(q, "'nro_siniestro'", Escape(sin.NroSiniestro))
    q = Replace$(q, "'diagnostico'", Escape(sin.Diagnostico))
    q = Replace$(q, "'prestador_medico'", Escape(sin.PrestadorMedico))
    q = Replace$(q, "'tipo_accidente'", Escape(sin.TipoAccidente))
    q = Replace$(q, "'tipo_tratamiento'", Escape(sin.TipoTratamiento))
    q = Replace$(q, "'tipo_gravedad'", Escape(sin.TipoGravedad))
    q = Replace$(q, "'renauda_tareas'", Escape(sin.RenaudaTareas))
    q = Replace$(q, "'gestor'", Escape(sin.Gestor))
    q = Replace$(q, "'id_art'", GetEntityId(sin.ART))
    q = Replace$(q, "'id_empleado_supervisor'", GetEntityId(sin.Supervisor))
    q = Replace$(q, "'id_accidente'", GetEntityId(sin.InformeAccidente))
    q = Replace$(q, "'id_sector'", GetEntityId(sin.Sector))

    Save = conectar.execute(q)

    If sin.id = 0 And Save Then
        sin.id = conectar.UltimoId2()
        Save = (sin.id <> 0)
    End If

    If Save And IsSomething(sin.InformeAccidente) Then
        Save = DAOInformeAccidente.Save(sin.InformeAccidente)
        If Save Then
            Save = conectar.execute("UPDATE siniestros_personal SET id_accidente = " & sin.InformeAccidente.id & " WHERE id = " & sin.id)
        End If
    End If

    If Save Then
        conectar.CommitTransaction
    Else
        conectar.RollBackTransaction
    End If

    Exit Function
E:
    Save = False
    conectar.RollBackTransaction
End Function


Public Function FindAll(Optional ByVal filter As String = vbNullString) As Collection
    Dim q As String
    q = "SELECT *" _
        & " FROM siniestros_personal sp" _
        & " LEFT JOIN personal p ON p.id = sp.id_empleado_asegurado" _
        & " LEFT JOIN personal p2 ON p2.id = sp.id_empleado_supervisor" _
        & " LEFT JOIN art a ON a.id = sp.id_art" _
        & " LEFT JOIN accidentes acc ON acc.id = sp.id_accidente" _
        & " LEFT JOIN sectores sec ON sec.id = sp.id_sector" _
        & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    Dim col As New Collection
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim sin As SiniestroPersonal

    While Not rs.EOF
        Set sin = Map(rs, fieldsIndex, "sp", "p", "p2", "a", "acc", "sec")
        col.Add sin, CStr(sin.id)
        rs.MoveNext
    Wend

    Set FindAll = col

End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaEmpleado As String = vbNullString, Optional tablaEmpleadoSupervisor As String = vbNullString, Optional tablaART As String = vbNullString, Optional tablaAccidente As String = vbNullString, Optional tablaSector As String = vbNullString) As SiniestroPersonal
    Dim s As SiniestroPersonal
    Dim id As Long

    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set s = New SiniestroPersonal
        s.id = id
        s.Diagnostico = GetValue(rs, indice, tabla, "diagnostico")
        s.FechaHoraOcurrido = GetValue(rs, indice, tabla, "fecha_ocurrido")
        s.Gestor = GetValue(rs, indice, tabla, "gestor")
        s.NroSiniestro = GetValue(rs, indice, tabla, "nro_siniestro")
        s.PrestadorMedico = GetValue(rs, indice, tabla, "prestador_medico")
        s.RenaudaTareas = GetValue(rs, indice, tabla, "renauda_tareas")
        s.TipoAccidente = GetValue(rs, indice, tabla, "tipo_accidente")
        s.TipoGravedad = GetValue(rs, indice, tabla, "tipo_gravedad")
        s.TipoTratamiento = GetValue(rs, indice, tabla, "tipo_tratamiento")
        If LenB(tablaEmpleado) > 0 Then Set s.Asegurado = DAOEmpleados.Map2(rs, indice, tablaEmpleado)
        If LenB(tablaEmpleadoSupervisor) > 0 Then Set s.Supervisor = DAOEmpleados.Map2(rs, indice, tablaEmpleadoSupervisor)
        If LenB(tablaART) > 0 Then Set s.ART = DAOART.Map(rs, indice, tablaART)
        If LenB(tablaAccidente) > 0 Then Set s.InformeAccidente = DAOInformeAccidente.Map(rs, indice, tablaAccidente)
        If LenB(tablaSector) > 0 Then Set s.Sector = DAOSectores.Map(rs, indice, tablaSector)
    End If

    Set Map = s
End Function


Public Function GetCantidadSiniestrosPorEmpleado(Optional ByRef idEmpleados As Collection) As Dictionary
    Dim diccionarioRetorno As New Dictionary

    Dim q As String
    q = "SELECT id_empleado_asegurado, COUNT(0) AS cant FROM siniestros_personal"

    If IsSomething(idEmpleados) Then
        q = q & " AND id_empleado_asegurado IN (" & funciones.JoinCollectionValues(idEmpleados, ", ") & ")"
    End If
    q = q & " GROUP BY id_empleado_asegurado"

    Dim rs As New Recordset
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        diccionarioRetorno.Add rs.Fields("id_empleado_asegurado").value, rs.Fields("cant").value
        rs.MoveNext
    Wend

    Set GetCantidadSiniestrosPorEmpleado = diccionarioRetorno
End Function
