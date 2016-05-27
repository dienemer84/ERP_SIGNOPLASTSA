Attribute VB_Name = "DAONotaNoConformidad"
Option Explicit

Public Function Imprimir(nnc As NotaNoConformidad)

    Dim Pieza As Pieza
    Set Pieza = DAOPieza.FindById(nnc.TiempoProceso.idPieza, FL_0)

    dsrNotaNoConformidad.Sections("sec4").Controls("lblReferencia").caption = nnc.TiempoProceso.idpedido & " / " & nnc.TiempoProceso.item & " / " & Pieza.nombre
    dsrNotaNoConformidad.Sections("sec4").Controls("lblTareaOrigen").caption = nnc.TareaOrigen.id & " - " & nnc.TareaOrigen.Tarea & " (" & nnc.TareaOrigen.Sector.Sector & ")"
    dsrNotaNoConformidad.Sections("sec4").Controls("lblTareaDestino").caption = nnc.TiempoProceso.Tarea.id & " - " & nnc.TiempoProceso.Tarea.Tarea & " (" & nnc.TiempoProceso.Tarea.Sector.Sector & ")"
    Dim Cap As String
    If nnc.UsuarioResolucionador Is Nothing Then Cap = vbNullString Else Cap = nnc.UsuarioResolucionador.usuario

    dsrNotaNoConformidad.Sections("sec4").Controls("lblAprobador").caption = Cap
    dsrNotaNoConformidad.Sections("sec4").Controls("lblAccion").caption = nnc.AccionTomada
    dsrNotaNoConformidad.Sections("sec4").Controls("lblFalla").caption = nnc.descripcion
    dsrNotaNoConformidad.Sections("sec4").Controls("lblOriginador").caption = nnc.UsuarioCreador.usuario
    
    dsrNotaNoConformidad.Sections("sec4").Controls("lblNroNNc").caption = nnc.numero
dsrNotaNoConformidad.Title = nnc.numero
    Set dsrNotaNoConformidad.DataSource = conectar.RSFactory("select * from NotasNoConformidad limit 1")
    dsrNotaNoConformidad.PrintReport False
End Function

Public Function FindAll(Optional ByVal filter As String = "1 = 1") As Collection
    Dim q As String
    q = "SELECT *" _
        & " From NotasNoConformidad" _
        & " LEFT JOIN PlaneamientoTiemposProcesos ON (NotasNoConformidad.idTiemposProceso = PlaneamientoTiemposProcesos.id)" _
        & " LEFT JOIN usuarios ON (NotasNoConformidad.id_usuario_creador = usuarios.id)" _
        & " LEFT JOIN usuarios usuarios2 ON (NotasNoConformidad.id_usuario_aprobador = usuarios2.id)" _
        & " LEFT JOIN tareas tareas2 ON (NotasNoConformidad.id_tarea_origen = tareas2.id)" _
        & " LEFT JOIN sectores sectores2 ON (sectores2.id = tareas2.id_sector)" _
        & " LEFT JOIN personal ON (NotasNoConformidad.id_encargado = personal.id)" _
        & " LEFT JOIN personal personal2 ON (NotasNoConformidad.id_operario = personal2.id)" _
        & " LEFT JOIN detalles_pedidos ON (PlaneamientoTiemposProcesos.idDetallePedido = detalles_pedidos.id)" _
        & " LEFT JOIN tareas ON (PlaneamientoTiemposProcesos.codigoTarea = tareas.id)" _
        & " LEFT JOIN stock ON (detalles_pedidos.idPieza = stock.id)" _
        & " LEFT JOIN pedidos ON (detalles_pedidos.idPedido = pedidos.id)" _
        & " LEFT JOIN clientes ON (pedidos.idCliente = clientes.id)" _
        & " LEFT JOIN sectores ON (tareas.id_sector = sectores.id)" _
        & " LEFT JOIN Localidades ON (clientes.id_localidad = Localidades.ID)" _
        & " LEFT JOIN Provincia ON (Localidades.idProvincia = Provincia.ID)" _
        & " LEFT JOIN Pais ON (Provincia.idPais = Pais.ID)" _
        & " WHERE " & filter


    Dim col As New Collection
    Dim nnc As NotaNoConformidad

    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set nnc = Map(rs, idx, "NotasNoConformidad", "PlaneamientoTiemposProcesos", "usuarios", "usuarios2" _
                                                                                                , "tareas2", "sectores2", "personal", "personal2", "detalles_pedidos", "tareas", "stock", "pedidos", "clientes", "sectores")
        col.Add nnc, CStr(nnc.id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaTiempoProceso As String = vbNullString, _
                    Optional tablaUsuarioCreador As String = vbNullString, _
                    Optional tablaUsuarioAprobador As String = vbNullString, _
                    Optional tablaTareaOrigen As String = vbNullString, _
                    Optional tablaSectorOrigen As String = vbNullString, _
                    Optional tablaEncargado As String = vbNullString, _
                    Optional tablaOperario As String = vbNullString, _
                    Optional tablaDetalleOT As String = vbNullString, _
                    Optional tablaTareaTiempoProceso As String = vbNullString, _
                    Optional tablaPieza As String = vbNullString, _
                    Optional tablaOT As String = vbNullString, _
                    Optional tablaCliente As String = vbNullString, _
                    Optional tablaSectorTarea As String = vbNullString _
                    ) As NotaNoConformidad

    Dim nnc As NotaNoConformidad
    Dim id As Long: id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set nnc = New NotaNoConformidad
        nnc.id = id
nnc.Incidencias = GetValue(rs, indice, tabla, "incidencias")
        nnc.AccionTomada = GetValue(rs, indice, tabla, "accion")
        nnc.descripcion = GetValue(rs, indice, tabla, "descripcion")
        nnc.FechaResolucion = GetValue(rs, indice, tabla, "fecha_aprobacion")
        nnc.FechaCreacion = GetValue(rs, indice, tabla, "fecha_creacion")
        nnc.estado = GetValue(rs, indice, tabla, "estado")
        If LenB(tablaTiempoProceso) > 0 Then Set nnc.TiempoProceso = DAOTiemposProceso.Map(rs, indice, tablaTiempoProceso, tablaTareaTiempoProceso)
        If IsSomething(nnc.TiempoProceso) Then
            If IsSomething(nnc.TiempoProceso.Tarea) Then
                Set nnc.TiempoProceso.Tarea.Sector = DAOSectores.Map(rs, indice, tablaSectorTarea)
            End If
            Set nnc.TiempoProceso.DetalleOt = DAODetalleOrdenTrabajo.Map(rs, indice, tablaDetalleOT, tablaPieza)
            If IsSomething(nnc.TiempoProceso.DetalleOt) Then
                Set nnc.TiempoProceso.DetalleOt.OrdenTrabajo = DAOOrdenTrabajo.Map(rs, indice, tablaOT, , tablaCliente)
            End If

        End If


        If LenB(tablaTareaOrigen) > 0 Then Set nnc.TareaOrigen = DAOTareas.Map(rs, indice, tablaTareaOrigen, , tablaSectorOrigen)

        If LenB(tablaOperario) > 0 Then Set nnc.Operario = DAOEmpleados.Map2(rs, indice, tablaOperario)
        If LenB(tablaEncargado) > 0 Then Set nnc.Encargado = DAOEmpleados.Map2(rs, indice, tablaEncargado)

        If LenB(tablaUsuarioCreador) > 0 Then Set nnc.UsuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuarioCreador)
        If LenB(tablaUsuarioAprobador) > 0 Then Set nnc.UsuarioResolucionador = DAOUsuarios.Map(rs, indice, tablaUsuarioAprobador)


    End If

    Set Map = nnc
End Function


Public Function Guardar(nnc As NotaNoConformidad) As Boolean
    Dim q As String


    If nnc.id = 0 Then
        q = "INSERT INTO NotasNoConformidad" _
            & " (idTiemposProceso," _
            & " fecha_creacion," _
            & " fecha_aprobacion," _
            & " id_usuario_creador," _
            & " id_operario," _
            & " id_encargado," _
            & " descripcion," _
            & " estado," _
            & " accion," _
            & " id_usuario_aprobador," _
            & " id_tarea_origen)" _
            & " VALUES ('idTiemposProceso'," _
            & " 'fecha_creacion'," _
            & " 'fecha_aprobacion'," _
            & " 'id_usuario_creador'," _
            & " 'id_operario'," _
            & " 'id_encargado'," _
            & " 'descripcion'," _
            & " '0'," _
            & " 'accion'," _
            & " 'id_usuario_aprobador','incidencias'," _
            & " 'id_tarea_origen')"

        If CDbl(nnc.FechaCreacion) = 0 Then nnc.FechaCreacion = Now
    Else
        q = "Update NotasNoConformidad" _
            & " SET idTiemposProceso = 'idTiemposProceso'," _
            & " fecha_creacion = 'fecha_creacion'," _
            & " fecha_aprobacion = 'fecha_aprobacion'," _
            & " id_usuario_creador = 'id_usuario_creador'," _
            & " id_operario = 'id_operario'," _
            & " id_encargado = 'id_encargado'," _
            & " estado = 'estado'," _
            & " descripcion = 'descripcion'," _
            & " accion = 'accion'," _
            & " id_usuario_aprobador = 'id_usuario_aprobador'," _
            & "incidencias = 'incidencias'," _
            & " id_tarea_origen = 'id_tarea_origen'" _
            & " WHERE id = 'id'"

        q = Replace$(q, "'id'", conectar.GetEntityId(nnc))
    End If
    q = Replace$(q, "'idTiemposProceso'", conectar.GetEntityId(nnc.TiempoProceso))
    q = Replace$(q, "'fecha_creacion'", conectar.Escape(nnc.FechaCreacion))
    q = Replace$(q, "'fecha_aprobacion'", conectar.Escape(nnc.FechaResolucion))
    q = Replace$(q, "'id_usuario_creador'", conectar.GetEntityId(nnc.UsuarioCreador))
    q = Replace$(q, "'id_operario'", conectar.GetEntityId(nnc.Operario))
    q = Replace$(q, "'id_encargado'", conectar.GetEntityId(nnc.Encargado))
    q = Replace$(q, "'descripcion'", conectar.Escape(nnc.descripcion))
    q = Replace$(q, "'accion'", conectar.Escape(nnc.AccionTomada))
    q = Replace$(q, "'estado'", conectar.Escape(nnc.estado))
    q = Replace$(q, "'id_usuario_aprobador'", conectar.GetEntityId(nnc.UsuarioResolucionador))
    q = Replace$(q, "'id_tarea_origen'", conectar.GetEntityId(nnc.TareaOrigen))
      q = Replace$(q, "'incidencias'", conectar.Escape(nnc.Incidencias))

    Guardar = conectar.execute(q)
    If Guardar And nnc.id = 0 Then
        Dim id As Long
        Guardar = conectar.UltimoId("NotasNoConformidad", id)
        nnc.id = id
    End If
End Function

