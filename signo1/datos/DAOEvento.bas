Attribute VB_Name = "DAOEvento"
Option Explicit


Public Function GetEventBroadCastTypes() As Collection
    Dim eventos As New Collection
    Dim eventoDTO As Collection

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_ArchivoDetalleOrdenTrabajo
    eventoDTO.Add "Nuevo Archivo en Detalle OT"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_ArchivoOrdenTrabajo
    eventoDTO.Add "Nuevo Archivo en OT"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_ArchivoPieza
    eventoDTO.Add "Nuevo Archivo en Pieza"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_FacturaAprobada
    eventoDTO.Add "Factura Aprobada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_OrdenEntregaAprobada
    eventoDTO.Add "Orden Entrega Aprobada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_OrdenTrabajoAprobada
    eventoDTO.Add "OT Aprobada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))
    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_OrdenConAnticipoAprobada
    eventoDTO.Add "OT Aprobada con Anticipo"
    eventos.Add eventoDTO, CStr(eventoDTO(1))





    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_OrdenTrabajoModificada
    eventoDTO.Add "OT Modificada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_PresupuestoAprobado
    eventoDTO.Add "Presupuesto Aprobado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_RemitoAprobado
    eventoDTO.Add "Remito Aprobado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_OrdenTrabajoAnulada
    eventoDTO.Add "OT Anulada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_OrdenTrabajoActivada
    eventoDTO.Add "OT En Producción"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_PresupuestoEnviado
    eventoDTO.Add "Presupuesto Enviado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_PresupuestoAnulado
    eventoDTO.Add "Presupuesto Anulado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_PresupuestoCreado
    eventoDTO.Add "Presupuesto Creado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_RemitoAnulado
    eventoDTO.Add "Remito Anulado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_FacturaAnulada
    eventoDTO.Add "Factura Anulada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_RemitoCreado
    eventoDTO.Add "Remito Creado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_FacturaCreada
    eventoDTO.Add "Factura Creada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))


    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_IncidenciaDetalleOrdenTrabajo
    eventoDTO.Add "Nueva Incidencia en Detalle OT"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_IncidenciaOrdenTrabajo
    eventoDTO.Add "Nueva Incidencia en OT"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_IncidenciaPieza
    eventoDTO.Add "Nueva Incidencia en Pieza"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_RequerimientoCompraFinalizado
    eventoDTO.Add "Requerimiento de Compra Finalizado para Aprobación"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_RequerimientoCompraAprobado
    eventoDTO.Add "Requerimiento de Compra Aprobado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_RequerimientoCompraAnulado
    eventoDTO.Add "Requerimiento de Compra Anulado"
    eventos.Add eventoDTO, CStr(eventoDTO(1))

    Set eventoDTO = New Collection
    eventoDTO.Add TipoEventoBroadcast.TEB_PeticionOfertaCreada
    eventoDTO.Add "Petición de Oferta Creada"
    eventos.Add eventoDTO, CStr(eventoDTO(1))




    Set GetEventBroadCastTypes = eventos
End Function

Public Function Publish(ObjetoId As Long, tipoEvento As TipoEventoBroadcast, Optional value As ISuscriber) As Boolean
    On Error GoTo err1
    Dim descripcion As String
    Dim descripcion2 As String
    descripcion2 = vbNullString
    Dim Ot As OrdenTrabajo

    Dim fact As Factura
    Select Case tipoEvento
    Case TipoEventoBroadcast.TEB_ArchivoDetalleOrdenTrabajo
        Dim detaOT As DetalleOrdenTrabajo
        Set detaOT = DAODetalleOrdenTrabajo.FindById(ObjetoId)
        descripcion = "Se ha agregado un archivo al item [" & detaOT.item & "] de la OT Nº " & detaOT.OrdenTrabajo.Id

    Case TipoEventoBroadcast.TEB_ArchivoOrdenTrabajo
        'Dim ot As OrdenTrabajo
        'Set ot = DAOOrdenTrabajo.FindById(objetoId)
        descripcion = "Se ha agregado un archivo a la OT Nº " & ObjetoId
        Set Ot = DAOOrdenTrabajo.FindById(ObjetoId)
        If IsSomething(Ot) Then
            descripcion2 = "<b>Cliente:</b>  " & Ot.ClienteFacturar.razon & "<br>" & " <b>Descripcion:</b> " & Ot.descripcion
        End If


    Case TipoEventoBroadcast.TEB_ArchivoPieza
        descripcion = "Se ha agregado un archivo a la pieza [" & DAOPieza.FindById(ObjetoId, FL_0).nombre & "] - codigo de pieza [" & ObjetoId & "]"


    Case TipoEventoBroadcast.TEB_IncidenciaDetalleOrdenTrabajo
        Set detaOT = DAODetalleOrdenTrabajo.FindById(ObjetoId)
        descripcion = "Se ha agregado una incidencia al item [" & detaOT.item & "] de la OT Nº " & detaOT.OrdenTrabajo.Id
    Case TipoEventoBroadcast.TEB_IncidenciaOrdenTrabajo
        'Dim ot As OrdenTrabajo
        'Set ot = DAOOrdenTrabajo.FindById(objetoId)
        descripcion = "Se ha agregado una incidencia a la OT Nº " & ObjetoId
    Case TipoEventoBroadcast.TEB_IncidenciaPieza
        descripcion = "Se ha agregado una incidencia a la pieza [" & DAOPieza.FindById(ObjetoId, FL_0).nombre & "] - codigo de pieza [" & ObjetoId & "]"


    Case TipoEventoBroadcast.TEB_FacturaAprobada

        Set fact = DAOFactura.FindById(ObjetoId)
        descripcion = "Se ha aprobado el comprobante " & fact.GetShortDescription(False, True)
        descripcion2 = "<b>Cliente:</b>  " & fact.cliente.razon & "<br>" & " <b>OC:</b> " & fact.OrdenCompra

    Case TipoEventoBroadcast.TEB_FacturaAnulada
        Set fact = DAOFactura.FindById(ObjetoId)
        descripcion = "Se ha anulado el comprobante " & fact.GetShortDescription(False, True)
        descripcion2 = "<b>Cliente:</b>  " & fact.cliente.razon & "<br>" & " <b>OC:</b> " & fact.OrdenCompra


    Case TipoEventoBroadcast.TEB_FacturaCreada
        Set fact = DAOFactura.FindById(ObjetoId)
        descripcion = "Se ha creado el comprobante " & fact.GetShortDescription(False, True)
        descripcion2 = "<b>Cliente:</b>  " & fact.cliente.razon & "<br>" & " <b>OC:</b> " & fact.OrdenCompra




        'Case TipoEventoBroadcast.TEB_OrdenEntregaAprobada
        '    Dim ordenEntrega As daoor

    Case TipoEventoBroadcast.TEB_OrdenTrabajoAprobada
        descripcion = "Se ha aprobado la OT Nº " & ObjetoId

        Set Ot = DAOOrdenTrabajo.FindById(ObjetoId)
        If IsSomething(Ot) Then
            descripcion2 = "<b>Cliente:</b>  " & Ot.ClienteFacturar.razon & "<br>" & " <b>Descripcion:</b> " & Ot.descripcion
        End If

    Case TipoEventoBroadcast.TEB_OrdenTrabajoModificada
        descripcion = "Se ha modificado la OT Nº " & ObjetoId

        Set Ot = DAOOrdenTrabajo.FindById(ObjetoId)
        If IsSomething(Ot) Then
            descripcion2 = "<b>Cliente:</b>  " & Ot.ClienteFacturar.razon & "<br>" & " <b>Descripcion:</b> " & Ot.descripcion
        End If

    Case TipoEventoBroadcast.TEB_OrdenTrabajoAnulada
        descripcion = "Se ha anulado la OT Nº " & ObjetoId

        Set Ot = DAOOrdenTrabajo.FindById(ObjetoId)
        If IsSomething(Ot) Then
            descripcion2 = "<b>Cliente:</b>  " & Ot.ClienteFacturar.razon & "<br>" & " <b>Descripcion:</b> " & Ot.descripcion
        End If

    Case TipoEventoBroadcast.TEB_OrdenTrabajoActivada
        descripcion = "Se ha puesto en producción la OT Nº " & ObjetoId

        Set Ot = DAOOrdenTrabajo.FindById(ObjetoId)
        If IsSomething(Ot) Then
            descripcion2 = "<b>Cliente:</b>  " & Ot.ClienteFacturar.razon & "<br>" & " <b>Descripcion:</b> " & Ot.descripcion
        End If




    Case TipoEventoBroadcast.TEB_RemitoAprobado
        Dim rto As Remito
        Set rto = DAORemitoS.FindById(ObjetoId)
        descripcion = "Se ha aprobado el remito Nº " & rto.numero
        descripcion2 = "<b>Cliente:</b>  " & rto.cliente.razon & "<br>" & " <b>Descripcion:</b> " & rto.detalle

    Case TipoEventoBroadcast.TEB_RemitoAnulado
        Set rto = DAORemitoS.FindById(ObjetoId)
        descripcion = "Se ha anulado el remito Nº " & rto.numero
        descripcion2 = "<b>Cliente:</b>  " & rto.cliente.razon & "<br>" & " <b>Descripcion:</b> " & rto.detalle

    Case TipoEventoBroadcast.TEB_RemitoCreado
        Set rto = DAORemitoS.FindById(ObjetoId)
        descripcion = "Se ha creado el remito Nº " & rto.numero



    Case TipoEventoBroadcast.TEB_PresupuestoAnulado
        descripcion = "Se ha anulado el presupuesto Nº " & ObjetoId

    Case TipoEventoBroadcast.TEB_PresupuestoEnviado
        descripcion = "Se ha enviado el presupuesto Nº " & ObjetoId

    Case TipoEventoBroadcast.TEB_PresupuestoCreado
        descripcion = "Se ha creado el presupuesto Nº " & ObjetoId

    Case TipoEventoBroadcast.TEB_PresupuestoAprobado
        descripcion = "Se ha aprobado el presupuesto Nº " & ObjetoId

    Case TipoEventoBroadcast.TEB_RequerimientoCompraFinalizado
        descripcion = "Se ha finalizado el requerimiento de compra Nº " & ObjetoId & " y esta listo para su aprobación"

    Case TipoEventoBroadcast.TEB_RequerimientoCompraAprobado
        descripcion = "Se ha aprobado el requerimiento de compra Nº " & ObjetoId

    Case TipoEventoBroadcast.TEB_RequerimientoCompraAnulado
        descripcion = "Se ha anulado el requerimiento de compra Nº " & ObjetoId

    Case TipoEventoBroadcast.TEB_PeticionOfertaCreada
        descripcion = "Se ha creado la petición de oferta Nº " & ObjetoId
    Case TipoEventoBroadcast.TEB_OrdenConAnticipoAprobada
        descripcion = "Se ha aprobado la Orden de Trabajo con Anticipo Nº " & ObjetoId

        Set Ot = DAOOrdenTrabajo.FindById(ObjetoId)
        If IsSomething(Ot) Then
            descripcion2 = "<b>Cliente:</b>  " & Ot.cliente.razon & "<br>" & " <b>Descripcion:</b> " & Ot.descripcion & "<br>" & " <b>Anticipo:</b> " & Ot.Anticipo & "%" & "<b>Facturado: " & Ot.AnticipoFacturado
        End If

    Case Else
        descripcion = "Ocurrió un evento desconocido"
    End Select


    Dim q As String
    q = "INSERT INTO eventos" _
      & " (id_usuario_involucrado,descripcion,id_tipo_evento,id_objeto_involucrado) Values" _
      & " ('id_usuario_involucrado','descripcion','id_tipo_evento','id_objeto_involucrado')"

    q = Replace$(q, "'id_usuario_involucrado'", funciones.GetUserObj.Id)
    q = Replace$(q, "'descripcion'", conectar.Escape(descripcion))
    q = Replace$(q, "'id_tipo_evento'", tipoEvento)
    q = Replace$(q, "'id_objeto_involucrado'", ObjetoId)

    Publish = conectar.execute(q)


    'envio de emails a los que observan los eventos
    Dim col As New Collection
    Set col = FindUsersByEventType(tipoEvento)

    Dim u As clsUsuario

    For Each u In col
        If Permisos.sistemaUsuarioActivo Then

            Dim mailHelper As New clsMailHelper
            If Not value Is Nothing Then
                Channel.AgregarSuscriptor value, EnvioMail_
            End If
            If mailHelper.isEmail(funciones.GetUserObj.Empleado.email) Then
                If Not value Is Nothing Then
                    mailHelper.EnviarMailEvento descripcion, funciones.GetUserObj.usuario, u.Empleado.email, descripcion2, value

                Else

                    mailHelper.EnviarMailEvento descripcion, funciones.GetUserObj.usuario, u.Empleado.email, descripcion2

                End If

            End If
        End If

    Next u
    Exit Function
err1:



End Function




Public Function FindAllByUser(idUsuario As Long, Optional ByVal unreadOnly As Boolean = False, Optional filtro As String = vbNullString) As Collection
    Dim q As String

    q = "SELECT *" _
      & " FROM eventos e" _
      & " INNER JOIN eventos_suscripciones es ON es.id_tipo_evento = e.id_tipo_evento" _
      & " LEFT JOIN usuarios u ON u.id = e.id_usuario_involucrado" _
      & " LEFT JOIN personal p ON p.id = u.idEmpleado" _
      & " LEFT JOIN eventos_lecturas el ON (el.id_evento = e.id AND (el.id_usuario = " & idUsuario & " OR el.id_usuario IS NULL))" _
      & " Where es.id_usuario = " & idUsuario & " AND e.fecha_creacion >= es.fecha_suscripcion"

    If unreadOnly Then
        q = q & " AND el.id IS NULL"
    End If

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & " ORDER BY e.fecha_creacion DESC"

    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim eventos As New Collection
    Dim EVENTO As EVENTO

    While Not rs.EOF
        Set EVENTO = DAOEvento.Map(rs, fieldsIndex, "e", "u", "el", "p")
        If Not funciones.BuscarEnColeccion(eventos, CStr(EVENTO.Id)) Then eventos.Add EVENTO, CStr(EVENTO.Id)
        rs.MoveNext
    Wend

    Set FindAllByUser = eventos

End Function

Public Function FindUsersByEventType(tipoEvento As TipoEventoBroadcast)
    Dim usuarios As New Collection
    Dim q As String

    q = "SELECT * from " _
      & "  sp.eventos_suscripciones es " _
      & " LEFT JOIN sp.usuarios u ON u.id = es.id_usuario" _
      & " LEFT JOIN sp.personal p ON p.id = u.idEmpleado" _
      & " Where es.id_tipo_evento = " & tipoEvento



    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim eventos As New Collection
    Dim u As clsUsuario
    Dim E As clsEmpleado
    While Not rs.EOF

        Set u = New clsUsuario
        Set u = DAOUsuarios.Map(rs, fieldsIndex, "u")


        u.Empleado = DAOEmpleados.Map2(rs, fieldsIndex, "p")
        usuarios.Add u
        rs.MoveNext
    Wend

    Set FindUsersByEventType = usuarios

End Function



Public Function Map(rs As Recordset, fieldsIndex As Dictionary, tablaEvento As String, tablaUsuario As String, tablaLectura As String, tablaEmpleado As String) As EVENTO
    Dim tmpEvento As EVENTO
    Dim Id As Variant
    Id = GetValue(rs, fieldsIndex, tablaEvento, "id")

    If Id <> 0 Then
        Set tmpEvento = New EVENTO
        tmpEvento.Id = Id
        tmpEvento.descripcion = GetValue(rs, fieldsIndex, tablaEvento, "descripcion")
        tmpEvento.FechaCreacion = GetValue(rs, fieldsIndex, tablaEvento, "fecha_creacion")
        tmpEvento.IdObjetoInvolucrado = GetValue(rs, fieldsIndex, tablaEvento, "id_objeto_involucrado")
        tmpEvento.tipoEvento = GetValue(rs, fieldsIndex, tablaEvento, "id_tipo_evento")

        If LenB(tablaUsuario) > 0 Then Set tmpEvento.UsuarioInvolucrado = DAOUsuarios.Map(rs, fieldsIndex, tablaUsuario, tablaEmpleado)

        If GetValue(rs, fieldsIndex, tablaLectura, "id") Then
            Dim lectura As New LecturaEvento
            lectura.Id = GetValue(rs, fieldsIndex, tablaLectura, "id")
            lectura.FechaLectura = GetValue(rs, fieldsIndex, tablaLectura, "fecha_lectura")
            lectura.idUsuario = GetValue(rs, fieldsIndex, tablaLectura, "id_usuario")
            'set lectura.Usuario

            tmpEvento.Lecturas.Add lectura, CStr(lectura.idUsuario)
        End If
    End If

    Set Map = tmpEvento
End Function


Public Function Read(evento_id As Long) As Boolean
    Dim q As String
    q = "INSERT INTO eventos_lecturas (id_evento, id_usuario) VALUES ('id_evento', 'id_usuario') ON duplicate KEY UPDATE id_evento=id_evento"
    q = Replace$(q, "'id_evento'", evento_id)
    q = Replace$(q, "'id_usuario'", funciones.GetUserObj.Id)
    Read = conectar.execute(q)
End Function

Public Function GetEventBroadCastTypesSuscribedForUser(idUsuario As Long) As Dictionary
    Dim rs As Recordset
    Dim tiposEvento As New Dictionary
    Dim q As String
    q = "SELECT id_tipo_evento FROM eventos_suscripciones WHERE id_usuario = " & idUsuario
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        tiposEvento.Add CStr(rs!id_tipo_evento), rs!id_tipo_evento
        rs.MoveNext
    Wend
    Set GetEventBroadCastTypesSuscribedForUser = tiposEvento
End Function

Public Function AddBroadCastTypesSuscribedForUser(idUsuario As Long, types As Dictionary) As Boolean
    On Error GoTo E
    Dim q As String
    Dim Tipo As Variant

    conectar.BeginTransaction

    q = "DELETE FROM eventos_suscripciones Where id_usuario = " & idUsuario
    If types.count > 0 Then q = q & " and id_tipo_evento NOT IN (" & funciones.JoinDictionaryKeyValues(types, ", ") & ")"
    If Not conectar.execute(q) Then GoTo E

    For Each Tipo In types
        q = "INSERT INTO eventos_suscripciones (id_tipo_evento, id_usuario)" _
          & " VALUES ('id_tipo_evento', 'id_usuario') ON DUPLICATE KEY UPDATE id_usuario = " & idUsuario

        q = Replace$(q, "'id_tipo_evento'", Tipo)
        q = Replace$(q, "'id_usuario'", idUsuario)

        If Not conectar.execute(q) Then GoTo E
    Next Tipo

    conectar.CommitTransaction
    AddBroadCastTypesSuscribedForUser = True
    Exit Function
E:
    conectar.RollBackTransaction
    AddBroadCastTypesSuscribedForUser = False
End Function
