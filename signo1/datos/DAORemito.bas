Attribute VB_Name = "DAORemitoS"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_OBSERVACIONES As String = "observaciones_cabecera"
Public Const CAMPO_LUGAR_ENTREGA As String = "datos_entrega_footer"
Public Const CAMPO_DETALLE As String = "detalle"
Public Const CAMPO_NUMERO As String = "numero"
Public Const CAMPO_ESTADO As String = "estado"
Public Const CAMPO_ESTADO_FACTURADO As String = "estadoFacturado"
Public Const CAMPO_FECHA As String = "fecha"
Public Const CAMPO_IMPRESO As String = "impreso"
Public Const TABLA_REMITO As String = "rto"
Public Const TABLA_CLIENTE As String = "cli"
Public Const TABLA_USUARIO_CREADOR As String = "u1"
Public Const TABLA_USUARIO_APROBADOR As String = "u2"
Public Const TABLA_CONTACTO As String = "cont"




Public Function ImprimirBultos(T As Remito) As Boolean
    On Error GoTo err1
    Dim x As Integer
    Dim Obj As PageSet.PrinterControl
    Set Obj = New PrinterControl
    Obj.ChngOrientationLandscape
    For x = 1 To T.CantidadBultos

        dsrCantidadBultos.Sections("fondo").Controls("bulto").caption = "BULTO " & x & " de " & T.CantidadBultos
        dsrCantidadBultos.Sections("fondo").Controls("remito").caption = "Remito Nº: " & T.numero
        dsrCantidadBultos.Sections("fondo").Controls("descripcion").caption = "Descripción: " & T.detalle
        Set dsrCantidadBultos.DataSource = conectar.RSFactory("Select 1")
        dsrCantidadBultos.PrintReport False

        ImprimirBultos = True
    Next x
    Obj.ReSetOrientation
    Exit Function

err1:
    ImprimirBultos = False
    Obj.ReSetOrientation

End Function


Public Function Save(T As Remito, Optional Cascade As Boolean = False, Optional NotificarObserver As Boolean = True) As Boolean
    conectar.BeginTransaction
    Save = Guardar(T, Cascade, NotificarObserver)
    If Not Save Then GoTo err1
    conectar.CommitTransaction
    Exit Function
err1:
    Save = False
    conectar.RollBackTransaction
End Function
Public Function Guardar(T As Remito, Optional Cascade As Boolean = False, Optional NotificarObserver As Boolean = True) As Boolean
    Dim q As String
    Guardar = True

    Dim Nueva As Boolean
    If T.Id = 0 Then
        Nueva = True
        '        q = "INSERT INTO remitos (detalle, idCliente,  fecha,  estado,  estadoFacturado,  impreso,  idContacto," _
                 '          & "idUsuario, numero,idUsuarioAprobador) Values (" _
                 '          & conectar.Escape(T.detalle) & ", " _
                 '          & conectar.GetEntityId(T.cliente) & ", " _
                 '          & conectar.Escape(T.FEcha) & ", " _
                 '          & conectar.Escape(T.estado) & "," _
                 '          & "0," _
                 '          & conectar.Escape(T.EstadoFacturado) & ", " _
                 '          & conectar.GetEntityId(T.contacto) & "," _
                 '          & conectar.GetEntityId(T.usuarioCreador) & ", " _
                 '          & conectar.Escape(T.numero) & ", " _
                 '          & conectar.GetEntityId(T.usuarioAprobador) & ")"

        q = "INSERT INTO remitos (observaciones_cabecera, datos_entrega_footer, detalle, idCliente,  fecha,  estado,  estadoFacturado,  impreso,  idContacto," _
          & "idUsuario, numero,idUsuarioAprobador) Values (" _
          & conectar.Escape(T.observaciones) & ", " _
          & conectar.Escape(T.lugarEntrega) & ", " _
          & conectar.Escape(T.detalle) & ", " _
          & conectar.GetEntityId(T.cliente) & ", " _
          & conectar.Escape(T.FEcha) & ", " _
          & conectar.Escape(T.estado) & "," _
          & "0," _
          & conectar.Escape(T.EstadoFacturado) & ", " _
          & conectar.GetEntityId(T.contacto) & "," _
          & conectar.GetEntityId(T.usuarioCreador) & ", " _
          & conectar.Escape(T.numero) & ", " _
          & conectar.GetEntityId(T.usuarioAprobador) & ")"

    Else
        Nueva = False
        q = "Update remitos " _
          & "SET " _
          & "observaciones_cabecera = " & conectar.Escape(T.observaciones) & " ," _
          & "datos_entrega_footer = " & conectar.Escape(T.lugarEntrega) & " ," _
          & "detalle = " & conectar.Escape(T.detalle) & " ," _
          & "idCliente =" & conectar.GetEntityId(T.cliente) & " ," _
          & "fecha = " & conectar.Escape(T.FEcha) & " ," _
          & "estado =" & conectar.Escape(T.estado) & "," _
          & "estadoFacturado =" & conectar.Escape(T.EstadoFacturado) & "," _
          & "idContacto = " & conectar.GetEntityId(T.contacto) & "," _
          & "numero = " & conectar.Escape(T.numero) & "," _
          & "idUsuario = " & conectar.GetEntityId(T.usuarioCreador) & "," _
          & "cantidad_bultos = " & conectar.Escape(T.CantidadBultos) & "," _
          & "idUsuarioAprobador = " & conectar.GetEntityId(T.usuarioAprobador) & " Where " _
          & "id =" & T.Id

    End If
    If Not conectar.execute(q) Then GoTo err1

    If T.Id = 0 Then
        T.Id = conectar.UltimoId2
    End If
    If Cascade Then
        If Not conectar.execute("DELETE FROM entregas WHERE remito=" & T.Id) Then GoTo err1

        Dim deta As remitoDetalle
        For Each deta In T.Detalles
            deta.Id = 0
            deta.Remito = T.Id
            If Not DAORemitoSDetalle.Guardar(deta) Then GoTo err1
        Next


        Dim evento2 As New clsEventoObserver

        Set evento2.Elemento = T.Detalles
        evento2.EVENTO = agregar_

        Set evento2.Originador = Nothing
        evento2.Tipo = RemitosDetalle_
        Channel.Notificar evento2, RemitosDetalle_
    End If


    If NotificarObserver Then
        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = T
        If Nueva Then
            EVENTO.EVENTO = agregar_
        Else
            EVENTO.EVENTO = modificar_
        End If
        Set EVENTO.Originador = Nothing
        EVENTO.Tipo = Remitos_
        Channel.Notificar EVENTO, Remitos_
    End If



    Exit Function

err1:
    Guardar = False
End Function

Public Function ProximoRemito() As Long
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select max(numero)+1 as proximo from remitos")
    If Not rs.EOF And Not rs.BOF Then ProximoRemito = rs!proximo
End Function

Public Function CambiarEstadoFacturable(T As Remito) As Boolean
    On Error GoTo err1
    CambiarEstadoFacturable = True
    Dim estAnt As EstadoRemitoFacturado
    estAnt = T.EstadoFacturado
    conectar.BeginTransaction
    Dim deta As remitoDetalle
    If T.estado <> RemitoAnulado Then    'si no esta anulado
        If T.EstadoFacturado = RemitoNoFacturado Then     ' no facturado
            If MsgBox("¿Está seguro de marcar este remito como No Facturable?", vbYesNo, "Confirmar") = vbYes Then
                If Not conectar.execute("update remitos set estadoFacturado=3 where id=" & T.Id) Then GoTo err1


                T.EstadoFacturado = RemitoNoFacturable
                Set T.Detalles = DAORemitoSDetalle.FindAllByRemito(T.Id)
                For Each deta In T.Detalles


                    If Not DAORemitoSDetalle.CambiarEstadoFacturable(False, deta) Then GoTo err1
                Next deta
            End If
        ElseIf T.EstadoFacturado = RemitoNoFacturable Then
            If MsgBox("¿Está seguro de marcar este remito como Facturable?", vbYesNo, "Confirmar") = vbYes Then
                If Not conectar.execute("update remitos set estadoFacturado=0 where id=" & T.Id) Then GoTo err1
                T.EstadoFacturado = RemitoNoFacturado
                Set T.Detalles = DAORemitoSDetalle.FindAllByRemito(T.Id)
                For Each deta In T.Detalles
                    If Not DAORemitoSDetalle.CambiarEstadoFacturable(True, deta) Then GoTo err1
                Next deta
            End If
        End If
    End If
    If Not DAORemitoS.CambiarEstadoFacturado(T.Id, DAORemitoS.AnalizarEstadoFacturado(T.Id)) Then GoTo err1
    conectar.CommitTransaction

    Exit Function

err1:
    CambiarEstadoFacturable = False
    T.EstadoFacturado = estAnt
    conectar.RollBackTransaction
End Function

Public Function FindByNumero(numero As Long) As Remito
    On Error GoTo err1
    Set FindByNumero = FindAll("and rto.Numero=" & numero)(1)
    Exit Function
err1:
    Set FindByNumero = Nothing
End Function


Public Function FindById(Id As Long) As Remito
    On Error GoTo err1
    Set FindById = FindAll("and rto.id=" & Id)(1)
    Exit Function
err1:
    Set FindById = Nothing
End Function

Public Function FindAll(Optional filter As String = vbNullString) As Collection
    Dim strsql As String
    Dim indice As Dictionary
    Dim rs As Recordset
    Dim col As New Collection
    strsql = "SELECT * FROM remitos rto LEFT JOIN clientes cli ON rto.idCliente=cli.id LEFT JOIN Localidades  ON cli.id_localidad = Localidades.ID   LEFT JOIN Provincia  ON cli.id_provincia = Provincia.ID   LEFT JOIN usuarios u1 ON rto.idUsuario=u1.id LEFT JOIN usuarios u2 ON rto.IdUsuarioAprobador=u2.id LEFT JOIN contactos cont ON rto.idContacto=cont.id WHERE 1=1 "
    'strsql = "SELECT * FROM remitos rto LEFT JOIN clientes cli ON rto.idCliente=cli.id LEFT JOIN Localidades  ON cli.id_localidad = Localidades.ID   LEFT JOIN Provincia  ON cli.id_provincia = Provincia.ID   LEFT JOIN usuarios u1 ON rto.idUsuario=u1.id LEFT JOIN usuarios u2 ON rto.IdUsuarioAprobador=u2.id LEFT JOIN contactos cont ON rto.idContacto=cont.id LEFT JOIN entregas e ON e.Remito=rto.id LEFT JOIN detalles_pedidos dp ON dp.id=e.idDetallePedido  WHERE 1=1 "


    If Len(filter) > 0 Then strsql = strsql & " " & filter
    strsql = strsql & " ORDER BY rto.numero DESC"

    Set rs = conectar.RSFactory(strsql)

    conectar.BuildFieldsIndex rs, indice
    Dim Re As Remito

    While Not rs.EOF
        Set Re = DAORemitoS.Map(rs, indice, TABLA_REMITO, TABLA_CLIENTE, TABLA_USUARIO_CREADOR, TABLA_USUARIO_APROBADOR, TABLA_CONTACTO)
        col.Add Re, CStr(Re.Id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function




Public Function Map(ByRef rs As Recordset, ByRef indice As Dictionary, ByRef tabla As String, Optional ByRef tablaCliente As String, Optional ByRef tablaUsuCreador As String, Optional ByRef TablaUsuAprobador As String, Optional ByRef tablaContacto As String) As Remito

    Dim Remito As Remito
    Dim Id As Variant
    Id = GetValue(rs, indice, tabla, CAMPO_ID)

    If Id > 0 Then
        Set Remito = New Remito
        Remito.Id = Id
        Remito.observaciones = GetValue(rs, indice, tabla, DAORemitoS.CAMPO_OBSERVACIONES)
        Remito.lugarEntrega = GetValue(rs, indice, tabla, DAORemitoS.CAMPO_LUGAR_ENTREGA)
        Remito.detalle = GetValue(rs, indice, tabla, DAORemitoS.CAMPO_DETALLE)
        Remito.numero = GetValue(rs, indice, tabla, DAORemitoS.CAMPO_NUMERO)
        Remito.estado = GetValue(rs, indice, tabla, CAMPO_ESTADO)
        Remito.EstadoFacturado = GetValue(rs, indice, tabla, CAMPO_ESTADO_FACTURADO)
        Remito.FEcha = GetValue(rs, indice, tabla, CAMPO_FECHA)
        Remito.CantidadBultos = GetValue(rs, indice, tabla, "cantidad_bultos")
        Remito.ControlCargaImpresiones = GetValue(rs, indice, tabla, "control_carga")
        If LenB(tablaCliente) > 0 Then Set Remito.cliente = DAOCliente.Map(rs, indice, tablaCliente, , "Localidades", "", "Provincia")
        If LenB(tablaUsuCreador) > 0 Then Set Remito.usuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuCreador)
        If LenB(TablaUsuAprobador) > 0 Then Set Remito.usuarioAprobador = DAOUsuarios.Map(rs, indice, TablaUsuAprobador)
        If LenB(tablaContacto) > 0 Then Set Remito.contacto = DAOContacto.Map(rs, indice, tablaContacto)
    End If

    Set Map = Remito
End Function
Public Function CambiarEstadoFacturado(T As Long, estadoNuevo As EstadoRemitoFacturado) As Boolean
    CambiarEstadoFacturado = True
    On Error GoTo err1

    CambiarEstadoFacturado = conectar.execute("update remitos set estadoFacturado=" & estadoNuevo & " where id=" & T)
    Dim msg As String
    If estadoNuevo = RemitoFacturadoTotal Then
        msg = "remito Facturado total"
    ElseIf estadoNuevo = RemitoNoFacturable Then
        msg = "remito no facturable"
    ElseIf estadoNuevo = RemitoNoFacturado Then
        msg = "remito no facturado"
    ElseIf estadoNuevo = RemitoFacturadoParcial Then
        msg = "remito facturado parcial"

    End If


    Dim rto As Remito
    Set rto = DAORemitoS.FindById(T)
    DAORemitoHistorico.agregar rto, msg

    Dim EVENTO As New clsEventoObserver
    Set EVENTO.Elemento = rto
    EVENTO.Tipo = Remitos_
    EVENTO.EVENTO = modificar_
    Set EVENTO.Originador = Nothing

    Channel.Notificar EVENTO, Remitos_


    Exit Function



err1:

    CambiarEstadoFacturado = False
End Function

Public Function AnalizarEstadoFacturado(idRto As Long) As EstadoRemitoFacturado
    Dim deta As remitoDetalle
    Dim cf As Long
    Dim cnf As Long
    Dim ct As Long
    Dim c As Long
    c = 0
    ct = 0
    cf = 0
    cnf = 0
    Dim rto As Remito
    Set rto = DAORemitoS.FindAll("and " & DAORemitoS.TABLA_REMITO & ".id=" & idRto)(1)
    If Not IsNull(rto) Then Set rto.Detalles = DAORemitoSDetalle.FindAllByRemito(rto.Id)

    If Not IsNull(rto.Detalles) Then
        For Each deta In rto.Detalles
            ct = ct + 1
            If deta.facturable Then c = c + 1
            If Not deta.facturable Then cnf = cnf + 1
            If deta.Facturado And deta.facturable Then cf = cf + 1
        Next deta


        If ct = cnf Then
            AnalizarEstadoFacturado = RemitoNoFacturable
        Else
            ct = ct - cnf



            If cf = 0 And ct > 0 Then
                AnalizarEstadoFacturado = RemitoNoFacturado
            ElseIf cf + cnf = ct + cnf Then
                AnalizarEstadoFacturado = RemitoFacturadoTotal
            ElseIf ct + cnf > cf + cnf Then
                AnalizarEstadoFacturado = RemitoFacturadoParcial
            End If
        End If
    End If
End Function

Public Function Anular(Remito As Remito) As Boolean
    On Error GoTo erranu
    Dim FEcha As Date
    Dim Autor As Long
    Dim tra As Boolean
    Dim canti As Long
    Dim rs_s As Recordset
    Dim estado_p As Integer
    Dim estado_nuevo As Integer
    Dim detalle As remitoDetalle
    conectar.BeginTransaction
    Anular = True
    tra = True
    Dim est_ant As EstadoRemito


    If Remito.estado = RemitoAnulado Then Err.Raise 1911, , "No se puede anular un remito ya anulado!"

    If Remito.EstadoFacturado = RemitoFacturadoParcial Or Remito.EstadoFacturado = RemitoFacturadoTotal Then
        MsgBox "Remito facturado total o parcialmente!" & Chr(10) & "Por favor primero anule la/s facturas involucradas!", vbInformation, "Información"
        GoTo erranu
    End If

    Remito.estado = RemitoAnulado

    If Not DAORemitoS.Guardar(Remito, False) Then GoTo erranu

    Set Remito.Detalles = DAORemitoSDetalle.FindAllByRemito(Remito.Id)

    For Each detalle In Remito.Detalles
        canti = detalle.Cantidad    'FIX 08-02-2010 | para que reste la cantidad entregada

        'resto la cantidad entregada
        If detalle.Origen = 1 Then


            Set rs_s = conectar.RSFactory("select estado from pedidos where id=" & detalle.idpedido)
            estado_p = rs_s!estado

            If estado_p = 3 Then estado_nuevo = 3
            If estado_p = 4 Then estado_nuevo = 3
            If estado_p = 2 Then estado_nuevo = 2


            If Not conectar.execute("update pedidos set estado=" & estado_nuevo & " where id=" & detalle.idpedido) Then GoTo erranu
            Autor = funciones.getUser
            FEcha = funciones.datetimeFormateada(Now)

            If Not conectar.execute("insert into historial_pedido (idPedido,nota,fecha,autor) values (" & detalle.idpedido & ",'Pedido abierto por anulación de remito','" & FEcha & "'," & Autor & ")") Then GoTo erranu


            'resto desde detalles_pedidos
            If detalle.idDetallePedido > 0 Then     'solo si no es concepto
                If Not conectar.execute("update detalles_pedidos set cantidad_entregada=cantidad_entregada-" & canti & " where id=" & detalle.idDetallePedido) Then GoTo erranu
                If Not DAODetalleOrdenTrabajo.SaveCantidad(detalle.idDetallePedido, -canti, CantidadEntregada_, 0, Remito.Id, 0, 0, 0) Then GoTo erranu

            End If
        ElseIf detalle.Origen = 2 Then

            'tengo que ver si la OE esta cerrada,
            'si esta cerrada hay q abrirla para poder remitar lo anulado
            If Not conectar.execute("Update PedidosEntregas set estado=2 where id=" & detalle.idpedido) Then GoTo erranu
            'resto desde detallesPedidosEntregas
            If Not conectar.execute("update detallesPedidosEntregas set entregados=entregados-" & canti & " where id=" & detalle.idDetallePedido) Then GoTo erranu
        End If
    Next detalle

    If Not DAORemitoHistorico.agregar(Remito, "Remito anulado") Then GoTo erranu

    tra = False
    conectar.CommitTransaction
    DAOEvento.Publish Remito.Id, TipoEventoBroadcast.TEB_RemitoAnulado
    Exit Function
erranu:
    If tra Then conectar.RollBackTransaction
    Anular = False
    Remito.estado = est_ant
    If Err.Number = 1911 Then MsgBox Err.Description
End Function

Public Function aprobar(Remito As Remito) As Boolean
    On Error GoTo errh44

    If Remito.ControlCargaImpresiones = 0 Then
        MsgBox "Aún no imprimio la planilla de control de carga" & Chr(10) & "Por favor, realice el control antes de aprobar.", vbInformation, "Información"
        Exit Function
    End If



    'controlo si el seguimiento esta hecho
    Set Remito.Detalles = DAORemitoSDetalle.FindAllByRemito(Remito.Id, True, True)
    Dim deta As remitoDetalle
    Dim segui As Boolean
    Dim cantok As Boolean
    cantok = True
    segui = True
    Dim Items As String
    For Each deta In Remito.Detalles



        If deta.Origen = OrigenRemitoConcepto Then

        Else
            Set deta.DetallePedido = DAODetalleOrdenTrabajo.FindById(deta.DetallePedido.Id, True, True, False)
            If (deta.DetallePedido.Cantidad_Fabricada + deta.DetallePedido.ReservaStock) - deta.DetallePedido.Cantidad_Entregada >= deta.Cantidad Then
            Else
                segui = False
            End If

            If (deta.DetallePedido.CantidadPedida - deta.DetallePedido.CantidadConsumida) < deta.Cantidad Then
                cantok = False
                Items = Items & " " & deta.DetallePedido.item
            End If
        End If
    Next

    If Not segui Then
        MsgBox "No hay piezas fabricadas disponibles para poder entregar.", vbCritical, "Error"
        Exit Function
    End If

    If Not cantok Then
        MsgBox "No hay  cantidad pedida disponible para poder entregar." & Chr(10) & " items: " & Items, vbCritical, "Error"
        Exit Function
    End If





    Dim est_ant As EstadoRemito
    aprobar = True
    est_ant = Remito.estado
    Remito.estado = RemitoAprobado
    Set Remito.usuarioAprobador = funciones.GetUserObj




    Dim Ot As OrdenTrabajo
    For Each deta In Remito.Detalles



        If deta.Origen = OrigenRemitoOt Then
            Set Ot = DAOOrdenTrabajo.FindById(deta.idpedido)
            If Ot.Anticipo = 100 And Ot.AnticipoFacturado Then
                'si se facturo todo x adelantado, se marcad como facturado.
                Remito.EstadoFacturado = RemitoFacturadoTotal
            End If


            'antes de aprobar tengo q volver a validar, por un tema de concurrencia.


            If deta.DetallePedido.CantidadEntregada + deta.Cantidad > deta.DetallePedido.CantidadPedida Then
                MsgBox "No puede entregar más de lo que tiene pedido!, por favor revea el remito!", vbInformation
                Exit Function

            Else
                conectar.execute "update detalles_pedidos set cantidad_entregada=cantidad_entregada+" & deta.Cantidad & " Where idPedido=" & deta.idpedido & " and id=" & deta.idDetallePedido
                DAODetalleOrdenTrabajo.SaveCantidad deta.idDetallePedido, deta.Cantidad, CantidadEntregada_, 0, Remito.Id, 0, 0, 0
                If Ot.Anticipo > 0 And Ot.AnticipoFacturado Then
                    deta.Facturado = True
                    deta.facturable = True
                End If


            End If

        ElseIf deta.Origen = OrigenRemitooe Then
            conectar.execute "update detallesPedidosEntregas set entregados=entregados+" & deta.Cantidad & " Where id=" & deta.idDetallePedido
        End If




    Next deta

    If Not DAORemitoS.Save(Remito, False) Then GoTo errh44
    If Not DAORemitoHistorico.agregar(Remito, "REMITO APROBADO") Then GoTo errh44

    DAOEvento.Publish Remito.Id, TipoEventoBroadcast.TEB_RemitoAprobado

    Exit Function
errh44:
    Remito.estado = est_ant
    Remito.usuarioAprobador = Nothing
    aprobar = False
End Function



Public Function ImprimirControlCarga(rto As Remito) As Boolean
    On Error GoTo err1
    dsrControlCarga.Sections("section4").Controls.item("lblRemitoNumero").caption = "Remito Nº: " & rto.numero
    dsrControlCarga.Sections("section4").Controls.item("lblCliente").caption = "Cliente: " & rto.cliente.razon
    dsrControlCarga.Sections("section4").Controls.item("lblDetalleRemito").caption = "Observaciones: " & rto.detalle

    Dim r As New Recordset
    With r
        .Fields.Append "cantidad", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "origen", adVarChar, 20, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "item", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "nota", adVarChar, 255, adFldUpdatable   ' And adFldIsNullable
        .Fields.Append "observaciones", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable

    End With
    r.Open


    Set rto.Detalles = DAORemitoSDetalle.FindAllByRemito(rto.Id, , True)

    Dim deta As remitoDetalle
    For Each deta In rto.Detalles
        r.AddNew
        r!Cantidad = IIf(deta.Cantidad = 0, "", funciones.FormatearDecimales(deta.Cantidad))
        r!Origen = deta.VerOrigen
        r!observaciones = deta.observaciones
        If deta.Origen = OrigenRemitoConcepto Then
            r!item = "000"
        Else
            r!item = deta.DetallePedido.item
        End If

        If deta.Origen = OrigenRemitoOt Or deta.Origen = OrigenRemitoAplicado Then
            '            r!Nota = deta.DetallePedido.Nota
            If deta.DetallePedido.Nota = deta.observaciones Then
                r!Nota = ""
            Else
                r!Nota = deta.DetallePedido.Nota
            End If
        End If

        r!detalle = deta.VerElemento
        r.Update
    Next

    Set dsrControlCarga.DataSource = r
    dsrControlCarga.PrintReport True
    conectar.execute "update remitos set control_carga=control_carga+1 where id = " & rto.Id
    rto.ControlCargaImpresiones = rto.ControlCargaImpresiones + 1


    dsrDatosDespacho.Sections("section4").Controls.item("lblRemitoNumero").caption = "Remito Nº: " & rto.numero
    dsrDatosDespacho.Sections("section4").Controls.item("lblCliente").caption = "Cliente: " & rto.cliente.razon
    dsrDatosDespacho.Sections("section4").Controls.item("lblDetalleRemito").caption = "Observaciones: " & rto.detalle
    Set dsrDatosDespacho.DataSource = conectar.RSFactory("select 1")
    dsrDatosDespacho.PrintReport False

    Exit Function
err1:
    MsgBox Err.Description, vbCritical, "Error"
End Function


Public Function ImprimirRemito(IdRemito As Long) As Boolean
    Dim observaciones
    Dim nroCli
    Dim cli
    Dim direccion
    Dim Cuit
    Dim ivva
    Dim detalle
    Dim observaciones_cabecera
    Dim datos_entrega_footer
    Dim mes
    Dim dia
    Dim anio
    Dim client
    Dim localidad
    Dim fe As Date
    Dim ci
    On Error GoTo err91
    Dim tra As Boolean
    Dim contac
    Dim contacto
    Dim contacto1
    Dim contacto2
    tra = True
    Dim rs As New Recordset
    Dim ImprimirRemitoNuevo As Boolean
    Dim comp
    ImprimirRemitoNuevo = True
    Set rs = conectar.RSFactory("select r.numero, r.cantidad_bultos,r.idContacto,r.id,r.detalle AS detalleRro,r.fecha,r.estado,c.domicilio,c.ciudad,c.cuit,c.id AS idcliente,i.detalle,c.razon, r.observaciones_cabecera,r.datos_entrega_footer from remitos r inner join clientes c on r.idcliente=c.id inner join AdminConfigIVA i on i.idIVA=c.iva where r.id= " & IdRemito)
    'Printer.Orientation = 1

    Dim client2 As New clsCliente

    Dim cant_bultos As Integer
    Dim Obj As PageSet.PrinterControl
    Set Obj = New PrinterControl
    Obj.ChngOrientationPortrait


    Set client2 = DAOCliente.BuscarPorID(rs!idCliente)
    comp = "REMITO"

    Printer.CurrentY = 648
    Printer.CurrentX = 6800
    Printer.Font.Size = 12
    '    Printer.Print comp

    Printer.CurrentY = 600
    Printer.CurrentX = 8000
    Printer.Font.Size = 6
    Printer.Print "Control R-" & IdRemito & " | " & rs!numero


    If rs.EOF Or rs.BOF Then Exit Function
    fe = rs!FEcha
    nroCli = rs!idCliente
    cant_bultos = rs!cantidad_bultos
    cli = rs!razon
    direccion = rs!Domicilio

    Cuit = rs!Cuit
    ivva = rs!detalle
    detalle = rs!detalleRro
    Dim strsql As String
    'posiciono la fecha
    mes = Month(fe)
    dia = Day(fe)
    anio = Year(fe)
    'cli = Format(nroCli, "0000") & " - " & cli
    client = Format(nroCli, "0000") & " - " & cli

    observaciones_cabecera = rs!observaciones_cabecera
    datos_entrega_footer = rs!datos_entrega_footer

    Printer.Font.Size = 14
    Printer.Line (8800, 1400)-(10100, 1400)
    Printer.Line (8800, 1900)-(10100, 1900)
    Printer.CurrentY = 1500
    Printer.CurrentX = 8900
    Printer.Print Format(dia, "00") & "/" & Format(mes, "00") & "/" & Format(anio - 2000, "00")


    Printer.Font.Size = 10
    Printer.Font = "arial"
    'posiciono los datos del cliente

    Printer.CurrentY = 3800
    Printer.Font.Size = 11
    Printer.Print Tab(4);
    Printer.Print "Señor/es: ";
    Printer.FontBold = True
    Printer.Font.Size = 13
    Printer.Print truncar(client, 48)
    Printer.Font.Size = 11
    Printer.Print Tab(4);
    Printer.FontBold = False
    Printer.Print "I.V.A.: ";
    Printer.FontBold = True
    Printer.Print truncar(ivva, 50);
    Printer.Print Tab(65);
    Printer.FontBold = False
    Printer.Print "C.U.I.T.: ";
    Printer.FontBold = True
    Printer.Print truncar(Cuit, 50)
    Printer.Print Tab(4);
    Printer.FontBold = False
    Printer.Print "Domicilio: ";
    Printer.FontBold = True
    Printer.Print truncar(direccion, 50);
    Printer.Print Tab(65);
    Printer.FontBold = False
    Printer.Print "Provincia: ";
    Printer.FontBold = True
    Printer.Print truncar(client2.provincia.nombre, 30)
    Printer.FontBold = False

    Printer.Print Tab(70);
    Printer.FontBold = False
    Printer.Print "Localidad: ";
    Printer.FontBold = True
    Printer.Print truncar(client2.localidad.nombre, 30)
    Printer.FontBold = False
    Printer.Print Tab(4);
    Printer.Print "O/C: ";
    Printer.FontBold = True
    Printer.Print truncar(detalle, 80);
    Printer.FontBold = False
    Printer.Print Tab(4);
    Printer.Print "Observaciones: "
    '    Printer.Print truncar(observaciones_cabecera, 150)
    Printer.Print Tab(4);
    Printer.Font.Size = 11
    Printer.Print observaciones_cabecera
    'detalle y encabezado de detalle de la factura


    Printer.Font.Size = 9.5
    'detalle y encabezado de detalle de la factura
    Printer.CurrentY = 6800
    Printer.Print Tab(6);
    Printer.Print "Cant";
    Printer.Print Tab(15);
    Printer.Print "Item";

    Printer.Print Tab(22);
    Printer.Print "Origen";
    Printer.Print Tab(35);
    Printer.Print "Detalle";


    ci = 0
    Printer.CurrentY = 6900
    contac = 0
    If rs!idContacto > 0 Then
        contac = 1
        strsql = "select * from contactos where id=" & rs!idContacto
        Set rs = conectar.RSFactory(strsql)
        If Not rs.EOF And Not rs.BOF Then
            contacto = UCase(rs!nombre)
            contacto1 = UCase(rs!direccion) & " - " & UCase(rs!localidad)
            contacto2 = UCase(rs!detalle)
        End If
    End If


    strsql = "select e.observaciones,dp.item,e.origen,e.id,e.valor,e.idPedido,sum(e.cantidad) as cantidad,s.detalle,e.origen from entregas e,detalles_pedidos dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito =" & IdRemito & " and e.origen=1 group by e.id union all select e.observaciones,'000' as item,e.origen,e.id,e.valor,e.idPedido,sum(e.cantidad) as cantidad ,s.detalle,e.origen from entregas e,detallesPedidosEntregas dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito=" & IdRemito & " and e.origen=2 group by e.id union all select e.observaciones, '000' as item,e.facturado,e.id,e.valor,e.idPedido,e.cantidad,e.concepto as detalle,e.origen from entregas e  where  e.remito=" & IdRemito & " and (e.origen=3 or e.origen=4)"
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        ci = ci + 1
        Printer.Print Tab(6);
        Printer.Print Format(Math.Round(rs!Cantidad, 2), "0.00");
        Printer.Print Tab(15);
        Dim it
        Dim ite
        Dim ori
        it = rs!item
        If it = -1 Then
            ite = ci
        Else
            ite = it
        End If
        Printer.Print Format(ite, "000");
        Printer.Print Tab(22);
        If rs!Origen = 1 Then
            ori = "O/T "
        ElseIf rs!Origen = 2 Then
            ori = "O/E "
        ElseIf rs!Origen = 3 Then
            ori = "Concepto"
        ElseIf rs!Origen = 4 Then    'concepto aplicado
            ori = "OTA"
        End If

        If rs!Origen = 3 Then
            Printer.Print ori;
            Printer.Print Tab(35);
            Printer.Print rs!detalle;
        Else
            Printer.Print ori & Format(rs!idpedido, "0000");

            Printer.Print Tab(35);
            Printer.Print Format(rs!item, "000") & " ";
            Printer.Print rs!detalle   ';

        End If

        If LenB(rs!observaciones) > 0 Then
            Printer.Print Tab(35);
            Printer.Print rs!observaciones
        End If

        rs.MoveNext
    Wend
    Dim fuente

    If cant_bultos > 0 Then
        Printer.CurrentY = 13500
        fuente = Printer.Font.Size
        Printer.Font.Size = 12
        Printer.Print Tab(8); "Son " & cant_bultos & " bultos"
        Printer.Font.Size = fuente


    End If
    If contac = 1 Then

        Printer.CurrentY = 14301
        fuente = Printer.Font.Size
        Printer.Font.Size = 6
        '   Printer.Print Tab(20)

        Printer.Print Tab(8); contacto
        Printer.Print Tab(8); contacto1
        Printer.Print Tab(8); contacto2
        Printer.Font.Size = fuente
    End If


    Printer.CurrentY = 14500
    Printer.CurrentX = 900

    Dim texto_largo As String
    Dim caracteres_max As Integer

    texto_largo = datos_entrega_footer
    caracteres_max = 30
 
    
    If Len(texto_largo) > caracteres_max Then
        Printer.Print Tab(4);
        Printer.Font.Size = 10
        Printer.Print Left(texto_largo, caracteres_max)
        Printer.Print Tab(4);
        Printer.Font.Size = 10
        Printer.Print Mid(texto_largo, caracteres_max + 1) & vbCrLf
    Else
        Printer.Print Tab(4);
        Printer.Font.Size = 10
        Printer.Print texto_largo
    End If

    Printer.EndDoc
    tra = False
    conectar.execute "update remitos set impreso=impreso+1 where id=" & IdRemito
    Exit Function
err91:
    If tra Then ImprimirRemitoNuevo = False
    MsgBox Err.Description
End Function
