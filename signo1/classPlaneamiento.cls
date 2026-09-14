VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classPlaneamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim claseS As New classStock
Dim idPiezaFabriacada As Long
Dim cuit


Public Property Get TipoOrden() As TipoOt
    TipoOrden = rs!tipo_orden
End Property
Public Property Get IdMoneda() As Integer
    IdMoneda = rs!IdMoneda
End Property
Public Property Get FechaEntrega() As Date
    FechaEntrega = rs!FechaEntrega
End Property
Public Property Get FechaFin() As Date
    If IsNull(rs!FechaCerrado) Then
        FechaFin = 0
    Else
        FechaFin = rs!FechaCerrado
    End If
End Property
Public Property Get Fabricados() As Long
    Fabricados = rs!cantidad_fabricados
End Property
Public Property Get pedidos() As Long
    pedidos = rs!Cantidad
End Property
Public Property Get reservados() As Long
    reservados = rs!reserva_stock
End Property
Public Property Get idPiezaFabricada() As Long
    idPiezaFabriacada = rs!idPieza
End Property
Public Property Get descripcion() As String
    descripcion = rs!descripcion
End Property
Public Property Get estadoOT() As Integer
    estadoOT = rs!estado
End Property
Public Property Get cantidadItems() As Long
    cantidadItems = rs!cantidadItems
End Property
Public Property Get cliente() As String
    cliente = rs!cliente
End Property
Public Property Get count() As Long
    count = rs!count
End Property

Public Property Get idCliente() As Long
    idCliente = rs!idCliente
End Property

Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion    ' New ADODB.Connection


    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutar_consulta(consulta As String)
    On Error GoTo err1
    Set rs = conectar.RSFactory(consulta)
    Exit Function
err1:
    MsgBox "Se produjo un error: " & Err.Description
End Function
Public Function ejecutar_consulta2(consulta As String)
    On Error GoTo err2
    Set rs2 = conectar.RSFactory(consulta)
    Exit Function
err2:
    MsgBox "Se produjo un error: " & Err.Description
End Function


Public Function ExistePedido(IdPedido As Long) As Integer
    Me.ejecutar_consulta "Select count(id) as canti from pedidos where id=" & IdPedido
    If rs!canti = 1 Then
        ExistePedido = 1
    ElseIf rs!canti = 0 Then
        ExistePedido = 0
    Else
        ExistePedido = -1
    End If
End Function

Public Function informePedidoContaduria(Id As Integer, DIALOGO As Boolean, Optional t1 As OrdenTrabajo) As Boolean     '1- enviar 2-imprimir

    Dim T As OrdenTrabajo
    Set T = DAOOrdenTrabajo.FindById(t1.Id)

    Dim tot As Double
    Dim subto As Double
    Dim rInci As Recordset
    On Error GoTo err2
    informePedidoContaduria = True
    Dim strsql As String, strsql2 As String
    strsql = "select p.id_ot_padre,c.cuit,p.idMoneda,c.razon,p.dto,p.id,p.descripcion,p.fechaEntrega as fe, p.fechaCreado, p.nroPresupuesto from pedidos p, clientes c where p.idcliente=c.id and p.id=" & Id
    Me.ejecutar_consulta strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
    End If
    If c = 1 Then
        strsql = "select count(id) as inci from Incidencias where idReferencia=" & Id & " and origen=2"
        Set rInci = conectar.RSFactory(strsql)
        If Not rs.EOF And Not rs.BOF Then
            If rInci!inci > 0 Then
                pedido_contaduria.Sections("pie").Controls("lblIncidencias").caption = "Existen " & rInci!inci & " incidencias para esta OT. Controlar en primera instancia"
            Else
                pedido_contaduria.Sections("pie").Controls("lblIncidencias").caption = "No hay incidencias al momento de emitir la OT, controlar periodicamente."
            End If
        End If


        pedido_contaduria.Sections("pie").Controls("lblNoAprobado").Visible = (T.estado = EstadoOT_Desactivado) Or (T.estado = EstadoOT_Pendiente)

        Dim EsMarco As Boolean
        Dim contrato_id As Long
        Dim DependeDeMarco As Boolean
        nroOT = Format("00000", rs!Id)
        ref = rs!descripcion

        DependeDeMarco = (rs!id_ot_padre > 0)
        Monedas = funciones.queMoneda(rs!IdMoneda)
        EsMarco = (rs!id_ot_padre = -1)
        contrato_id = rs!id_ot_padre

        If DependeDeMarco Then
            ref = ref & " (OT Contrato Abierto " & contrato_id & ")"
        End If

        razon = rs!razon
        entrega = rs!fe
        Creado = rs!fechaCreado
        presu = rs!NroPresupuesto
        cuit = rs!cuit
        Dim tit As String

        If EsMarco Then
            tit = "Contrato Abierto Nº: "
        Else
            tit = "Orden de Trabajo Nº: "
        End If


        pedido_contaduria.Sections("cabeza").Controls("lblOT").caption = tit & T.Id
        pedido_contaduria.Sections("cabeza").Controls("lblCliente").caption = razon
        pedido_contaduria.Sections("cabeza").Controls("lblReferencia").caption = ref
        pedido_contaduria.Sections("cabeza").Controls("lblfechaEntrega").caption = entrega
        pedido_contaduria.Sections("cabeza").Controls("lblClienteFacturar").caption = T.ClienteFacturar.razon
        dto = rs!dto
        dtos = 1 - (dto / 100)
        strsql = "select dp.descuento,dp.item, dp.cantidad as cantidad, dp.precio as unitario, dp.precio,(dp.cantidad *dp.precio*(1-(dp.descuento/100))) as totaldto,(dp.precio*(1-(dp.descuento/100))) as stotaldto,(dp.precio*dp.cantidad) as stotal, s.detalle,dp.nota from detalles_pedidos dp, stock s where dp.idPieza=s.id and dp.idPedido=" & nroOT
        Me.ejecutar_consulta strsql
        TotFab = 0
        TotRes = 0
        TotIt = 0
        c = 0
        Dim subtodto As Double
        subtodto = 0
        While Not rs.EOF
            subto = subto + (rs!Precio * rs!Cantidad)

            subtodto = subtodto + funciones.RedondearDecimales(rs!Precio * rs!Cantidad * (1 - (rs!Descuento / 100)))
            '            'debug.print (funciones.RedondearDecimales(rs!stotaldto) * rs!Cantidad)
            c = c + 1
            rs.MoveNext
        Wend
        tot = subto * dtos
        tot = funciones.RedondearDecimales(tot, 2)

        pedido_contaduria.Sections("pie").Controls("lblSubTotal").caption = Monedas & " " & funciones.RedondearDecimales(subto, 2)
        pedido_contaduria.Sections("pie").Controls("lbldto").caption = dto & "%"
        pedido_contaduria.Sections("pie").Controls("lblTotal").caption = Monedas & " " & funciones.RedondearDecimales(subtodto, 2)
        pedido_contaduria.Sections("s3").Controls("lblbarcode").caption = "*" & Format(Id, "0000") & "*"




        pedido_contaduria.Sections("pie").Controls("lblAnticipo").caption = T.CondicionesComercialesAnticipo    ' fp_anticipo
        pedido_contaduria.Sections("pie").Controls("lblSaldo").caption = T.CondicionesComercialesSaldo



        anticipo_Facturado = "Anticipo aún no facturado."



        If T.EsHija Then

            strEsHija = " OT marco " & T.OTMarcoIdPadre
            Set T = DAOOrdenTrabajo.FindById(T.OTMarcoIdPadre)

        End If



        If T.AnticipoFacturado Then
            Dim F As New Factura
            Set F = DAOFactura.FindById(T.AnticipoFacturadoIdFactura)
            anticipo_Facturado = "Anticipo facturado.  Comprobante: " & F.GetShortDescription(False, True)

            If LenB(strEsHija) Then
                antipo_facturado = anticipo_Facturado & strEsHija
            End If
        End If
        If T.Anticipo = 0 Then anticipo_Facturado = "Sin Anticipo"

        pedido_contaduria.Sections("pie").Controls("lblAnticipoFacturado").caption = anticipo_Facturado


        If T.Anticipo <= 0 Then
            pedido_contaduria.Sections("pie").Controls("lblAnticipo").Visible = False
            pedido_contaduria.Sections("pie").Controls("lblAnticipo_header").Visible = False
        Else
            pedido_contaduria.Sections("pie").Controls("lblAnticipo").Visible = True
            pedido_contaduria.Sections("pie").Controls("lblAnticipo_header").Visible = True
        End If
        pedido_contaduria.Sections("pie").Controls("lblSaldo").Visible = True
        pedido_contaduria.Sections("pie").Controls("lblSaldo_header").Visible = True
        If c > 0 Then rs.MoveFirst
        Set pedido_contaduria.DataSource = rs

        '22/01/2021 DNEMER
        'Comento el envio de la impresión directa para que primero se muestre la vista previa de la hoja de Contaduria.
        'Entonces ahora va a mostrar la vista previa desde donde se puede enviar o no a IMPRIMIR

        'pedido_contaduria.PrintReport False

        pedido_contaduria.Show 1

        If Not DAOOrdenTrabajoHistorial.agregar(T, "Caratula de Contaduria previsualizado") Then
            GoTo err2
        End If

    End If
    Exit Function
err2:
    MsgBox Err.Description

    Resume Next
    informePedidoContaduria = False
End Function


Public Function RealizarEntrega(modo, Remito As Long, Optional CantAEntregar As Double, Optional idDetallePedido As Long, Optional IdPedido As Long, Optional otoe = 1, Optional vector, Optional valor2, Optional deta As DetalleOrdenTrabajo, Optional mostrar_obs As Boolean = False, Optional ot_header As String = vbNullString) As Boolean
    Dim rto As Remito
    Set rto = DAORemitoS.FindById(CLng(Remito))
    Dim Nota As String

    'modos
    'modo 1: solo una pieza
    'modo 2: todo el pedido


    'remito = DAORemitoS.FindById(remito).id

    Dim aentregar As Double
    Dim tra As Boolean
    RealizarEntrega = True
    On Error GoTo err32
    If otoe = 1 Then Origen = 1    'la entrega se hizo desde una OT
    If otoe = 2 Then Origen = 2    'la entrga se hizo desde una OE


    If modo = 1 Then
        cn.BeginTrans
        tra = True
        fec = funciones.datetimeFormateada(Now)
        'cambio la cant entregada

        If otoe = 1 Then Me.ejecutar_consulta "select cantidad_entregada as entregados,cantidad,precio as valor from detalles_pedidos where id=" & idDetallePedido
        If otoe = 2 Then Me.ejecutar_consulta "select entregados,cantidad,vale as valor from detallesPedidosEntregas where id=" & idDetallePedido


        If otoe = 1 Then Me.ejecutar_consulta "select cantidad_entregada as entregados,cantidad,precio as valor from detalles_pedidos where id=" & idDetallePedido


        valor = rs!valor
        Cantidad = deta.CantidadPedida   'rs!Cantidad
        Entregados = deta.CantidadEntregada  'rs!Entregados
        If Cantidad - Entregados >= CantAEntregar Then
            If otoe = 1 Then

                '                cn.execute "update detalles_pedidos set cantidad_entregada=cantidad_entregada+" & CantAEntregar & " Where id=" & idDetallePedido
                '                DAODetalleOrdenTrabajo.SaveCantidad idDetallePedido, CantAEntregar, CantidadEntregada_, 0

            End If

            If otoe = 2 Then cn.execute "update detallesPedidosEntregas set entregados=entregados+" & CantAEntregar & " Where id=" & idDetallePedido


            If rto.CantidadDeLineasActuales >= funciones.itemsPorRemito Then
                MsgBox "El remito posee más elementos de los permitidos."
            Else
                'creo el registro de entrega

                If LenB(ot_header) > 0 Then
                    Nota = ot_header & " | " & deta.Nota
                Else
                    Nota = deta.Nota
                End If

                If mostrar_obs Then
                    cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor,observaciones) values (" & idDetallePedido & "," & CantAEntregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ",'" & UCase(Nota) & "')"
                Else
                    cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor) values (" & idDetallePedido & "," & CantAEntregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ")"
                End If
            End If
        Else
            MsgBox "No puede entregar mas de lo que tiene en la orden!", vbCritical, "Información"
        End If
        tra = False
        cn.CommitTrans


    ElseIf modo = 2 Then
        tra = True
        cn.BeginTrans

        'cambio la cant entregada por el total solicitado
        'busco dato para crear el registro de entregas
        If otoe = 1 Then
            Me.ejecutar_consulta "select nota,precio as valor,id,cantidad as cantAEntregar, cantidad_entregada as entregados from detalles_pedidos where idPedido=" & IdPedido
            '      cn.execute "update detalles_pedidos set cantidad_entregada=(cantidad-cantidad_entregada)+cantidad_entregada Where idPedido=" & idpedido
            Dim col As New Collection
            Set col = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(IdPedido)
            Dim det As DetalleOrdenTrabajo

        Else
            Me.ejecutar_consulta "select precio as valor,id,cantidad as entregada,entregados from detallePedidosEntrega where idPedido=" & IdPedido
            '      cn.execute "update detallesPedidosEntregas set entregados=entregados+" & CantAEntregar & " Where id=" & idDetallePedido
        End If
        'actualizo el detalle_pedido con el nro de rto para las entregas
        fec = funciones.datetimeFormateada(Now)
        While Not rs.EOF
            CantAEntregar = rs!CantAEntregar
            Nota = rs!Nota


            If LenB(ot_header) > 0 Then
                Nota = ot_header & " | " & rs!Nota
            End If

            Entregados = rs!Entregados
            If CantAEntregar > Entregados Then
                valor = rs!valor
                If mostrar_obs Then
                    cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor,observaciones) values (" & idDetallePedido & "," & CantAEntregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ",'" & deta.Nota & "')"
                Else
                    cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor) values (" & rs!Id & "," & CantAEntregar - Entregados & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ")"
                End If
            End If
            rs.MoveNext
        Wend
        tra = False
        cn.CommitTrans
    ElseIf modo = 3 Then    'entrega multiple pero no total
        'tengo que comprobar que no haya problemas!!!!!!!!!!!!
        fec = funciones.datetimeFormateada(Now)
        tra = True
        cn.BeginTrans

        cantITems = UBound(vector) + 1



        'Set rs3 = conectar.RSFactory("select count(id) as cr from entregas where remito=" & CLng(Remito))


        If rto.CantidadDeLineasActuales >= funciones.itemsPorRemito Then
            MsgBox "El remito posee más elementos de los permitidos. Por favor, anulelo"
            RealizarEntrega = False
            cn.RollbackTrans
            Exit Function
        Else
            If rto.CantidadDeLineasActuales + CLng(cantITems) > funciones.itemsPorRemito Then
                MsgBox "Remitando estos elementos, el remito superaria lo permitido"
                RealizarEntrega = False
                cn.RollbackTrans
                Exit Function
            End If
        End If
        Dim r As Recordset

        For i = 0 To UBound(vector)
            idDetallePedido = vector(i)
            'cambio la cant entregada por el total solicitado
            'busco dato para crear el registro de entregas
            If otoe = 1 Then
                Me.ejecutar_consulta "select nota,precio,id,cantidad, cantidad_entregada, cantidad_fabricados from detalles_pedidos where idPedido=" & IdPedido & " and id=" & idDetallePedido
                Nota = rs!Nota
                If LenB(ot_header) > 0 Then
                    Nota = ot_header & " | " & rs!Nota
                End If
                valor = rs!Precio
                Cantidad = rs!Cantidad
                ' Cantidad_Entregada = rs!Cantidad_Entregada

                Cantidad_Entregada = DAODetalleOrdenTrabajo.GetCantidad(idDetallePedido, CantidadEntregada_)


                cantidad_fabricados = DAODetalleOrdenTrabajo.GetCantidad(idDetallePedido, CantidadFabricada_)

                If cantidad_fabricados > Cantidad Then
                    aentregar = Cantidad - Cantidad_Entregada
                Else
                    aentregar = cantidad_fabricados - Cantidad_Entregada    'cantidad - (cantidad_entregada) - cantidad_fabricados

                End If




                If Cantidad > Cantidad_Entregada Then    'hay pendientes de entregar
                    If cantidad_fabricados > Cantidad_Entregada Then  'hay productos disponibles para entregar
                        '                  cn.execute "update detalles_pedidos set cantidad_entregada=cantidad_entregada+" & aentregar & " Where idPedido=" & idpedido & " and id=" & idDetallePedido
                        '                  DAODetalleOrdenTrabajo.SaveCantidad idDetallePedido, aentregar, CantidadEntregada_, 0
                        If mostrar_obs Then
                            cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor,observaciones) values (" & idDetallePedido & "," & aentregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ",'" & UCase(Nota) & "')"
                        Else
                            cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor) values (" & rs!Id & "," & aentregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ")"
                        End If

                    End If
                End If

            Else

                'si la cantidad pedida menos la cantidad entregada es mayor a lo que piden, procedo c la entrega
                Me.ejecutar_consulta "select vale,id,cantidad, entregados from detallesPedidosEntregas where idPedidoEntrega=" & IdPedido & " and id=" & idDetallePedido

                If Not rs.EOF And Not rs.BOF Then

                    valor = rs!vale
                    Cantidad = rs!Cantidad
                    Cantidad_Entregada = rs!Entregados
                    If Cantidad > Cantidad_Entregada Then
                        aentregar = Cantidad - Cantidad_Entregada
                    End If
                    If aentregar > 0 Then
                        cn.execute "update detallesPedidosEntregas set entregados=entregados+" & aentregar & " Where id=" & idDetallePedido

                        If mostrar_obs Then
                            cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor,observaciones) values (" & idDetallePedido & "," & CantAEntregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ",'" & UCase(Nota) & "')"
                        Else
                            cn.execute "insert into entregas (idDetallePedido, cantidad, Remito, fecha,origen,idPedido,valor) values (" & rs!Id & "," & aentregar & "," & CLng(Remito) & ",'" & fec & "'," & Origen & "," & IdPedido & "," & valor & ")"
                        End If


                        'Me.ejecutar_consulta "select vale,id,cantidad as entregada,entregados from detallesPedidosEntrega where idPedido=" & idpedido & " and id= " & idDetallePedido
                    End If
                End If
            End If
        Next i
        tra = False
        cn.CommitTrans
    End If
    Exit Function
err32:
    cn.RollbackTrans
    Resume
    RealizarEntrega = False
End Function


Private Sub Class_Terminate()

    Set rs = Nothing
    Set cn = Nothing
    Set rs2 = Nothing

End Sub


Public Function estaTodoFabricado(IdPedido As Long) As Boolean
'si está todo fabricado,
'cambio el estado del pedido.
    Dim esta As Boolean
    Me.ejecutar_consulta "select dp.retirado,dp.cantidad,dp.cantidad_fabricados,dp.reserva_stock from detalles_pedidos dp, pedidos p where dp.idPedido=p.id and p.id=" & IdPedido
    esta = True
    estaTodoFabricado = esta
    While Not rs.EOF
        fabricado = rs!cantidad_fabricados
        stock = rs!reserva_stock
        Cantidad = rs!Cantidad
        If fabricado + stock < Cantidad Then
            esta = False
            estaTodoFabricado = esta
            Exit Function
        End If
        rs.MoveNext
    Wend
    estaTodoFabricado = esta
    'If esta Then Me.cambiar_estado idpedido, 3 -> ver
    Exit Function
    estaTodoFabricado = False
End Function



Public Function imprimirOrdenEntrega(nroOE As Long) As Boolean
    On Error GoTo err2
    imprimirOrdenEntrega = True
    Dim strsql As String, strsql2 As String
    strsql = "select c.domicilio,c.cuit,c.razon,p.referencia,p.fecha,p.fechaCreado from PedidosEntregas p inner join clientes c on p.idCliente=c.id where p.id=" & nroOE
    Me.ejecutar_consulta strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c > 0 Then
        rs.MoveFirst
        ref = rs!referencia
        fecEnt = rs!FEcha
        fecCre = rs!fechaCreado
        razon = rs!razon
        dire = rs!Domicilio
        cuit = rs!cuit
        ordenEntrega.Sections("cabeza").Controls("lblOT").caption = nroOE
        ordenEntrega.Sections("cabeza").Controls("lblCliente").caption = razon
        ordenEntrega.Sections("cabeza").Controls("lblCliente2").caption = razon
        ordenEntrega.Sections("cabeza").Controls("lblfechaEntrega").caption = fecEnt
        ordenEntrega.Sections("cabeza").Controls("lblReferencia").caption = ref

    End If
    strsql2 = "select pe.id ,s.detalle, pe.cantidad,pe.vale from detallesPedidosEntregas pe inner join stock s on pe.idPieza = s.id where pe.IdPedidoEntrega=" & nroOE
    Me.ejecutar_consulta2 strsql2
    tt = 0
    Set ordenEntrega.DataSource = rs2

    ordenEntrega.PrintReport 1
    Set ordenEntrega.DataSource = Nothing
    Exit Function
err2:
    MsgBox Err.Description
    imprimirOrdenEntrega = False
End Function

Public Function verificarOrdenEntrega(IDOE As Long, Optional CambiarEstado As Boolean = False) As Boolean
    On Error GoTo errh23
    verificarOrdenEntrega = True
    Me.ejecutar_consulta "select cantidad,entregados from detallesPedidosEntregas where idPedidoEntrega=" & IDOE
    While Not rs.EOF
        If rs!Entregados < rs!Cantidad Then
            'si algún item tiene menos entregados que la cant. solicitada
            'devuelvo false y salgo de la función
            verificarOrdenEntrega = False
            Exit Function
        End If
        rs.MoveNext
    Wend

    If CambiarEstado Then
        'estado 2: entrega completa. PROCESO FINALIZADO.
        cn.execute "update PedidosEntregas set estado=3 where id=" & IDOE
    End If
    Exit Function
errh23:
    verificarOrdenEntrega = False
End Function


Public Function informePieza2(idPiezaPedido, DIALOGO As Boolean, Optional conjunto = False, Optional idPiezaHija = -1, Optional pos = 1, Optional CantElemConj = 0, Optional PosConj = 0, Optional cantEnConj = 1, Optional detalle = Empty, Optional idPiezaPedidoConjunto = -1) As Boolean    '1- enviar 2-imprimir
    On Error GoTo err2
    informePieza2 = True
    Dim strsql As String, strsql2 As String


    If Not conjunto Then
        strsql = "select dp.impresiones_ruta as copias, dp.nota_produccion,dp.reserva_stock,dp.fechaEntrega as ei,dp.cantidad,dp.item,s.id as pizza,s.detalle,dp.nota ,c.razon,p.id,p.descripcion,p.fechaEntrega from stock s,detalles_pedidos dp, pedidos p,clientes c where p.idcliente=c.id and s.id=dp.idpieza and dp.id=" & idPiezaPedido & " and p.id=dp.idpedido"
    Else
        strsql = "select dp.impresiones_ruta as copias,dp.nota_produccion,dp.reserva_stock,dp.fechaEntrega as ei, dp.cantidad,dp.item,s.id as pizza,c.razon,p.id,p.descripcion,p.fechaEntrega from stock s,detalles_pedidos dp, pedidos p,clientes c where p.idcliente=c.id and s.id=dp.idpieza and dp.id=" & idPiezaPedido & " and p.id=dp.idpedido"

    End If
    Me.ejecutar_consulta strsql
    copia = rs!COPIAS
    If conjunto Then
        idPiezaPadre = rs!pizza
    End If
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c > 0 Then
        rs.MoveFirst
    End If

    While Not rs.EOF
        nroOT = Format("0000", rs!Id)
        ref = rs!descripcion
        eni = rs!ei
        fecEnt = rs!FechaEntrega
        nota_produccion = UCase(rs!nota_produccion)
        razon = rs!razon
        item = rs!item
        If conjunto Then
            Pieza = detalle
            Nota = emtpy
        Else
            Pieza = rs!detalle
            Nota = rs!Nota
            Pieza = Trim(Pieza) & " " & Trim(Nota)

        End If
        idPieza = rs!pizza
        Cantidad = rs!Cantidad
        Cantidad = rs!Cantidad - rs!reserva_stock


        Me.ejecutar_consulta2 "select count(id) as cantITEMS from detalles_pedidos where idPedido=" & nroOT

        Items = nroOT & "." & pos & "/" & rs2!cantITems


        If conjunto Then
            barcode = "." & idPiezaPedido & "." & idPiezaPedidoConjunto
        Else
            barcode = "." & idPiezaPedido & ".0000"
        End If

        pedido_pieza2.Sections("cabeza").Controls("lblOT").caption = Items
        pedido_pieza2.Sections("cabeza").Controls("barcode").caption = "*" & Format(idPiezaPedido, "00000000") & "*"
        pedido_pieza2.Sections("cabeza").Controls("lblCliente").caption = razon
        pedido_pieza2.Sections("cabeza").Controls("lblReferencia").caption = ref
        pedido_pieza2.Sections("observar").Controls("lblNota").caption = IIf(IsNull(Nota), vbNullString, Nota)
        pedido_pieza2.Sections("cabeza").Controls("lblfechaentrega").caption = fecEnt
        pedido_pieza2.Sections("cabeza").Controls("entregait").caption = eni
        pedido_pieza2.Sections("cabeza").Controls("lblDetalle").caption = Pieza
        pedido_pieza2.Sections("observar").Controls("copia").caption = copia

        If Cantidad = 0 Then
            pedido_pieza2.Sections("cabeza").Controls("lbldeStock").caption = "De Stock"
        Else
            pedido_pieza2.Sections("cabeza").Controls("lbldeStock").caption = Empty
        End If


        pedido_pieza2.Sections("cabeza").Controls("lblItem").caption = "Item " & item
        Cantidad = Cantidad * cantEnConj

        If conjunto Then
            Cantidad = cantEnConj
            Me.ejecutar_consulta2 "select detalle as det from stock where id=" & idPiezaHija
            While Not rs2.EOF
                masdet = rs2!det
                rs2.MoveNext
            Wend
            pedido_pieza2.Sections("cabeza").Controls("lblMasDetalle").caption = "Elemento: " & PosConj & "/" & CantElemConj & " " & masdet

            'Exit Function
        End If
        pedido_pieza2.Sections("cabeza").Controls("lblCantidad").caption = Cantidad
        pedido_pieza2.Sections("cabeza").Controls("lblMasDetalle").caption = Empty
        pedido_pieza2.Sections("observar").Controls("lblnota").caption = IIf(IsNull(nota_produccion), vbNullString, nota_produccion)



        Set rs2 = conectar.RSFactory("select count(id) as c from sp_archivos.archivos where origen=1 and idPieza=" & idPieza)
        If Not rs2.EOF And Not rs2.BOF Then
            Cantidad = rs2!c
        End If

        Set rs2 = conectar.RSFactory("select count(id) as c from Incidencias where origen=3 and idReferencia=" & idPieza)
        If Not rs2.EOF And Not rs2.BOF Then
            cantidad_inci = rs2!c
        End If



        If Cantidad = 1 Then
            archivos = "Existe " & Cantidad & " archivo asociado. Por favor Verificar"
        ElseIf Cantidad > 1 Then
            archivos = "Existen " & Cantidad & " archivos asociados. Por favor Verificar"
        Else
            archivos = "No existen archivos asociados. Controlar periodicamente"
        End If

        If cantidad_inci = 1 Then
            inci = "Existe 1 incidencia. Por favor Verificar"
        ElseIf Cantidad > 1 Then
            inci = "Existen " & cantidad_inci & " Incidencias. Por favor Verificar"
        Else
            inci = "No existen incidencias. Controlar periodicamente"
        End If

        pedido_pieza2.Sections("pie").Controls("archivos_asociados").caption = archivos
        pedido_pieza2.Sections("pie").Controls("incidencias").caption = inci


        If Not conjunto Then

            strsql2 = "SELECT concat('*00',t.id,'" & barcode & "','*') as codigo, t.cantxproc,t.tarea,s.sector from tareas t,sectores s,PlaneamientoTiemposProcesos dm where dm.codigoTarea=t.id and s.id=t.id_sector and dm.idDetallePedido=" & idPiezaPedido
        Else

            strsql2 = "SELECT concat('*00',t.id,'" & barcode & "','*') as codigo, t.cantxproc,t.tarea,s.sector from tareas t,sectores s,PlaneamientoTiemposProcesos dm where dm.codigoTarea=t.id and s.id=t.id_sector and dm.idDetallePedido=" & idPiezaPedidoConjunto
        End If
        Me.ejecutar_consulta2 strsql2

        Set pedido_pieza2.DataSource = rs2
        pedido_pieza2.PrintReport False
        'pedido_pieza2.Show 1
        rs.MoveNext
    Wend
    rs.Close
    rs2.Close
    Exit Function
err2:
    MsgBox Err.Description
    informePieza2 = False
End Function


Public Function generarOE(FEcha As Date, referencia As String, idCliente As Long, ByRef nroOE) As Boolean
    Dim ok As Boolean, tra As Boolean, locka As Boolean
    On Error GoTo errh2

    generarOE = True

    'ahora si lockeo la tabla
    cn.BeginTrans
    tra = True
    'CN.Execute "LOCK TABLES stock WRITE, PedidosEntregas WRITE, detallesPedidosEntregas WRITE"
    locka = True
    fec = funciones.datetimeFormateada(FEcha)
    usu = funciones.getUser
    fecNow = funciones.datetimeFormateada(Now)
    'inserto el principal de la orden de entrega
    cn.execute "insert into PedidosEntregas (referencia,fecha,usuario,idCliente,fechaCreado) values ('" & referencia & "', '" & fec & "'," & usu & "," & idCliente & ",'" & fecNow & "')"
    'obtengo el id del pedido creado
    'MsgBox "insertado el dato"
    Me.ejecutar_consulta "select last_insert_id() as ultid from PedidosEntregas"
    'MsgBox "selecionado el id"
    ultid = rs!ultid
    nroOE = ultid

    locka = False
    'CN.Execute "UNLOCK TABLES"
    tra = False
    cn.CommitTrans
    'Set rs = Nothing
    Exit Function
errh2:
    generarOE = False
    MsgBox Err.Description

    If tra Then cn.RollbackTrans
    'If locka Then CN.Execute "UNLOCK TABLES"
End Function


Public Function editarOE(lst As ListView, FEcha As Date, referencia As String, idCliente As Long, nroOE, IdMoneda As Integer) As Boolean
    Dim ok As Boolean, tra As Boolean, locka As Boolean
    On Error GoTo errh2
    Dim l
    ok = True
    editarOE = True
    'rutina que lee los productos que componen la O/E con sus cantidades
    'comparo contra el stock real. Y si esta todo ok para generar el remito
    'lockeo la tabla para operar.
    For l = 1 To lst.ListItems.count
        ids = CLng(lst.ListItems(l).Tag)
        Me.ejecutar_consulta "select cantidad from stock where id=" & ids
        If rs!Cantidad < CLng(lst.ListItems(l).ListSubItems(1)) Then
            ok = False
        End If
    Next l

    If ok Then
        'ahora si lockeo la tabla
        cn.BeginTrans
        tra = True
        'CN.Execute "LOCK TABLES stock WRITE, PedidosEntregas WRITE, detallesPedidosEntregas WRITE"
        locka = True
        fec = funciones.datetimeFormateada(FEcha)
        usu = funciones.getUser
        fecNow = funciones.datetimeFormateada(Now)
        'inserto el principal de la orden de entrega
        'CN.Execute "insert into PedidosEntregas (referencia,fecha,usuario,idCliente,fechaCreado) values ('" & referencia & "', '" & fec & "'," & usu & "," & idCliente & ",'" & fecNow & "')"
        cn.execute "update PedidosEntregas  set referencia='" & referencia & "', fecha='" & fec & "',usuario=" & usu & ", idCliente=" & idCliente & ",idMoneda=" & IdMoneda & " where id=" & nroOE



        ultid = nroOE

        'elimino los datos anteriores, para dejar los nuevos
        cn.execute "delete from detallesPedidosEntregas where idPedidoEntrega=" & nroOE
        For l = 1 To lst.ListItems.count
            iDst = CLng(lst.ListItems(l).Tag)
            Cantidad = CLng(lst.ListItems(l).ListSubItems(1))
            valor = CDbl(lst.ListItems(l).ListSubItems(2))

            'actualizo el stock
            'CN.Execute "update stock set cantidad=cantidad-" & CLng(lst.ListItems(l).ListSubItems(2)) & " where id=" & idst
            'creo el contenido de la O/E
            cn.execute "insert into detallesPedidosEntregas (idPedidoEntrega,idPieza,cantidad,vale) values (" & nroOE & "," & iDst & "," & Cantidad & "," & valor & ")"
        Next l
        locka = False
        'CN.Execute "UNLOCK TABLES"
        tra = False
        cn.CommitTrans
        'Set rs = Nothing
    Else
        editarOE = False
    End If
    Exit Function
errh2:
    editarOE = False
    If tra Then cn.RollbackTrans
    'If locka Then CN.Execute "UNLOCK TABLES"
End Function

Public Function ejecutarComando(comando) As Boolean
    On Error GoTo err45
    ejecutarComando = True
    cn.execute comando
    Exit Function
err45:
    ejecutarComando = False
    MsgBox Err.Description
End Function


Public Function precio_pieza2(idPieza, ByRef IdMoneda) As Double
    Me.ejecutar_consulta "select id_moneda_precio, precio_definido from stock where id=" & idPieza

    If Not rs.EOF And Not rs.BOF Then
        IdMoneda = rs!id_moneda_precio
        prec_def = rs!precio_definido
    End If
    If prec_def > 0 Then
        'si hay un precio definido, utilizo este
        precio_pieza = prec_def

    Else
        'si no hay un precio definido busco el máximo entre cotizados y fabricados
        'fabricados
        Me.ejecutar_consulta "Select max(precio) as maximo from detalles_pedidos where idPieza=" & idPieza
        If Not rs.EOF And Not rs.BOF Then
            maxiot = rs!MAXIMO
        Else
            maxiot = 0
        End If
        'cotizados
        Me.ejecutar_consulta "Select max(valorUnitarioManual) as maximo from detalle_presupuesto where idPieza=" & idPieza
        If Not rs.EOF And Not rs.BOF Then
            maxipres = rs!MAXIMO
        Else
            maxipres = 0
        End If

        If maxipres > 0 And maxiot > 0 Then

            If maxipres > maxiot Then
                precio_pieza = maxipres
            ElseIf maxipres < maxiot Then
                precio_pieza = maxiot
            ElseIf maxipres = maxiot Then
                precio_pieza = maxipres
            End If
        ElseIf maxipres <= 0 Then
            maxiot = 0
            precio_pieza = maxiot
        ElseIf maxiot <= 0 Then
            maxipres = 0
            precio_pieza = maxipres
        End If

    End If
End Function

Public Function CantidadFabricada(idPieza As Long, Optional cantidadNueva As Double = 0) As Double
    Me.ejecutar_consulta "Select sum(cantidad) as c from detalles_pedidos where idPieza=" & idPieza
    If Not rs.EOF And Not rs.BOF Then
        c = rs!c
        If IsNull(rs!c) Then c = 0
        CantidadFabricada = c + cantidadNueva
    Else
        CantidadFabricada = -1
    End If
End Function

Public Function aplicarRemitoAOT(IdPedido As Long, idEntrega As Long, IdDetalle As Long, Cantidad As Double) As Boolean
    On Error GoTo err3
    erro = True
    cn.BeginTrans

    Dim rto As remitoDetalle
    Set rto = DAORemitoSDetalle.FindById(idEntrega)

    cn.execute "update entregas set idPedido=" & IdPedido & ",idDetallePedido=" & IdDetalle & ",origen=4 where id=" & idEntrega

    cn.execute "update detalles_pedidos set cantidad_entregada=cantidad_entregada+" & Cantidad & " where id=" & IdDetalle

    If Not DAODetalleOrdenTrabajo.SaveCantidad(IdDetalle, Cantidad, CantidadEntregada_, 0, rto.Remito, 0, 0, 0) Then GoTo err3

    cn.CommitTrans

    aplicarRemitoAOT = erro

    Exit Function
err3:
    erro = False
    aplicarRemitoAOT = erro
    cn.RollbackTrans
End Function
Public Function definirPrecios(idPieza, valor, IdMoneda) As Boolean
    On Error GoTo err4
    definirPrecios = True
    cn.execute "update stock set precio_definido=" & valor & ",fecha_precio_definido='" & funciones.dateFormateada(Now) & "',id_moneda_precio=" & IdMoneda & " where id=" & idPieza
    Exit Function
err4:
    definirPrecios = False
End Function


Public Function procesos_definidos(IdPedido) As Boolean
    procesos_definidos = True
    Dim rs1 As Recordset
    Set rs1 = conectar.RSFactory("Select procesos_definidos from detalles_pedidos where idPedido=" & IdPedido)
    While Not rs1.EOF And Not rs1.BOF
        If rs1!procesos_definidos = 0 Then
            procesos_definidos = False
            Exit Function
        End If

        rs1.MoveNext
    Wend
End Function
Public Sub etiquetas_informe(codigo, linea1, linea2, linea3, linea4, linea5)
    Cantidad = 1
    For x = 1 To Cantidad
        Printer.CurrentY = 100
        Printer.Font = "IDAUTOMATIONHC39M"
        Printer.Font.Size = "20"
        x = Printer.Width
        mitad = x / 2
        xval = Printer.TextWidth(codigo) / 2
        puntox = mitad - xval
        Printer.CurrentX = puntox
        Printer.Print codigo
        Printer.Font = "times new roman"
        Printer.Font.Size = "18"
        Printer.Print linea1
        Printer.Print linea2
        Printer.Print linea3
        Printer.Print linea4
        Printer.Print linea5
        Printer.EndDoc
    Next x

End Sub



Public Function informePiezaMaterialesPDF(Id As Long, Origen As Integer, DIALOGO As Boolean) As Boolean   '1- enviar 2-imprimir
    On Error GoTo err2

    Dim Obj As PageSet.PrinterControl
    Set Obj = New PrinterControl


    Dim r As Recordset
    informePiezaMateriales2 = True
    Dim strsql As String, strsql2 As String
    If Origen = 2 Then
        titulo = "Presupuesto:"
        strsql = "select c.razon,p.id,p.detalle,p.fechaEntrega as fe from presupuestos p,clientes c where p.idcliente=c.id and p.id=" & Id
    ElseIf Origen = 1 Then
        titulo = "Pedido:"
        strsql = "select c.razon,p.id,p.descripcion as detalle,p.fechaEntrega as fe from pedidos p,clientes c where p.idcliente=c.id and p.id=" & Id
    ElseIf Origen = 3 Then
        titulo = "Conjunto:"
        strsql = "select c.razon,s.detalle from stock s inner join clientes c on s.id_cliente=c.id where s.id=" & Id
    End If

    Me.ejecutar_consulta strsql


    ref = rs!detalle
    If Origen = 1 Or Origen = 2 Then
        fecEnt = rs!fe
        nroOT = Format("00000", rs!Id)
    Else
        fecEnt = Empty

    End If
    razon = rs!razon


    Materiales.Sections("cabeza").Controls("lblOT").caption = titulo & ": " & nroOT
    Materiales.Sections("cabeza").Controls("lblCliente").caption = razon
    Materiales.Sections("cabeza").Controls("lblReferencia").caption = ref
    Materiales.Sections("s3").Controls("lblbarcode").caption = "*" & Format(Id, "0000") & "*"




    Set r = rs_materiales(Id, Origen)


    Set Materiales.DataSource = r
    Obj.ChngOrientationLandscape
    Materiales.Show 1
    Obj.ReSetOrientation    'This resets the printer to portrait.
    Exit Function
err2:
    MsgBox Err.Description
    informePiezaMateriales2 = False
    Obj.ReSetOrientation
End Function

Public Function rs_materiales(Id As Long, Origen As Integer) As Recordset
    On Error GoTo err4
    Dim Ot As Boolean, presu As Boolean, otro As Boolean
    Dim r_1 As New Recordset
    Dim r_piezas As Recordset
    Dim r_mat As New Recordset
    With r_mat
        .Fields.Append "idMaterial", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "cantidad", adDouble, 20, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "codigo", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "m2kg", adDouble, 20, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "cantUnitario", adDouble, 20, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "descripcion", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "Rubro", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "grupo", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "espesor", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "unidad", adVarChar, 20, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "valor", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "fecha", adVarChar, 20, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "moneda", adVarChar, 20, adFldUpdatable    ' And adFldIsNullable

    End With
    r_mat.Open





    If Origen = 1 Then    'de ot
        Ot = True
        otro = False
        presu = False
    ElseIf Origen = 2 Then
        Ot = False
        otro = False
        presu = True

    ElseIf Origen = 3 Then
        Ot = False
        otro = True
        presu = False

    End If
    Dim stock As New classStock
    Set r_piezas = stock.ListaPiezas(Id, Ot, presu, otro)
    'r_piezas.MoveFirst
    While Not r_piezas.EOF
        Pieza = CLng(r_piezas!idPieza)
        cantidad_p = CDbl(r_piezas!Cantidad)
        'strsql = "select m.codigo,id_material, largo,ancho,largoTerm,AnchoTerm,dm.cantidad,id_unidad,pesoxunidad from desarrollo_material dm inner join materiales m on dm.id_material=m.id  where id_pieza=" & pieza
        'strsql = " select m.descripcion,r.rubro,g.grupo,m.espesor,m.codigo,id_material, largo,ancho,largoTerm,AnchoTerm,dm.cantidad,id_unidad,pesoxunidad,r.rubro,g.grupo from desarrollo_material dm, rubros r, grupos g, materiales m where dm.id_material=m.id and m.id_rubro = r.id and  m.id_grupo=g.id and id_pieza=" & pieza
        'strsql = " select vm.valor_unitario,vm.fecha_actualizacion,m.descripcion,r.rubro,g.grupo,m.espesor,m.codigo,vm.id_material, largo,ancho,largoTerm,AnchoTerm,dm.cantidad,id_unidad,pesoxunidad,r.rubro,g.grupo from desarrollo_material dm, rubros r, grupos g, materiales m, valores_MATERIALES vm where dm.id_material=m.id and m.id_rubro = r.id and  m.id_grupo=g.id and vm.id_material=m.id and id_pieza=" & pieza
        strsql = " select mo.nombre_corto,m.valor_unitario,m.fecha_valor as  fecha_actualizacion,m.descripcion,r.rubro,g.grupo,m.espesor,m.codigo,m.id as id_material, dm.largo,dm.ancho,largoTerm,AnchoTerm,dm.cantidad,id_unidad,pesoxunidad,r.rubro,g.grupo from desarrollo_material dm, rubros r, grupos g, materiales m, AdminConfigMonedas mo where m.id_moneda=mo.id and dm.id_material=m.id and m.id_rubro = r.id and  m.id_grupo=g.id and  id_pieza=" & Pieza
        Set r_1 = conectar.RSFactory(strsql)
        While Not r_1.EOF
            IdMaterial = r_1!id_material
            Largo = r_1!Largo
            Ancho = r_1!Ancho
            Grupo = r_1!Grupo
            descrip = r_1!descripcion
            rubro = r_1!rubro
            Grupo = r_1!Grupo
            Espesor = r_1!Espesor
            largop = r_1!LargoTerm
            anchop = r_1!AnchoTerm
            Cantidad = r_1!Cantidad * cantidad_p
            PesoXUnidad = r_1!PesoXUnidad
            moneda = r_1!Nombre_corto
            codigo = r_1!codigo
            id_Unidad = r_1!id_Unidad
            valor = r_1!valor_unitario
            fec_act = r_1!FEcha_actualizacion
            'tengo que buscar en el RS temporal a ver si ya agregé el material
            'si está agregado lo sumo
            If id_Unidad = 3 Then    'ml
                medida = largop / 1000 * Cantidad
                unidad = Math.Round(medida, 2)
                totUnit = Math.Round(unidad * PesoXUnidad, 2)
                UN = "Ml"

            ElseIf id_Unidad = 1 Then    'kg
                medida = PesoXUnidad * Cantidad
                unidad = Math.Round(medida, 2)
                totUnit = Math.Round(unidad, 2)
                UN = "Kg"
            ElseIf id_Unidad = 2 Then    'm2
                medida = (anchop * largop) / 1000000
                unidad = Math.Round(medida * Cantidad, 2)
                totUnit = Math.Round(unidad * PesoXUnidad, 2)
                UN = "M2"
            ElseIf id_Unidad = 4 Then    'uni
                medida = Cantidad
                unidad = Math.Round(medida, 2)
                totUnit = Math.Round(unidad, 2)
                UN = "Un"
            End If
            agregarARSmateriales r_mat, id_Unidad, IdMaterial, unidad, totUnit, codigo, descrip, rubro, Grupo, Espesor, UN, valor, fec_act, moneda

            r_1.MoveNext
        Wend


        r_piezas.MoveNext
    Wend

    r_mat.MoveFirst

    Set rs_materiales = r_mat
    Exit Function
err4:

End Function

Private Sub agregarARSmateriales(r_mat As Recordset, unidad, IdMaterial, m2kg, totUnit, codigo, descripcion, rubro, Grupo, Espesor, UN, valor_unit, fec_act, moneda)
'busco en el rs
    If r_mat.RecordCount > 0 Then r_mat.MoveFirst

    r_mat.Find "idmaterial=" & IdMaterial

    If r_mat.EOF Then
        'agrego
        r_mat.AddNew
        r_mat!IdMaterial = IdMaterial
        r_mat!codigo = codigo
        r_mat!CANTuNITARIO = Math.Round(totUnit, 2)
        r_mat!m2kg = Math.Round(m2kg, 2)
        r_mat!rubro = UCase(rubro)
        r_mat!Grupo = UCase(Grupo)
        r_mat!descripcion = UCase(descripcion)
        r_mat!unidad = UN
        r_mat!Espesor = Espesor
        r_mat!valor = valor_unit
        r_mat!FEcha = fec_act
        r_mat!moneda = moneda
        r_mat.Update
    Else
        'edito

        r_mat!CANTuNITARIO = r_mat!CANTuNITARIO + totUnit
        r_mat!m2kg = Math.Round(r_mat!m2kg + m2kg, 2)
        r_mat.Update



        'sumo
    End If



End Sub


Public Function AprobarOrdenEntrega(ByVal IDOE As Long) As Boolean

    On Error GoTo errHandler

    Dim rsDetalle As ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim rsOE As ADODB.Recordset

    Dim idPieza As Long
    Dim cantidadOE As Double
    Dim cantidadStock As Double
    Dim valor As Double

    Dim idMonedaOE As Long
    Dim idUsuario As Long

    Dim asumirPrecio As Boolean
    Dim transActiva As Boolean

    Dim fechaAprobacion As String
    Dim fechaPrecio As String
    Dim fechaMovimiento As String

    Dim strCantidad As String
    Dim strValor As String

    Dim sql As String
    Dim afectados As Variant


    AprobarOrdenEntrega = False


    '--------------------------------------------------
    ' VALIDAR QUE EXISTA LA OE
    '--------------------------------------------------
    Set rsOE = cn.execute( _
        "SELECT estado, idMoneda " & _
        "FROM PedidosEntregas " & _
        "WHERE id = " & CStr(IDOE))


    If rsOE.EOF Then

        MsgBox "No existe la Orden de Entrega Nro. " & _
               CStr(IDOE) & ".", _
               vbExclamation, _
               "Orden de Entrega"

        Exit Function

    End If


    '--------------------------------------------------
    ' VALIDAR ESTADO
    '--------------------------------------------------
    If CLng(rsOE!estado) <> EstadoOrdenEntrega.Pendiente Then

        MsgBox "La Orden de Entrega Nro. " & _
               CStr(IDOE) & _
               " ya no se encuentra pendiente.", _
               vbExclamation, _
               "Orden de Entrega"

        Exit Function

    End If


    If Not IsNull(rsOE!IdMoneda) Then
        idMonedaOE = CLng(rsOE!IdMoneda)
    Else
        idMonedaOE = 0
    End If


    '--------------------------------------------------
    ' OBTENER DETALLE
    '--------------------------------------------------
    Set rsDetalle = cn.execute( _
        "SELECT idPieza, cantidad, vale " & _
        "FROM detallesPedidosEntregas " & _
        "WHERE idPedidoEntrega = " & CStr(IDOE))


    If rsDetalle.EOF Then

        MsgBox "La Orden de Entrega Nro. " & _
               CStr(IDOE) & _
               " no tiene piezas cargadas." & vbCrLf & _
               "No puede ser aprobada.", _
               vbExclamation, _
               "Orden de Entrega"

        Exit Function

    End If


    '--------------------------------------------------
    ' VALIDAR STOCK DE TODAS LAS PIEZAS
    '--------------------------------------------------
    rsDetalle.MoveFirst

    Do While Not rsDetalle.EOF

        idPieza = CLng(rsDetalle!idPieza)
        cantidadOE = CDbl(rsDetalle!Cantidad)


        Set rsStock = cn.execute( _
            "SELECT cantidad " & _
            "FROM stock " & _
            "WHERE id = " & CStr(idPieza))


        If rsStock.EOF Then

            MsgBox "No se encontró en stock la pieza ID " & _
                   CStr(idPieza) & ".", _
                   vbCritical, _
                   "Orden de Entrega"

            Exit Function

        End If


        cantidadStock = CDbl(rsStock!Cantidad)


        If cantidadStock < cantidadOE Then

            MsgBox "No hay stock suficiente para aprobar la O/E." & _
                   vbCrLf & vbCrLf & _
                   "Pieza ID: " & CStr(idPieza) & vbCrLf & _
                   "Stock disponible: " & _
                   funciones.FormatearDecimales(cantidadStock, 2) & vbCrLf & _
                   "Cantidad requerida: " & _
                   funciones.FormatearDecimales(cantidadOE, 2), _
                   vbExclamation, _
                   "Orden de Entrega"

            Exit Function

        End If


        rsDetalle.MoveNext

    Loop


    '--------------------------------------------------
    ' CONSULTAR SI ASUME LOS PRECIOS
    '--------------------------------------------------
    asumirPrecio = False

    If MsgBox( _
        "¿Desea asumir el valor de las piezas de esta O/E " & _
        "como precio definido en el stock?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Precio definido") = vbYes Then

        asumirPrecio = True

    End If


    '--------------------------------------------------
    ' COMENZAR TRANSACCION
    '--------------------------------------------------
    cn.BeginTrans
    transActiva = True


    idUsuario = funciones.getUser
    fechaAprobacion = funciones.datetimeFormateada(Now)


    '--------------------------------------------------
    ' APROBAR CABECERA
    '--------------------------------------------------
    sql = _
        "UPDATE PedidosEntregas SET " & _
        "estado = " & CStr(EstadoOrdenEntrega.Aprobado) & ", " & _
        "IdUsuarioAprobado = " & CStr(idUsuario) & ", " & _
        "fechaAprobado = '" & fechaAprobacion & "' " & _
        "WHERE id = " & CStr(IDOE) & _
        " AND estado = " & CStr(EstadoOrdenEntrega.Pendiente)


    cn.execute sql, afectados


    If CLng(afectados) <> 1 Then

        Err.Raise vbObjectError + 2500, _
                  "classPlaneamiento.AprobarOrdenEntrega", _
                  "La O/E cambió de estado mientras se intentaba aprobar."

    End If


    '--------------------------------------------------
    ' HISTORIAL
    '--------------------------------------------------
    cn.execute _
        "INSERT INTO historico_PedidoEntrega " & _
        "(usuario, nota, fecha) VALUES (" & _
        CStr(idUsuario) & ", " & _
        "'O/E Aprobada', " & _
        "'" & fechaAprobacion & "')"


    '--------------------------------------------------
    ' DESCONTAR STOCK
    '--------------------------------------------------
    rsDetalle.MoveFirst

    Do While Not rsDetalle.EOF

        idPieza = CLng(rsDetalle!idPieza)
        cantidadOE = CDbl(rsDetalle!Cantidad)

        If IsNull(rsDetalle!vale) Then
            valor = 0
        Else
            valor = CDbl(rsDetalle!vale)
        End If


        strCantidad = Replace$(CStr(cantidadOE), ",", ".")
        strValor = Replace$(CStr(valor), ",", ".")


        If asumirPrecio Then

            fechaPrecio = funciones.dateFormateada(Now)

            sql = _
                "UPDATE stock SET " & _
                "cantidad = cantidad - " & strCantidad & ", " & _
                "precio_definido = " & strValor & ", " & _
                "fecha_precio_definido = '" & fechaPrecio & "', " & _
                "id_moneda_precio = " & CStr(idMonedaOE) & " " & _
                "WHERE id = " & CStr(idPieza)

        Else

            sql = _
                "UPDATE stock SET " & _
                "cantidad = cantidad - " & strCantidad & " " & _
                "WHERE id = " & CStr(idPieza)

        End If


        cn.execute sql


        '--------------------------------------------------
        ' MOVIMIENTO DE STOCK
        '--------------------------------------------------
        fechaMovimiento = Format$(Date, "yyyy/mm/dd")

        cn.execute _
            "INSERT INTO movimiento_stock " & _
            "(id_pieza, cantidad, operacion, fecha, nota) VALUES (" & _
            CStr(idPieza) & ", " & _
            strCantidad & ", " & _
            "3, " & _
            "'" & fechaMovimiento & "', " & _
            CStr(IDOE) & ")"


        rsDetalle.MoveNext

    Loop


    '--------------------------------------------------
    ' CONFIRMAR
    '--------------------------------------------------
    cn.CommitTrans
    transActiva = False

    AprobarOrdenEntrega = True

    Exit Function


errHandler:

    If transActiva Then
        cn.RollbackTrans
    End If

    AprobarOrdenEntrega = False

    MsgBox "Error al aprobar la Orden de Entrega Nro. " & _
           CStr(IDOE) & "." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Function

