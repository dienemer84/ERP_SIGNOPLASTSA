Attribute VB_Name = "DAODetalleOrdenTrabajo"
Option Explicit
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ITEM As String = "item"
Public Const CAMPO_CANTIDAD_PEDIDA As String = "cantidad"
Public Const CAMPO_NOMBRE_PIEZA_HISTORICO As String = "detalle_pieza_temporal"
Public Const CAMPO_NOTA As String = "nota"
Public Const CAMPO_RESERVA_STOCK As String = "reserva_stock"
Public Const CAMPO_CANTIDAD_FABRICADOS As String = "cantidad_fabricados"
Public Const CAMPO_RETIRADO As String = "retirado"
Public Const CAMPO_PRECIO As String = "precio"
Public Const CAMPO_CANTIDAD_ENTREGADA As String = "cantidad_entregada"
Public Const CAMPO_FECHA_ENTREGA As String = "fechaEntrega"
Public Const CAMPO_CANTIDAD_FACTURADA As String = "cantidad_facturada"
Public Const CAMPO_NOTA_PRODUCCION As String = "nota_produccion"
Public Const CAMPO_CANT_IMPRESIONES_RUTA As String = "impresiones_ruta"
Public Const CAMPO_PRECIO_MODIFICADO As String = "precio_modificado"
Public Const CAMPO_ESTADO_PROCESO As String = "procesos_definidos"
Public Const CAMPO_PEDIDO_ID As String = "idPedido"
Public Const CAMPO_PIEZA_ID As String = "idPieza"
Public Const TABLA_DETALLE_PEDIDO As String = "dp"
Public Const CAMPO_ID_PRESU = "id_presupuesto_origen"
Public Const CAMPO_DESCUENTO = "descuento"

Public Sub EnviarAStock(detalle As DetalleOrdenTrabajo, Cantidad As Double)
    On Error GoTo err1
    'reload
    Set detalle = DAODetalleOrdenTrabajo.FindById(detalle.Id)

    'valido nuevamente
    If detalle.CantidadPedida >= detalle.CantidadEnviadasAStock + Cantidad Then
        detalle.CantidadEnviadasAStock = detalle.CantidadEnviadasAStock + Cantidad
        conectar.BeginTransaction
        Save (detalle)
        conectar.CommitTransaction
    Else
        Err.Raise 2020, "Asignación de Stock", "La cantidad de stock que se intenta asignar supera a la cantidad disponible"
    End If
    Exit Sub
err1:
    conectar.RollBackTransaction
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function CountPrintedLabels(Id As Long) As Boolean
    On Error GoTo err1
    conectar.execute "update detalles_pedidos set etiquetas_impresas = etiquetas_impresas+1 where id=" & Id
    CountPrintedLabels = True
    Exit Function
err1:
    CountPrintedLabels = False
End Function

Public Function FindById(IdPedido As Long, Optional withEntregados As Boolean = False, Optional withFabricados As Boolean = False, Optional withFacturados As Boolean = False) As DetalleOrdenTrabajo
    Dim col As Collection
    Set col = FindAll(TABLA_DETALLE_PEDIDO & "." & CAMPO_ID & "=" & IdPedido, withEntregados, withFabricados, withFacturados)

    If col.count > 0 Then
        Set FindById = col.item(1)
    Else
        Set FindById = Nothing
    End If

End Function


Public Function PendientesEntregaPorPieza(idPieza As Long) As Double

'busco todas las ordenes activas donde se est'e fabricando esta pieza
    Dim sql As String
    sql = " SELECT SUM(dp.cantidad)-SUM(dpc.cantidad)  as pendientes FROM detalles_pedidos_cantidad dpc " _
        & "INNER JOIN detalles_pedidos dp ON dpc.id_detalle_pedido=dp.id " _
        & "Where tipo_cantidad = 2 And id_detalle_pedido " _
        & "IN (SELECT dp.id FROM pedidos p INNER JOIN detalles_pedidos dp ON dp.idPedido=p.id  WHERE p.estado=2 AND dp.idPieza = " & idPieza & ") " _
        & " GROUP BY dp.idPieza"

    Dim rs As New Recordset
    Set rs = RSFactory(sql)
    If Not rs.EOF And Not rs.BOF Then
        PendientesEntregaPorPieza = rs!pendientes
    End If



End Function



Public Function FindAllByPieza(piezasId As Collection) As Collection
    Dim filter As String
    filter = "{detalle_ot}.{idPieza} IN ({idPieza_value})"
    filter = Replace$(filter, "{detalle_ot}", DAODetalleOrdenTrabajo.TABLA_DETALLE_PEDIDO)
    filter = Replace$(filter, "{idPieza}", DAODetalleOrdenTrabajo.CAMPO_PIEZA_ID)
    filter = Replace$(filter, "{idPieza_value}", funciones.JoinCollectionValues(piezasId, ", ", "id"))

    Set FindAllByPieza = DAODetalleOrdenTrabajo.FindAll(filter)
End Function


Public Function FindAllByOrdenTrabajo(orden_trabajo_id As Long, Optional withEntregados As Boolean = False, Optional withFabricados As Boolean = False, Optional withFacturados As Boolean = False, Optional WithTareasFinalizadas As Boolean = False) As Collection
    Dim filter As String
    filter = "{detalle_ot}.{ot_id} = {ot_id_value}"
    filter = Replace$(filter, "{detalle_ot}", DAODetalleOrdenTrabajo.TABLA_DETALLE_PEDIDO)
    filter = Replace$(filter, "{ot_id}", DAODetalleOrdenTrabajo.CAMPO_PEDIDO_ID)
    filter = Replace$(filter, "{ot_id_value}", orden_trabajo_id)
    Set FindAllByOrdenTrabajo = DAODetalleOrdenTrabajo.FindAll(filter, withEntregados, withFabricados, withFacturados, , , , WithTareasFinalizadas)
End Function


Public Function FindConjuntoById(Id As Long) As DetalleOTConjuntoDTO
    Dim col As Collection
    Set col = FindAllConjunto(, , , , Id)    ', , "dpc.id = " & id)
    If col.count > 0 Then
        Set FindConjuntoById = col.item(1)
    Else
        Set FindConjuntoById = Nothing
    End If
End Function


Public Function FindConjuntoByPiezas(piezasId As Collection) As Collection
    Set FindConjuntoByPiezas = FindAllConjunto(, , "dpc.idPieza IN (" & funciones.JoinCollectionValues(piezasId, ", ") & ")")
End Function

'si se llama solo especificando idDetallePedido (detalle_pedido.id) viene todo el conjunto en plano
'para que venga por partes se debe filtrar por idPiezaPadre tambien
Public Function FindAllConjunto(Optional ByVal idDetallePedido As Long = 0, Optional ByVal idPiezaPadre As Long = 0, Optional ByRef filter As String = vbNullString, Optional withDesarrolloManoObra As Boolean = False, Optional idDetalleConjunto As Long = 0) As Collection

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim detaOrdenesTrabajo As New Collection

    q = "SELECT" _
      & " dpc.*," _
      & " s.*," _
      & " dp.*" _
      & " FROM detalles_pedidos_conjuntos dpc" _
      & " LEFT JOIN stock s" _
      & " ON s.id = dpc.idPieza" _
      & " LEFT JOIN detalles_pedidos dp" _
      & " ON dp.id = dpc.idDetalle_pedido" _
      & " WHERE 1 = 1"

    If idDetallePedido > 0 Then
        q = q & " AND dpc.idDetalle_Pedido = " & idDetallePedido
    End If

    If idDetalleConjunto > 0 Then
        q = q & " And dpc.id = " & idDetalleConjunto
    End If
    If idPiezaPadre > 0 Then
        q = q & " And dpc.idPiezaPadre = " & idPiezaPadre
    End If

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    'Set FindAll = New Collection

    Const piezaTabla As String = "s"
    Dim tmpDeta As DetalleOTConjuntoDTO

    While Not rs.EOF
        Set tmpDeta = DAODetalleOrdenTrabajo.MapConjunto(rs, fieldsIndex, "dpc", piezaTabla, "dp")

        If withDesarrolloManoObra Then
            Set tmpDeta.Pieza.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(tmpDeta.Pieza.Id)
        End If

        detaOrdenesTrabajo.Add tmpDeta, CStr(tmpDeta.Id)
        rs.MoveNext
    Wend

    Set FindAllConjunto = detaOrdenesTrabajo

End Function

Public Function FindAll(Optional ByRef filter As String = vbNullString, _
                        Optional Entregados As Boolean = False, _
                        Optional Fabricados As Boolean = False, _
                        Optional Facturados As Boolean = False, _
                        Optional withColEntregados As Boolean = False, _
                        Optional withColFabricados As Boolean = False, _
                        Optional withColFacturados As Boolean = False, _
                        Optional WithTareasFinalizadas As Boolean = False _
                      ) As Collection

    On Error GoTo err1
    Dim tickStart As Double
    Dim tickend As Double
    '  tickStart = GetTickCount
    Dim rs As ADODB.Recordset
    Dim q As String

    Dim detaOrdenesTrabajo As New Collection

    q = "SELECT * " _

    If Facturados Then
    
            q = q & ",IFNULL((SELECT SUM(cantidad) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadFacturada_ & " AND id_detalle_pedido=dp.id),0) AS FacturadosCantidad"
            q = q & ",IFNULL((SELECT SUM(((monto * cantidad)/1)*tipo_cambio) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadFacturada_ & " AND id_detalle_pedido=dp.id),0) AS FacturadosMonto"
    End If

    If Entregados Then
        q = q & ",IFNULL((SELECT SUM(cantidad) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadEntregada_ & " AND id_detalle_pedido=dp.id),0) AS EntregadosCantidad"
    End If

    If Fabricados Then
        q = q & ",IFNULL((SELECT SUM(cantidad) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadFabricada_ & " AND id_detalle_pedido=dp.id),0) AS FabricadosCantidad"
    End If

    If WithTareasFinalizadas Then
        'deberia traer las tareas finalizadas y las tareas totales para poder sacr un porcentaje de avance de la OT
        q = q & ", COUNT(ptp.id) AS CantidadTareas,  SUM(ptp.fechaFin>0) AS CantidadTareasFinalizadas "
    End If
    q = q & " FROM detalles_pedidos dp"

    q = q & " LEFT JOIN stock s ON s.id = dp.idPieza"

    If WithTareasFinalizadas Then

        q = q & " LEFT JOIN PlaneamientoTiemposProcesos ptp ON ptp.idDetallePedido=dp.id "
    End If


    q = q & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    If WithTareasFinalizadas Then
        q = q & " group by dp.id "
        
    End If

    q = q & " ORDER BY dp.item ASC"
    'If filter = "dp.IdDetalleOtPadre = 36415" Then Stop

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAll = New Collection


    Const piezaTabla As String = "s"
    Dim tmpDeta As DetalleOrdenTrabajo

    While Not rs.EOF
        Set tmpDeta = DAODetalleOrdenTrabajo.Map(rs, fieldsIndex, TABLA_DETALLE_PEDIDO, piezaTabla, Entregados, Fabricados, Facturados, WithTareasFinalizadas)


        If withColEntregados Then
            Set tmpDeta.colCantidadesEntregadas = MapCantidad(tmpDeta.Id, CantidadEntregada_, tmpDeta.OrdenTrabajo.IdMoneda)
        End If
        If withColFabricados Then
            Set tmpDeta.colCantidadesFabricadas = MapCantidad(tmpDeta.Id, CantidadFabricada_, tmpDeta.OrdenTrabajo.IdMoneda)
        End If
        If withColFacturados Then
            Set tmpDeta.colCantidadesFacturadas = MapCantidad(tmpDeta.Id, CantidadFacturada_, tmpDeta.OrdenTrabajo.IdMoneda)
        End If

        detaOrdenesTrabajo.Add tmpDeta, CStr(tmpDeta.Id)
        rs.MoveNext
    Wend

    Set FindAll = detaOrdenesTrabajo
    Exit Function

err1:
    Set FindAll = Nothing
End Function


Public Function MapConjunto(ByRef rs As Recordset, _
                            ByRef fieldsIndex As Dictionary, _
                            ByRef tableNameOrAlias As String, _
                            Optional ByRef piezaTableNameOrAlias As String = vbNullString, _
                            Optional ByRef detallePedidoTableNameOrAlias As String = vbNullString _
                          ) As DetalleOTConjuntoDTO

    Dim tmpDeta As DetalleOTConjuntoDTO
    Dim Id As Variant
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, "id")

    If Id > 0 Then
        Set tmpDeta = New DetalleOTConjuntoDTO

        tmpDeta.Id = Id

        tmpDeta.Cantidad = GetValue(rs, fieldsIndex, tableNameOrAlias, "CantidadPieza")
        tmpDeta.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, "procesos_definidos")
        tmpDeta.idDetallePedido = GetValue(rs, fieldsIndex, tableNameOrAlias, "idDetalle_pedido")
        tmpDeta.IdPedido = GetValue(rs, fieldsIndex, tableNameOrAlias, "idPedido")
        tmpDeta.IdentificadorPosicion = GetValue(rs, fieldsIndex, tableNameOrAlias, "identificador_posicion")
        
        tmpDeta.CantidadTotalStatic = GetValue(rs, fieldsIndex, tableNameOrAlias, "cantidad_total_static")
        
'''        tmpDeta.CantidadRecibida = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_cant_recibida")
'''        tmpDeta.CantidadFabricada = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_cant_fabricada")
'''        tmpDeta.CantidadScrap = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_cant_scrap")
'''
'''        tmpDeta.FechaInicio = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_fecha_inicio")
'''        tmpDeta.FechaFin = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_fecha_fin")
'''
'''        tmpDeta.Recibio = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_recibio")
'''        tmpDeta.SiguienteProceso = GetValue(rs, fieldsIndex, tableNameOrAlias, "a_siguiente_proceso")

        If LenB(piezaTableNameOrAlias) > 0 Then Set tmpDeta.Pieza = DAOPieza.Map(rs, fieldsIndex, piezaTableNameOrAlias)
        If LenB(detallePedidoTableNameOrAlias) > 0 Then Set tmpDeta.DetalleRaiz = DAODetalleOrdenTrabajo.Map(rs, fieldsIndex, detallePedidoTableNameOrAlias)

    End If

    Set MapConjunto = tmpDeta

End Function

Public Function Map(ByRef rs As Recordset, _
                    ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef piezaTableNameOrAlias As String = vbNullString, _
                    Optional Entregados As Boolean = False, _
                    Optional Fabricados As Boolean = False, _
                    Optional Facturados As Boolean = False, _
                    Optional WithTareasFinalizadas As Boolean = False _
                  ) As DetalleOrdenTrabajo

    Dim tmpDetaOrdenTrabajo As DetalleOrdenTrabajo
    Dim Id As Variant
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODetalleOrdenTrabajo.CAMPO_ID)

    If Id > 0 Then
        Set tmpDetaOrdenTrabajo = New DetalleOrdenTrabajo

        tmpDetaOrdenTrabajo.Id = Id
        tmpDetaOrdenTrabajo.item = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ITEM)
        tmpDetaOrdenTrabajo.CantidadPedida = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CANTIDAD_PEDIDA)
        tmpDetaOrdenTrabajo.NombrePiezaHistorico = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NOMBRE_PIEZA_HISTORICO)
        tmpDetaOrdenTrabajo.Nota = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NOTA)
        tmpDetaOrdenTrabajo.ReservaStock = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_RESERVA_STOCK)
        tmpDetaOrdenTrabajo.CantidadFabricadosStatic = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CANTIDAD_FABRICADOS)
        tmpDetaOrdenTrabajo.Retirado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_RETIRADO)
        tmpDetaOrdenTrabajo.Precio = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PRECIO)
        tmpDetaOrdenTrabajo.CantidadEntregadaStatic = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CANTIDAD_ENTREGADA)
        tmpDetaOrdenTrabajo.FechaEntrega = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_ENTREGA)
        tmpDetaOrdenTrabajo.CantidadFacturadaStatic = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CANTIDAD_FACTURADA)
        tmpDetaOrdenTrabajo.NotaProduccion = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NOTA_PRODUCCION)
        tmpDetaOrdenTrabajo.CantidadImpresionesDeRuta = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CANT_IMPRESIONES_RUTA)
        tmpDetaOrdenTrabajo.PrecioModificado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PRECIO_MODIFICADO)
        tmpDetaOrdenTrabajo.EstadoProceso = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ESTADO_PROCESO)
        tmpDetaOrdenTrabajo.idDetalleOtPadre = GetValue(rs, fieldsIndex, tableNameOrAlias, "IdDetalleOtPadre")
        tmpDetaOrdenTrabajo.EtiquetasImpresas = GetValue(rs, fieldsIndex, tableNameOrAlias, "etiquetas_impresas")
        tmpDetaOrdenTrabajo.idPresupuestoOrigen = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID_PRESU)
        tmpDetaOrdenTrabajo.CantidadEnviadasAStock = GetValue(rs, fieldsIndex, tableNameOrAlias, "cantidad_a_stock")
        
        
        'pseudo proxy
        Set tmpDetaOrdenTrabajo.OrdenTrabajo = New OrdenTrabajo
        tmpDetaOrdenTrabajo.OrdenTrabajo.Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PEDIDO_ID)

        If LenB(piezaTableNameOrAlias) > 0 Then Set tmpDetaOrdenTrabajo.Pieza = DAOPieza.Map(rs, fieldsIndex, piezaTableNameOrAlias)



        tmpDetaOrdenTrabajo.Descuento = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DESCUENTO)

        If WithTareasFinalizadas Then
            tmpDetaOrdenTrabajo.CantidadTareas = GetValue(rs, fieldsIndex, vbNullString, "CantidadTareas")
            tmpDetaOrdenTrabajo.CantidadTareasFinalizadas = GetValue(rs, fieldsIndex, vbNullString, "CantidadTareasFinalizadas")
        End If


        If Facturados Then
            tmpDetaOrdenTrabajo.CantidadFacturada = GetValue(rs, fieldsIndex, vbNullString, "FacturadosCantidad")
            tmpDetaOrdenTrabajo.MontoFacturado = GetValue(rs, fieldsIndex, vbNullString, "FacturadosMonto")

        End If

        If Entregados Then
            tmpDetaOrdenTrabajo.CantidadEntregada = GetValue(rs, fieldsIndex, "", "EntregadosCantidad")
        End If

        If Fabricados Then
            tmpDetaOrdenTrabajo.CantidadFabricados = GetValue(rs, fieldsIndex, "", "FabricadosCantidad")
        End If

    End If






    Set Map = tmpDetaOrdenTrabajo

End Function

Public Function Save(deta As DetalleOrdenTrabajo) As Boolean
    Dim q As String

    If deta.NotaProduccion = Empty Then deta.NotaProduccion = " "

    If deta.NotaProduccion = Empty Then
        deta.NotaProduccion = " "
    End If



    If deta.Id = 0 Then
        q = q & "INSERT INTO detalles_pedidos" _
          & " (item," _
          & " idPedido," _
          & " idPieza," _
          & " cantidad, " _
          & " detalle_pieza_temporal," _
          & " nota," _
          & " reserva_stock," _
          & " cantidad_fabricados," _
          & " retirado," _
          & " precio," _
          & " cantidad_entregada," _
          & " fechaEntrega," _
          & " cantidad_facturada," _
          & " nota_produccion," _
          & " impresiones_ruta," _
          & " precio_modificado," _
          & " procesos_definidos, IdDetalleOtPadre,id_presupuesto_origen,descuento,id_moneda,cantidad_a_stock)" _
          & " VALUES (" _
          & conectar.Escape(deta.item) & "," _
          & conectar.GetEntityId(deta.OrdenTrabajo) & "," _
          & conectar.GetEntityId(deta.Pieza) & "," _
          & conectar.Escape(deta.CantidadPedida) & "," _
          & conectar.Escape(deta.NombrePiezaHistorico) & ","

        Dim d As Integer
        d = conectar.Escape(deta.CantidadImpresionesDeRuta)

        q = q & "'" & deta.Nota & "'," _
          & conectar.Escape(deta.ReservaStock) & "," _
          & conectar.Escape(deta.CantidadFabricados) & "," _
          & conectar.Escape(deta.Retirado) & "," _
          & conectar.Escape(deta.Precio) & "," _
          & conectar.Escape(deta.CantidadEntregada) & "," _
          & conectar.Escape(deta.FechaEntrega) & "," _
          & conectar.Escape(deta.CantidadFacturada) & "," _
          & "'" & deta.NotaProduccion & "'," _
          & d & "," _
          & conectar.Escape(deta.PrecioModificado) & "," _
          & conectar.Escape(deta.EstadoProceso) & ", " & deta.idDetalleOtPadre & "," & Escape(deta.idPresupuestoOrigen) & "," & Escape(deta.Descuento) & "," & Escape(deta.IdMoneda) & "," & Escape(deta.CantidadEnviadasAStock) & ")"

        Save = conectar.execute(q)

        Dim Id As Long

        If Save Then
            conectar.UltimoId "detalles_pedidos", Id
            deta.Id = Id

            If deta.Pieza.EsConjunto Then
                Dim st As New classStock
                Dim r_arbol As Recordset
                'traigo el recordset con el arbol en plano
                Set r_arbol = st.ArbolConjunto(deta.Pieza.Id)
                'cargo el detalle_pedido de conjuntos para teenr el despiece separado
                While Not r_arbol.EOF
                    q = "insert into detalles_pedidos_conjuntos  (idPedido, idDetalle_pedido, IdPiezaPadre, idPieza, esConjunto, cantidad, reserva_stock,cantidadPieza,cantidad_total_static,identificador_posicion)"
                    q = q & " values (" & deta.OrdenTrabajo.Id & "," & deta.Id & "," & r_arbol!idPiezaPadre & "," & r_arbol!idPieza & "," & r_arbol!conjunto & "," & r_arbol!Cantidad & "," & 0 & "," & r_arbol!Cantidad & "," & r_arbol!Cantt * deta.CantidadPedida & ",'" & r_arbol!id_pos & "')"
                    Save = conectar.execute(q)

                    r_arbol.MoveNext
                Wend
            End If

        End If

    Else

        q = "update detalles_pedidos" _
          & " SET" _
          & " item = " & conectar.Escape(deta.item) & " ," _
          & " idPedido = " & conectar.GetEntityId(deta.OrdenTrabajo) & " ," _
          & " idPieza = " & conectar.GetEntityId(deta.Pieza) & "," _
          & " cantidad = " & conectar.Escape(deta.CantidadPedida) & " ," _
          & " detalle_pieza_temporal = " & conectar.Escape(deta.NombrePiezaHistorico) & " ," _
          & " nota = " & conectar.Escape(deta.Nota) & " ," _
          & " reserva_stock = " & conectar.Escape(deta.ReservaStock) & " ," _
          & " cantidad_fabricados = " & conectar.Escape(deta.CantidadFabricados) & " ," _
          & " retirado = " & conectar.Escape(deta.Retirado) & " ," _
          & " precio = " & conectar.Escape(deta.Precio) & " ," _
          & " cantidad_entregada = " & conectar.Escape(deta.CantidadEntregada) & " ," _
          & " fechaEntrega = " & conectar.Escape(deta.FechaEntrega) & " ," _
          & " cantidad_facturada = " & conectar.Escape(deta.CantidadFacturada) & " ," _
          & " nota_produccion = " & conectar.Escape(deta.NotaProduccion) & " ," _
          & " impresiones_ruta = " & conectar.Escape(deta.CantidadImpresionesDeRuta) & " ," _
          & " precio_modificado = " & conectar.Escape(deta.PrecioModificado) & " ," _
          & " procesos_definidos = " & conectar.Escape(deta.EstadoProceso) & ", " _
          & " id_presupuesto_origen = " & conectar.Escape(deta.idPresupuestoOrigen) & ", " _
          & " idDetalleOtPadre = " & conectar.Escape(deta.idDetalleOtPadre) & "," _
          & " id_moneda = " & conectar.Escape(deta.IdMoneda) & "," _
          & " descuento = " & conectar.Escape(deta.Descuento) _
          & " Where" _
          & " id = " & deta.Id



        Save = conectar.execute(q)
    End If



End Function




Public Function MapCantidad(id_detalle As Long, Tipo As TipoCantidadOT, idMonedaOrden As Long) As Collection
    Dim strsql As String
    Dim rs As Recordset
    Dim Cant As clsDetalleOrdenTrabajoCantidades
    Dim cole As New Collection
    strsql = "select * from detalles_pedidos_cantidad where id_detalle_pedido=" & id_detalle & " and tipo_cantidad =" & Tipo
    Set rs = conectar.RSFactory(strsql)

    While Not rs.EOF And Not rs.BOF
        Set Cant = New clsDetalleOrdenTrabajoCantidades
        Cant.Cantidad = rs!Cantidad
        Cant.FEcha = rs!FEcha
        Cant.Monto = rs!Monto
        Cant.Tipo = rs!tipo_cantidad
        Cant.TipoCambio = rs!tipo_cambio
        Cant.IdMoneda = rs!id_moneda
        Cant.MontoSegunMoneda = MonedaConverter.ConvertirForzado2(Cant.Monto, Cant.IdMoneda, idMonedaOrden, Cant.TipoCambio)


        cole.Add Cant
        rs.MoveNext
    Wend
    Set MapCantidad = cole
End Function

Public Function arreglarCagada()
    On Error GoTo erro
    conectar.BeginTransaction

    Dim str As String
    Dim rs As Recordset
    Dim deta As FacturaDetalle

    Dim str1 As String
    Dim rs1 As Recordset
    Dim F As Factura

    Dim col As New Collection

    Set col = DAOOrdenTrabajo.FindAll(, , , , True)

    Dim Ot As OrdenTrabajo
    Dim det As DetalleOrdenTrabajo

    For Each Ot In col

        For Each det In Ot.detalles

            det.IdMoneda = Ot.moneda.Id
            ''If ot.Moneda.Id = 1 Then Stop
            DAODetalleOrdenTrabajo.Save det


        Next

    Next


    'Set rs = conectar.RSFactory("select id_detalle_pedido from detalles_pedidos_cantidad where tipo_cantidad=2 and tipo_cambio=0")
    '
    '
    'Set rs = conectar.RSFactory("SELECT *  FROM detalles_pedidos_cantidad WHERE tipo_cantidad=2 AND id_detalle_pedido >0 ")
    'While Not rs.EOF And Not rs.BOF
    '
    'If rs!id_detalle_pedido > 0 Then
    'str1 = "SELECT  a.idFactura, a.id  as id_detalle_Factura FROM AdminFacturasDetalleNueva a LEFT JOIN entregas e ON a.`idEntrega` = e.id WHERE e.`idDetallePedido` =" & rs!id_detalle_pedido
    'Set rs1 = conectar.RSFactory(str1)
    '
    'If Not rs1.BOF And Not rs1.EOF Then
    'Set F = DAOFactura.FindById(rs1!idFactura)
    'Set deta = DAOFacturaDetalles.FindAll("AdminFacturasDetalleNueva.id=" & rs1!id_detalle_Factura)(1)
    'End If
    'If IsSomething(deta) And IsSomething(F) Then
    'If Not conectar.execute("update detalles_pedidos_cantidad  set  tipo_cambio_ajuste=" & F.TipoCambioAjuste & ", tipo_cambio = " & F.CambioAPatron & ", id_moneda= " & F.Moneda.Id & ", id_comprobante = " & deta.Id & ", id_comprobante_entrega=" & deta.DetalleRemitoId & " where id_detalle_pedido=" & rs!id_detalle_pedido & " and tipo_cantidad=2") Then GoTo erro
    '
    'Else
    '
    'Stop
    'End If
    '
    'End If
    'rs.MoveNext
    'Wend


    conectar.CommitTransaction
    Exit Function
erro:
    conectar.RollBackTransaction
End Function


Public Function SaveCantidad(id_detalle As Long, Cantidad As Double, Tipo As TipoCantidadOT, Monto As Double, id_comprobante As Long, id_moneda As Long, tipo_cambio As Double, tipo_cambio_ajuste As Double) As Boolean
    Dim strsql As String
    On Error GoTo err1
    SaveCantidad = True
    'habria que validar q no se vaya a negativo...


    strsql = "INSERT INTO detalles_pedidos_cantidad (id_detalle_pedido, cantidad, fecha, tipo_cantidad, monto,id_comprobante,id_moneda,tipo_cambio)VALUES ( " & id_detalle & "," & Cantidad & "," & conectar.Escape(Now) & "," & Tipo & "," & Monto & "," & id_comprobante & "," & id_moneda & " ," & tipo_cambio & ")"
    SaveCantidad = conectar.execute(strsql)

    Exit Function
err1:
    SaveCantidad = False
End Function
Public Function GetCantidad(id_detalle As Long, Tipo As TipoCantidadOT) As Double
    Dim strsql As String
    Dim rs As Recordset
    strsql = "select SUM(cantidad) AS cantidad from detalles_pedidos_cantidad where id_detalle_pedido=" & id_detalle & " and tipo_cantidad=" & Tipo
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        GetCantidad = IIf(IsNumeric(rs!Cantidad), rs!Cantidad, 0)
    Else
        GetCantidad = -1
    End If
End Function
Public Sub CalcularPorcentajeAvanceYPromedioFabricado(idDetallePedido As Long, ByRef Porcentaje As Double, ByRef promedio As Double)
    Dim tmpRS As Recordset
    Dim porcentajeAvanceTareas As Double
    Dim q As String

    'primero calculo a primer nivel del detalle pedido, despues me meto si es que es conjunto


    'calculo el porcentaje de avance de todas las tareas del detalle del pedido

    q = "SELECT" _
      & " dp.cantidad AS cantidad_pedida," _
      & " ptp.codigoTarea," _
      & " SUM(ptpd.cantidad_procesada) AS cantidad_procesada," _
      & " udf_hasta_cien(((SUM(ptpd.cantidad_procesada) / dp.cantidad) * 100)) AS porcentaje_realizado_tarea," _
      & " ((SELECT (1 / COUNT(ptp2.id)) * 100 FROM PlaneamientoTiemposProcesos ptp2 WHERE ptp2.iddetallepedido = dp.id) * udf_hasta_cien(((SUM(ptpd.cantidad_procesada) / dp.cantidad) * 100))) / 100  AS porcentaje_tarea_pedido" _
      & " FROM detalles_pedidos dp" _
      & " INNER JOIN PlaneamientoTiemposProcesos ptp" _
      & " ON ptp.idDetallePedido = dp.id" _
      & " LEFT JOIN PlaneamientoTiemposProcesosDetalle ptpd" _
      & " ON ptpd.idTiemposProcesos = ptp.id" _
      & " Where dp.id = " & idDetallePedido _
      & " GROUP BY ptp.id"

    porcentajeAvanceTareas = -1
    Set tmpRS = conectar.RSFactory(q)
    While Not tmpRS.EOF
        If Not IsNull(tmpRS!porcentaje_tarea_pedido) Then
            porcentajeAvanceTareas = porcentajeAvanceTareas + tmpRS!porcentaje_tarea_pedido
        End If
        tmpRS.MoveNext
    Wend


    'calculo la cantidad de piezas que fabrique hasta ahora del detalle pedido, una pieza esta compuesta por todas las tareas

    Dim minTemp As Long
    'Dim cantidadTareasDeDetallePedido As Long
    'cantidadTareasDeDetallePedido = 0
    'Dim contadorTareasConAvance As Long
    'contadorTareasConAvance = 0
    'Dim piezasRealizadasDeDetallePedido As Long
    'piezasRealizadasDeDetallePedido = 0

    q = "SELECT (SELECT COUNT(ptp2.id) FROM PlaneamientoTiemposProcesos ptp2 WHERE ptp2.idDetallePedido = dp.id) AS cantidad_tareas," _
      & " IFNULL(SUM(ptpd.cantidad_procesada), 0) As cantidad_piezas_procesadas_tarea" _
      & " FROM detalles_pedidos dp" _
      & " INNER JOIN PlaneamientoTiemposProcesos ptp" _
      & " ON ptp.idDetallePedido = dp.id" _
      & " LEFT JOIN PlaneamientoTiemposProcesosDetalle ptpd" _
      & " ON ptpd.idTiemposProcesos = ptp.id" _
      & " Where dp.id = " & idDetallePedido _
      & " GROUP BY ptp.id"

    Set tmpRS = conectar.RSFactory(q)
    minTemp = -1
    While Not tmpRS.EOF
        If minTemp = -1 Then    'para inicializar con algo
            minTemp = tmpRS!cantidad_piezas_procesadas_tarea
        End If

        If tmpRS!cantidad_piezas_procesadas_tarea < minTemp Then
            minTemp = tmpRS!cantidad_piezas_procesadas_tarea
        End If

        tmpRS.MoveNext
    Wend

    Porcentaje = Math.Round(IIf(porcentajeAvanceTareas = -1, 0, porcentajeAvanceTareas), 2)
    promedio = minTemp

    'me fijo si es conjunto el detalle pedido
    Dim detallePedidoEsConjunto As Boolean
    q = "select 1 from detalles_pedidos dp inner join stock s on dp.idpieza = s.id where dp.id=" & idDetallePedido & " AND s.conjunto = 0"    '0 es conjunto
    Set tmpRS = conectar.RSFactory(q)
    detallePedidoEsConjunto = Not tmpRS.EOF    'si trajo algo es porque es conjunto


    If detallePedidoEsConjunto Then
        'por cada pieza del detalle pedido, calculo el avance
        Dim cantidadPiezasConjunto As Long
        q = "SELECT COUNT(1) as cantidadpiezas FROM detalles_pedidos_conjuntos dpc WHERE dpc.iddetalle_pedido = " & idDetallePedido
        Set tmpRS = conectar.RSFactory(q)
        cantidadPiezasConjunto = tmpRS!cantidadpiezas

        Dim porcentajeTotalDelConjunto As Double
        porcentajeTotalDelConjunto = 0

        Dim tmpRS2 As Recordset
        Dim tmpContador As Double

        q = "SELECT * FROM detalles_pedidos_conjuntos dpc WHERE dpc.iddetalle_pedido = " & idDetallePedido
        Set tmpRS = conectar.RSFactory(q)

        Dim cantPiezasMinima As Long    'Double
        Dim minimoTmp As Long
        Dim tmpCantPosible As Long    'Double

        cantPiezasMinima = -1

        While Not tmpRS.EOF
            q = "SELECT" _
              & " dp.cantidad AS cantidad_pedida," _
              & " ptp.codigoTarea," _
              & " SUM(ptpd.cantidad_procesada) AS cantidad_procesada," _
              & " udf_hasta_cien(((SUM(ptpd.cantidad_procesada) / dp.cantidad) * 100)) AS porcentaje_realizado_tarea," _
              & " ((SELECT (1 / COUNT(ptp2.id)) * 100 FROM PlaneamientoTiemposProcesos ptp2 WHERE ptp2.iddetallepedido = dp.id) * udf_hasta_cien(((SUM(ptpd.cantidad_procesada) / dp.cantidad) * 100))) / 100  AS porcentaje_tarea_pedido" _
              & " FROM detalles_pedidos_conjuntos dp" _
              & " INNER JOIN PlaneamientoTiemposProcesos ptp" _
              & " ON ptp.idDetallePedido = dp.id" _
              & " LEFT JOIN PlaneamientoTiemposProcesosDetalle ptpd" _
              & " ON ptpd.idTiemposProcesos = ptp.id" _
              & " Where dp.id = " & tmpRS!Id & " And ptp.conjunto = 1" _
              & " GROUP BY ptp.id"


            tmpContador = 0
            Set tmpRS2 = conectar.RSFactory(q)
            While Not tmpRS2.EOF
                If Not IsNull(tmpRS2!porcentaje_tarea_pedido) Then
                    tmpContador = tmpContador + tmpRS2!porcentaje_tarea_pedido
                End If
                tmpRS2.MoveNext
            Wend

            'si 100% -> 4% (25 piezas 100/25)
            '12.5% en 4% = 0.05
            porcentajeTotalDelConjunto = porcentajeTotalDelConjunto + ((tmpContador * (100 / cantidadPiezasConjunto)) / 100)

            'calculo el prom de piezas de lo que integra el conjunto
            q = "SELECT" _
              & " (SELECT" _
              & " count(ptp2.id)" _
              & " FROM PlaneamientoTiemposProcesos ptp2" _
              & " WHERE ptp2.idDetallePedido = dp.id) AS cantidad_tareas," _
              & " IFNULL(SUM(ptpd.cantidad_procesada), 0) As cantidad_piezas_procesadas_tarea" _
              & " FROM detalles_pedidos_conjuntos dp" _
              & " INNER JOIN PlaneamientoTiemposProcesos ptp" _
              & " ON ptp.idDetallePedido = dp.id" _
              & " LEFT JOIN PlaneamientoTiemposProcesosDetalle ptpd" _
              & " ON ptpd.idTiemposProcesos = ptp.id" _
              & " Where dp.id = " & tmpRS!Id & " And ptp.conjunto = 1" _
              & " GROUP BY ptp.id"

            minimoTmp = -1
            Set tmpRS2 = conectar.RSFactory(q)
            While Not tmpRS2.EOF
                If minimoTmp = -1 Then minimoTmp = tmpRS2!cantidad_piezas_procesadas_tarea
                If tmpRS2!cantidad_piezas_procesadas_tarea < minimoTmp Then
                    minimoTmp = tmpRS2!cantidad_piezas_procesadas_tarea
                End If
                tmpRS2.MoveNext
            Wend


            tmpCantPosible = 0
            If tmpRS!Cantidad > 0 Then
                tmpCantPosible = Int(minimoTmp / tmpRS!Cantidad)
            End If

            If cantPiezasMinima = -1 Then
                cantPiezasMinima = tmpCantPosible
            End If

            If tmpCantPosible < cantPiezasMinima Then
                cantPiezasMinima = tmpCantPosible
            End If


            tmpRS.MoveNext
        Wend
        If cantPiezasMinima = -1 Then cantPiezasMinima = 0


        'meto lo que medio el porcentaje de avante de tareas del conjunto en el avance de la pieza principal del detalle pedido

        If porcentajeAvanceTareas = -1 Then    'no tiene tareas
            Porcentaje = Math.Round(porcentajeTotalDelConjunto, 2)
        Else
            porcentajeAvanceTareas = (porcentajeTotalDelConjunto * porcentajeAvanceTareas) / 100
            Porcentaje = Math.Round(porcentajeAvanceTareas, 2)    'retorno
        End If



        'falta el promedio de piezas
        'si mintemp = -1, no tiene tareas la pieza principal
        If minTemp = -1 Then
            promedio = cantPiezasMinima
        Else
            promedio = IIf(cantPiezasMinima < minTemp, cantPiezasMinima, minTemp)
        End If
    End If
End Sub



Public Function FindBestPriceByPiezaId(idPieza As Long) As Double
    Dim rs As Recordset
    Dim q As String
    Dim MAXIMO As Double
    Dim ultimo As Double
    q = "SELECT MAX(precio) as max FROM detalles_pedidos WHERE idPieza=" & idPieza

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF And Not rs.BOF Then
        MAXIMO = rs!max
    Else
        MAXIMO = 0
    End If

    q = "SELECT  dp.precio  FROM pedidos p INNER JOIN detalles_pedidos dp ON dp.idPedido=p.id  WHERE idPieza=" & idPieza & "  ORDER BY fechaAprobado DESC LIMIT 1"

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF And Not rs.BOF Then

        ultimo = rs!Precio
    Else
        ultimo = rs!Precio
    End If

    If MAXIMO > ultimo Then FindBestPriceByPiezaId = MAXIMO
    If ultimo <= MAXIMO Then FindBestPriceByPiezaId = ultimo





End Function



