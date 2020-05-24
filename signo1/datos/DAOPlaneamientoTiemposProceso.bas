Attribute VB_Name = "DAOTiemposProceso"

Option Explicit
Dim rs As Recordset
Dim EsConjunto As Boolean
Public Const CAMPO_ITEM As String = "item"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ID_PEDIDO As String = "idPedido"
Public Const CAMPO_ID_PIEZA As String = "idPieza"
Public Const CAMPO_ID_DETALLE_PEDIDO As String = "idDetallePedido"
Public Const CAMPO_ID_DETALLE_PEDIDO_CONJUNTO As String = "idDetallePedidoConj"

Public Const CAMPO_TIEMPO_COTIZADO As String = "TiempoCotizado"
Public Const CAMPO_OPERARIOS_COTIZADO As String = "OperariosCotizado"
Public Const TABLA_TAREA As String = "t"

'me dice si determinada tarea puede finalizarse ya que tiene detalles y la cant procesada es mayor a cero o no tiene detalles
'1 = puede finalizar
'-1 = no puede por tiempos sin terminar
'-2 = no puede por que la cantidad procesada es igual a 0
'-3 = ya esta finalizada
Public Function CanFinalize(idTiempoProceso As Long) As Integer
    Dim tp As PlaneamientoTiempoProceso
    Set tp = FindById(idTiempoProceso)
    If IsSomething(tp) Then
        If tp.FINALIZADO Then
            CanFinalize = -3
            Exit Function
        End If
    End If

    Dim col As Collection
    Dim det As PlaneamientoTiempoProcesoDetalle
    Set col = DAOTiemposProcesosDetalles.FindAllByTiempoProceso(idTiempoProceso)

    Dim todosFinalizados As Boolean: todosFinalizados = True
    Dim cantProc As Double: cantProc = 0

    If col.count <> 0 Then
        For Each det In col
            todosFinalizados = todosFinalizados And (CDbl(det.FechaFinTarea) <> 0)
            cantProc = cantProc + det.CantidadProcesada
        Next det

        If Not todosFinalizados Then
            CanFinalize = -1
        ElseIf cantProc = 0 Then
            CanFinalize = -2
        Else
            CanFinalize = 1
        End If

        Exit Function
    End If

    CanFinalize = 1
End Function

Public Function Finalize(idTiempoProceso As Long) As Boolean
    If CanFinalize(idTiempoProceso) Then
        Dim q As String
        q = "UPDATE PlaneamientoTiemposProcesos SET fechaFin = NOW() WHERE id = " & idTiempoProceso
        Finalize = conectar.execute(q)
    Else
        Finalize = False
    End If
End Function


Public Function GetSectoresByIdPedido(idpedido As Long) As Collection
    Dim index As Dictionary
    Set rs = conectar.RSFactory("select sec.* FROM PlaneamientoTiemposProcesos p INNER JOIN tareas t ON t.id=p.codigoTarea INNER JOIN sectores sec ON t.id_sector=sec.id  WHERE idPedido=" & idpedido & " GROUP BY sec.id")
    conectar.BuildFieldsIndex rs, index
    Dim col As New Collection
    While Not rs.EOF
        col.Add DAOSectores.Map(rs, index, DAOSectores.TABLA_SECTOR), CStr(rs.Fields(index("sec.id")).value)
        rs.MoveNext
    Wend
    Set GetSectoresByIdPedido = col
End Function


Public Function FindById(id As Long) As PlaneamientoTiempoProceso
    Dim col As Collection
    Set col = FindAll("ptp.id = " & id, True)

    If col.count > 0 Then
        Set FindById = col.item(1)
    Else
        Set FindById = Nothing
    End If

End Function

Public Function GetAvancesOTBySectorAndTareaFinalizada(Ot As Long) As Collection
    Dim filter As String
    Dim orden_trabajo_id As Object
    Dim q As String

    q = "SELECT " _
        & " s.*, " _
        & " SUM(1) AS TotalPorSector,  SUM(ptp.fechaFin > 0) AS TotalFinalizado ,  (1-(  SUM(ptp.fechaFin > 0)/SUM(1))) AS IndiceTotalFinalizado " _
        & "From" _
        & "PlaneamientoTiemposProcesos ptp   LEFT JOIN tareas t     ON ptp.codigoTarea = t.id " _
        & "LEFT JOIN sectores s     ON t.`id_sector` = s.`id`  LEFT JOIN detalles_pedidos dp " _
        & " ON ptp.idDetallePedido = dp.id Where  1=1 and  ptp.idPedido in (" & Ot & ") GROUP BY s.id "

    Dim index As Dictionary
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, index
    Dim col As New Collection
    Dim P As DTOSectoresTiempo

    While Not rs.EOF
        Set P = New DTOSectoresTiempo
        Set P.Sector = DAOSectores.Map(rs, index, "sectores")

        col.Add P, CStr(P.Sector.id)
        rs.MoveNext
    Wend

    Set GetAvancesOTBySectorAndTareaFinalizada = col
End Function
Public Function FindAll(Optional filter As String = vbNullString, Optional WithDetalle As Boolean = False, Optional withPlanificacion As Boolean = False) As Collection
    Dim q As String
    q = "SELECT *"
    'If WithDetalle Then q = q & ",ptpd.*,u.*,p.*"

    q = q & " FROM PlaneamientoTiemposProcesos ptp" _
        & " LEFT JOIN tareas t" _
        & " ON t.id = ptp.codigoTarea" _
        & " LEFT JOIN sectores sec" _
        & " ON sec.id = t.id_sector" _

If WithDetalle Then
        q = q & " LEFT JOIN PlaneamientoTiemposProcesosDetalle ptpd  ON ptpd.idTiemposProcesos=ptp.id"
        q = q & " LEFT JOIN personal p ON ptpd.legajo=p.id "
    End If

    If withPlanificacion Then
        q = q & " LEFT JOIN PlaneamientoTiemposProcesosPlanificacion ptpp  ON ptpp.id_ptp=ptp.id "
    End If

    q = q & " WHERE 1 = 1 "
    If LenB(filter) > 0 Then q = q & " AND " & filter

    Dim tmpTiempoProc As PlaneamientoTiempoProceso

    Dim index As Dictionary
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, index
    Dim col As New Collection

    While Not rs.EOF

        If funciones.BuscarEnColeccion(col, CStr(rs.Fields(index("ptp.id")))) Then
            Set tmpTiempoProc = col.item(CStr(rs.Fields(index("ptp.id"))))
        Else
            Set tmpTiempoProc = New PlaneamientoTiempoProceso
            Set tmpTiempoProc.Detalles = New Collection

            tmpTiempoProc.id = rs.Fields(index("ptp.id"))
            tmpTiempoProc.idDetallePedido = rs.Fields(index("ptp.idDetallePedido"))
            tmpTiempoProc.idDetallePedidoConj = rs.Fields(index("ptp.idDetallePedidoConj"))
            tmpTiempoProc.idpedido = rs.Fields(index("ptp.idPedido"))
            tmpTiempoProc.TiempoCotizado = GetValue(rs, index, "ptp", DAOTiemposProceso.CAMPO_TIEMPO_COTIZADO)
            tmpTiempoProc.OperariosCotizado = GetValue(rs, index, "ptp", DAOTiemposProceso.CAMPO_OPERARIOS_COTIZADO)
            tmpTiempoProc.item = GetValue(rs, index, "ptp", DAOTiemposProceso.CAMPO_ITEM)
            tmpTiempoProc.idPieza = rs.Fields(index("ptp.idPieza"))
            tmpTiempoProc.FechaFin = GetValue(rs, index, "ptp", "fechaFin")
            tmpTiempoProc.EsConjunto = GetValue(rs, index, "ptp", "conjunto")
            tmpTiempoProc.Observacion = GetValue(rs, index, "ptp", "observacion_agregado")


            If withPlanificacion Then Set tmpTiempoProc.Planificacion = DAOTiempoProcesoPlanificado.Map(rs, index, "ptpp", "ptp")
            Set tmpTiempoProc.Tarea = DAOTareas.Map(rs, index, "t", , "sec")


        End If

        If WithDetalle Then
            If Not IsNull(rs.Fields(index("ptpd.id"))) Then
                tmpTiempoProc.Detalles.Add DAOTiemposProcesosDetalles.Map(rs, index, "ptpd", , "p"), CStr(rs.Fields(index("ptpd.id")))
            End If
        End If

        If Not funciones.BuscarEnColeccion(col, CStr(rs.Fields(index("ptp.id")))) Then
            col.Add tmpTiempoProc, CStr(tmpTiempoProc.id)
        End If
        rs.MoveNext
    Wend

    Set FindAll = col
End Function


Public Function FindAllByDetallePedidoId(idDetallePedido As Long, Optional idDetallePedidoConj As Long = 0, Optional WithDetalle As Boolean = False, Optional withPlanificacion As Boolean = False, Optional tareaId As Long = 0) As Collection
    Dim F As String
    If idDetallePedidoConj = 0 Then
        F = " ptp.IdDetallePedido = " & idDetallePedido & " AND  ptp.idDetallePedidoConj = 0"
    Else
        F = " ptp.idDetallePedidoConj = " & idDetallePedidoConj
    End If
    If tareaId <> 0 Then
        F = F & " AND ptp.codigoTarea = " & tareaId
    End If
    Set FindAllByDetallePedidoId = FindAll(F, WithDetalle, withPlanificacion)
End Function

'no se tendria que utilizar mas, ahora usar FindAllByDetallePedidoId
Public Function FindAllByDetallePedidoIdAndPiezaId(idDetallePedido As Long, idPieza As Long, Optional WithDetalle As Boolean = False, Optional withPlanificacion As Boolean = False) As Collection
    Dim F As String
    F = " ptp.idPieza = " & idPieza & " And ptp.IdDetallePedido = " & idDetallePedido
    Set FindAllByDetallePedidoIdAndPiezaId = FindAll(F, WithDetalle, withPlanificacion)
End Function

Public Function SectorColl2RS(sectores As Collection) As ADODB.Recordset
    Dim r As New Recordset
    With r
        .Fields.Append "idsector", adVarChar, 255, adFldUpdatable      ' And adFldIsNullable"
        .Fields.Append "sector", adVarChar, 255, adFldUpdatable     ' And adFldIsNullab
    End With

    Dim sec As clsSector
    r.Open
    For Each sec In sectores
        r.AddNew
        r.Fields("idsector").value = sec.id
        r.Fields("sector").value = sec.Sector
        r.MoveNext
    Next sec

    Set SectorColl2RS = r
End Function

Public Function Map(ByRef rs As Recordset, ByRef indice As Dictionary, ByRef tabla As String, Optional ByRef tablaTarea As String, Optional tablaPlanificacion As String = vbNullString, Optional tablaSectores As String = vbNullString) As PlaneamientoTiempoProceso

    Dim ptp As PlaneamientoTiempoProceso
    Dim id As Variant
    id = GetValue(rs, indice, tabla, CAMPO_ID)

    If id > 0 Then
        Set ptp = New PlaneamientoTiempoProceso
        ptp.id = id
        ptp.TiempoCotizado = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_TIEMPO_COTIZADO)
        ptp.OperariosCotizado = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_OPERARIOS_COTIZADO)
        ptp.idDetallePedido = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_ID_DETALLE_PEDIDO)
        ptp.idpedido = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_ID_PEDIDO)
        ptp.idPieza = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_ID_PIEZA)
        ptp.item = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_ITEM)
        ptp.EsConjunto = GetValue(rs, indice, tabla, "conjunto")
        ptp.idDetallePedidoConj = GetValue(rs, indice, tabla, DAOTiemposProceso.CAMPO_ID_DETALLE_PEDIDO_CONJUNTO)

        If LenB(tablaTarea) > 0 Then Set ptp.Tarea = DAOTareas.Map(rs, indice, tablaTarea, , tablaSectores)
        If LenB(tablaPlanificacion) > 0 Then Set ptp.Planificacion = DAOTiempoProcesoPlanificado.Map(rs, indice, "ptpp", "ptp")

    End If

    Set Map = ptp
End Function


Private Function CargarTiempoProceso(deta As DetalleOrdenTrabajo, Pieza As Pieza, Optional detalleOrdenTrabajoConjunto As DetalleOTConjuntoDTO = Nothing) As Boolean
    Dim desa As DesarrolloManoObra
    Dim col As New Collection

    Dim q As String
    Dim ptpId As Long
    Dim tpPlanif As TiempoProcesoPlanificado

    CargarTiempoProceso = True

    If IsSomething(detalleOrdenTrabajoConjunto) Then


        For Each desa In detalleOrdenTrabajoConjunto.Pieza.desarrollosManoObra
            q = "insert into PlaneamientoTiemposProcesos (idPedido, idPieza, idDetallePedido, idDetallePedidoConj, codigoTarea, OperariosCotizado,TiempoCotizado,item,conjunto,observacion_agregado) values " _
                & "( " & deta.OrdenTrabajo.id & "," & Pieza.id & "," & deta.id & "," & detalleOrdenTrabajoConjunto.id & "," & desa.Tarea.id & "," & conectar.Escape(desa.Cantidad) & "," & conectar.Escape(desa.Tiempo) & "," & conectar.Escape(deta.item) & "," & Abs(deta.Pieza.EsConjunto) & "," & Escape(desa.detalle) & ")"

            If conectar.execute(q) Then
                ptpId = 0
                ptpId = conectar.UltimoId2()
                If ptpId = 0 Then
                    GoTo err1
                Else
                    conectar.execute "DELETE FROM PlaneamientoTiemposProcesosPlanificacion where id_ptp = " & ptpId
                    'ver
                    If deta.EstadoProceso = EstProcDetOT_ProcesoDefinido Or deta.EstadoProceso = EstProcDetOT_ProcesoNoDefinido Then
                        Set tpPlanif = New TiempoProcesoPlanificado
                        tpPlanif.idTiempoProceso = ptpId
                        tpPlanif.Color = 65280
                        tpPlanif.Critica = False
                        tpPlanif.Inicio = Date + TimeSerial(7, 0, 0)
                        tpPlanif.Fin = DateAdd("n", desa.Tiempo * desa.Cantidad, tpPlanif.Inicio)
                        DAOTiempoProcesoPlanificado.Save tpPlanif
                    End If
                End If
            Else
                GoTo err1
            End If


        Next desa
    Else

        For Each desa In Pieza.desarrollosManoObra

            q = "insert into PlaneamientoTiemposProcesos (idPedido, idPieza, idDetallePedido, codigoTarea, OperariosCotizado,TiempoCotizado,item,conjunto,observacion_agregado) values " _
                & "( " & deta.OrdenTrabajo.id & "," & Pieza.id & "," & deta.id & "," & desa.Tarea.id & "," & desa.Cantidad & "," & desa.Tiempo & "," & conectar.Escape(deta.item) & "," & Abs(deta.Pieza.EsConjunto) & "," & Escape(desa.detalle) & ")"

            If conectar.execute(q) Then
                ptpId = 0
                ptpId = conectar.UltimoId2()
                If ptpId = 0 Then
                    GoTo err1
                Else
                    conectar.execute "DELETE FROM PlaneamientoTiemposProcesosPlanificacion where id_ptp = " & ptpId
                    'ver
                    If deta.EstadoProceso = EstProcDetOT_ProcesoDefinido Or deta.EstadoProceso = EstProcDetOT_ProcesoNoDefinido Then
                        Set tpPlanif = New TiempoProcesoPlanificado
                        tpPlanif.idTiempoProceso = ptpId
                        tpPlanif.Color = 65280
                        tpPlanif.Critica = False
                        tpPlanif.Inicio = Date + TimeSerial(7, 0, 0)
                        tpPlanif.Fin = DateAdd("n", desa.Tiempo * desa.Cantidad, tpPlanif.Inicio)
                        DAOTiempoProcesoPlanificado.Save tpPlanif
                    End If
                End If
            Else
                GoTo err1
            End If
        Next desa

    End If



    If Pieza.EsConjunto Then
        Set col = DAODetalleOrdenTrabajo.FindAllConjunto(deta.id, Pieza.id, , True)

        For Each detalleOrdenTrabajoConjunto In col
            CargarTiempoProceso = CargarTiempoProceso(deta, detalleOrdenTrabajoConjunto.Pieza, detalleOrdenTrabajoConjunto)
        Next detalleOrdenTrabajoConjunto
    End If

    Exit Function
err1:
    CargarTiempoProceso = False
End Function
Public Function crear(Pedido As OrdenTrabajo, Optional progressbar As Object) As Boolean


    On Error GoTo errt
    crear = True

    Dim oper_cot As Long
    Dim time_cut As Double
    Dim item As String



    Dim deta As DetalleOrdenTrabajo
    Dim desa As DesarrolloManoObra
    Dim Pieza As Pieza


    EsConjunto = False
    conectar.execute "delete from PlaneamientoTiemposProcesos where idPedido=" & Pedido.id
    progressbar.min = 0
    progressbar.max = Pedido.Detalles.count
    Dim c As Long
    c = 0
    For Each deta In Pedido.Detalles
        c = c + 1
        progressbar.value = c
        Set deta.Pieza = DAOPieza.FindById(deta.Pieza.id, FL_4, True)

        If Not CargarTiempoProceso(deta, deta.Pieza) Then GoTo errt


    Next deta


    Exit Function
errt:
    crear = False

End Function

Public Function GetAvancesHsPorOTs(OtsId As Collection) As Collection



    Dim q As String
    q = "SELECT " _
        & " ifnull(SUM((TIME_TO_SEC(TIMEDIFF(ptpd.fin,ptpd.inico)) / 60) / 60),0) AS sum_horas," _
        & " t.id_sector," _
        & " ptp.codigoTarea" _
        & " FROM PlaneamientoTiemposProcesos ptp" _
        & " INNER JOIN PlaneamientoTiemposProcesosDetalle ptpd" _
        & " ON ptpd.idTiemposProcesos = ptp.id" _
        & " INNER JOIN tareas t" _
        & " ON t.id = ptp.codigoTarea" _
        & " Where 1 = 1"

    Dim orden_trabajo_id As Variant
    Dim filter As String
    For Each orden_trabajo_id In OtsId
        filter = filter & orden_trabajo_id & ", "
    Next orden_trabajo_id

    If OtsId.count > 0 Then
        q = q & " AND  ptp.idpedido IN (" & Left$(filter, Len(filter) - 2) & ")"
    End If
    q = q & " GROUP BY t.id_sector, ptp.codigoTarea"


    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)

    Dim sectoresTiempo As New Collection

    Dim tiempoSectorDTO As DTOSectoresTiempo
    Dim tareaTiempoDTO As DTOTareaTiempo

    Dim tarea_id As String
    Dim sector_id As String

    While Not rs.EOF
        tarea_id = rs!codigoTarea
        sector_id = rs!id_sector


        If BuscarEnColeccion(sectoresTiempo, sector_id) Then
            Set tiempoSectorDTO = sectoresTiempo.item(sector_id)
        Else
            Set tiempoSectorDTO = New DTOSectoresTiempo
            Set tiempoSectorDTO.Sector = DAOSectores.GetById(sector_id)
            sectoresTiempo.Add tiempoSectorDTO, sector_id
        End If


        If BuscarEnColeccion(tiempoSectorDTO.ListaDtoTareaTiempo, tarea_id) Then
            Set tareaTiempoDTO = tiempoSectorDTO.ListaDtoTareaTiempo.item(tarea_id)
        Else
            Set tareaTiempoDTO = New DTOTareaTiempo
            Set tareaTiempoDTO.Tarea = DAOTareas.FindById(CLng(tarea_id))
            tiempoSectorDTO.ListaDtoTareaTiempo.Add tareaTiempoDTO, tarea_id
        End If

        tareaTiempoDTO.Tiempo = tareaTiempoDTO.Tiempo + rs!sum_horas

        rs.MoveNext
    Wend

    Set GetAvancesHsPorOTs = sectoresTiempo
End Function
