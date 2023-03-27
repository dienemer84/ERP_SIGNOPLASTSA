Attribute VB_Name = "DAOTiemposProcesosDetalles"
Option Explicit
Dim rs As Recordset
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_LEGAJO As String = "legajo"
Public Const CAMPO_FECHA As String = "fecha"
Public Const CAMPO_INICIO As String = "inico"
Public Const CAMPO_FIN As String = "fin"
Public Const CAMPO_CANTIDAD_PROCESADA As String = "cantidad_procesada"
Public Const CAMPO_ID_PLAN_TIEMPO_PROCESO As String = "idTiemposProcesos"
Public Const TABLA_PERSONAL As String = "per"
Public Const TABLA_TIEMPO_PROCESO_DETALLE As String = "ptpd"
Public Const TABLA_USUARIO As String = "usu"


Public Function FindAllByTiempoProceso(tiempoProcesoId As Long) As Collection
    Dim F As String: F = "ptpd." & DAOTiemposProcesosDetalles.CAMPO_ID_PLAN_TIEMPO_PROCESO & " = " & tiempoProcesoId & " and inico <> '0000-00-00 00:00:00'"
    Set FindAllByTiempoProceso = DAOTiemposProcesosDetalles.FindAll(F)
End Function


Public Function FindById(Id As Long) As PlaneamientoTiempoProcesoDetalle
    Dim F As String: F = "ptpd." & DAOTiemposProcesosDetalles.CAMPO_ID & " = " & Id
    Dim col As Collection
    Set col = DAOTiemposProcesosDetalles.FindAll(F)
    If col.count > 0 Then
        Set FindById = col.item(1)
    Else
        Set FindById = Nothing
    End If
End Function

Public Function FindFirstWithoutFinishByEmpleadoIdAndTiempoProceso(empleadoId As Long, tiempoProcesoId As Long) As PlaneamientoTiempoProcesoDetalle
    Dim F As String
    F = "ptpd.legajo = " & empleadoId & " AND fin IS NULL  and inico <> '0000-00-00 00:00:00' AND ptpd.idTiemposProcesos = " & tiempoProcesoId
    Dim det As PlaneamientoTiempoProcesoDetalle
    Dim col As Collection
    Set col = FindAll(F)

    If col.count = 0 Then
        Set FindFirstWithoutFinishByEmpleadoIdAndTiempoProceso = Nothing
    Else
        Set FindFirstWithoutFinishByEmpleadoIdAndTiempoProceso = col.item(1)
    End If

End Function

Public Function FindAllWithoutFinish() As Collection
    Dim F As String
    F = "fin IS NULL and inico <> '0000-00-00 00:00:00'"
    Dim det As PlaneamientoTiempoProcesoDetalle
    Dim col As Collection
    Set col = FindAll(F)

    If col.count = 0 Then
        Set FindAllWithoutFinish = Nothing
    Else
        Set FindAllWithoutFinish = col
    End If

End Function
Public Function FindAllWithoutFinishByEmpleado(IdEmpleado As Long) As Collection
    Dim F As String
    F = "fin IS NULL  and inico <> '0000-00-00 00:00:00' and per.id=" & IdEmpleado
    Dim det As PlaneamientoTiempoProcesoDetalle
    Dim col As Collection
    Set col = FindAll(F)

    If col.count = 0 Then
        Set FindAllWithoutFinishByEmpleado = Nothing
    Else
        Set FindAllWithoutFinishByEmpleado = col
    End If

End Function

Public Function FindAllAsignedNotFinishedByEmpleado(IdEmpleado As Long) As Collection
    Dim F As String
    F = "ptpd.legajo = " & IdEmpleado & " AND inico = '0000-00-00 00:00:00' AND ptp.fechaFin = '0000-00-00 00:00:00'"
    Set FindAllAsignedNotFinishedByEmpleado = FindAll(F)
End Function

Public Function FindAllAsignedByEmpleadoAndTiempoProcesoAndTareaId(IdEmpleado As Long, tiempoProcesoId As Dictionary, tareaId As Long) As Collection
    Dim F As String
    F = "idTiemposProcesos IN (" & funciones.JoinDictionaryKeyValues(tiempoProcesoId, ", ") & ") AND ptpd.legajo = " & IdEmpleado & " AND ptp.codigoTarea = " & tareaId & " AND inico = '0000-00-00 00:00:00'"
    Set FindAllAsignedByEmpleadoAndTiempoProcesoAndTareaId = FindAll(F)
End Function


Public Function FindFirstWithoutFinishByEmpleadoId(empleadoId As Long) As PlaneamientoTiempoProcesoDetalle
    Dim F As String
    F = "ptpd.legajo = " & empleadoId & " AND fin IS NULL  and inico <> '0000-00-00 00:00:00'"
    Dim det As PlaneamientoTiempoProcesoDetalle
    Dim col As Collection
    Set col = FindAll(F)
    If col.count = 0 Then
        Set FindFirstWithoutFinishByEmpleadoId = Nothing
    Else
        Set FindFirstWithoutFinishByEmpleadoId = col.item(1)
    End If
End Function

Public Function FindAllPeriodosConProceso(periodos As TipoPeriodo) As Collection
    Dim dto As DTOTiempoProcesoDetalle
    Dim col As New Collection
    Dim q As String
    If periodos = TipoPeriodoMes Then
        q = "SELECT distinct MONTH(inico) as mes, year(inico) as anio FROM PlaneamientoTiemposProcesosDetalle WHERE MONTH(inico)=MONTH(fin) AND YEAR(inico)=YEAR(fin) AND (inico >0 OR fin > 0)  "
    ElseIf periodos = TipoPeriodoAño Then
        q = "SELECT distinct YEAR(inico) as anio FROM PlaneamientoTiemposProcesosDetalle WHERE YEAR(inico)=YEAR(fin) AND (inico >0 OR fin > 0)  "
    End If
    Set rs = conectar.RSFactory(q)
    Dim i As Long
    i = 0
    While Not rs.EOF
        i = i + 1
        Set dto = New DTOTiempoProcesoDetalle
        dto.Año = rs!anio
        If periodos = TipoPeriodoMes Then
            dto.mes = rs!mes
            dto.mostrar = dto.Año & " - " & enums.EnumPeriodo(dto.mes)
        Else
            dto.mostrar = dto.Año
            dto.mes = 0
        End If
        dto.indice = i
        col.Add dto, str(i)
        rs.MoveNext
    Wend
    Set FindAllPeriodosConProceso = col
End Function

Public Function FindAllPorPeriodoAgrupado(filter As String) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim indice As Dictionary
    Dim strsql As String
    strsql = "SELECT *,ifnull(SUM(TIMESTAMPDIFF(SECOND,inico,fin)/60)/60,0) AS suma_dif " _
           & "FROM PlaneamientoTiemposProcesosDetalle ptpd " _
           & "LEFT JOIN PlaneamientoTiemposProcesos ptp " _
           & "ON ptpd.idTiemposProcesos = ptp.id " _
           & " LEFT JOIN personal per ON per.id = ptpd.legajo" _
           & " where 1=1 "

    If Len(filter) > 0 Then strsql = strsql & " AND " & filter

    strsql = strsql & " group by ptpd.legajo"


    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, indice
    Dim det As PlaneamientoTiempoProcesoDetalle

    While Not rs.EOF And Not rs.BOF
        Set det = Map(rs, indice, TABLA_TIEMPO_PROCESO_DETALLE, , DAOTiemposProcesosDetalles.TABLA_PERSONAL, "ptp")
        det.PlaneamientoTiempoProceso.TiempoTotalReal = rs!suma_dif

        col.Add det, CStr(det.Id)
        rs.MoveNext
    Wend
    Set FindAllPorPeriodoAgrupado = col
    Exit Function
err1:
    MsgBox Err.Description
    Set FindAllPorPeriodoAgrupado = New Collection
End Function



Public Function FindAll(filter As String, Optional ptpExtraJoin As String = vbNullString) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim indice As Dictionary
    Dim strsql As String
    strsql = "SELECT  *" _
           & " FROM PlaneamientoTiemposProcesosDetalle ptpd" _
           & " LEFT JOIN usuarios usu ON ptpd.id_usuario = usu.id" _
           & " LEFT JOIN personal per ON per.id = ptpd.legajo" _
           & " LEFT JOIN PlaneamientoTiemposProcesos ptp ON ptp.Id = ptpd.idTiemposProcesos " & ptpExtraJoin

    strsql = strsql & " LEFT JOIN tareas t ON t.id = ptp.codigoTarea" _
           & " LEFT JOIN sectores s ON s.id = t.id_sector" _
           & " WHERE 1=1"
    If Len(filter) > 0 Then strsql = strsql & " AND " & filter



    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, indice
    Dim det As PlaneamientoTiempoProcesoDetalle

    While Not rs.EOF And Not rs.BOF
        Set det = Map(rs, indice, TABLA_TIEMPO_PROCESO_DETALLE, TABLA_USUARIO, "per", "ptp", "t", "s")
        col.Add det, CStr(det.Id)
        rs.MoveNext
    Wend
    Set FindAll = col
    Exit Function
err1:
    MsgBox Err.Description
    Set FindAll = New Collection
End Function

Public Function Map(ByRef rs As Recordset, ByRef indice As Dictionary, ByRef tabla As String, Optional ByRef tablaUsuario As String = vbNullString, Optional ByRef tablaPersonal As String = vbNullString, Optional ByRef tablaPTP As String = vbNullString, Optional ByRef tablaTarea As String = vbNullString, Optional ByRef tablaSectores As String = vbNullString) As PlaneamientoTiempoProcesoDetalle

    Dim ptpd As PlaneamientoTiempoProcesoDetalle
    Dim Id As Variant
    Id = GetValue(rs, indice, tabla, CAMPO_ID)

    If Id > 0 Then
        Set ptpd = New PlaneamientoTiempoProcesoDetalle
        ptpd.Id = Id
        ptpd.CantidadProcesada = GetValue(rs, indice, tabla, DAOTiemposProcesosDetalles.CAMPO_CANTIDAD_PROCESADA)
        ptpd.FechaCarga = GetValue(rs, indice, tabla, DAOTiemposProcesosDetalles.CAMPO_FECHA)
        ptpd.FechaFinTarea = GetValue(rs, indice, tabla, DAOTiemposProcesosDetalles.CAMPO_FIN)
        ptpd.FechaInicioTarea = GetValue(rs, indice, tabla, DAOTiemposProcesosDetalles.CAMPO_INICIO)
        ptpd.IdPlaneamientoTiempoProceso = GetValue(rs, indice, tabla, DAOTiemposProcesosDetalles.CAMPO_ID_PLAN_TIEMPO_PROCESO)
        ptpd.legajo = GetValue(rs, indice, tabla, DAOTiemposProcesosDetalles.CAMPO_LEGAJO)
        If LenB(tablaUsuario) > 0 Then Set ptpd.usuario = DAOUsuarios.Map(rs, indice, tablaUsuario)
        If LenB(tablaPersonal) > 0 Then Set ptpd.Empleado = DAOEmpleados.Map2(rs, indice, tablaPersonal)
        If LenB(tablaPTP) > 0 Then Set ptpd.PlaneamientoTiempoProceso = DAOTiemposProceso.Map(rs, indice, tablaPTP, tablaTarea, , tablaSectores)
    End If

    Set Map = ptpd
End Function


Public Function Save(tpd As PlaneamientoTiempoProcesoDetalle, Optional ByVal replaceInicio As Boolean = False) As Boolean

    Dim q As String

    If tpd.Id = 0 Then
        q = "INSERT INTO PlaneamientoTiemposProcesosDetalle" _
          & " (idTiemposProcesos," _
          & " legajo," _
          & " fecha," _
          & " inico," _
          & " fin," _
          & " cantidad_procesada," _
          & " id_usuario)" _
          & " Values (" _
          & conectar.Escape(tpd.IdPlaneamientoTiempoProceso) & "," _
          & conectar.GetEntityId(tpd.Empleado) & "," _
          & conectar.Escape(Now) & "," _
          & conectar.Escape(IIf(replaceInicio, "'0000-00-00 00:00:00'", tpd.FechaInicioTarea)) & "," _
          & conectar.Escape(tpd.FechaFinTarea) & "," _
          & conectar.Escape(tpd.CantidadProcesada) & "," _
          & conectar.Escape(funciones.GetUserObj().Id) & ")"
    Else
        q = "UPDATE PlaneamientoTiemposProcesosDetalle" _
          & " SET" _
          & " inico = " & conectar.Escape(tpd.FechaInicioTarea) & "," _
          & " fin = " & conectar.Escape(tpd.FechaFinTarea) & "," _
          & " cantidad_procesada = " & conectar.Escape(tpd.CantidadProcesada) _
          & " Where id = " & tpd.Id
    End If

    Save = conectar.execute(q)
End Function
Public Function FindPromedioByTareaOfPieza(idTarea As Long, idPieza As Long) As Double
    Dim A As Double
    Dim q As String
    On Error Resume Next
    q = " SELECT " _
      & "SUM(TIMESTAMPDIFF(SECOND,inico,fin)/60) AS b,SUM(cantidad_procesada), t.codigoTarea, t.idPieza, IF(ta.cantxproc>0,(SUM(TIMESTAMPDIFF(SECOND,inico,fin)/60))/SUM(d.cantidad_procesada),SUM(TIMESTAMPDIFF(SECOND,inico,fin)/60)) AS promedio " _
      & " FROM PlaneamientoTiemposProcesosDetalle d " _
      & " INNER JOIN PlaneamientoTiemposProcesos t" _
      & " ON d.idTiemposProcesos = t.id" _
      & " INNER JOIN tareas ta ON " _
      & " t.codigoTarea = ta.Id " _
      & " Where t.estado = 1" _
      & " AND t.idpieza = " & idPieza _
      & " and t.codigoTarea=" & idTarea _
      & " GROUP BY t.idPieza, t.codigoTarea"
    A = 0
    A = conectar.RSFactory(q)!promedio

    FindPromedioByTareaOfPieza = A
End Function
