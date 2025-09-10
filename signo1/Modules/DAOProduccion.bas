Attribute VB_Name = "DAOProduccion"
Option Explicit
Public LastError As String

Public Function Save(r As clsFilaPlanoRow) As Boolean
    On Error GoTo err1
    Dim q As String
    Dim ra As Long

    ' Intentar actualizar primero
    q = "UPDATE sp.detalles_pedidos_conjuntos_avance SET " & _
        "a_cant_recibida=" & EscapeNum(r.CantRecibida) & "," & _
        "a_cant_fabricada=" & EscapeNum(r.CantFabricada) & "," & _
        "a_cant_scrap=" & EscapeNum(r.CantScrap) & "," & _
        "a_fecha_inicio=" & EscapeDate(r.FechaInicio) & "," & _
        "a_fecha_fin=" & EscapeDate(r.FechaFin) & "," & _
        "a_recibio=" & EscapeNum(r.UsuarioRecibio) & "," & _
        "a_siguiente_proceso=" & EscapeStr(r.ProcesoSiguiente) & _
        " WHERE id_detalle_pedido=" & EscapeNum(r.IdTabla) & _
        " AND id_sector=" & EscapeNum(r.idSector)

    conectar.ExecuteRa q, ra

    ' Si no existía registro, insertar
    If ra = 0 Then
        q = "INSERT INTO sp.detalles_pedidos_conjuntos_avance " & _
            "(id_detalle_pedido,id_sector,a_cant_recibida,a_cant_fabricada,a_cant_scrap," & _
            "a_fecha_inicio,a_fecha_fin,a_recibio,a_siguiente_proceso) VALUES (" & _
            EscapeNum(r.IdTabla) & "," & EscapeNum(r.idSector) & "," & _
            EscapeNum(r.CantRecibida) & "," & EscapeNum(r.CantFabricada) & "," & _
            EscapeNum(r.CantScrap) & "," & EscapeDate(r.FechaInicio) & "," & _
            EscapeDate(r.FechaFin) & "," & EscapeNum(r.UsuarioRecibio) & "," & _
            EscapeStr(r.ProcesoSiguiente) & ")"
        conectar.execute q
    End If

    Save = True
    Exit Function

err1:
    LastError = Err.Description
    Save = False
End Function


Public Function FindAllConjuntoProduccion( _
                    Optional ByVal idDetallePedido As Long = 0, _
                    Optional ByVal idPiezaPadre As Long = 0, _
                    Optional ByRef filter As String = vbNullString, _
                    Optional ByVal withDesarrolloManoObra As Boolean = False, _
                    Optional ByVal idDetalleConjunto As Long = 0, _
                    Optional ByVal idSector As Long = 0 _
                ) As Collection

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim detaOrdenesTrabajo As New Collection

    q = ""
    q = q & "SELECT dpc.*, s.*, dp.*, dpca.* " & _
            "FROM detalles_pedidos_conjuntos dpc " & _
            "LEFT JOIN stock s ON s.id = dpc.idPieza " & _
            "LEFT JOIN detalles_pedidos dp ON dp.id = dpc.idDetalle_pedido " & _
            "LEFT JOIN ( " & _
            "   SELECT id_detalle_pedido, id_sector, MAX(id) AS max_id " & _
            "   FROM detalles_pedidos_conjuntos_avance " & _
            IIf(idSector > 0, "   WHERE id_sector = " & idSector & " ", "") & _
            "   GROUP BY id_detalle_pedido, id_sector " & _
            ") last ON last.id_detalle_pedido = dpc.id " & _
            "LEFT JOIN detalles_pedidos_conjuntos_avance dpca ON dpca.id = last.max_id " & _
            "WHERE 1=1 "
    
    If idDetallePedido > 0 Then q = q & " AND dpc.idDetalle_Pedido = " & idDetallePedido
    If idDetalleConjunto > 0 Then q = q & " AND dpc.id = " & idDetalleConjunto
    If idPiezaPadre > 0 Then q = q & " AND dpc.idPiezaPadre = " & idPiezaPadre
    If LenB(filter) > 0 Then q = q & " AND " & filter

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    'Set FindAll = New Collection

    Const piezaTabla As String = "s"
    Dim tmpDeta As DetalleOTConjuntoDTO

    While Not rs.EOF
        Set tmpDeta = DAOProduccion.MapConjuntoProduccion(rs, fieldsIndex, "dpc", piezaTabla, "dp", "dpca")

        detaOrdenesTrabajo.Add tmpDeta, CStr(tmpDeta.Id)
        rs.MoveNext
    Wend

    Set FindAllConjuntoProduccion = detaOrdenesTrabajo

End Function

Public Function MapConjuntoProduccion(ByRef rs As Recordset, _
                            ByRef fieldsIndex As Dictionary, _
                            ByRef tableNameOrAlias As String, _
                            Optional ByRef piezaTableNameOrAlias As String = vbNullString, _
                            Optional ByRef detallePedidoTableNameOrAlias As String = vbNullString, _
                            Optional ByRef ProduccionAlias As String = vbNullString _
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
        tmpDeta.idPedido = GetValue(rs, fieldsIndex, tableNameOrAlias, "idPedido")
        tmpDeta.IdentificadorPosicion = GetValue(rs, fieldsIndex, tableNameOrAlias, "identificador_posicion")
        tmpDeta.CantidadTotalStatic = GetValue(rs, fieldsIndex, tableNameOrAlias, "cantidad_total_static")
        
        tmpDeta.CantidadRecibida = GetValue(rs, fieldsIndex, ProduccionAlias, "a_cant_recibida")
        tmpDeta.CantidadFabricada = GetValue(rs, fieldsIndex, ProduccionAlias, "a_cant_fabricada")
        tmpDeta.CantidadScrap = GetValue(rs, fieldsIndex, ProduccionAlias, "a_cant_scrap")
        tmpDeta.FechaInicio = GetValue(rs, fieldsIndex, ProduccionAlias, "a_fecha_inicio")
        tmpDeta.FechaFin = GetValue(rs, fieldsIndex, ProduccionAlias, "a_fecha_fin")
        tmpDeta.Recibio = GetValue(rs, fieldsIndex, ProduccionAlias, "a_recibio")
        tmpDeta.SiguienteProceso = GetValue(rs, fieldsIndex, ProduccionAlias, "a_siguiente_proceso")
        
        If LenB(piezaTableNameOrAlias) > 0 Then Set tmpDeta.Pieza = DAOPieza.Map(rs, fieldsIndex, piezaTableNameOrAlias)
        If LenB(detallePedidoTableNameOrAlias) > 0 Then Set tmpDeta.DetalleRaiz = DAODetalleOrdenTrabajo.Map(rs, fieldsIndex, detallePedidoTableNameOrAlias)

    End If

    Set MapConjuntoProduccion = tmpDeta

End Function


Public Function FindAvanceSimple(ByVal idPedido As Long, _
                                 ByVal idPieza As Long, _
                                 ByVal idSector As Long, _
                                 Optional ByVal FallbackCualquierSector As Boolean = False) As AvanceSimpleDTO
    Dim q As String, rs As ADODB.Recordset
    Dim a As AvanceSimpleDTO

    ' 1) intenta por sector
    q = "SELECT * FROM detalles_pedidos_conjuntos_avance dpca " & _
        "WHERE dpca.id_detalle_pedido=" & idPedido & _
        " AND id_sector=" & idSector & " ORDER BY id DESC LIMIT 1"
    Set rs = conectar.RSFactory(q)

    ' 2) si no hay y querés fallback, busca sin sector
    If (rs Is Nothing Or rs.EOF) And FallbackCualquierSector Then
        q = "SELECT * FROM detalles_pedidos_conjuntos_avance dpca " & _
            "WHERE dpca.id_detalle_pedido=" & idPedido & _
            " ORDER BY id DESC LIMIT 1"
        Set rs = conectar.RSFactory(q)
    End If

    If Not (rs Is Nothing) Then
        If Not rs.EOF Then
            a.CantRecibida = NzDbl(rs!cant_recibida)
            a.CantFabricada = NzDbl(rs!cant_fabricada)
            a.CantScrap = NzDbl(rs!cant_scrap)
            a.FechaInicio = IIf(IsNull(rs!fecha_inicio), Null, rs!fecha_inicio)
            a.FechaFin = IIf(IsNull(rs!fecha_fin), Null, rs!fecha_fin)
            a.Recibio = NzLng(rs!Recibio)
            a.SiguienteProceso = NzStr(rs!siguiente_proceso)
        End If
    End If

    FindAvanceSimple = a
End Function

' Helpers
Private Function NzLng(v As Variant) As Long
    If IsNull(v) Or v = "" Then NzLng = 0 Else NzLng = CLng(v)
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNull(v) Or v = "" Then NzDbl = 0 Else NzDbl = CDbl(v)
End Function

Private Function NzStr(v As Variant) As String
    If IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function NzDate(v As Variant) As Variant
    If IsDate(v) Then NzDate = CDate(v) Else NzDate = Null
End Function

