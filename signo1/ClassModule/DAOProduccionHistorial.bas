Attribute VB_Name = "DAOProduccionHistorial"
Option Explicit

Dim rs As ADODB.Recordset

'''Public Function getAllByIdPIeza(id_pieza As Long) As Collection
'''    Dim col As New Collection
'''    Dim A As clsHistorial
'''    Set rs = conectar.RSFactory("select * from detalles_pedidos_conjuntos_avance_historial where idPieza=" & id_pieza)
'''    While Not rs.EOF
'''        Set A = New clsHistorial
'''        A.FEcha = rs!FEcha
'''        A.mensaje = rs!nota
'''        A.usuario = DAOUsuarios.GetById(rs!idUsuario)
'''        col.Add A
'''
'''        rs.MoveNext
'''    Wend
'''    Set A = Nothing
'''    Set getAllByIdFactura = col
'''End Function



Public Function Agregar(ByVal r As clsFilaPlanoRow, _
                        ByVal Accion As String, _
                        ByVal Nota As String, _
                        ByRef prev As AvanceSimpleDTO) As Boolean
    On Error GoTo err1

    Dim q As String, ra As Long, uid As Long
    uid = conectar.GetEntityId(funciones.GetUserObj)

    q = "INSERT INTO detalles_pedidos_conjuntos_avance_historial (" & _
        "id_pedido,id_detalle,id_pieza,id_sector,usuario_operacion,usuario_recibio," & _
        "cant_recibida_old,cant_recibida_new,cant_fabricada_old,cant_fabricada_new," & _
        "cant_scrap_old,cant_scrap_new,fecha_inicio_old,fecha_inicio_new,fecha_fin_old,fecha_fin_new," & _
        "proceso_old,proceso_new,accion,nota,fecha) VALUES (" & _
        EscapeNum(r.IdPedido) & "," & _
        EscapeNum(r.IdTabla) & "," & _
        EscapeNum(r.IdPiezaPedido) & "," & _
        EscapeNum(r.idSector) & "," & _
        EscapeNum(uid) & "," & _
        EscapeNum(r.UsuarioRecibio) & "," & _
        EscapeNum(prev.CantRecibida) & "," & EscapeNum(r.CantRecibida) & "," & _
        EscapeNum(prev.CantFabricada) & "," & EscapeNum(r.CantFabricada) & "," & _
        EscapeNum(prev.CantScrap) & "," & EscapeNum(r.CantScrap) & "," & _
        EscapeDate(prev.FechaInicio) & "," & EscapeDate(r.FechaInicio) & "," & _
        EscapeDate(prev.FechaFin) & "," & EscapeDate(r.FechaFin) & "," & _
        EscapeStr(prev.SiguienteProceso) & "," & EscapeStr(r.ProcesoSiguiente) & "," & _
        EscapeStr(UCase$(Accion)) & "," & _
        EscapeStr(Nota) & "," & _
        "CURRENT_TIMESTAMP" & _
        ")"

    conectar.execute q
    Agregar = (ra > 0)
    Exit Function
err1:
    Agregar = False
    MsgBox (Err.Number & " - " & Err.Description)
End Function


Public Function GetAllByPieza(ByVal id_pieza As Long, _
                              Optional ByVal id_pedido As Long = 0, _
                              Optional ByVal id_sector As Long = 0, _
                              Optional ByVal topN As Long = 0) As Collection
    On Error GoTo err1

    Dim col As New Collection
    Dim rs As Object
    Dim sql As String

    sql = "SELECT * FROM detalles_pedidos_conjuntos_avance_historial " & _
          "WHERE id_detalle = " & EscapeNum(id_pieza)

    If id_pedido > 0 Then sql = sql & " AND id_pedido = " & EscapeNum(id_pedido)
    If id_sector > 0 Then sql = sql & " AND id_sector = " & EscapeNum(id_sector)

    sql = sql & " ORDER BY fecha DESC"
    If topN > 0 Then sql = sql & " LIMIT " & CLng(topN)

    Set rs = conectar.RSFactory(sql)

    Do While Not rs.EOF
        Dim h As clsHistorialProduccion
        Set h = New clsHistorialProduccion

        ' Campos clave
        h.IdPedido = NzLngF(rs, "id_pedido")
        h.IdDetalle = NzLngF(rs, "id_detalle")
        h.idPieza = NzLngF(rs, "id_pieza")
        h.Sector = NzLngF(rs, "id_sector")
        h.FEcha = NzDateF(rs, "fecha")
        h.Nota = NzStrF(rs, "nota")
        h.Accion = NzStrF(rs, "accion")

        ' Usuarios (objetos o IDs, según tu clase)
'''        h.Usuario = DAOUsuarios.GetById(NzLngF(rs, "usuario_operacion"))
'''        h.UsuarioRecibio = DAOUsuarios.GetById(NzLngF(rs, "usuario_recibio"))

'''        ' Viejos y nuevos
'''        h.CantRecibidaOld = NzDblF(rs, "cant_recibida_old")
'''        h.CantRecibidaNew = NzDblF(rs, "cant_recibida_new")
'''        h.CantFabricadaOld = NzDblF(rs, "cant_fabricada_old")
'''        h.CantFabricadaNew = NzDblF(rs, "cant_fabricada_new")
'''        h.CantScrapOld = NzDblF(rs, "cant_scrap_old")
'''        h.CantScrapNew = NzDblF(rs, "cant_scrap_new")
'''
'''        h.FechaInicioOld = NzDateF(rs, "fecha_inicio_old")
'''        h.FechaInicioNew = NzDateF(rs, "fecha_inicio_new")
'''        h.FechaFinOld = NzDateF(rs, "fecha_fin_old")
'''        h.FechaFinNew = NzDateF(rs, "fecha_fin_new")
'''
'''        h.ProcesoOld = NzStrF(rs, "proceso_old")
'''        h.ProcesoNew = NzStrF(rs, "proceso_new")

        col.Add h
        rs.MoveNext
    Loop

    Set GetAllByPieza = col
    Exit Function
err1:
    Set GetAllByPieza = Nothing
End Function

' Helpers de lectura segura desde Recordset
Private Function NzLngF(ByVal rs As Object, ByVal fld As String) As Long
    If Not IsNull(rs.Fields(fld).value) Then NzLngF = CLng(rs.Fields(fld).value)
End Function

Private Function NzDblF(ByVal rs As Object, ByVal fld As String) As Double
    If Not IsNull(rs.Fields(fld).value) Then NzDblF = CDbl(rs.Fields(fld).value)
End Function

Private Function NzStrF(ByVal rs As Object, ByVal fld As String) As String
    If Not IsNull(rs.Fields(fld).value) Then NzStrF = CStr(rs.Fields(fld).value) Else NzStrF = ""
End Function

Private Function NzDateF(ByVal rs As Object, ByVal fld As String) As Variant
    If Not IsNull(rs.Fields(fld).value) Then NzDateF = CDate(rs.Fields(fld).value) Else NzDateF = Null
End Function
