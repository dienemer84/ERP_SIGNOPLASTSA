Attribute VB_Name = "DAOProduccionHistorial"
Option Explicit

Dim rs As ADODB.Recordset


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
        h.IdSector = NzLngF(rs, "id_sector")
        h.FEcha = NzDateF(rs, "fecha")
        h.ObservacionOld = NzStrF(rs, "observacion_old")
        h.ObservacionNew = NzStrF(rs, "observacion_new")
        h.Accion = NzStrF(rs, "accion")

        ' Usuarios como objetos
        Set h.UsuarioOperacion = DAOUsuarios.GetById(NzLngF(rs, "usuario_operacion"))
        Set h.UsuarioRecibio = DAOUsuarios.GetById(NzLngF(rs, "usuario_recibio"))

        ' Viejos y nuevos
        h.CantRecibidaOld = NzDblF(rs, "cant_recibida_old")
        h.CantRecibidaNew = NzDblF(rs, "cant_recibida_new")
        h.CantFabricadaOld = NzDblF(rs, "cant_fabricada_old")
        h.CantFabricadaNew = NzDblF(rs, "cant_fabricada_new")
        h.CantScrapOld = NzDblF(rs, "cant_scrap_old")
        h.CantScrapNew = NzDblF(rs, "cant_scrap_new")

        h.FechaInicioOld = NzDateF(rs, "fecha_inicio_old")
        h.FechaInicioNew = NzDateF(rs, "fecha_inicio_new")
        h.FechaFinOld = NzDateF(rs, "fecha_fin_old")
        h.FechaFinNew = NzDateF(rs, "fecha_fin_new")
        
        h.HoraInicioOld = NzDateF(rs, "hora_inicio_old")
        h.HoraInicioNew = NzDateF(rs, "hora_inicio_new")
        h.HoraFinOld = NzDateF(rs, "hora_fin_old")
        h.HoraFinNew = NzDateF(rs, "hora_fin_new")
        
        Set h.ProcesoOld = DAOSectores.GetByIdModulo(NzLngF(rs, "proceso_old"))
        Set h.ProcesoNew = DAOSectores.GetByIdModulo(NzLngF(rs, "proceso_new"))
        
        col.Add h
        
        rs.MoveNext
        
    Loop

    Set GetAllByPieza = col
    Exit Function

err1:
    Set GetAllByPieza = Nothing
End Function


Public Function agregar(ByVal r As clsFilaPlanoRow, _
                        ByVal Accion As String, _
                        ByRef prev As AvanceSimpleDTO) As Boolean
    On Error GoTo err1

    Dim uid As Long
    Dim ra As Long
    Dim cols As String
    Dim vals As String
    Dim q As String

    uid = conectar.GetEntityId(funciones.GetUserObj)

    cols = "id_pedido,id_detalle,id_pieza,id_sector,usuario_operacion,usuario_recibio,"
    cols = cols & "cant_recibida_old,cant_recibida_new,cant_fabricada_old,cant_fabricada_new,"
    cols = cols & "cant_scrap_old,cant_scrap_new,fecha_inicio_old,fecha_inicio_new,fecha_fin_old,fecha_fin_new,"
    cols = cols & "hora_inicio_old,hora_inicio_new,hora_fin_old,hora_fin_new,"
    cols = cols & "proceso_old,proceso_new,almacen_old,almacen_new,"
    cols = cols & "observacion_old,observacion_new,accion,fecha"

    vals = EscapeNum(r.IdPedido) & "," & _
           EscapeNum(r.IdTabla) & "," & _
           EscapeNum(r.idPiezaPedido) & "," & _
           EscapeNum(r.IdSector) & "," & _
           EscapeNum(uid) & "," & _
           EscapeNum(NzLng(r.UsuarioRecibio)) & ","

    vals = vals & EscapeNum(NzDbl(prev.CantRecibida)) & "," & EscapeNum(NzDbl(r.CantRecibida)) & ","
    vals = vals & EscapeNum(NzDbl(prev.CantFabricada)) & "," & EscapeNum(NzDbl(r.CantFabricada)) & ","
    vals = vals & EscapeNum(NzDbl(prev.CantScrap)) & "," & EscapeNum(NzDbl(r.CantScrap)) & ","

    vals = vals & EscapeDate(NzDateVar(prev.FechaInicio)) & "," & EscapeDate(NzDateVar(r.FechaInicio)) & ","
    vals = vals & EscapeDate(NzDateVar(prev.FechaFin)) & "," & EscapeDate(NzDateVar(r.FechaFin)) & ","

    vals = vals & EscapeTime(NzDateVar(prev.HoraInicio)) & "," & EscapeTime(NzDateVar(r.HoraInicio)) & ","
    vals = vals & EscapeTime(NzDateVar(prev.HoraFin)) & "," & EscapeTime(NzDateVar(r.HoraFin)) & ","

    vals = vals & EscapeStr(NzStr(prev.SiguienteProceso)) & "," & EscapeStr(NzStr(r.ProcesoSiguiente)) & ","
    vals = vals & EscapeNum(NzLng(prev.Almacen)) & "," & EscapeNum(NzLng(r.Almacen)) & ","
    vals = vals & EscapeStr(NzStr(prev.Observaciones)) & "," & EscapeStr(NzStr(r.Observaciones)) & ","

    vals = vals & EscapeStr(UCase$(Accion)) & ",CURRENT_TIMESTAMP"

    q = "INSERT INTO detalles_pedidos_conjuntos_avance_historial (" & cols & ") VALUES (" & vals & ")"

    conectar.ExecuteRa q, ra
    agregar = (ra > 0)
    Exit Function

err1:
    agregar = False
    MsgBox "agregar(): " & Err.Number & " - " & Err.Description, vbExclamation, "Historial de Avance"
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
    On Error Resume Next
    If IsNull(rs.Fields(fld).value) Or Trim(rs.Fields(fld).value & "") = "" Then
        NzDateF = 0          ' equivale a 30/12/1899
    Else
        NzDateF = CDate(rs.Fields(fld).value)
    End If
End Function

Private Function NzLng(v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzLng = 0
    Else
        NzLng = CLng(v)
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNull(v) Or v = "" Then
        NzDbl = 0
    Else
        NzDbl = CDbl(v)
    End If
End Function

Private Function NzStr(v As Variant) As String
    If IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Private Function NzDateVar(v As Variant) As Variant
    If IsDate(v) Then
        NzDateVar = CDate(v)
    Else
        NzDateVar = Null
    End If
End Function

Private Function EscapeTime(ByVal d As Variant) As String
    If IsNull(d) Or d = 0 Or (VarType(d) = vbString And Trim$(d) = "") Then
        EscapeTime = "NULL"
    Else
        ' ? MySQL / SQL Server (TIME o VARCHAR):
        EscapeTime = "'" & Format$(CDate(d), "hh:nn:ss") & "'"
        ' ? Si fuera Access/Jet, cambiá por:
        ' EscapeTime = "#" & Format$(CDate(d), "hh:nn:ss") & "#"
    End If
End Function
