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

