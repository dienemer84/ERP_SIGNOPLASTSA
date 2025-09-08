Attribute VB_Name = "DAOProduccion"
Option Explicit
Public LastError As String

Public Function Save(r As clsFilaPlanoRow) As Boolean
    On Error GoTo err1
    Dim q As String

    If r.IdPiezaPedido = 0 Then
        q = "INSERT INTO pedidos_produccion_carga " & _
            "(id_pedido,id_pedido_pieza,cant_recibida,cant_fabricada,cant_scrap,fecha_inicio,fecha_fin,recibio,siguiente_proceso) VALUES (" & _
            Escape(r.IdPedido) & "," & _
            Escape(r.IdPiezaPedido) & "," & _
            Escape(r.CantRecibida) & "," & _
            Escape(r.CantFabricada) & "," & _
            Escape(r.CantScrap) & "," & _
            EscapeDate(r.FechaInicio) & "," & _
            EscapeDate(r.FechaFin) & "," & _
            Escape(r.UsuarioRecibio) & "," & _
            Escape(r.ProcesoSiguiente) & ")"
    Else
    q = "UPDATE pedidos_produccion_carga SET " & _
        "id_pedido = " & EscapeNum(r.IdPedido) & "," & _
        "id_pedido_pieza = " & EscapeNum(r.IdPiezaPedido) & "," & _
        "cant_recibida = " & EscapeNum(r.CantRecibida) & "," & _
        "cant_fabricada = " & EscapeNum(r.CantFabricada) & "," & _
        "cant_scrap = " & EscapeNum(r.CantScrap) & "," & _
        "fecha_inicio = " & EscapeDate(r.FechaInicio) & "," & _
        "fecha_fin = " & EscapeDate(r.FechaFin) & "," & _
        "recibio = " & EscapeNum(r.UsuarioRecibio) & "," & _
        "siguiente_proceso = " & EscapeStr(r.ProcesoSiguiente) & _
        " WHERE id_pedido = " & EscapeNum(r.Id) & " and id_pedido_pieza = " & EscapeNum(r.IdPiezaPedido)
        End If

    conectar.execute (q)
    Save = True
    Exit Function
err1:
    Save = False
End Function

