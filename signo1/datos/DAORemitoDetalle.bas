Attribute VB_Name = "DAORemitoSDetalle"
Option Explicit
Public Const TABLA_ENTREGA As String = "e"
Public Const TABLA_DETALLE_PEDIDO As String = "dp"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_OTDETALLE_ID As String = "idDetallePedido"
Public Const CAMPO_OT_ID As String = "idPedido"
Public Const CAMPO_IDREMITO As String = "Remito"
Public Const CAMPO_CANTIDAD = "cantidad"
Public Const CAMPO_FECHA As String = "fecha"
Public Const CAMPO_ORIGEN As String = "origen"
Public Const CAMPO_FACTURADO As String = "facturado"
Public Const CAMPO_CONCEPTO As String = "concepto"
Public Const CAMPO_FACTURABLE As String = "facturable"
Public Const CAMPO_VALOR As String = "valor"
Public Const CAMPO_VALOR_MODIFICADO As String = "ModifValor"
Public Const CAMPO_OBSERVACIONES = "observaciones"

Public Function Guardar(deta As remitoDetalle) As Boolean
    On Error GoTo err1
    Guardar = True
    Dim q As String
    Dim n As Boolean
    If deta.Id = 0 Then
        n = True
        q = " INSERT INTO entregas  (" _
          & "idPedido, " _
          & "idDetallePedido," _
          & "cantidad, " _
          & "Remito, " _
          & "fecha, " _
          & "origen, " _
          & "facturado, " _
          & "valor, " _
          & "concepto, " _
          & "ModifValor, " _
          & "facturable, observaciones ) " _
          & "Values  ( " _
          & conectar.Escape(deta.idpedido) & "," _
          & conectar.Escape(deta.idDetallePedido) & "," _
          & conectar.Escape(deta.Cantidad) & "," _
          & conectar.Escape(deta.Remito) & "," _
          & conectar.Escape(deta.FEcha) & "," _
          & conectar.Escape(deta.Origen) & "," _
          & conectar.Escape(deta.Facturado) & "," _
          & conectar.Escape(deta.Valor) & "," _
          & conectar.Escape(deta.Concepto) & "," _
          & conectar.Escape(deta.ValorModificado) & "," _
          & conectar.Escape(deta.facturable) & "," & conectar.Escape(deta.observaciones) & ")"

    Else
        n = False
        q = "update sp.entregas   SET " _
          & "idPedido = 'idPedido' ," _
          & "idDetallePedido = 'idDetallePedido' , " _
          & "cantidad = 'cantidad' , " _
          & " Remito = 'Remito' , " _
          & " fecha = 'fecha' , " _
          & " origen = 'origen' ," _
          & " facturado = 'facturado' ," _
          & "valor = 'valor' , " _
          & "concepto = 'concepto' ," _
          & "ModifValor = 'ModifValor' ," _
          & "observaciones = 'observaciones' ," _
          & "facturable = 'facturable'  Where  id = 'id' "

        q = Replace$(q, "'id'", conectar.Escape(deta.Id))
        q = Replace$(q, "'idPedido'", conectar.Escape(deta.idpedido))
        q = Replace$(q, "'idDetallePedido'", conectar.Escape(deta.idDetallePedido))
        q = Replace$(q, "'cantidad'", conectar.Escape(deta.Cantidad))
        q = Replace$(q, "'Remito'", conectar.Escape(deta.Remito))
        q = Replace$(q, "'fecha'", conectar.Escape(deta.FEcha))
        q = Replace$(q, "'origen'", conectar.Escape(deta.Origen))
        q = Replace$(q, "'facturado'", conectar.Escape(deta.Facturado))
        q = Replace$(q, "'valor'", conectar.Escape(deta.Valor))
        q = Replace$(q, "'concepto'", conectar.Escape(deta.Concepto))
        q = Replace$(q, "'ModifValor'", conectar.Escape(deta.ValorModificado))
        q = Replace$(q, "'facturable'", conectar.Escape(deta.facturable))
        q = Replace$(q, "'observaciones'", conectar.Escape(deta.observaciones))

    End If
    If Not conectar.execute(q) Then GoTo err1
    Dim EVENTO As New clsEventoObserver


    Exit Function
err1:
    Guardar = False
End Function

Public Function Delete(detalle As remitoDetalle) As Boolean

    Delete = conectar.execute("delete from entregas where id=" & detalle.Id)  ' And DAODetalleOrdenTrabajo.SaveCantidad(detalle.DetallePedido.Id, detalle.Cantidad * -1, CantidadEntregada_, 0, detalle.Remito)


' Stop


End Function

Public Function FindAllByRemito(IdRemito As Long, Optional withCantidadEntregada As Boolean = False, Optional WithDetallePedido As Boolean = False) As Collection
    Set FindAllByRemito = FindAll("AND " & TABLA_ENTREGA & "." & CAMPO_IDREMITO & "=" & IdRemito, withCantidadEntregada, WithDetallePedido)
End Function

Public Function FindAllByDetallePedido(idDP As Long) As Collection
    Dim q As String
    q = "AND " & TABLA_ENTREGA & "." & CAMPO_OTDETALLE_ID & "=" & idDP



    Set FindAllByDetallePedido = FindAll(q, True, False)
End Function

Public Function FindById(Id As Long) As remitoDetalle
    If Id = -1 Then
        Set FindById = Nothing
    Else
        Dim col As New Collection
        Set col = FindAll("AND " & TABLA_ENTREGA & "." & CAMPO_ID & "=" & Id, True, True)
        If col.count > 0 Then
            Set FindById = col.item(1)
        Else
            Set FindById = Nothing
        End If

    End If
End Function

Public Function FindAll(Optional filtro As String = vbNullString, Optional WithCantidadEntregadas As Boolean = False, Optional WithDetallePedido As Boolean = False) As Collection
    Dim indice As Dictionary
    Dim rs As Recordset
    Dim col As New Collection
    Dim strsql As String
    strsql = "SELECT * FROM entregas e LEFT JOIN detalles_pedidos dp ON e.idDetallePedido=dp.id LEFT JOIN remitos r ON r.id = e.Remito WHERE 1=1 "
    If LenB(filtro) > 0 Then strsql = strsql & filtro
    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, indice

    Dim detalle As remitoDetalle


    While Not rs.EOF
        Set detalle = New remitoDetalle

        Set detalle = Map(rs, indice, TABLA_ENTREGA, "r")
        If detalle.idpedido > 0 And WithDetallePedido Then Set detalle.DetallePedido = DAODetalleOrdenTrabajo.FindById(detalle.idDetallePedido, True, False, False)

        col.Add detalle
        rs.MoveNext
    Wend


    Set FindAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tablaEntrega As String, Optional tablaRemito As String = vbNullString) As remitoDetalle
    Dim Id As Variant
    Id = GetValue(rs, indice, tablaEntrega, CAMPO_ID)
    If Id > 0 Then
        Dim dr As remitoDetalle
        Set dr = New remitoDetalle
        dr.Origen = GetValue(rs, indice, tablaEntrega, CAMPO_ORIGEN)

        dr.Concepto = GetValue(rs, indice, tablaEntrega, CAMPO_CONCEPTO)
        dr.Cantidad = GetValue(rs, indice, tablaEntrega, CAMPO_CANTIDAD)
        dr.idDetallePedido = GetValue(rs, indice, tablaEntrega, CAMPO_OTDETALLE_ID)
        dr.facturable = GetValue(rs, indice, tablaEntrega, CAMPO_FACTURABLE)
        dr.FEcha = GetValue(rs, indice, tablaEntrega, CAMPO_FECHA)
        dr.Facturado = GetValue(rs, indice, tablaEntrega, CAMPO_FACTURADO)
        dr.idpedido = GetValue(rs, indice, tablaEntrega, CAMPO_OT_ID)
        dr.Remito = GetValue(rs, indice, tablaEntrega, CAMPO_IDREMITO)
        dr.Valor = GetValue(rs, indice, tablaEntrega, CAMPO_VALOR)
        dr.ValorModificado = GetValue(rs, indice, tablaEntrega, CAMPO_VALOR_MODIFICADO)
        dr.observaciones = GetValue(rs, indice, tablaEntrega, CAMPO_OBSERVACIONES)
        dr.Id = Id

        If LenB(tablaRemito) > 0 Then
            Set dr.RemitoAlQuePertenece = DAORemitoS.Map(rs, indice, tablaRemito)
        End If

        Set Map = dr
    End If

End Function


Public Function CambiarEstadoFacturable(nuevoEstado As Boolean, deta As remitoDetalle) As Boolean
    On Error GoTo err1
    CambiarEstadoFacturable = conectar.execute("update entregas set facturable=" & Abs(nuevoEstado) & " where id=" & deta.Id)



    deta.facturable = Not deta.facturable
    Dim est As EstadoRemitoFacturado
    est = DAORemitoS.AnalizarEstadoFacturado(deta.Remito)
    If Not DAORemitoS.CambiarEstadoFacturado(deta.Remito, est) Then CambiarEstadoFacturable = False
    Exit Function
err1:
End Function






