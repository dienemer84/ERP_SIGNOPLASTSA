Attribute VB_Name = "DAOOrdenTrabajo"
Option Explicit
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_DESCRIPCION As String = "descripcion"
Public Const CAMPO_FECHA_ENTREGA As String = "fechaEntrega"
Public Const CAMPO_FECHA_CREADO As String = "fechaCreado"
Public Const CAMPO_ACTIVO As String = "activo"
Public Const CAMPO_ENTREGADO As String = "entregado"
Public Const CAMPO_FECHA_CERRADO As String = "fechaCerrado"
Public Const CAMPO_DESCUENTO As String = "dto"
Public Const CAMPO_FECHA_APROBADO As String = "fechaAprobado"
Public Const CAMPO_FECHA_MODIFICADO As String = "fechaModificado"
Public Const CAMPO_ESTADO As String = "estado"
Public Const CAMPO_NUMERO_PRESUPUESTO As String = "nroPresupuesto"
Public Const CAMPO_ANTICIPO As String = "anticipo"
Public Const CAMPO_ANTICIPO_FACTURADO As String = "anticipo_facturado"
Public Const CAMPO_FORMA_PAGO_ANTICIPO As String = "forma_pago_anticipo"
Public Const CAMPO_MISMA_FECHA_DETALLES As String = "misma_fecha_entrega_detalles"
Public Const CAMPO_ANTICIPO_DIAS As String = "anticipo_dias"
Public Const CAMPO_SALDO_DIAS As String = "saldo_dias"
Public Const CAMPO_FORMA_PAGO_SALDO As String = "forma_pago_saldo"
Public Const CAMPO_CLIENTE_ID As String = "idCliente"
Public Const TABLA_PEDIDO As String = "p"
Public Const CAMPO_ANTICIPO_FACTURA_ID As String = "id_anticipo_factura"
Public Const TABLA_CANTIDADES As String = "canti"
Public Const CAMPO_CANTIDADES_CANTIDAD As String = "cantidad"
Public Const CAMPO_CANTIDADES_FECHA As String = "fecha"
Public Const CAMPO_CANTIDADES_TIPO As String = "tipo_cantidad"


Public Enum TipoCantidadOT
    CantidadEntregada_ = 1
    CantidadFacturada_ = 2
    CantidadFabricada_ = 3
End Enum

Public Sub ExportarExcelResumenGeneral(idOt As Long)
  On Error GoTo E

Dim Ot As OrdenTrabajo
Dim deta As DetalleOrdenTrabajo
Set Ot = DAOOrdenTrabajo.FindById(idOt)
Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True, True, True)

    Dim xlWorkbook As New Excel.Workbook
    Dim xlWorksheet As New Excel.Worksheet
    Dim xlApplication As New Excel.Application


    Dim tareas As New Dictionary


    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 2).value = "Orden de Trabajo Nº " & Ot.IdFormateado
    xlWorksheet.Cells(2, 2).value = "Cliente " & Ot.cliente.razon
    
        xlWorksheet.Range(xlWorksheet.Cells(1, 2), xlWorksheet.Cells(1, 9)).Merge
        xlWorksheet.Range(xlWorksheet.Cells(2, 2), xlWorksheet.Cells(2, 9)).Merge
    xlWorksheet.Range(xlWorksheet.Cells(1, 2), xlWorksheet.Cells(1, 9)).HorizontalAlignment = xlLeft
    xlWorksheet.Range(xlWorksheet.Cells(2, 2), xlWorksheet.Cells(2, 9)).HorizontalAlignment = xlLeft
    xlWorksheet.Cells(3, 2).value = "Item"
    xlWorksheet.Cells(3, 3).value = "Pieza"
    xlWorksheet.Cells(3, 4).value = "Cantidad"
    xlWorksheet.Cells(3, 5).value = "Entregadas"
          
    Dim Entregas As New Collection
    Dim initialPos As Long: initialPos = 3
    Dim yPos As Long: yPos = initialPos
    Dim entrega As remitoDetalle
    For Each deta In Ot.Detalles
    yPos = yPos + 1
    
    Set deta.OrdenTrabajo = Ot
    exportaDetalle yPos, deta, xlWorksheet
    
     
    Next deta
    
    
    
    
    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True
    ''

    Dim ruta As String
    ruta = Environ$("TEMP")
    If LenB(ruta) = 0 Then ruta = Environ$("TMP")
    If LenB(ruta) = 0 Then ruta = App.path
    ruta = ruta & "\" & funciones.CreateGUID() & ".xls"

    xlWorkbook.SaveAs ruta

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    ShellExecute -1, "open", ruta, "", "", 4

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing
    
Exit Sub
E:

End Sub

Private Sub exportaDetalle(yPos, deta As DetalleOrdenTrabajo, xlWorksheet)
Dim deta2 As DetalleOrdenTrabajo
Dim origenypos As Long

origenypos = yPos

    xlWorksheet.Cells(yPos, 2).value = deta.item
     xlWorksheet.Cells(yPos, 3).value = deta.Pieza.nombre
     xlWorksheet.Cells(yPos, 4).value = deta.CantidadPedida
     xlWorksheet.Cells(yPos, 5).value = deta.Cantidad_Entregada
  xlWorksheet.Cells(yPos, 6).value = "Remito"
        xlWorksheet.Cells(yPos, 7).value = "Fecha"
          xlWorksheet.Cells(yPos, 8).value = "Cantidad"
    
If deta.OrdenTrabajo.EsMarco Then

For Each deta2 In deta.DetallesHijasMarcoPadre
   yPos = yPos + 1
   exportaDetalleHijo yPos, deta2, xlWorksheet
    
Next
Else
yPos = yPos + 1
 exportarEntregas yPos, deta.id, xlWorksheet
 End If
 
yPos = yPos + 1
  xlWorksheet.Range(xlWorksheet.Cells(origenypos, 2), xlWorksheet.Cells(yPos, 9)).BorderAround xlContinuous, xlMedium
End Sub

Private Sub exportaDetalleHijo(yPos, deta As DetalleOrdenTrabajo, xlWorksheet)
Dim deta2 As DetalleOrdenTrabajo
    Dim rs As Recordset
    Dim strsql As String
    Dim origenypos As Long
    origenypos = yPos
      strsql = "Select descripcion from pedidos where id=" & deta.OrdenTrabajo.id
    Set rs = conectar.RSFactory(strsql)

                deta.OrdenTrabajo.descripcion = rs!descripcion

    xlWorksheet.Cells(yPos, 2).value = ""
     xlWorksheet.Cells(yPos, 3).value = "OT " & deta.OrdenTrabajo.id & " - " & deta.OrdenTrabajo.descripcion
     xlWorksheet.Cells(yPos, 4).value = deta.CantidadPedida
     xlWorksheet.Cells(yPos, 5).value = deta.Cantidad_Entregada

If deta.OrdenTrabajo.EsMarco Then

For Each deta2 In deta.DetallesHijasMarcoPadre
    
   exportaDetalleHijo yPos, deta2, xlWorksheet
    
Next
Else
 exportarEntregas yPos, deta.id, xlWorksheet
 End If
 xlWorksheet.Range(xlWorksheet.Cells(origenypos, 3), xlWorksheet.Cells(yPos, 8)).BorderAround xlContinuous
End Sub


Private Sub exportarEntregas(yPos, detalleid As Long, xlWorksheet)
 Dim Entregas As New Collection
 Dim facturas As Collection
 Dim origenypos As Long
 origenypos = yPos
 Dim entrega As remitoDetalle
 Set Entregas = DAORemitoSDetalle.FindAllByDetallePedido(detalleid)
    
     If Entregas.count > 0 Then
       For Each entrega In Entregas
       If (entrega.RemitoAlQuePertenece.estado = RemitoAprobado) Then
       
       
       xlWorksheet.Cells(yPos, 6).value = entrega.RemitoAlQuePertenece.numero
       xlWorksheet.Cells(yPos, 7).value = entrega.FEcha
       xlWorksheet.Cells(yPos, 8).value = entrega.Cantidad
       yPos = yPos + 1
      End If
      Next entrega
     
     End If
     yPos = yPos - 1
       xlWorksheet.Range(xlWorksheet.Cells(origenypos, 6), xlWorksheet.Cells(yPos, 8)).BorderAround xlContinuous, xlThin
End Sub
Public Sub ImprimirFaltantesFacturacion(Ot As OrdenTrabajo)
    On Error GoTo err1
    dsrFaltantesFacturar.Sections("Sec4").Controls.item("lblOT").caption = "Orden de Trabajo Nº: " & Ot.id & " | " & Ot.descripcion
    dsrFaltantesFacturar.Sections("Sec4").Controls.item("lblCliente").caption = "Cliente: " & Ot.cliente.razon
    dsrFaltantesFacturar.Sections("Sec4").Controls.item("lblTitulo").caption = "FALTANTES DE FACTURACION OT Nº " & Ot.id


    Dim r As New Recordset
    With r
        .Fields.Append "item", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "nota", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "entregados", adDouble, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "pedidos", adDouble, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "faltantes", adDouble, 255, adFldUpdatable            ' And adFldIsNullable
        .Fields.Append "codigo", adVarChar, 255, adFldUpdatable            ' And adFldIsNullable
        .Fields.Append "PU", adVarChar, 255, adFldUpdatable
        .Fields.Append "PT", adVarChar, 255, adFldUpdatable
        .Fields.Append "importe_faltante", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
    End With
    r.Open
    Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True)
    Dim deta As DetalleOrdenTrabajo

    Dim cod As String
    Dim Total As Double
    Total = 0
    For Each deta In Ot.Detalles
        r.AddNew
        r!item = deta.item
        r!Nota = deta.Nota
        r!detalle = " [ " & deta.Pieza.UnidadMedida & " ] " & deta.Pieza.nombre
        r!Entregados = deta.Cantidad_Entregada
        r!pedidos = deta.CantidadPedida
        r!faltantes = r!pedidos - r!Entregados
        r!pu = funciones.FormatearDecimales(deta.Precio)
        r!Pt = funciones.FormatearDecimales(r!pu * r!faltantes)

        Total = Total + r!Pt

        cod = "*R" & Format(deta.id, "00000000") & "*"
        r!codigo = cod
        r.Update
    Next
    dsrFaltantesFacturar.Sections("SEC5").Controls.item("lblTotalFaltante").caption = Ot.moneda.NombreCorto & " " & funciones.FormatearDecimales(Total)

    Set dsrFaltantesFacturar.DataSource = r
    dsrFaltantesFacturar.PrintReport True
    Exit Sub
err1:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
Public Sub ImprimirPreconteo(Ot As OrdenTrabajo)
    On Error GoTo err1
    dsrPreconteo.Sections("section4").Controls.item("lblOT").caption = "Orden de Trabajo Nº: " & Ot.id & " | " & Ot.descripcion
    dsrPreconteo.Sections("section4").Controls.item("lblCliente").caption = "Cliente: " & Ot.cliente.razon
    dsrPreconteo.Sections("section4").Controls.item("lblTitulo").caption = "PRECONTEO DE ENTREGA OT Nº " & Ot.id

    Dim r As New Recordset
    With r
        .Fields.Append "item", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "nota", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "entregados", adDouble, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "pedidos", adDouble, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "faltantes", adDouble, 255, adFldUpdatable            ' And adFldIsNullable
        .Fields.Append "codigo", adVarChar, 255, adFldUpdatable            ' And adFldIsNullable
        .Fields.Append "X", adVarChar, 255, adFldUpdatable            ' And adFldIsNullable
    End With
    r.Open
    Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True)
    Dim deta As DetalleOrdenTrabajo

    Dim cod As String

    For Each deta In Ot.Detalles
        r.AddNew
        r!item = deta.item
        r!Nota = deta.Nota
        r!detalle = " [ " & deta.Pieza.UnidadMedida & " ] " & deta.Pieza.nombre
        r!Entregados = deta.Cantidad_Entregada
        r!pedidos = deta.CantidadPedida
        r!faltantes = r!pedidos - r!Entregados
        If r!faltantes = 0 Then
            r!x = "XXXX"
        Else
            r!x = ""
        End If
        cod = "*R" & Format(deta.id, "00000000") & "*"
        r!codigo = cod
        r.Update
    Next
    Set dsrPreconteo.DataSource = r
    dsrPreconteo.PrintReport True
    Exit Sub
err1:
    MsgBox Err.Description, vbCritical, "Error"
End Sub



Public Sub ImprimirRuta(Ot As OrdenTrabajo, detOT As DetalleOrdenTrabajo, Optional detOTConj As DetalleOTConjuntoDTO = Nothing, Optional posOnConj As String = vbNullString)
    pedido_pieza2.Sections("cabeza").Controls("lblOT").caption = Format("0000", Ot.id) & " | " & detOT.item  ' "." & 1 & "/" & Ot.detalles.count     '1 = pos
    If IsSomething(detOTConj) Then
        pedido_pieza2.Sections("cabeza").Controls("barcode").caption = "*R" & Format(detOTConj.id, "00000000") & "*"
    Else
        pedido_pieza2.Sections("cabeza").Controls("barcode").caption = "*R" & Format(detOT.id, "00000000") & "*"
    End If
    pedido_pieza2.Sections("cabeza").Controls("lblCliente").caption = Ot.cliente.razon
    pedido_pieza2.Sections("cabeza").Controls("lblReferencia").caption = Ot.descripcion
    pedido_pieza2.Sections("observar").Controls("lblNota").caption = detOT.Nota
    pedido_pieza2.Sections("cabeza").Controls("lblfechaentrega").caption = Ot.FechaEntrega
    pedido_pieza2.Sections("cabeza").Controls("entregait").caption = detOT.FechaEntrega
    pedido_pieza2.Sections("cabeza").Controls("lblDetalle").caption = detOT.Pieza.nombre
    pedido_pieza2.Sections("observar").Controls("copia").caption = detOT.CantidadImpresionesDeRuta
    pedido_pieza2.Sections("cabeza").Controls("lbldeStock").caption = IIf((detOT.CantidadPedida - detOT.ReservaStock) = 0, "De Stock", vbNullString)
    pedido_pieza2.Sections("cabeza").Controls("lblItem").caption = "Item " & detOT.item

    If Not detOTConj Is Nothing Then
        pedido_pieza2.Sections("cabeza").Controls("lblMasDetalle").caption = detOTConj.IdentificadorPosicion & " - " & detOTConj.Pieza.nombre     '"Elemento: " & PosConj & "/" & CantElemConj & " " & detOTConj.Pieza.nombre
    Else
        pedido_pieza2.Sections("cabeza").Controls("lblMasDetalle").caption = vbNullString
    End If

    If detOTConj Is Nothing Then
        pedido_pieza2.Sections("cabeza").Controls("lblCantidad").caption = detOT.CantidadPedida
    Else
        pedido_pieza2.Sections("cabeza").Controls("lblCantidad").caption = (detOTConj.CantidadTotalStatic)
    End If
    pedido_pieza2.Sections("observar").Controls("lblnota").caption = detOT.Nota
    Dim idPieza As Long
    If detOTConj Is Nothing Then
        idPieza = detOT.Pieza.id
    Else
        idPieza = detOTConj.Pieza.id
    End If

    Dim Cant As Long
    Dim texto As String

    '''''''''archivos
    Cant = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas, Array(idPieza)).item(idPieza)
    If Cant = 0 Then
        texto = "No existen archivos asociados. Controlar periodicamente"
    Else
        texto = "Existe/n " & Cant & " archivo/s asociado/s. Por favor Verificar"
    End If
    pedido_pieza2.Sections("pie").Controls("archivos_asociados").caption = texto

    '''''''''incidencias
    Cant = DAOIncidencias.GetCantidadIncidenciasPorReferencia(OA_Piezas, Array(idPieza)).item(idPieza)
    If Cant = 0 Then
        texto = "No existen incidencias asociadas. Controlar periodicamente"
    Else
        texto = "Existe/n " & Cant & " incidencia/s asociada/s. Por favor Verificar"
    End If
    pedido_pieza2.Sections("pie").Controls("incidencias").caption = texto

    If IsSomething(detOT) And Not IsSomething(detOTConj) Then
        conectar.execute "UPDATE detalles_pedidos SET impresiones_ruta = " & (detOT.CantidadImpresionesDeRuta + 1) & " WHERE id = " & detOT.id
    End If

    Dim tmpId As Long
    If detOTConj Is Nothing Then
        tmpId = detOT.id
    Else
        tmpId = detOTConj.id
    End If

    If IsSomething(detOTConj) Then
        Set pedido_pieza2.DataSource = conectar.RSFactory("SELECT concat('*', dm.id,'*') as codigo, t.cantxproc,t.tarea,s.sector from tareas t,sectores s,PlaneamientoTiemposProcesos dm where dm.codigoTarea=t.id and s.id=t.id_sector and dm.idDetallePedidoConj=" & tmpId)
    Else
        Set pedido_pieza2.DataSource = conectar.RSFactory("SELECT concat('*', dm.id,'*') as codigo, t.cantxproc,t.tarea,s.sector from tareas t,sectores s,PlaneamientoTiemposProcesos dm where dm.codigoTarea=t.id and s.id=t.id_sector and dm.idDetallePedidoConj = 0 and dm.idDetallePedido=" & tmpId)
    End If

    DAOOrdenTrabajoHistorial.agregar detOT.OrdenTrabajo, "RUTA " & detOT.id & " IMPRESA"

    pedido_pieza2.PrintReport False
End Sub

Public Function PonerEnProduccion(T As OrdenTrabajo) As Boolean    'activar viejo
    Dim estado_ant As EstadoOrdenTrabajo
    estado_ant = T.estado
    conectar.BeginTransaction
    If T.estado = EstadoOT_EnProceso Then
        GoTo err1
    Else
        T.estado = EstadoOT_EnProceso
        If Not Guardar(T, False) Then GoTo err1
    End If

    If Not DAOOrdenTrabajoHistorial.agregar(T, "OT En producción") Then GoTo err1


    PonerEnProduccion = True
    conectar.CommitTransaction

    DAOEvento.Publish T.id, TipoEventoBroadcast.TEB_OrdenTrabajoActivada
    Exit Function
err1:
    T.estado = estado_ant
    conectar.RollBackTransaction
End Function




Private Function CambiarEstado(T As OrdenTrabajo, estadoNuevo As EstadoOrdenTrabajo) As Boolean
    On Error GoTo e4
    Dim est_anterior As EstadoOrdenTrabajo
    Dim fec_anterior As Date
    Dim usu_anterior As clsUsuario
    est_anterior = T.estado
    fec_anterior = T.FechaModificado
    Set usu_anterior = T.UsuarioModificado
    T.estado = estadoNuevo
    Set T.UsuarioModificado = funciones.GetUserObj
    T.FechaModificado = Now

    CambiarEstado = Guardar(T)
    If CambiarEstado = False Then GoTo e4
    CambiarEstado = True
    Exit Function
e4:

    T.estado = est_anterior
    T.FechaModificado = fec_anterior
    Set T.UsuarioModificado = usu_anterior
    CambiarEstado = False
End Function

Public Function DescontarReservaDetalle(detalle As DetalleOrdenTrabajo, Cant As Double) As Boolean
On Error GoTo err1
conectar.BeginTransaction
Dim P As Pieza
Set P = DAOPieza.FindById(detalle.Pieza.id, FL_0)
If P.CantidadStock >= Cant Then

        detalle.ReservaStock = detalle.ReservaStock + Cant
        If Not DAOPieza.ModificarStock(P, ModificarStock_BajaOT, Cant, , detalle.OrdenTrabajo.id) Then
                Err.Raise 8002, "Reserva Stock", "Imposible Descontar la Reserva"
        End If
        
        If Not DAODetalleOrdenTrabajo.Save(detalle) Then
                Err.Raise 8002, "Reserva Stock", "Imposible Descontar la Reserva"
        End If
            
            
Else
    Err.Raise 8001, "Reserva Stock", "Stock Insuficiente"
End If

conectar.CommitTransaction
Exit Function

err1:
conectar.RollBackTransaction
Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function DescontarReserva(Ot As OrdenTrabajo) As Boolean
    On Error GoTo err441
    Dim det As DetalleOrdenTrabajo
    Dim Pieza As Pieza
    Dim ok As Boolean
    'valido x ultima vez

    ok = True
    For Each det In Ot.Detalles
        Set Pieza = DAOPieza.FindById(det.Pieza.id, FL_0)
        If Pieza.CantidadStock < det.ReservaStock Then
            ok = False
            Exit For
        End If
    Next

    If Not ok Then GoTo err441

    For Each det In Ot.Detalles
        If det.ReservaStock > 0 Then
            If Not DAOPieza.ModificarStock(Pieza, ModificarStock_BajaOT, det.ReservaStock, , Ot.id) Then
                GoTo err441
            End If
        End If

    Next

    DescontarReserva = True
    Exit Function
err441:

    DescontarReserva = False

End Function

Public Function HacerEditable(T As OrdenTrabajo) As Boolean
    Dim estado_ant As EstadoOrdenTrabajo
    Dim fecha_aprob_ant As Date
    Dim usu_aprob_ant As clsUsuario
    Dim deta As DetalleOrdenTrabajo
    estado_ant = T.estado
    fecha_aprob_ant = T.fechaAprobado
    Set usu_aprob_ant = T.UsuarioAprobado

    conectar.BeginTransaction
    If T.estado = EstadoOT_Pendiente Then
        GoTo err1
    Else
        T.estado = EstadoOT_Pendiente
        T.fechaAprobado = 0
        Set T.UsuarioAprobado = Nothing
        T.StockDescontado = False

        If estado_ant = EstadoOT_EnEspera Then
            For Each deta In T.Detalles

                If Not DAOPieza.ModificarStock(deta.Pieza, ModificarStock_AltaOT, deta.ReservaStock, , T.id) Then GoTo err1
                deta.ReservaStock = 0
            Next
        End If

        If Not Guardar(T, False) Then
            GoTo err1
        Else
            If Not DAOOrdenTrabajoHistorial.agregar(T, "OT lista para editar") Then GoTo err1

        End If
    End If
    conectar.CommitTransaction
    HacerEditable = True
    Exit Function
err1:
    conectar.RollBackTransaction
    T.estado = estado_ant
    Set T.UsuarioAprobado = usu_aprob_ant
    T.fechaAprobado = fecha_aprob_ant

End Function

Public Function FindById(ByRef id As Long) As OrdenTrabajo
    Dim col As Collection
    Set col = FindAll("p.id = " & id)
    If col.count > 0 Then
        Set FindById = col(1)
    Else
        Set FindById = Nothing
    End If
End Function
Public Function FindAll(Optional ByRef filter As String = vbNullString, _
                        Optional withColEntregados As Boolean = False, _
                        Optional withColFacturados As Boolean = False, _
                        Optional withColFabricados As Boolean = False, _
                        Optional withDetalles As Boolean = False, _
                        Optional withEntregados As Boolean = False, _
                        Optional withFacturados As Boolean = False, _
                        Optional withFabricados As Boolean = False, _
                        Optional withPorcentajeTareasFinalizadas As Boolean = False _
                        ) As Collection

    Dim rs As ADODB.Recordset
    Dim q As String
    Dim ordenesTrabajo As New Collection
    Dim deta As New DetalleOrdenTrabajo
    q = "SELECT" _
        & " * "



    If withDetalles And withFacturados Then
        q = q & ",IFNULL((SELECT SUM(cantidad) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadFacturada_ & " AND id_detalle_pedido=dp.id),0) AS FacturadosCantidad"
        q = q & ",IFNULL((SELECT SUM(((monto * cantidad)/1)*tipo_cambio) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadFacturada_ & " AND id_detalle_pedido=dp.id),0) AS FacturadosMonto"
    End If

    If withDetalles And withEntregados Then
        q = q & ",IFNULL((SELECT SUM(cantidad) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadEntregada_ & " AND id_detalle_pedido=dp.id),0) AS EntregadosCantidad"
    End If

    If withDetalles And withFabricados Then
        q = q & ",IFNULL((SELECT SUM(cantidad) FROM detalles_pedidos_cantidad WHERE tipo_cantidad=" & TipoCantidadOT.CantidadFabricada_ & " AND id_detalle_pedido=dp.id),0) AS FabricadosCantidad"
    End If


    q = q & " FROM pedidos p" _
        & " LEFT JOIN clientes c" _
        & " ON (c.id = p.idCliente)" _
        & " LEFT JOIN clientes c2" _
        & " ON (c2.id = p.idClienteFacturar)" _
        & " LEFT JOIN AdminConfigMonedas m" _
        & " ON (m.id = p.idMoneda)" _
        & " LEFT JOIN usuarios u" _
        & " ON (u.id = p.idUsuario)" _
        & " LEFT JOIN usuarios u1" _
        & " ON (u1.id = p.idUsuarioAprobado)" _
        & " LEFT JOIN usuarios u2" _
        & " ON (u2.id = p.idUsuarioModificado)" _
        & " LEFT JOIN usuarios u3" _
        & " ON (u3.id = p.idUsuarioFinalizado)"
    q = q & " LEFT JOIN AdminConfigIVA iva" _
        & " ON (iva.idIVA = c.iva)" _
        & " LEFT JOIN AdminConfigFacturasTipos tfact" _
        & " ON (tfact.id = iva.tipo_factura)" _
        & " LEFT JOIN Localidades ON (c.id_localidad = Localidades.ID)" _
        & " LEFT JOIN Provincia ON (Localidades.idProvincia = Provincia.ID)" _
        & " LEFT JOIN Pais ON (Provincia.idPais = Pais.ID)" _


If withDetalles Then
        q = q & " LEFT JOIN detalles_pedidos dp   ON (dp.idPedido=p.id) " _
            & " LEFT JOIN stock ON dp.idPieza=stock.id "
    End If

    If withDetalles And withPorcentajeTareasFinalizadas Then

    End If


    q = q & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If
    q = q & " ORDER BY p.id DESC"
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAll = New Collection
    Const monedaTabla As String = "m"
    Const clienteTabla As String = "c"
    Const clienteFacturarTabla As String = "c2"
    Const ivaTabla As String = "iva"
    Const tipoFacturaTabla As String = "tfact"
    Const usuarioTabla As String = "u"
    Const usuarioAprobadoTabla As String = "u1"
    Const usuarioModificadoTabla As String = "u2"
    Const usuarioFinalizadoTabla As String = "u3"
    Dim esta As Boolean
    Dim Ot As OrdenTrabajo
    While Not rs.EOF
        Set Ot = Map(rs, fieldsIndex, TABLA_PEDIDO, monedaTabla, clienteTabla, usuarioTabla, usuarioAprobadoTabla, usuarioModificadoTabla, usuarioFinalizadoTabla, ivaTabla, tipoFacturaTabla, clienteFacturarTabla)
        esta = False
        If BuscarEnColeccion(ordenesTrabajo, CStr(Ot.id)) Then
            esta = True
            Set Ot = ordenesTrabajo.item(CStr(Ot.id))
        End If


        'If Ot.Id = 1836 Then Stop
        If withDetalles Then
            Set deta = DAODetalleOrdenTrabajo.Map(rs, fieldsIndex, "dp", "stock", withEntregados, withFabricados, withFacturados, withPorcentajeTareasFinalizadas)
            deta.idpedido = Ot.id
            deta.IdMoneda = Ot.moneda.id
            If IsSomething(deta) Then Ot.Detalles.Add deta, CStr(deta.id)
        End If



        If Not BuscarEnColeccion(ordenesTrabajo, CStr(Ot.id)) Then
            ordenesTrabajo.Add Ot, CStr(Ot.id)
        End If

        rs.MoveNext
    Wend
    Set FindAll = ordenesTrabajo
End Function



Public Function Map(ByRef rs As Recordset, _
                    ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef monedaTableNameOrAlias As String = vbNullString, _
                    Optional ByRef clienteTableNameOrAlias As String = vbNullString, _
                    Optional ByRef usuarioTableNameOrAlias As String = vbNullString, _
                    Optional ByRef usuarioAprobadoTableNameOrAlias As String = vbNullString, _
                    Optional ByRef usuarioModificadoTableNameOrAlias As String = vbNullString, _
                    Optional ByRef usuarioFinalizadoTableNameOrAlias As String = vbNullString, _
                    Optional ByRef ivaTableNameOrAlias As String = vbNullString, _
                    Optional ByRef tipoFacturaTableNameOrAlias As String = vbNullString, _
                    Optional ByRef clienteFacturarTableNameOrAlias As String = vbNullString _
                    ) As OrdenTrabajo

    Dim tmpOrdenTrabajo As OrdenTrabajo
    Dim id As Variant
    id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If id > 0 Then
        Set tmpOrdenTrabajo = New OrdenTrabajo
        tmpOrdenTrabajo.id = id


        tmpOrdenTrabajo.AnticipoFacturadoIdFactura = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ANTICIPO_FACTURA_ID)
        tmpOrdenTrabajo.descripcion = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DESCRIPCION)
        tmpOrdenTrabajo.FechaEntrega = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_ENTREGA)
        tmpOrdenTrabajo.fechaCreado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_CREADO)
        tmpOrdenTrabajo.Activa = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ACTIVO)
        tmpOrdenTrabajo.Entregada = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ENTREGADO)
        tmpOrdenTrabajo.FechaCerrado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_CERRADO)
        tmpOrdenTrabajo.Descuento = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_DESCUENTO)
        tmpOrdenTrabajo.fechaAprobado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_APROBADO)
        tmpOrdenTrabajo.FechaModificado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_MODIFICADO)
        tmpOrdenTrabajo.estado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ESTADO)
        tmpOrdenTrabajo.TipoOrden = GetValue(rs, fieldsIndex, tableNameOrAlias, "tipo_orden")
        tmpOrdenTrabajo.NroPresupuesto = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NUMERO_PRESUPUESTO)
        tmpOrdenTrabajo.Anticipo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ANTICIPO)
        tmpOrdenTrabajo.AnticipoFacturado = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ANTICIPO_FACTURADO)
        tmpOrdenTrabajo.FormaDePagoAnticipo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FORMA_PAGO_ANTICIPO)
        tmpOrdenTrabajo.MismaFechaEntregaParaDetalles = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_MISMA_FECHA_DETALLES)
        tmpOrdenTrabajo.CantDiasAnticipo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ANTICIPO_DIAS)
        tmpOrdenTrabajo.CantDiasSaldo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_SALDO_DIAS)
        tmpOrdenTrabajo.FormaDePagoSaldo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FORMA_PAGO_SALDO)
        tmpOrdenTrabajo.IdMoneda = GetValue(rs, fieldsIndex, tableNameOrAlias, "idMoneda")

        'marcos
        tmpOrdenTrabajo.OTMarcoIdPadre = GetValue(rs, fieldsIndex, tableNameOrAlias, "id_ot_padre")
        tmpOrdenTrabajo.FechaInicioMarco = GetValue(rs, fieldsIndex, tableNameOrAlias, "marco_fecha_inicio")
        tmpOrdenTrabajo.FechaFinMarco = GetValue(rs, fieldsIndex, tableNameOrAlias, "marco_fecha_fin")
        tmpOrdenTrabajo.UltimaFechaActualizacionPrecios = GetValue(rs, fieldsIndex, tableNameOrAlias, "ultima_fecha_actualizacion_precios")
        tmpOrdenTrabajo.MontoTopeMarco = GetValue(rs, fieldsIndex, tableNameOrAlias, "marco_monto_tope")
        tmpOrdenTrabajo.ContaduriaImpreso = GetValue(rs, fieldsIndex, tableNameOrAlias, "contaduria_impreso")

        If LenB(clienteTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.cliente = DAOCliente.Map(rs, fieldsIndex, clienteTableNameOrAlias, ivaTableNameOrAlias, "Localidades", "Pais", "Provincia")
        If LenB(clienteFacturarTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.ClienteFacturar = DAOCliente.Map(rs, fieldsIndex, clienteFacturarTableNameOrAlias, ivaTableNameOrAlias)
        If LenB(monedaTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.moneda = DAOMoneda.Map(rs, fieldsIndex, monedaTableNameOrAlias)
        If LenB(usuarioTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.usuario = DAOUsuarios.Map(rs, fieldsIndex, usuarioTableNameOrAlias)
        If LenB(usuarioAprobadoTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.UsuarioAprobado = DAOUsuarios.Map(rs, fieldsIndex, usuarioAprobadoTableNameOrAlias)
        If LenB(usuarioModificadoTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.UsuarioModificado = DAOUsuarios.Map(rs, fieldsIndex, usuarioModificadoTableNameOrAlias)
        If LenB(usuarioFinalizadoTableNameOrAlias) > 0 Then Set tmpOrdenTrabajo.UsuarioFinalizado = DAOUsuarios.Map(rs, fieldsIndex, usuarioFinalizadoTableNameOrAlias)
    End If

    Set Map = tmpOrdenTrabajo

End Function




Public Function Save(Ot As OrdenTrabajo) As Boolean
    conectar.BeginTransaction
    If Guardar(Ot, True) Then
        conectar.CommitTransaction
        Save = True
    Else
        conectar.RollBackTransaction
        Save = False
    End If
End Function


Public Function Guardar(Ot As OrdenTrabajo, Optional Cascade As Boolean = True, Optional publishEvent As Boolean = False) As Boolean
    On Error GoTo E
    Dim q As String
    Dim esUpdate As Boolean

    If Ot.id = 0 Then
        q = "INSERT INTO pedidos" _
            & " (descripcion," _
            & " idCliente,idClienteFacturar," _
            & " fechaEntrega," _
            & " fechaCreado," _
            & " nroPresupuesto," _
            & " estado,tipo_orden," _
            & " activo," _
            & " entregado," _
            & " idUsuario," _
            & " dto," _
            & " idMoneda," _
            & " idUsuarioAprobado," _
            & " idUsuarioModificado," _
            & " idUsuarioFinalizado," _
            & " stockDescontado," _
            & " anticipo," _
            & " anticipo_facturado," _
            & " forma_pago_anticipo," _
            & " misma_fecha_entrega_detalles," _
            & " anticipo_dias,id_anticipo_factura,"
        q = q & " saldo_dias," _
            & " forma_pago_saldo, id_ot_padre, marco_fecha_inicio, marco_fecha_fin, marco_monto_tope)" _
            & " Values" _
            & " (" & conectar.Escape(UCase(Ot.descripcion)) & "," _
            & conectar.GetEntityId(Ot.cliente) & "," _
            & conectar.GetEntityId(Ot.ClienteFacturar) & "," _
            & conectar.Escape(Ot.FechaEntrega) & "," _
            & conectar.Escape(Ot.fechaCreado) & "," _
            & conectar.Escape(Ot.NroPresupuesto) & "," _
            & conectar.Escape(Ot.estado) & "," & conectar.Escape(Ot.TipoOrden) & "," _
            & conectar.Escape(Ot.Activa) & "," _
            & conectar.Escape(Ot.Entregada) & "," _
            & conectar.GetEntityId(Ot.usuario) & "," _
            & conectar.Escape(Ot.Descuento) & "," _
            & conectar.GetEntityId(Ot.moneda) & "," _
            & conectar.GetEntityId(Ot.UsuarioAprobado) & "," _
            & conectar.GetEntityId(Ot.UsuarioModificado) & "," _
            & conectar.GetEntityId(Ot.UsuarioFinalizado) & "," _
            & conectar.Escape(Ot.StockDescontado) & "," _
            & conectar.Escape(Ot.Anticipo) & "," _
            & conectar.Escape(Ot.AnticipoFacturado) & ","
        q = q & conectar.Escape(Ot.FormaDePagoAnticipo) & "," _
            & conectar.Escape(Ot.MismaFechaEntregaParaDetalles) & "," _
            & conectar.Escape(Ot.CantDiasAnticipo) & "," _
            & conectar.Escape(Ot.AnticipoFacturadoIdFactura) & "," _
            & conectar.Escape(Ot.CantDiasSaldo) & "," _
            & conectar.Escape(Ot.FormaDePagoSaldo) & "," _
            & conectar.Escape(Ot.OTMarcoIdPadre) & "," _
            & conectar.Escape(Ot.FechaInicioMarco) & "," _
            & conectar.Escape(Ot.FechaFinMarco) & ", " & conectar.Escape(Ot.MontoTopeMarco) & ")"

        Guardar = conectar.execute(q)

        Dim id As Long

        If Guardar Then
            conectar.UltimoId "pedidos", id
            Ot.id = id
            If Ot.id <> 0 Then DAOOrdenTrabajoHistorial.agregar Ot, "Pedido creado"
        End If
    Else
        esUpdate = True

        q = "update pedidos" _
            & " SET" _
            & " descripcion = " & conectar.Escape(Ot.descripcion) & " ," _
            & " idCliente = " & conectar.GetEntityId(Ot.cliente) & " , idClienteFacturar = " & conectar.GetEntityId(Ot.ClienteFacturar) & " ," _
            & " fechaEntrega = " & conectar.Escape(Ot.FechaEntrega) & " ," _
            & " fechaCreado = " & conectar.Escape(Ot.fechaCreado) & " ," _
            & " nroPresupuesto = " & conectar.Escape(Ot.NroPresupuesto) & " ," _
            & " estado = " & conectar.Escape(Ot.estado) & " , estado = " & conectar.Escape(Ot.estado) & " ," _
            & " activo = " & conectar.Escape(Ot.Activa) & " ," _
            & " entregado = " & conectar.Escape(Ot.Entregada) & " ," _
            & " fechaCerrado = " & conectar.Escape(Ot.FechaCerrado) & " ," _
            & " idUsuario = " & conectar.GetEntityId(Ot.usuario) & " ," _
            & " dto = " & conectar.Escape(Ot.Descuento) & " ," _
            & " idMoneda = " & conectar.GetEntityId(Ot.moneda) & " ," _
            & " fechaAprobado = " & conectar.Escape(Ot.fechaAprobado) & " ," _
            & " idUsuarioAprobado = " & conectar.GetEntityId(Ot.UsuarioAprobado) & " ," _
            & " fechaModificado = " & conectar.Escape(Now) & " ," _
            & " idUsuarioModificado = " & funciones.GetUserObj.id & " ," _
            & " idUsuarioFinalizado = " & conectar.GetEntityId(Ot.UsuarioFinalizado) & " ," _
            & " stockDescontado = " & conectar.Escape(Ot.StockDescontado) & " ," _
            & " anticipo = " & conectar.Escape(Ot.Anticipo) & " ," _
            & " anticipo_facturado = " & conectar.Escape(Ot.AnticipoFacturado) & " ," _
            & " forma_pago_anticipo = " & conectar.Escape(Ot.FormaDePagoAnticipo) & " ," _
            & " misma_fecha_entrega_detalles = " & conectar.Escape(Ot.MismaFechaEntregaParaDetalles) & " ," _
            & " anticipo_dias = " & conectar.Escape(Ot.CantDiasAnticipo) & " ,"
        q = q & " saldo_dias = " & conectar.Escape(Ot.CantDiasSaldo) & " ," _
            & " forma_pago_saldo = " & conectar.Escape(Ot.FormaDePagoSaldo) & ", " _
            & " id_ot_padre = " & conectar.Escape(Ot.OTMarcoIdPadre) & ", " _
            & " marco_fecha_inicio = " & conectar.Escape(Ot.FechaInicioMarco) & ", " _
            & " marco_fecha_fin = " & conectar.Escape(Ot.FechaFinMarco) & ", " _
            & " marco_monto_tope = " & conectar.Escape(Ot.MontoTopeMarco) & "," _
            & " id_anticipo_factura = " & conectar.Escape(Ot.AnticipoFacturadoIdFactura) _
            & " WHERE" _
            & " id = " & Ot.id

        Guardar = conectar.execute(q)
    End If
    Dim c As Long
    c = 0
    If Guardar Then
        c = Ot.FechasPreciosMarco.count     'fuerzo la recarga
        conectar.execute "DELETE FROM pedidos_fechas_precios WHERE id_ot_marco = " & Ot.id
        If Ot.EsMarco Then
            Dim FechaPrecio As Variant
            For Each FechaPrecio In Ot.FechasPreciosMarco
                conectar.execute "INSERT INTO pedidos_fechas_precios VALUES (" & Ot.id & ", " & conectar.Escape(FechaPrecio) & ")"
            Next FechaPrecio
        End If

        If Cascade Then


            Dim deta As DetalleOrdenTrabajo
            conectar.execute "DELETE FROM detalles_pedidos WHERE idPedido = " & Ot.id
            conectar.execute "DELETE FROM detalles_pedidos_conjuntos WHERE idPedido = " & Ot.id
            c = 0
            For Each deta In Ot.Detalles
                If Not IsSomething(deta.OrdenTrabajo) Then
                    Set deta.OrdenTrabajo = Ot
                End If
                c = c + 1
                'Debug.Print c
                deta.id = 0    'asi me lo toma como insert ==> (deja de hardcordear raulo) => comela toda nicocolasba -> sr, si sr.
                deta.Descuento = Ot.Descuento    '>>>> paso el descuento de la OT al dto del detalle(cuando se graba)
                deta.IdMoneda = Ot.moneda.id   '>>> pongo la momneda en el detalle, para hacer informes 10-10-13
                If Not DAODetalleOrdenTrabajo.Save(deta) Then GoTo E

            Next deta
            DAOOrdenTrabajoHistorial.agregar Ot, "Pedido editado"
        End If

        If esUpdate And publishEvent Then
            DAOEvento.Publish Ot.id, TipoEventoBroadcast.TEB_OrdenTrabajoModificada
        End If
    Else
        GoTo E
    End If

    Exit Function
E:
    Guardar = False
    
End Function
Public Function AprobarOT(T As OrdenTrabajo, Optional progressbar As Object) As Boolean
   ' If T.TipoOrden = OT_ENTREGA Then
   ' MsgBox "Orden de entrega, consultar con sistemas"
  '  Exit Function
'End If
    
    
    conectar.BeginTransaction
    Dim claseP As New classPlaneamiento
    AprobarOT = True
    Dim usu_ant As clsUsuario
    Dim fecha_ant As Date
    Dim estado_ant As EstadoOrdenTrabajo

    Set usu_ant = T.UsuarioAprobado
    fecha_ant = T.fechaAprobado
    estado_ant = T.estado

    T.fechaAprobado = Now
    Set T.UsuarioAprobado = funciones.GetUserObj

    If CambiarEstado(T, EstadoOT_EnEspera) Then

        If DAOTiemposProceso.crear(T, progressbar) Then
            If Not T.StockDescontado Then
                If Not DescontarReserva(T) Then
                    GoTo err44
                End If
            End If
            If Not DAOOrdenTrabajoHistorial.agregar(T, "OT Aprobada") Then
                GoTo err44


            Else
                conectar.CommitTransaction
                DAOEvento.Publish T.id, TipoEventoBroadcast.TEB_OrdenTrabajoAprobada

                If T.Anticipo > 0 Then
                    DAOEvento.Publish T.id, TipoEventoBroadcast.TEB_OrdenConAnticipoAprobada
                End If
                AprobarOT = True
            End If
        Else
            GoTo err44
        End If
    Else
        GoTo err44
    End If

    Exit Function
err44:
    AprobarOT = False
    Set T.UsuarioAprobado = usu_ant
    T.fechaAprobado = fecha_ant
    T.estado = estado_ant
    conectar.RollBackTransaction
End Function
Public Function desactivar(T As OrdenTrabajo) As Boolean
    Dim estado_ant As EstadoOrdenTrabajo
    estado_ant = T.estado
    conectar.BeginTransaction
    If T.estado = EstadoOT_Desactivado Then
        GoTo err1
    Else
        T.estado = EstadoOT_Desactivado
        If Not Guardar(T) Then GoTo err1
    End If

    desactivar = True
    conectar.CommitTransaction

    DAOEvento.Publish T.id, TipoEventoBroadcast.TEB_OrdenTrabajoAnulada

    Exit Function
err1:
    T.estado = estado_ant
    conectar.RollBackTransaction
End Function



Public Function informePedidoPlaneamiento(id As Long, DIALOGO As Boolean, rsSectores As ADODB.Recordset) As Boolean     '1- enviar 2-imprimir
    Dim s As New classStock
    On Error GoTo err2
    informePedidoPlaneamiento = True
    Dim strsql As String, strsql2 As String
    Dim rs As ADODB.Recordset
    Set rs = conectar.RSFactory("select idPieza, cantidad from detalles_pedidos  where idPedido=" & id)
    Dim c As Long
    c = 0
    While Not rs.EOF
        c = c + s.TiemposPieza(rs!idPieza, rs!Cantidad)
        rs.MoveNext
    Wend
    Dim tim As String
    tim = " Min."
    If c > 60 Then
        c = c / 60
        tim = " Hs."
    End If
    pedido_planeamiento.Sections("pata").Controls("lblCMO").caption = Math.Round(c, 2) & tim
    Set s = Nothing
    Set rs = conectar.RSFactory("select c.cuit,c.razon,p.id,p.descripcion,p.fechaEntrega as fe, p.fechaCreado, p.nroPresupuesto from pedidos p, clientes c where p.idcliente=c.id and p.id=" & id)
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
    End If

    If c = 1 Then
        'presu = rs!NroPresupuesto
        'Cuit = rs!Cuit
        pedido_planeamiento.Sections("cabeza").Controls("lblOT").caption = Format("00000", rs!id)
        pedido_planeamiento.Sections("cabeza").Controls("lblCliente").caption = rs!razon
        pedido_planeamiento.Sections("cabeza").Controls("lblReferencia").caption = rs!descripcion
        pedido_planeamiento.Sections("cabeza").Controls("lblfechaEntrega").caption = rs!fe
        pedido_planeamiento.Sections("cabeza").Controls("lblfechaCreado").caption = rs!fechaCreado
        pedido_planeamiento.Sections("s3").Controls("lblbarcode").caption = "*" & Format(id, "0000") & "*"
        Dim RS_1 As Recordset
        'lo mando yo
        'Set RS_1 = claseS.resumenTareasPedido(Id)
        Set RS_1 = rsSectores

        'TotFab = 0
        'TotRes = 0
        'TotIt = 0
        c = 0

        If c > 0 Then rs.MoveFirst
        Set pedido_planeamiento.DataSource = RS_1
        'pedido_planeamiento.PrintReport DIALOGO
        pedido_planeamiento.Show 1
    End If
    Exit Function
err2:
    MsgBox Err.Description
    informePedidoPlaneamiento = False
End Function


Public Function informePiezaMateriales(id As Long, Origen As Integer, DIALOGO As Boolean, Optional colListado As Collection) As Boolean    '1- enviar 2-imprimir
    On Error GoTo err2

    Dim Obj As PageSet.PrinterControl
    Set Obj = New PrinterControl
    Dim titulo As String
    Dim strsql As String
    Dim A As String
    Dim r As Recordset
    informePiezaMateriales = True
    Dim strsql2 As String
    If Origen = 2 Then
        titulo = "Presupuesto:"
        strsql = "select c.razon,p.id,p.detalle,p.fechaEntrega as fe from presupuestos p,clientes c where p.idcliente=c.id and p.id=" & id
    ElseIf Origen = 1 Then
        Dim fecha_entre As String
        titulo = "Pedido:"
        strsql = "select c.razon,p.id,p.descripcion as detalle,p.fechaEntrega as fe from pedidos p,clientes c where p.idcliente=c.id and p.id=" & id

    ElseIf Origen = 3 Then
        titulo = "Conjunto:"
        strsql = "select c.razon,s.detalle, s.id as id from stock s inner join clientes c on s.id_cliente=c.id where s.id=" & id

    ElseIf Origen = 4 And IsSomething(colListado) Then

        Dim lista As String
        lista = funciones.JoinCollectionValues(colListado, ",", "Id")
        Dim ListaPiezas As New Dictionary

        Dim dp As DetalleOrdenTrabajo
        Dim s As Pieza
        Dim P As OrdenTrabajo
        For Each P In colListado
            Set P.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(P.id)
            For Each dp In P.Detalles
                ListaPiezas.Add dp.Pieza, dp.CantidadPedida
            Next dp
        Next P


        titulo = "Materiales en común"
        strsql = "select c.razon,s.detalle, s.id as id from stock s inner join clientes c on s.id_cliente=c.id where s.id in (" & lista & ")"

    End If


    Dim rs As ADODB.Recordset
    Set rs = conectar.RSFactory(strsql)

    If Origen = 1 And Not rs.EOF Then
        rs.MoveFirst
        fecha_entre = Format(rs!fe, "dd-mm-yyyy")
    End If
    'Me.ejecutar_consulta strsql



    Dim totTit As String
    Dim razon As String
    Dim detalle As String
    totTit = titulo & " " & Format("00000", rs!id)

    If Origen = 1 Then
        totTit = totTit & " (Entrega: " & fecha_entre & ")"
    ElseIf Origen = 4 Then
        totTit = "OTs " & funciones.JoinCollectionValues(colListado, ",", "Id")
    End If


    Materiales.Sections("cabeza").Controls("lblOT").caption = totTit

    If Origen = 4 Then
        razon = "Varias OT"
    Else
        razon = rs!razon
    End If

    Materiales.Sections("cabeza").Controls("lblCliente").caption = razon

    If Origen = 4 Then
        detalle = "Varios"
    Else
        detalle = rs!detalle
    End If
    Materiales.Sections("cabeza").Controls("lblReferencia").caption = detalle
    Materiales.Sections("s3").Controls("lblbarcode").caption = "*" & Format(id, "0000") & "*"



    If Origen = 4 Then
        Set r = rs_materiales(id, Origen, ListaPiezas)
    Else
        Set r = rs_materiales(id, Origen)
    End If
    Set Materiales.DataSource = r
    Obj.ChngOrientationLandscape
    Materiales.Show 1
    Obj.ReSetOrientation    'This resets the printer to portrait.
    Exit Function
err2:
    MsgBox Err.Description
    informePiezaMateriales = False
    Obj.ReSetOrientation
End Function

Private Function rs_materiales(id As Long, Origen As Integer, Optional ListaPiezas As Dictionary) As Recordset
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

    ElseIf Origen = 4 Then
        Ot = False
        otro = False
        presu = False
    End If


    Dim stock As New classStock
    If Origen = 4 Then
        Set r_piezas = stock.ListaPiezas(id, Ot, presu, otro, , ListaPiezas)
    Else
        Set r_piezas = stock.ListaPiezas(id, Ot, presu, otro)
    End If
    'r_piezas.MoveFirst


    Dim Pieza As Long
    Dim cantidad_p As Long
    Dim strsql As String


    Dim IdMaterial As Long
    Dim Largo As Double
    Dim Ancho As Double
    Dim Grupo As String
    Dim descrip As String
    Dim rubro As String
    Dim Espesor As Double
    Dim largop As Double
    Dim anchop As Double
    Dim Cantidad As Long
    Dim PesoXUnidad As Double
    Dim moneda As String
    Dim codigo As String
    Dim id_Unidad As Long
    Dim Valor As Double
    Dim fec_act As Date


    Dim medida As Double
    Dim unidad As Double
    Dim totUnit As Double
    Dim UN As String

    While Not r_piezas.EOF
        Pieza = CLng(r_piezas!idPieza)
        cantidad_p = CDbl(r_piezas!Cantidad)

        strsql = " select mo.nombre_corto,m.valor_unitario,m.fecha_valor as fecha_actualizacion,m.descripcion,r.rubro,g.grupo,m.espesor,m.codigo,m.id as id_material, dm.largo,dm.ancho,largoTerm,AnchoTerm,dm.cantidad,id_unidad,pesoxunidad,r.rubro,g.grupo from desarrollo_material dm, rubros r, grupos g, materiales m, AdminConfigMonedas mo where m.id_moneda=mo.id and dm.id_material=m.id and m.id_rubro = r.id and  m.id_grupo=g.id and id_pieza=" & Pieza
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
            Valor = r_1!valor_unitario
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
            agregarARSmateriales r_mat, id_Unidad, IdMaterial, unidad, totUnit, codigo, descrip, rubro, Grupo, Espesor, UN, Valor, fec_act, moneda

            r_1.MoveNext
        Wend


        r_piezas.MoveNext
    Wend

    If r_mat.RecordCount > 0 Then r_mat.MoveFirst


    Set rs_materiales = r_mat
    Exit Function
err4:
    MsgBox Err.Description, vbCritical
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
        r_mat!Valor = valor_unit
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
Public Function InformePiezasFabricadas(T As OrdenTrabajo) As Boolean
    On Error GoTo err2
    Dim rInci As Recordset
    InformePiezasFabricadas = True
    Dim strsql As String, strsql2 As String
    Dim rs As Recordset
    'Set rs = conectar.RSFactory(strsql)
    If Not T Is Nothing Then
        strsql = "select count(id) as inci from Incidencias where idReferencia=" & T.id & " and origen=2"
        Set rInci = conectar.RSFactory(strsql)


        dsrResumenFabricacion.Sections("cabeza").Controls("Etiqueta2").caption = IIf(T.EsMarco, "Contrato Marco Nº:", "OT Nº") & " " & Format("00000", T.id)
        dsrResumenFabricacion.Sections("cabeza").Controls("lblCliente").caption = T.cliente.razon   'r s!Razon
        dsrResumenFabricacion.Sections("cabeza").Controls("lblReferencia").caption = T.descripcion    ' rs!Descripcion
        dsrResumenFabricacion.Sections("cabeza").Controls("lblfechaEntrega").caption = T.FechaEntrega    ' rs!fe
        dsrResumenFabricacion.Sections("cabeza").Controls("lblfechaCreado").caption = T.fechaCreado  'rs!FechaCreado
        dsrResumenFabricacion.Sections("s3").Controls("lblbarcode").caption = "*" & Format(T.id, "0000") & "*"


        strsql = "SELECT  if(s.conjunto=-1,'Un','Conj') as unidad ,dp.cantidad,dp.item,  s.detalle, dp.nota, " _
                 & " (SELECT  SUM(cantidad)  FROM detalles_pedidos dp2  Where dp2.idPieza = dp.idPieza       AND dp2.id <> dp.id) AS cantidad_fabricada " _
                 & " FROM detalles_pedidos dp  LEFT JOIN stock s   ON dp.idPieza = s.id " _
                 & " Where dp.idpedido =  " & T.id


        Set rs = conectar.RSFactory(strsql)
        Set dsrResumenFabricacion.DataSource = rs
        dsrResumenFabricacion.Show 1



    End If
    Exit Function
err2:
    MsgBox Err.Description
    InformePiezasFabricadas = False
End Function

Public Function imprimirEtiquetas(idpedido As Long) As Boolean
    On Error GoTo err44
    imprimirEtiquetas = True
    Dim rs_temp As Recordset
    Dim r As Recordset

    Set rs_temp = New Recordset
    With rs_temp
        .Fields.Append "idPedido", adVarChar, 255, adFldUpdatable     ' And adFldIsNullable
        .Fields.Append "oc", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "item", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullableç
        .Fields.Append "cantidad", adVarChar, 255, adFldUpdatable    ' And adFldIsNullableç
        .Fields.Append "cantidad_total", adVarChar, 255, adFldUpdatable    ' And adFldIsNullableç
        .Fields.Append "cliente", adVarChar, 255, adFldUpdatable    ' And adFldIsNullableç

    End With
    rs_temp.Open

    Dim Cant As Long
    Dim oc As String
    Dim it As String
    Dim detalle As String
    Dim razon As String
    Dim cantAimprimir As Long
    Dim Aimprimir As Long

    Dim P As Long

    Set r = conectar.RSFactory("select p.id,p.descripcion,c.razon,dp.item,s.detalle,dp.cantidad,dp.cantidad_fabricados,dp.cantidad_entregada from pedidos p inner join detalles_pedidos dp on dp.idPedido=p.id inner join clientes c on p.idCliente=c.id inner join stock s on dp.idPieza=s.id where idPedido=" & idpedido)
    While Not r.EOF
        Cant = r!Cantidad
        oc = r!descripcion
        it = r!item
        detalle = r!detalle
        razon = r!razon

        cantAimprimir = Int(Cant / 4)
        Aimprimir = Cant Mod 4
        If Aimprimir > 0 Then
            cantAimprimir = cantAimprimir + 1
        End If



        For P = 1 To cantAimprimir

            With rs_temp
                .AddNew
                !cliente = razon
                !idpedido = idpedido
                !oc = oc
                !detalle = detalle
                !item = it
                !Cantidad = P
                !cantidad_total = Cant

                .Update
            End With
        Next P

        r.MoveNext
    Wend

    rs_temp.MoveFirst
    Set etiquetas.DataSource = rs_temp
    etiquetas.Show
    Exit Function
err44:
    imprimirEtiquetas = False
End Function
Public Function informePedido(T As OrdenTrabajo, DIALOGO As Boolean, Sector) As Boolean      '1- enviar 2-imprimir
    On Error GoTo err2
    Dim rInci As Recordset
    informePedido = True
    Dim strsql As String, strsql2 As String
    Dim rs As Recordset
    'Set rs = conectar.RSFactory(strsql)
    If Not T Is Nothing Then
        strsql = "select count(id) as inci from Incidencias where idReferencia=" & T.id & " and origen=2"
        Set rInci = conectar.RSFactory(strsql)

        If rInci!inci > 0 Then
            informeOT.Sections("pata").Controls("lblIncidencias").caption = "Existen " & rInci!inci & " incidencias para esta OT. Controlar en primera instancia"
        Else
            informeOT.Sections("pata").Controls("lblIncidencias").caption = "No hay incidencias al momento de emitir la OT, controlar periodicamente."
        End If

        informeOT.Sections("cabeza").Controls("Etiqueta2").caption = IIf(T.EsMarco, "Contrato Marco Nº:", "OT Nº") & " " & Format("00000", T.id)
        informeOT.Sections("cabeza").Controls("lblCliente").caption = T.cliente.razon   'r s!Razon
        informeOT.Sections("cabeza").Controls("lblReferencia").caption = T.descripcion    ' rs!Descripcion
        informeOT.Sections("cabeza").Controls("lblfechaEntrega").caption = T.FechaEntrega    ' rs!fe

        informeOT.Sections("cabeza").Controls("lblSector").caption = Sector

        strsql = "select CONCAT('*R', LPAD(CONVERT(dp.id,CHAR(8)),8,'0'),'*') AS codigo,dp.fechaEntrega as entregaItem,dp.item, dp.cantidad as stotal, dp.reserva_stock as reserva,dp.cantidad-dp.reserva_stock as fabricar, s.detalle,dp.nota, if(s.conjunto=-1,'Un','Conj') as unidad from detalles_pedidos dp, stock s where dp.idPieza=s.id and dp.idPedido=" & Format("00000", T.id)    'nroOT

        'Me.ejecutar_consulta strsql
        Set rs = conectar.RSFactory(strsql)
        Dim TotFab As Long
        Dim TotRes As Long
        Dim TotIt As Long

        TotFab = 0
        TotRes = 0
        TotIt = 0

        While Not rs.EOF
            TotFab = TotFab + rs!fabricar
            TotRes = TotRes + rs!reserva
            rs.MoveNext
        Wend

        informeOT.Sections("pata").Controls("lblCantRes").caption = TotRes & " Un."
        informeOT.Sections("pata").Controls("lblCantIt").caption = T.Detalles.count
        informeOT.Sections("pata").Controls("lblCantFab").caption = TotFab & " Un."

        If T.EsHija Then
            informeOT.Sections("pata").Controls("lblContratoMarcoDependencia").caption = "Forma parte del Contrato Marco Nº " & T.OTMarcoIdPadre
        Else
            informeOT.Sections("pata").Controls("lblContratoMarcoDependencia").caption = vbNullString
        End If


        informeOT.Sections("s3").Controls("lblbarcode").caption = "*" & Format(T.id, "0000") & "*"
        Set informeOT.DataSource = rs

        'informeOT.Show
        informeOT.PrintReport DIALOGO


    End If
    Exit Function
err2:
    MsgBox Err.Description
    informePedido = False
End Function

Public Function grillaTiempos(T As OrdenTrabajo, DIALOGO As Boolean, Sector) As Boolean      '1- enviar 2-imprimir
    On Error GoTo err2
    grillaTiempos = True
    Dim strsql As String, strsql2 As String

    Dim rs As Recordset


    If Not T Is Nothing Then
        gruilla_tiempos.Sections("cabeza").Controls("lblOT").caption = Format("00000", T.id)
        gruilla_tiempos.Sections("cabeza").Controls("lblCliente").caption = T.cliente.razon    ' rs!Razon
        gruilla_tiempos.Sections("cabeza").Controls("lblReferencia").caption = T.descripcion    ' rs!Descripcion
        gruilla_tiempos.Sections("cabeza").Controls("lblfechaEntrega").caption = T.FechaEntrega    ' rs!fe
        gruilla_tiempos.Sections("cabeza").Controls("lblfechaCreado").caption = T.fechaCreado    ' rs!FechaCreado
        gruilla_tiempos.Sections("cabeza").Controls("lblSector").caption = Sector
        Set rs = conectar.RSFactory("select count(id) from pedidos limit 1")
        Set gruilla_tiempos.DataSource = rs
        gruilla_tiempos.PrintReport DIALOGO

    End If
    Exit Function
err2:
    MsgBox Err.Description
    grillaTiempos = False
End Function

Public Function GetDTOSectoresTiempo(ordenes_trabajo_ids As Collection) As Collection
    'On Error GoTo e

    If ordenes_trabajo_ids.count = 0 Then
        Set GetDTOSectoresTiempo = New Collection
        Exit Function
    End If

    Dim orden_trabajo_id As Variant
    Dim ordenes As Collection
    Dim filter As String
    Dim Ot As OrdenTrabajo
    Dim deta As DetalleOrdenTrabajo
    Dim detallesHijos As Collection
    Dim detaHijo As DetalleOTConjuntoDTO
    Dim proceso As PlaneamientoTiempoProceso
    Dim procesos As Collection

    Dim sectoresTiempo As New Collection

    For Each orden_trabajo_id In ordenes_trabajo_ids
        filter = filter & orden_trabajo_id & ", "
    Next orden_trabajo_id
    If ordenes_trabajo_ids.count > 0 Then
        filter = " p.id IN (" & Left$(filter, Len(filter) - 2) & ")"
    End If

    Set ordenes = DAOOrdenTrabajo.FindAll(filter)

    For Each Ot In ordenes
        Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, , True, , True)
        For Each deta In Ot.Detalles
            'Set procesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(deta.Id, deta.pieza.Id)
            Set procesos = DAOTiemposProceso.FindAllByDetallePedidoId(deta.id)

            For Each proceso In procesos
                AddTarea sectoresTiempo, proceso.Tarea, deta.CantidadPedida, proceso.OperariosCotizado, proceso.TiempoCotizado, deta.Cantidad_Fabricada, proceso

            Next proceso

            If deta.Pieza.EsConjunto Then
                Set detallesHijos = DAODetalleOrdenTrabajo.FindAllConjunto(deta.id, deta.Pieza.id)
                For Each detaHijo In detallesHijos
                    ProcesarDetalleOT sectoresTiempo, detaHijo, deta.id, deta.CantidadPedida, deta.Cantidad_Fabricada
                Next detaHijo
            End If
        Next deta
    Next Ot


    Set GetDTOSectoresTiempo = sectoresTiempo
    'Exit Function
    'e:
    'Stop

End Function

Private Sub ProcesarDetalleOT(sectoresTiempo As Collection, deta As DetalleOTConjuntoDTO, detaPadreId As Long, cantPadre As Double, cantidadFabricada As Double)
    Dim proceso As PlaneamientoTiempoProceso
    Dim procesos As Collection


    'Set procesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(deta.Id, deta.pieza.Id)
    Set procesos = DAOTiemposProceso.FindAllByDetallePedidoId(detaPadreId, deta.id)

    For Each proceso In procesos
        AddTarea sectoresTiempo, proceso.Tarea, deta.Cantidad * cantPadre, proceso.OperariosCotizado, proceso.TiempoCotizado, cantidadFabricada, proceso

    Next proceso

    If deta.Pieza.EsConjunto Then
        Dim Detalles As Collection
        Dim tmpDeta As DetalleOTConjuntoDTO
        Set Detalles = DAODetalleOrdenTrabajo.FindAllConjunto(detaPadreId, deta.Pieza.id)

        For Each tmpDeta In Detalles
            ProcesarDetalleOT sectoresTiempo, tmpDeta, detaPadreId, cantPadre, cantidadFabricada
        Next tmpDeta
    End If


End Sub

Private Sub AddTarea(sectoresTiempo As Collection, Tarea As clsTarea, CantidadPedida As Double, OperariosCotizado As Long, TiempoCotizado As Double, Optional cantidadFabricada As Double, Optional proceso As PlaneamientoTiempoProceso)
    Dim tiempoSectorDTO As DTOSectoresTiempo
    Dim tareaTiempoDTO As DTOTareaTiempo
    Dim Tiempo As Double
    Dim TiempoPendiente As Double
    Dim cantRestante As Double

    If Tarea.CantPorProc = 1 Then
        Tiempo = (OperariosCotizado * TiempoCotizado * CantidadPedida) / 60

        cantRestante = CantidadPedida - cantidadFabricada
        If cantRestante < 0 Then cantRestante = 0
        TiempoPendiente = (OperariosCotizado * TiempoCotizado * cantRestante) / 60
    Else
        Tiempo = (OperariosCotizado * TiempoCotizado) / 60
        If cantidadFabricada > 0 Then
            TiempoPendiente = 0
        Else
            TiempoPendiente = (OperariosCotizado * TiempoCotizado) / 60
        End If
    End If

    If BuscarEnColeccion(sectoresTiempo, CStr(Tarea.SectorID)) Then
        Set tiempoSectorDTO = sectoresTiempo.item(CStr(Tarea.SectorID))
    Else
        Set tiempoSectorDTO = New DTOSectoresTiempo
        Set tiempoSectorDTO.Sector = DAOSectores.GetById(Tarea.SectorID)

        sectoresTiempo.Add tiempoSectorDTO, CStr(Tarea.Sector.id)

    End If


    If BuscarEnColeccion(tiempoSectorDTO.ListaDtoTareaTiempo, CStr(Tarea.id)) Then
        Set tareaTiempoDTO = tiempoSectorDTO.ListaDtoTareaTiempo.item(CStr(Tarea.id))

    Else
        Set tareaTiempoDTO = New DTOTareaTiempo
        Set tareaTiempoDTO.Tarea = Tarea





        tiempoSectorDTO.ListaDtoTareaTiempo.Add tareaTiempoDTO, CStr(Tarea.id)
        tiempoSectorDTO.ListaDtoTareaTiempoPendiente.Add tareaTiempoDTO, CStr(Tarea.id)

    End If


    tareaTiempoDTO.CantidadTareas = tareaTiempoDTO.CantidadTareas + 1

    If proceso.FINALIZADO Then
        tareaTiempoDTO.CantidadTareasFinalizadas = tareaTiempoDTO.CantidadTareasFinalizadas + 1
    End If

    tareaTiempoDTO.Tiempo = tareaTiempoDTO.Tiempo + Tiempo
    tareaTiempoDTO.TiempoPendiente = tareaTiempoDTO.TiempoPendiente + TiempoPendiente


End Sub

Public Function otsMarco(idCliente As Long) As Collection
    Set otsMarco = DAOOrdenTrabajo.FindAll("p.id_ot_padre = -1 and p.estado = " & EstadoOrdenTrabajo.EstadoOT_EnProceso & " and p.idCliente=" & idCliente)
End Function


Public Sub ExcelListadoTareas(ruta As String, Ot As OrdenTrabajo)
    On Error GoTo E

    Dim deta As DetalleOrdenTrabajo
    Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id)

    Dim xlWorkbook As New Excel.Workbook
    Dim xlWorksheet As New Excel.Worksheet
    Dim xlApplication As New Excel.Application


    Dim tareas As New Dictionary


    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 1).value = "Orden de Trabajo Nº " & Ot.IdFormateado
    xlWorksheet.Cells(2, 2).value = "Tareas"
    xlWorksheet.Cells(3, 1).value = "Piezas"

    Dim initialPos As Long: initialPos = 4

    For Each deta In Ot.Detalles
        ProcesarDetalle tareas, initialPos, 0, xlWorksheet, deta
    Next deta


    'tareas
    xlWorksheet.Range(xlWorksheet.Cells(2, 2), xlWorksheet.Cells(2, tareas.count + 1)).Merge
    xlWorksheet.Range(xlWorksheet.Cells(2, 2), xlWorksheet.Cells(2, tareas.count + 1)).HorizontalAlignment = xlCenter
    EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(2, 2), xlWorksheet.Cells(2, tareas.count + 1))

    'pieza
    EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(3, 1), xlWorksheet.Cells(3, 1))
    xlWorksheet.Range(xlWorksheet.Cells(3, 1), xlWorksheet.Cells(3, 1)).HorizontalAlignment = xlCenter
    xlWorksheet.Range(xlWorksheet.Cells(3, 1), xlWorksheet.Cells(3, 1)).VerticalAlignment = xlCenter


    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True

    xlWorksheet.PageSetup.PrintTitleRows = "$1:$3"    'para que al imprimir queden las columnas fijas
    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)


    xlWorkbook.SaveAs ruta

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit


    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    Exit Sub
E:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub EncuadrarCelda(ByRef rango As Range)
    rango.Borders.item(xlEdgeTop).Weight = xlThin
    rango.Borders.item(xlEdgeBottom).Weight = xlThin
    rango.Borders.item(xlEdgeLeft).Weight = xlThin
    rango.Borders.item(xlEdgeRight).Weight = xlThin

End Sub


Private Sub ProcesarDetalle(tareas As Dictionary, pos As Long, depth As Long, xlWorksheet As Excel.Worksheet, Optional DetalleOt As DetalleOrdenTrabajo = Nothing, Optional DetalleOtConjunto As DetalleOTConjuntoDTO = Nothing)

    Dim TiempoProceso As PlaneamientoTiempoProceso
    Dim tiemposProcesos As New Collection
    Dim detasDto As Collection
    Dim NombrePieza As String


    If IsSomething(DetalleOt) Then
        'Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(DetalleOt.Id, DetalleOt.pieza.Id)
        Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoId(DetalleOt.id)
        NombrePieza = "(Item " & DetalleOt.item & ") - " & DetalleOt.Pieza.nombre
        Set detasDto = DAODetalleOrdenTrabajo.FindAllConjunto(DetalleOt.id, DetalleOt.Pieza.id)
    Else
        'Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(DetalleOtConjunto.Id, DetalleOtConjunto.pieza.Id)
        Set tiemposProcesos = DAOTiemposProceso.FindAllByDetallePedidoId(0, DetalleOtConjunto.id)
        NombrePieza = DetalleOtConjunto.Pieza.nombre
        Set detasDto = DAODetalleOrdenTrabajo.FindAllConjunto(DetalleOtConjunto.idDetallePedido, DetalleOtConjunto.Pieza.id)
    End If

    xlWorksheet.Cells(pos, 1).value = String(depth * 6, " ") + " - " & NombrePieza
    EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(pos, 1), xlWorksheet.Cells(pos, 1))

    For Each TiempoProceso In tiemposProcesos
        If Not tareas.Exists(CStr(TiempoProceso.Tarea.id)) Then
            tareas.Add CStr(TiempoProceso.Tarea.id), tareas.count + 2
            xlWorksheet.Cells(3, tareas.item(CStr(TiempoProceso.Tarea.id))) = TiempoProceso.Tarea.Tarea
            EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(3, tareas.item(CStr(TiempoProceso.Tarea.id))), xlWorksheet.Cells(3, tareas.item(CStr(TiempoProceso.Tarea.id))))
            xlWorksheet.Range(xlWorksheet.Cells(3, tareas.item(CStr(TiempoProceso.Tarea.id))), xlWorksheet.Cells(3, tareas.item(CStr(TiempoProceso.Tarea.id)))).Orientation = 90
        End If
        'xlWorksheet.Cells(pos, tareas.Item(CStr(TiempoProceso.Tarea.id))) = "X"
        xlWorksheet.Range(xlWorksheet.Cells(pos, tareas.item(CStr(TiempoProceso.Tarea.id))), xlWorksheet.Cells(pos, tareas.item(CStr(TiempoProceso.Tarea.id)))).Interior.Color = VBA.Information.RGB(194, 194, 194)
        EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(pos, tareas.item(CStr(TiempoProceso.Tarea.id))), xlWorksheet.Cells(pos, tareas.item(CStr(TiempoProceso.Tarea.id))))

        If TiempoProceso.FechaFin <> 0 Then
            xlWorksheet.Cells(pos, tareas.item(CStr(TiempoProceso.Tarea.id))) = "X"
        End If
    Next TiempoProceso

    pos = pos + 1

    Dim dtoConjunto As DetalleOTConjuntoDTO
    For Each dtoConjunto In detasDto
        ProcesarDetalle tareas, pos, depth + 1, xlWorksheet, , dtoConjunto
    Next

End Sub

Function Cerrar(Ot As OrdenTrabajo, Optional a_stock = False) As Boolean
    On Error GoTo err33
    Dim astock As Integer
    Dim deta As DetalleOrdenTrabajo

    Dim tra As Boolean
    Cerrar = True

    If Ot.estado = EstadoOT_ProcesoCompleto Or Ot.estado = EstadoOT_EnProceso Then

        If Not a_stock Then
            'si no va a stock, se genera la descarga de la OT y la actualizacion de stock
            conectar.BeginTransaction
            tra = True
            'CAMBIO EL ESTADO
            Ot.estado = EstadoOT_Finalizado
            Ot.FechaCerrado = Now
            DAOOrdenTrabajo.Guardar Ot, False
            DAOOrdenTrabajoHistorial.agregar Ot, "Finalización del proceso productivo"

            For Each deta In Ot.Detalles
                If deta.CantidadFabricados > deta.CantidadPedida Then
                    If Not DAOPieza.ModificarStock(deta.Pieza, ModificarStock_AltaOT, deta.CantidadFabricados - deta.CantidadPedida) Then GoTo err33
                End If
            Next deta

            conectar.CommitTransaction
            tra = False

        Else

            If Ot.TodoFabricado Then
                'si va a stock, cambio el estado y modifico stock
                'tengo que chequear que esté todo fabricado para poder cerrar la orden
                'tengo que controlar que haya un saldo de productos fabricados (en relacion a lo entregado)
                'para q pueda salir a stock

                conectar.BeginTransaction
                tra = True

                Ot.estado = EstadoOT_Finalizado
                Ot.FechaCerrado = Now
                DAOOrdenTrabajo.Guardar Ot, False
                DAOOrdenTrabajoHistorial.agregar Ot, "Finalización del proceso productivo"

                'actualizo stock pasando todos los productos fbricados a stock
                'y el nro de remito deberia quedar en -1 porq no slio de fábrica


                For Each deta In Ot.Detalles
                    If (deta.CantidadFabricados - deta.CantidadEntregada) > 0 Then
                        If Not DAOPieza.ModificarStock(deta.Pieza, ModificarStock_AltaOT, deta.CantidadFabricados - deta.CantidadEntregada, , Ot.id) Then GoTo err33


                        If deta.Pieza.EsConjunto Then
                            Dim detaOTDto As DetalleOTConjuntoDTO
                            'si es conjunto tengo que abrirlo y actualizar todas las piezas x la cantidad de todo el conjunto en plano
                            For Each detaOTDto In DAODetalleOrdenTrabajo.FindAllConjunto(deta.id)
                                If Not DAOPieza.ModificarStock(detaOTDto.Pieza, ModificarStock_AltaOT, (deta.CantidadFabricados - deta.CantidadEntregada) * detaOTDto.Cantidad, , Ot.id) Then GoTo err33
                            Next detaOTDto
                        End If

                    End If
                Next deta

                conectar.CommitTransaction
                tra = False


            Else
                MsgBox "Para poder cerrar y enviar a stock debería tener todo fabricado", vbCritical, "Error"
                Cerrar = False
            End If
        End If
    ElseIf Ot.estado = EstadoOT_Finalizado Then
        MsgBox "El pedido se encuentra cerrado.", vbCritical, "Error"
        Cerrar = False
    End If

    Exit Function
err33:
    If tra Then conectar.RollBackTransaction
    Cerrar = False
End Function


Public Function CopiarOT(orig As OrdenTrabajo) As Boolean
    On Error GoTo err1
    CopiarOT = True

    Set orig.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(orig.id)
    Dim nue As New OrdenTrabajo
    Dim deta As DetalleOrdenTrabajo
    Dim ndeta As DetalleOrdenTrabajo
    nue.Activa = orig.Activa
    nue.Anticipo = orig.Anticipo
    nue.AnticipoFacturado = False
    nue.CantDiasAnticipo = orig.CantDiasAnticipo
    nue.CantDiasSaldo = orig.CantDiasSaldo
    Set nue.cliente = orig.cliente
    Set nue.ClienteFacturar = orig.ClienteFacturar
    nue.TipoOrden = orig.TipoOrden

    nue.descripcion = orig.descripcion
    nue.Descuento = orig.Descuento
    nue.Entregada = False
    nue.estado = EstadoOT_Pendiente

    nue.fechaCreado = Now
    nue.FormaDePagoAnticipo = orig.FormaDePagoAnticipo
    nue.FormaDePagoSaldo = orig.FormaDePagoSaldo
    nue.FechaEntrega = Now
    Set nue.moneda = orig.moneda


    For Each deta In orig.Detalles
        Set ndeta = New DetalleOrdenTrabajo
        ndeta.CantidadPedida = deta.CantidadPedida
        ndeta.Descuento = deta.Descuento
        ndeta.EstadoProceso = EstProcDetOT_AunNoDefinido
        ndeta.FechaEntrega = deta.FechaEntrega
        ndeta.item = deta.item
        ndeta.Nota = deta.Nota
        ndeta.NotaProduccion = deta.NotaProduccion
        Set ndeta.Pieza = deta.Pieza
        ndeta.Precio = deta.Precio
        nue.Detalles.Add ndeta
    Next

    If Not DAOOrdenTrabajo.Save(nue) Then GoTo err1


    Exit Function
err1:
    CopiarOT = False
End Function
