Attribute VB_Name = "DAOPieza"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "detalle"
Public Const CAMPO_CANTIDAD_STOCK As String = "cantidad"
Public Const CAMPO_ACTIVA As String = "estado"
Public Const CAMPO_ES_CONJUNTO As String = "conjunto"
Public Const CAMPO_UBICACION_STOCK As String = "detalle_stock"
Public Const CAMPO_PRECIO As String = "precio_definido"
Public Const CAMPO_FECHA_PRECIO As String = "fecha_precio_definido"
Public Const CAMPO_CANTIDAD_DE_PIEZAS_EN_CONJUNTO As String = "cantidad"
Public Const CAMPO_ID_MONEDA As String = "id_moneda_precio"
Public Const CAMPO_ID_CLIENTE As String = "id_cliente"
Public Const CAMPO_ID_PIEZA_ULTIMA_REVISION As String = "id_pieza_ultima_revision"
Public Const CAMPO_REVISION As String = "revision"

Public Const TABLA_PIEZA As String = "s"
Public Const TABLA_MONEDA1 As String = "m1"
Public Const TABLA_CLIENTE1 As String = "c1"
Public Enum TipoPieza
    TP_Conjunto = 0
    TP_Unidad = -1
    TP_Ambas = 999
End Enum

Public Enum FetchLevel
    FL_0 = 0
    FL_1 = 1
    FL_2 = 2
    FL_3 = 3
    FL_4 = 4
End Enum

Public Enum ModificarStockOperaciones
    ModificarStock_Ingreso = 0
    ModificarStock_AltaOT = 1
    ModificarStock_BajaOT = 2
    ModificarStock_BajaOE = 3
    ModificarStock_Baja = 4
End Enum

Public Enum ModificarStockMovimientos
    ModificarStockMovimientos_exitoso = 0
    ModificarStockMovimientos_error = -1
End Enum

Public Function FindById(ByRef Id As Long, _
                         fetch As FetchLevel, _
                         Optional includeDesarrolloManoObra As Boolean = False, _
                         Optional includeDesarrolloMaterial As Boolean = False, _
                         Optional withTiempoHistorico As Boolean = False _
                       ) As Pieza
    Dim col As Collection
    Set col = DAOPieza.FindAll(fetch, TABLA_PIEZA & ".id = " & Id, , includeDesarrolloManoObra, includeDesarrolloMaterial, withTiempoHistorico)
    If col.count > 0 Then
        Set FindById = col(1)
    Else
        Set FindById = Nothing
    End If
End Function

Public Function FindAll(fetch As FetchLevel, _
                        Optional filter As String = vbNullString, _
                        Optional Tipo As TipoPieza = TipoPieza.TP_Ambas, _
                        Optional includeDesarrolloManoObra As Boolean = False, _
                        Optional includeDesarrolloMaterial As Boolean = False, _
                        Optional withTiempoHistorico As Boolean = False, _
                        Optional onlyActives As Boolean = False, _
                        Optional orderByItemMarco As String = vbNullString _
                      ) As Collection
    Dim rs As ADODB.Recordset
    Dim q As String
    On Error GoTo err1
    Dim piezas As New Collection

    Dim tickStart As Double
    Dim tickend As Double
    'tickStart = GetTickCount

    q = "SELECT s.*, m1.*, c1.*"
    If fetch >= FL_1 Then q = q & ", sc.*, m2.*, c2.*, s2.*"
    If fetch >= FL_2 Then q = q & ", sc2.*, m3.*, c3.*, s3.*"
    If fetch >= FL_3 Then q = q & ", sc3.*, m4.*, c4.*, s4.*"
    If fetch >= FL_4 Then q = q & ", sc4.*, m5.*, c5.*, s5.*"

    q = q & " FROM stock s" _
      & " LEFT JOIN AdminConfigMonedas m1 ON m1.id = s.id_moneda_precio" _
      & " LEFT JOIN clientes c1 ON c1.id = s.id_cliente"
    If fetch >= FL_1 Then
        q = q & " LEFT JOIN stockConjuntos sc ON sc.idPiezaPadre = s.id" _
          & " LEFT JOIN stock s2 ON s2.id = sc.idPiezaHija" _
          & " LEFT JOIN AdminConfigMonedas m2 ON m2.id = s2.id_moneda_precio" _
          & " LEFT JOIN clientes c2 ON c2.id = s2.id_cliente"
    End If
    If fetch >= FL_2 Then
        q = q & " LEFT JOIN stockConjuntos sc2 ON sc2.idPiezaPadre = s2.id" _
          & " LEFT JOIN stock s3 ON s3.id = sc2.idPiezaHija" _
          & " LEFT JOIN AdminConfigMonedas m3 ON m3.id = s3.id_moneda_precio" _
          & " LEFT JOIN clientes c3 ON c3.id = s3.id_cliente"
    End If
    If fetch >= FL_3 Then
        q = q & " LEFT JOIN stockConjuntos sc3 ON sc3.idPiezaPadre = s3.id" _
          & " LEFT JOIN stock s4 ON s4.id = sc3.idPiezaHija" _
          & " LEFT JOIN AdminConfigMonedas m4 ON m4.id = s4.id_moneda_precio" _
          & " LEFT JOIN clientes c4 ON c4.id = s4.id_cliente"
    End If
    If fetch >= FL_4 Then
        q = q & " LEFT JOIN stockConjuntos sc4 ON sc4.idPiezaPadre = s4.id" _
          & " LEFT JOIN stock s5 ON s5.id = sc4.idPiezaHija" _
          & " LEFT JOIN AdminConfigMonedas m5 ON m5.id = s5.id_moneda_precio" _
          & " LEFT JOIN clientes c5 ON c5.id = s5.id_cliente"
    End If

    q = q & " WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    If Tipo <> TP_Ambas Then
        q = q & " AND s.conjunto = " & Tipo
    End If

    If onlyActives Then q = q & " AND s." & DAOPieza.CAMPO_ACTIVA & " = 1"

    If orderByItemMarco <> vbNullString Then
        q = q & " ORDER BY " & orderByItemMarco
    Else
        q = q & " ORDER BY s.id ASC"
    End If


    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim piezaIdIndex As Long: piezaIdIndex = fieldsIndex("s.id")
    Dim piezaId2Index As Long
    Dim piezaId3Index As Long
    Dim piezaId4Index As Long
    Dim piezaId5Index As Long

    Dim tablaCliente2 As String
    Dim tablaMoneda2 As String
    Dim tablaStock2 As String
    Dim tablaStockConjuntos As String
    If fetch >= FL_1 Then
        tablaCliente2 = "c2"
        tablaMoneda2 = "m2"
        tablaStock2 = "s2"
        tablaStockConjuntos = "sc"
        piezaId2Index = fieldsIndex("s2.id")
    End If

    Dim tablaCliente3 As String
    Dim tablaMoneda3 As String
    Dim tablaStock3 As String
    Dim tablaStockConjuntos2 As String
    If fetch >= FL_2 Then
        tablaCliente3 = "c3"
        tablaMoneda3 = "m3"
        tablaStock3 = "s3"
        tablaStockConjuntos2 = "sc2"
        piezaId3Index = fieldsIndex("s3.id")
    End If

    Dim tablaCliente4 As String
    Dim tablaMoneda4 As String
    Dim tablaStock4 As String
    Dim tablaStockConjuntos3 As String
    If fetch >= FL_3 Then
        tablaCliente4 = "c4"
        tablaMoneda4 = "m4"
        tablaStock4 = "s4"
        tablaStockConjuntos3 = "sc3"
        piezaId4Index = fieldsIndex("s4.id")
    End If

    Dim tablaCliente5 As String
    Dim tablaMoneda5 As String
    Dim tablaStock5 As String
    Dim tablaStockConjuntos4 As String
    If fetch >= FL_4 Then
        tablaCliente5 = "c5"
        tablaMoneda5 = "m5"
        tablaStock5 = "s5"
        tablaStockConjuntos4 = "sc4"
        piezaId5Index = fieldsIndex("s5.id")
    End If

    Dim tmpPieza As Pieza
    Dim tmpPieza2 As Pieza
    Dim lastPiezaId As Long: lastPiezaId = 0

    While Not rs.EOF

        If BuscarEnColeccion(piezas, CStr(rs.Fields(piezaIdIndex).value)) Then
            Set tmpPieza = piezas.item(CStr(rs.Fields(piezaIdIndex).value))
        Else
            Set tmpPieza = DAOPieza.Map(rs, fieldsIndex, TABLA_PIEZA, DAOPieza.TABLA_MONEDA1, DAOPieza.TABLA_CLIENTE1)
            If includeDesarrolloManoObra Then
                Set tmpPieza.desarrollosManoObra = New Collection
                Set tmpPieza.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(tmpPieza.Id, withTiempoHistorico)
            End If
            If includeDesarrolloMaterial Then
                Set tmpPieza.DesarrollosMaterial = New Collection
                Set tmpPieza.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(tmpPieza.Id)
            End If
        End If

        If fetch >= FL_1 And Not IsNull(rs.Fields(piezaId2Index).value) Then    ' no importa que no diga que no es conjunto, por las dudas
            If Not BuscarEnColeccion(tmpPieza.PiezasHijas, CStr(rs.Fields(piezaId2Index).value)) Then
                Set tmpPieza2 = DAOPieza.Map(rs, fieldsIndex, tablaStock2, tablaMoneda2, tablaCliente2, tablaStockConjuntos)
                If includeDesarrolloManoObra Then Set tmpPieza2.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(tmpPieza2.Id, withTiempoHistorico)
                If includeDesarrolloMaterial Then Set tmpPieza2.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(tmpPieza2.Id)

                tmpPieza.PiezasHijas.Add tmpPieza2, CStr(tmpPieza2.Id)
            End If
        End If

        If fetch >= FL_2 And Not IsNull(rs.Fields(piezaId3Index).value) Then
            If Not BuscarEnColeccion(tmpPieza.PiezasHijas.item(CStr(rs.Fields(piezaId2Index).value)).PiezasHijas, CStr(rs.Fields(piezaId3Index).value)) Then
                Set tmpPieza2 = DAOPieza.Map(rs, fieldsIndex, tablaStock3, tablaMoneda3, tablaCliente3, tablaStockConjuntos2)
                If includeDesarrolloManoObra Then Set tmpPieza2.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(tmpPieza2.Id, withTiempoHistorico)
                If includeDesarrolloMaterial Then Set tmpPieza2.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(tmpPieza2.Id)

                tmpPieza.PiezasHijas.item(CStr(rs.Fields(piezaId2Index).value)).PiezasHijas.Add tmpPieza2, CStr(tmpPieza2.Id)    ' CStr(rs.Fields(piezaId3Index).value)
            End If
        End If

        If fetch >= FL_3 And Not IsNull(rs.Fields(piezaId4Index).value) Then
            If Not BuscarEnColeccion(tmpPieza.PiezasHijas.item(CStr(rs.Fields(piezaId2Index).value)).PiezasHijas.item(CStr(rs.Fields(piezaId3Index).value)).PiezasHijas, CStr(rs.Fields(piezaId4Index).value)) Then
                Set tmpPieza2 = DAOPieza.Map(rs, fieldsIndex, tablaStock4, tablaMoneda4, tablaCliente4, tablaStockConjuntos3)
                If includeDesarrolloManoObra Then Set tmpPieza2.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(tmpPieza2.Id, withTiempoHistorico)
                If includeDesarrolloMaterial Then Set tmpPieza2.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(tmpPieza2.Id)

                tmpPieza.PiezasHijas.item(CStr(rs.Fields(piezaId2Index).value)).PiezasHijas.item(CStr(rs.Fields(piezaId3Index).value)).PiezasHijas.Add tmpPieza2, CStr(tmpPieza2.Id)  ' CStr(rs.Fields(piezaId4Index).value)
            End If
        End If

        If fetch >= FL_4 And Not IsNull(rs.Fields(piezaId5Index).value) Then
            If Not BuscarEnColeccion(tmpPieza.PiezasHijas.item(CStr(rs.Fields(piezaId2Index).value)).PiezasHijas.item(CStr(rs.Fields(piezaId3Index).value)).PiezasHijas.item(CStr(rs.Fields(piezaId4Index).value)).PiezasHijas, CStr(rs.Fields(piezaId5Index).value)) Then
                Set tmpPieza2 = DAOPieza.Map(rs, fieldsIndex, tablaStock5, tablaMoneda5, tablaCliente5, tablaStockConjuntos4)
                If includeDesarrolloManoObra Then Set tmpPieza2.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(tmpPieza2.Id, withTiempoHistorico)
                If includeDesarrolloMaterial Then Set tmpPieza2.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(tmpPieza2.Id)

                tmpPieza.PiezasHijas.item(CStr(rs.Fields(piezaId2Index).value)).PiezasHijas.item(CStr(rs.Fields(piezaId3Index).value)).PiezasHijas.item(CStr(rs.Fields(piezaId4Index).value)).PiezasHijas.Add tmpPieza2, CStr(tmpPieza2.Id)  ' CStr(rs.Fields(piezaId4Index).value)
            End If
        End If

        If Not BuscarEnColeccion(piezas, CStr(tmpPieza.Id)) Then
            piezas.Add tmpPieza, CStr(tmpPieza.Id)
        End If

        rs.MoveNext

    Wend


    'tickEnd = GetTickCount

    'Debug.Print tickEnd - tickStart, "ms elapsed"

    Set FindAll = piezas
    Exit Function
err1:
    Set FindAll = New Collection
    'Debug.Print Err.Description
    MsgBox Err.Description, vbCritical, "DAOPieza.FindAll()"
End Function

Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, Optional ByRef monedaTableNameOrAlias As String = vbNullString, Optional ByVal clienteTableNameOrAlias As String = vbNullString, Optional ByRef stockConjuntoTableNameOrAlias As String = vbNullString) As Pieza
    Dim P As Pieza
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If Id > 0 Then
        Set P = New Pieza
        P.Id = Id

        P.Activa = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_ACTIVA)
        If LenB(stockConjuntoTableNameOrAlias) > 0 Then P.Cantidad = GetValue(rs, fieldsIndex, stockConjuntoTableNameOrAlias, DAOPieza.CAMPO_CANTIDAD_DE_PIEZAS_EN_CONJUNTO)
        P.CantidadStock = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_CANTIDAD_STOCK)
        P.EsConjunto = CBool(GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_ES_CONJUNTO) + 1)
        P.FechaPrecio = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_FECHA_PRECIO)
        P.YaFabricada = GetValue(rs, fieldsIndex, tableNameOrAlias, "ya_fabricado")
        P.nombre = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_NOMBRE)
        P.Precio = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_PRECIO)
        P.UbicacionStock = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_UBICACION_STOCK)
        P.Revision = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_REVISION)
        P.IdPiezaUltimaRevision = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOPieza.CAMPO_ID_PIEZA_ULTIMA_REVISION)
        P.Complejidad = GetValue(rs, fieldsIndex, tableNameOrAlias, "tipo_complejidad")

        If LenB(monedaTableNameOrAlias) > 0 Then Set P.MonedaPrecio = DAOMoneda.Map(rs, fieldsIndex, monedaTableNameOrAlias)
        If LenB(clienteTableNameOrAlias) > 0 Then Set P.cliente = DAOCliente.Map(rs, fieldsIndex, clienteTableNameOrAlias)
    End If

    Set Map = P
End Function

Public Function ModificarStock(T As Pieza, operacion As ModificarStockOperaciones, Cantidad As Double, Optional ubicacion As String = Empty, Optional OTOEid As Long) As Boolean
    On Error GoTo err5
    Dim charO As String
    Dim Nota As Long
    'nota
    '0- ingreso existoso
    '-1 defectuoso/egreso sin causa
    '>0 nro. orden de trabajo

    If operacion = 1 Then charO = "+" Else charO = "-"
    conectar.execute "select * from stock where id =" & T.Id
    If IsEmpty(ubicacion) Then
        conectar.execute "update stock set cantidad = cantidad" & charO & Cantidad & " where id=" & T.Id
    Else
        conectar.execute "update stock set detalle_stock='" & UCase(ubicacion) & "',cantidad = cantidad" & charO & Cantidad & " where id=" & T.Id
    End If
    If operacion = ModificarStock_Ingreso Then
        Nota = 0
    ElseIf operacion = ModificarStock_Baja Then
        Nota = -1
    Else
        Nota = OTOEid
    End If
    conectar.execute "insert into movimiento_stock (id_pieza,cantidad,operacion,fecha,nota)values (" & T.Id & "," & Cantidad & "," & operacion & ",'" & Format(Date, "yyyy/mm/dd") & "'," & Nota & ")"
    ModificarStock = True
    Exit Function

err5:
    ModificarStock = False
End Function


Private Sub CargarTiempoDto(ByRef dm As DesarrolloManoObra, ByRef lista As Collection, Cantidad As Double)
    Dim dto1 As DTOTareaTiempo
    Dim tempo As Double
    If dm.Tarea.CantPorProc = 1 Then
        tempo = ((dm.Cantidad * dm.Tiempo * Cantidad)) / 60
    Else
        tempo = ((dm.Cantidad * dm.Tiempo) / 60)
    End If


    If BuscarEnColeccion(lista, CStr(dm.Tarea.Id)) Then
        Set dto1 = lista.item(CStr(dm.Tarea.Id))
        dto1.Tiempo = dto1.Tiempo + funciones.RedondearDecimales(tempo)
    Else
        Set dto1 = New DTOTareaTiempo
        Set dto1.Tarea = dm.Tarea
        dto1.Tiempo = funciones.RedondearDecimales(tempo)
        lista.Add dto1, CStr(dm.Tarea.Id)
    End If
End Sub

Private Sub recorre(col As Collection, Pieza As Pieza, Cantidad As Double)
    Dim dm As DesarrolloManoObra
    Dim dto As DTOSectoresTiempo
    Dim Pie As Pieza

    For Each dm In Pieza.desarrollosManoObra

        If BuscarEnColeccion(col, CStr(dm.Tarea.Sector.Id)) Then
            Set dto = col.item(CStr(dm.Tarea.Sector.Id))
            CargarTiempoDto dm, dto.ListaDtoTareaTiempo, Cantidad
        Else
            Set dto = New DTOSectoresTiempo
            Set dto.Sector = dm.Tarea.Sector
            CargarTiempoDto dm, dto.ListaDtoTareaTiempo, Cantidad
            col.Add dto, CStr(dm.Tarea.Sector.Id)
        End If

    Next dm


    For Each Pie In Pieza.PiezasHijas
        recorre col, Pie, Cantidad * Pie.Cantidad
    Next Pie
End Sub


Public Function ListaDTOTiempoPorSector(listadtopiezacantidad As Collection) As Collection
    Dim col As New Collection
    Dim dto1 As DTOPiezaCantidad

    Dim detallesOT As Collection

    For Each dto1 In listadtopiezacantidad

        If IsSomething(dto1.Pieza) Then
            Set dto1.Pieza = DAOPieza.FindById(dto1.Pieza.Id, FL_4, True, False, False)
        End If

        recorre col, dto1.Pieza, dto1.Cantidad

    Next dto1
    Set ListaDTOTiempoPorSector = col

End Function


Public Function Save(P As Pieza, Optional ByVal paraRevision As Boolean = False) As Boolean
    Dim q As String

    If P.Id <> 0 Then
        q = "Update {tabla} SET" _
          & " detalle = 'detalle' ," _
          & " id_cliente = 'id_cliente' ," _
          & " cantidad = 'cantidad' ," _
          & " estado = 'estado' ," _
          & " conjunto = 'conjunto' ," _
          & " ya_fabricado = 'ya_fabricado' ," _
          & " detalle_stock = 'detalle_stock' ," _
          & " precio_definido = 'precio_definido' ," _
          & " fecha_precio_definido = 'fecha_precio_definido' ," _
          & " id_moneda_precio = 'id_moneda_precio', " _
          & " id_pieza_ultima_revision = 'id_pieza_ultima_revision' ," _
          & " revision = 'revision'" _
          & " tipo_complejidad = 'tipo_complejidad'" _
          & " Where id = 'id'"
    Else
        q = "INSERT INTO {tabla}" _
          & " (tipo_complejidad,detalle, id_cliente, cantidad, estado, conjunto, detalle_stock, precio_definido, fecha_precio_definido," _
          & " id_moneda_precio, ya_fabricado, id_pieza_ultima_revision,revision)" _
          & " Values " _
          & " ('tipo_complejidad','detalle', 'id_cliente', 'cantidad', 'estado', 'conjunto', 'detalle_stock', 'precio_definido', 'fecha_precio_definido'," _
          & "'id_moneda_precio','ya_fabricado','id_pieza_ultima_revision','revision')"
    End If


    If paraRevision Then
        q = Replace$(q, "{tabla}", "stock_rev")
    Else
        q = Replace$(q, "{tabla}", "stock")
    End If

    q = Replace$(q, "'detalle'", conectar.Escape(P.nombre))
    q = Replace$(q, "'id_cliente'", conectar.GetEntityId(P.cliente))
    q = Replace$(q, "'cantidad'", conectar.Escape(P.Cantidad))
    q = Replace$(q, "'estado'", conectar.Escape(P.Activa))
    q = Replace$(q, "'conjunto'", IIf(P.EsConjunto, "0", "-1"))
    q = Replace$(q, "'ya_fabricado'", conectar.Escape(P.YaFabricada))
    q = Replace$(q, "'detalle_stock'", conectar.Escape(P.CantidadStock))
    q = Replace$(q, "'precio_definido'", conectar.Escape(P.Precio))
    q = Replace$(q, "'fecha_precio_definido'", conectar.Escape(P.FechaPrecio))
    q = Replace$(q, "'id_moneda_precio'", conectar.Escape(P.MonedaPrecio.Id))
    q = Replace$(q, "'id_pieza_ultima_revision'", conectar.Escape(P.IdPiezaUltimaRevision))
    q = Replace$(q, "'revision'", conectar.Escape(P.Revision))
    q = Replace$(q, "'tipo_complejidad'", conectar.Escape(P.Complejidad))

    q = Replace$(q, "'id'", conectar.Escape(P.Id))

    Save = conectar.execute(q)

    If Save And P.Id = 0 Then
        P.Id = conectar.UltimoId2()
        Save = (P.Id <> 0)
    End If

End Function
Public Function informePiezaArbol(idConjunto, pos)    'idConjunto=detalle pedido


    On Error GoTo err2
    informePiezaArbol = True
    Dim IDCONJ
    Dim item
    Dim FECHAENT
    Dim ENTREGAITEM
    Dim idDetallePedido
    Dim Nota
    Dim Ot
    Dim detail1
    Dim c2
    Dim cantidad2
    Dim idDetallesPedidosConjuntos2
    Dim detalle2
    Dim stock2
    Dim cantidad3
    Dim idDetallesPedidosConjuntos3
    Dim detalle3
    Dim stock3
    Dim c3
    Dim ref
    Dim CantConj
    Dim razon
    Dim idDetallesPedidosConjuntos
    Dim detallePieza
    Dim Items
    Dim barcode
    Dim c1
    Dim detalle
    Dim stock1
    Dim detail
    Dim esconj
    Dim Cantidad
    Dim COPIAS
    Dim stock As New classStock
    Dim STOCKCONJ
    Dim strsql As String, strsql2 As String
    Dim r_tmp As Recordset
    Dim rs1 As Recordset, rs2 As Recordset, rs3 As Recordset
    Dim Id As Long, id2 As Long, id3 As Long
    Dim cant1 As Double, cant2 As Double, cant3 As Double
    Dim detail2
    Dim detail3

    Set r_tmp = New Recordset
    With r_tmp
        .Fields.Append "idPieza", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "cantu", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "cantt", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "rama", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "stock", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "idDetallesPedidosConjuntos", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable

    End With
    r_tmp.Open
    Set rs1 = conectar.RSFactory("select dp.impresiones_ruta as copias,dp.id,dp.nota,p.fechaEntrega,p.descripcion,p.id as ot,c.razon,dp.fechaEntrega as entregaItem,dp.reserva_stock,dp.idpieza,dp.item,dp.cantidad,s.cantidad as stock,s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id inner join clientes c on s.id_cliente=c.id inner join pedidos p on p.id=dp.idPedido where dp.id=" & idConjunto)
    r_tmp.AddNew
    STOCKCONJ = rs1!stock
    IDCONJ = rs1!idPieza
    r_tmp!idPieza = IDCONJ
    item = rs1!item

    COPIAS = rs1!COPIAS

    FECHAENT = rs1!FechaEntrega
    ENTREGAITEM = rs1!ENTREGAITEM

    idDetallePedido = rs1!Id

    Nota = rs1!Nota

    Ot = rs1!Ot
    ref = rs1!descripcion
    CantConj = rs1!Cantidad

    Cantidad = rs1!Cantidad
    Cantidad = rs1!Cantidad - rs1!reserva_stock


    razon = rs1!razon
    idDetallesPedidosConjuntos = 0    'es madre, no tiene id acá. uso esta id solo para imprimir el barcode de ruta
    r_tmp!Cantt = Cantidad
    r_tmp!cantU = 1
    r_tmp!rama = 0
    detallePieza = rs1!detalle
    r_tmp!detalle = detallePieza
    r_tmp!stock = STOCKCONJ
    r_tmp.Update


    Set rs1 = conectar.RSFactory("select count(id) as cantITEMS from detalles_pedidos where idPedido=" & Ot)
    Items = Ot & "." & pos & "/" & rs1!cantITems

    barcode = Ot & "." & idConjunto
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lblItem").caption = "Item " & item
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lblCantidad").caption = CantConj
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lblOT").caption = Items
    pedido_pieza_ARBOL.Sections("cabeza").Controls("barcode").caption = barcode

    pedido_pieza_ARBOL.Sections("observar").Controls("copia").caption = COPIAS
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lblCliente").caption = razon
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lblReferencia").caption = ref
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lblfechaentrega").caption = FECHAENT
    pedido_pieza_ARBOL.Sections("cabeza").Controls("entregait").caption = ENTREGAITEM
    pedido_pieza_ARBOL.Sections("cabeza").Controls("lbldetalle").caption = detallePieza

    'strsql = "select sc.idPiezaHija,sc.cantidad,s.cantidad as stock,s.detalle from stockConjuntos sc inner join stock s on sc.idPiezaHija=s.id  where idPiezaPadre=" & idConj

    strsql = "select sc.id as id1,sc.esconjunto,sc.idPieza,sc.cantidadPieza as cantidad,s.cantidad as stock,s.detalle from detalles_pedidos_conjuntos sc inner join stock s on sc.idPieza=s.id  where idPiezaPadre=" & IDCONJ & " and idDetalle_pedido=" & idDetallePedido


    Set rs1 = conectar.RSFactory(strsql)
    c1 = 0
    While Not rs1.EOF    'itero sobre la primer rama del arbol
        c1 = c1 + 1
        Id = rs1!idPieza
        idDetallesPedidosConjuntos = rs1!id1
        Cantidad = rs1!Cantidad

        detalle = rs1!detalle
        stock1 = rs1!stock
        esconj = rs1!EsConjunto
        If esconj = 1 Then
            'si es conjunto, analizo la rama
            'strsql = "select * from stockConjuntos where idPiezaPadre=" & id

            'strsql = "select s.cantidad as stock,sc.idPiezaHija,sc.cantidad,s.detalle from stockConjuntos sc inner join stock s on sc.idPiezaHija=s.id  where idPiezaPadre=" & id
            strsql = "select sc.id as id1,sc.esConjunto,s.cantidad as stock,sc.idPieza,sc.cantidad,s.detalle from detalles_pedidos_conjuntos sc inner join stock s on sc.idPieza=s.id  where idPiezaPadre=" & Id & " and idDetalle_pedido=" & idDetallePedido

            r_tmp.AddNew
            r_tmp!idPieza = Id
            r_tmp!idDetallesPedidosConjuntos = idDetallesPedidosConjuntos
            r_tmp!Cantt = Cantidad * CantConj
            r_tmp!cantU = Cantidad
            r_tmp!stock = stock1
            r_tmp!rama = 1

            detail1 = c1 & " - " & detalle
            r_tmp!detalle = "    " & detail1
            r_tmp.Update
            Set rs2 = conectar.RSFactory(strsql)
            c2 = 0
            While Not rs2.EOF    'itero sobre la segunda rama del arbol
                c2 = c2 + 1
                id2 = rs2!idPieza
                cantidad2 = rs2!Cantidad
                esconj = rs2!EsConjunto
                idDetallesPedidosConjuntos2 = rs2!id1
                detalle2 = rs2!detalle
                stock2 = rs2!stock
                If esconj = 1 Then
                    r_tmp.AddNew
                    r_tmp!idPieza = id2
                    r_tmp!Cantt = cantidad2 * CantConj * Cantidad
                    r_tmp!cantU = cantidad2
                    r_tmp!idDetallesPedidosConjuntos = idDetallesPedidosConjuntos2
                    r_tmp!stock = stock2
                    r_tmp!rama = 1
                    detail2 = c1 & "." & c2 & " - " & detalle2
                    r_tmp!detalle = "        " & detail2
                    r_tmp.Update
                    'si es conjunto, analizo la rama
                    'strsql = "select * from stockConjuntos where idPiezaPadre=" & id2
                    strsql = "select sc.id as id1,sc.esConjunto,sc.idPieza,sc.cantidad,s.cantidad as stock,s.detalle from detalles_pedidos_conjuntos sc inner join stock s on sc.idPieza=s.id  where idPiezaPadre=" & id2 & " and idDetalle_pedido=" & idDetallePedido
                    Set rs3 = conectar.RSFactory(strsql)
                    c3 = 0
                    While Not rs3.EOF
                        c3 = c3 + 1
                        'en esta rama que es la ultima, todas las piezas deberian no ser conjunto
                        'asique directamente sumamos los costos
                        id3 = rs3!idPieza
                        cantidad3 = rs3!Cantidad
                        detalle3 = rs3!detalle
                        idDetallesPedidosConjuntos3 = rs3!id1
                        stock3 = rs3!stock
                        r_tmp.AddNew
                        r_tmp!idDetallesPedidosConjuntos = idDetallesPedidosConjuntos3
                        r_tmp!idPieza = id3
                        r_tmp!Cantt = cantidad3 * CantConj * cantidad2 * Cantidad
                        r_tmp!cantU = cantidad3
                        r_tmp!stock = stock3
                        r_tmp!rama = 0
                        detail3 = c1 & "." & c2 & "." & c3 & " - " & detalle3
                        r_tmp!detalle = "            " & detail3
                        r_tmp.Update
                        rs3.MoveNext
                    Wend
                Else
                    'si no es conjunto, acumulo el costo
                    r_tmp.AddNew
                    r_tmp!idPieza = id2
                    r_tmp!Cantt = cantidad2 * CantConj * Cantidad
                    r_tmp!idDetallesPedidosConjuntos = idDetallesPedidosConjuntos2
                    r_tmp!cantU = cantidad2
                    r_tmp!stock = stock2
                    r_tmp!rama = 0
                    detail2 = c1 & "." & c2 & " - " & detalle2
                    r_tmp!detalle = "        " & detail2
                    r_tmp.Update
                End If
                rs2.MoveNext
            Wend
        Else
            r_tmp.AddNew
            r_tmp!idPieza = Id
            r_tmp!Cantt = Cantidad * CantConj
            r_tmp!cantU = Cantidad
            r_tmp!idDetallesPedidosConjuntos = idDetallesPedidosConjuntos
            r_tmp!stock = stock1
            r_tmp!rama = 0
            detail = c1 & " - " & detalle
            r_tmp!detalle = "    " & detail
            r_tmp.Update
            'si no es conjunto, acumulo el costo
        End If
        rs1.MoveNext
    Wend
    r_tmp.MoveFirst
    Set pedido_pieza_ARBOL.DataSource = r_tmp
    pedido_pieza_ARBOL.PrintReport False
    'pedido_pieza_ARBOL.Show 1
    r_tmp.MoveFirst
    Dim canti As Long
    canti = r_tmp.RecordCount

    'releo el recordset del conjunto con el fin de imprimir las rutas

    Dim o As Long
    Dim cantTotal
    Dim A
    Dim idPieza As Long
    Dim idDetalle As Long
    For o = 1 To canti
        cantTotal = r_tmp!Cantt
        detalle = Trim(r_tmp!detalle)
        idPieza = r_tmp!idPieza
        idDetalle = r_tmp!idDetallesPedidosConjuntos
        A = IDCONJ


        '''''''''AHORA ESTO LOHAGO AFUERA!!!
        If CLng(idPieza) = IDCONJ Then
            idDetalle = 0    'ya que no tiene un registro en detallePedidoCOnj
            ''''''''''''Me.informePieza2 idConjunto, False, True, idPieza, pos, canti, o, cantTotal, detalle, idDetalle
        Else
            ''''''''''''Me.informePieza2 IdDetallePedido, False, True, idPieza, pos, canti, o, cantTotal, detalle, idDetalle
        End If

        r_tmp.MoveNext

    Next o


    'rs.Close
    'rs2.Close
    r_tmp.Close
    Exit Function
err2:
    MsgBox Err.Description
    informePiezaArbol = False
End Function


