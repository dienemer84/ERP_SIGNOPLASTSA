Attribute VB_Name = "DAOOrdenCompra"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_FECHA_CREACION As String = "FechaCreacion"
Public Const CAMPO_ESTADO As String = "estado"

Public Const CAMPO_ID_PROVEEEDOR As String = "idProveedor"

Public Const TABLA_ORDEN_COMPRA As String = "oc"

Public Function FindAll(Optional whereFilter As String = vbNullString, Optional includeDetalles As Boolean = False) As Collection
    Dim q As String
    Dim rs As Recordset
    Dim ordenes As New Collection

    q = "SELECT oc.*, p.*" _
        & " FROM ComprasOrdenes oc" _
        & " LEFT JOIN proveedores p ON p.id = oc.idProveedor" _
        & " WHERE 1=1"

    If LenB(whereFilter) > 0 Then q = q & " AND " & whereFilter

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim oc As OrdenCompra

    While Not rs.EOF
        'traer los detalles por afuera con el dao si son necesarios
        Set oc = DAOOrdenCompra.Map(rs, fieldsIndex, DAOOrdenCompra.TABLA_ORDEN_COMPRA, "p")
        ordenes.Add oc, CStr(oc.id)
        rs.MoveNext
    Wend

    Set FindAll = ordenes
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional ByVal tablaProveedor As String = vbNullString) As OrdenCompra
    Dim oc As OrdenCompra
    Dim id As Long
    id = GetValue(rs, indice, tabla, DAOOrdenCompra.CAMPO_ID)

    If id > 0 Then
        Set oc = New OrdenCompra
        oc.id = id

        oc.estado = GetValue(rs, indice, tabla, DAOOrdenCompra.CAMPO_ESTADO)
        oc.FechaCreacion = GetValue(rs, indice, tabla, DAOOrdenCompra.CAMPO_FECHA_CREACION)

        If LenB(tablaProveedor) > 0 Then Set oc.Proveedor = DAOProveedor.Map2(rs, indice, tablaProveedor)
    End If

    Set Map = oc
End Function

Public Function CrearOrdenCompra(po As clsPeticionOferta) As Boolean
    On Error GoTo E

    'recarga de detalles por si las moscas
    If po.detalle Is Nothing Then
        po.detalle = DAOPeticionOfertaDetalle.FindAll(po.numero)
    Else
        If po.detalle.count = 0 Then
            po.detalle = DAOPeticionOfertaDetalle.FindAll(po.numero)
        End If
    End If

    conectar.BeginTransaction

    Dim IdOrdenCompra As Long
    Dim IdDetalleOrdenCompra As Long
    Dim deta As clsPeticionOfertaDetalle
    Dim entrega As EntregaPetOfDetalle

    Dim q As String
    q = "INSERT INTO ComprasOrdenes" _
        & " (idProveedor, FechaCreacion, estado) VALUES (" _
        & po.Proveedor.id & ", " _
        & conectar.Escape(Date) & ", " _
        & 0 & ")"

    If conectar.execute(q) Then
        If conectar.UltimoId("ComprasOrdenes", IdOrdenCompra) Then
            If IdOrdenCompra = 0 Then
                GoTo E
            Else
                For Each deta In po.detalle

                    q = "INSERT INTO ComprasOrdenesDetalles" _
                        & " (id_orden_compra, id_peticion_oferta_detalle, valor, cantidad, descripcion) Values (" _
                        & IdOrdenCompra & ", " _
                        & deta.id & ", " _
                        & conectar.Escape(deta.Valor) & ", " _
                        & conectar.Escape(deta.Cantidad) & ", " _
                        & conectar.Escape(vbNullString) & ")"

                    If conectar.execute(q) Then
                        IdDetalleOrdenCompra = 0
                        If conectar.UltimoId("ComprasOrdenesDetalles", IdDetalleOrdenCompra) Then
                            If IdDetalleOrdenCompra <> 0 Then
                                For Each entrega In deta.Entregas
                                    q = "INSERT INTO ComprasOrdenesDetallesEntregas" _
                                        & "(fecha, cant, id_detalle_orden_compra) VALUES (" _
                                        & conectar.Escape(entrega.FEcha) & ", " _
                                        & entrega.Cantidad & ", " _
                                        & IdDetalleOrdenCompra & ")"

                                    If Not conectar.execute(q) Then GoTo E
                                Next entrega
                            Else
                                GoTo E
                            End If
                        Else
                            GoTo E
                        End If
                    Else
                        GoTo E
                    End If

                Next deta
            End If
        Else
            GoTo E
        End If
    Else
        GoTo E
    End If

    q = "UPDATE ComprasPeticionOferta SET estado = " & EstadoPO.OrdenCompraCreada & " WHERE id = " & po.numero
    If conectar.execute(q) Then
        po.estado = OrdenCompraCreada
        CrearOrdenCompra = True
        conectar.CommitTransaction
    Else
        GoTo E
    End If

    Exit Function
E:
    CrearOrdenCompra = False
    conectar.RollBackTransaction
    Resume Next
End Function

Public Function UpdateEstado(oc As OrdenCompra) As Boolean
    Dim q As String

    'q = "Update ComprasOrdenes" _
     & " SET" _
     & " idProveedor = " & conectar.GetEntityId(oc.Proveedor) & "," _
     & " FechaCreacion = " & conectar.Escape(oc.FechaCreacion) & "," _
     & " estado = " & oc.Estado _
     & " Where id = " & oc.Id

    q = "Update ComprasOrdenes" _
        & " SET" _
        & " estado = " & oc.estado _
        & " Where id = " & oc.id

    UpdateEstado = conectar.execute(q)
End Function
