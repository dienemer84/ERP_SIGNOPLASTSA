Attribute VB_Name = "DAOOrdenCompraDetalle"
Option Explicit


Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ID_ORDEN_COMPRA As String = "id_orden_compra"
Public Const CAMPO_ID_PETICION_OFERTA_DETALLE As String = "id_peticion_oferta_detalle"
Public Const CAMPO_VALOR As String = "valor"
Public Const CAMPO_CANTIDAD As String = "cantidad"
Public Const CAMPO_DESCRIPCION As String = "descripcion"

Public Const TABLA_ORDEN_COMPRA_DETALLE As String = "ocd"

Public Function FindAllByOrdenCompraId(IdOrdenCompra As Long) As Collection
    Dim F As String: F = " ocd.id_orden_compra = " & IdOrdenCompra
    Set FindAllByOrdenCompraId = FindAll(F)
End Function

Public Function FindAll(Optional whereFilter As String = vbNullString) As Collection
    Dim q As String
    Dim rs As Recordset
    Dim Detalles As New Collection

    q = "SELECT *" _
      & " FROM ComprasOrdenesDetalles ocd" _
      & " LEFT JOIN ComprasOrdenesDetallesEntregas ocde ON ocde.id_detalle_orden_compra = ocd.id" _
      & " LEFT JOIN ComprasRemitosDetalles dr ON dr.id_detalle_orden_compra = ocd.id" _
      & " WHERE 1 = 1"

    If LenB(whereFilter) > 0 Then q = q & " AND " & whereFilter

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim ocd As OrdenCompraDetalle
    Dim ocde As OrdenCompraDetalleEntrega
    Dim drto As RemitoProveedorDetalle

    While Not rs.EOF

        If funciones.BuscarEnColeccion(Detalles, CStr(GetValue(rs, fieldsIndex, "ocd", "id"))) Then
            Set ocd = Detalles.item(CStr(rs(fieldsIndex("ocd.id")).value))
        Else
            Set ocd = DAOOrdenCompraDetalle.Map(rs, fieldsIndex, DAOOrdenCompraDetalle.TABLA_ORDEN_COMPRA_DETALLE)
        End If


        If Not funciones.BuscarEnColeccion(ocd.DetallesRemitos, CStr(GetValue(rs, fieldsIndex, "dr", "id"))) Then
            Set drto = DAORemitoProveedorDetalle.Map(rs, fieldsIndex, "dr")
            If IsSomething(drto) Then ocd.DetallesRemitos.Add drto, CStr(drto.Id)
        End If

        If Not funciones.BuscarEnColeccion(ocd.Entregas, CStr(GetValue(rs, fieldsIndex, "ocde", "id"))) Then
            Set ocde = DAOOrdenCompraDetalleEntrega.Map(rs, fieldsIndex, "ocde")
            If IsSomething(ocde) Then ocd.Entregas.Add ocde, CStr(ocde.Id)
        End If

        If Not funciones.BuscarEnColeccion(Detalles, CStr(GetValue(rs, fieldsIndex, "ocd", "id"))) Then
            Detalles.Add ocd, CStr(ocd.Id)
        End If

        rs.MoveNext
    Wend

    Set FindAll = Detalles
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As OrdenCompraDetalle
    Dim ocd As OrdenCompraDetalle
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, DAOOrdenCompraDetalle.CAMPO_ID)

    If Id > 0 Then
        Set ocd = New OrdenCompraDetalle
        ocd.Id = Id

        ocd.Cantidad = GetValue(rs, indice, tabla, DAOOrdenCompraDetalle.CAMPO_CANTIDAD)
        ocd.descripcion = GetValue(rs, indice, tabla, DAOOrdenCompraDetalle.CAMPO_DESCRIPCION)
        ocd.IdOrdenCompra = GetValue(rs, indice, tabla, DAOOrdenCompraDetalle.CAMPO_ID_ORDEN_COMPRA)
        ocd.Valor = GetValue(rs, indice, tabla, DAOOrdenCompraDetalle.CAMPO_VALOR)
        ocd.IdPeticionOfertaDetalle = GetValue(rs, indice, tabla, DAOOrdenCompraDetalle.CAMPO_ID_PETICION_OFERTA_DETALLE)
    End If

    Set Map = ocd
End Function


Public Function Update(ocd As OrdenCompraDetalle) As Boolean
    Dim q As String

    q = "Update ComprasOrdenesDetalles" _
      & " SET" _
      & " valor = " & conectar.Escape(ocd.Valor) & " ," _
      & " cantidad = " & conectar.Escape(ocd.Cantidad) & " ," _
      & " descripcion = " & conectar.Escape(ocd.descripcion) _
      & " Where id = " & ocd.Id

    Update = conectar.execute(q)
End Function

