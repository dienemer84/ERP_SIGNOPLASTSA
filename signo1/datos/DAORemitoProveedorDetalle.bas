Attribute VB_Name = "DAORemitoProveedorDetalle"
Option Explicit

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As RemitoProveedorDetalle
    Dim rpd As RemitoProveedorDetalle
    Dim id As Long
    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set rpd = New RemitoProveedorDetalle
        rpd.id = id
        rpd.IdRemito = GetValue(rs, indice, tabla, "id_remito")
        rpd.IdDetalleOrdenCompra = GetValue(rs, indice, tabla, "id_detalle_orden_compra")
        rpd.Cantidad = GetValue(rs, indice, tabla, "cantidad")
    End If

    Set Map = rpd
End Function

