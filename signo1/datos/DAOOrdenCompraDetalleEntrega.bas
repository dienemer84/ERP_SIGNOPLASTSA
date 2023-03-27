Attribute VB_Name = "DAOOrdenCompraDetalleEntrega"
Option Explicit

Public Function Update(ent As OrdenCompraDetalleEntrega) As Boolean
    Dim q As String

    q = "Update ComprasOrdenesDetallesEntregas" _
      & " SET" _
      & " fecha = " & conectar.Escape(ent.FEcha) & " ," _
      & " cant = " & ent.Cantidad _
      & " WHERE id = " & ent.Id
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As OrdenCompraDetalleEntrega
    Dim ocde As OrdenCompraDetalleEntrega
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set ocde = New OrdenCompraDetalleEntrega
        ocde.Id = Id
        ocde.Cantidad = GetValue(rs, indice, tabla, "cant")
        ocde.FEcha = GetValue(rs, indice, tabla, "fecha")
        ocde.IdDetalleOrdenCompra = GetValue(rs, indice, tabla, "id_detalle_orden_compra")
    End If

    Set Map = ocde
End Function
