Attribute VB_Name = "DAOFacturaDetalles"
Option Explicit

Public Function Guardar(detalle As FacturaDetalle) As Boolean
    On Error GoTo err1
    Guardar = True
    Dim q As String
    If detalle.Id > 0 Then
        q = "Update sp.AdminFacturasDetalleNueva  SET " _
          & "idEntrega = 'idEntrega' ," _
          & "idFactura = 'idFactura' ," _
          & "estado = 'estado' ," _
          & "Valor = 'Valor' ," _
          & "detalle = 'detalle' ," _
          & "Cantidad = 'Cantidad' ," _
          & "iva = 'iva' ," _
          & "observaciones = 'observaciones' ," _
          & "IB = 'IB' ," _
          & "aplicada = 'aplicada' ," _
          & "descuento_anticipo = 'descuento_anticipo', idOt_anticipo = 'idOt_anticipo' " _
          & " Where id = 'id' "

        q = Replace$(q, "'id'", conectar.Escape(detalle.Id))
    Else
        q = " INSERT INTO sp.AdminFacturasDetalleNueva " _
          & "(idEntrega, idFactura,  estado, Valor, " _
          & "detalle,  Cantidad,    iva,   IB,  aplicada, observaciones,    porcentaje_descuento,descuento_anticipo,idOt_anticipo  )  Values " _
          & " ('idEntrega',  'idFactura',  'estado',  'Valor',  'detalle', " _
          & "'Cantidad', 'iva', 'IB', 'aplicada', 'observaciones', 'porcentaje_descuento', 'descuento_anticipo','idOt_anticipo')"
    End If



    'q = Replace$(q, "'idEntrega'", conectar.GetEntityId(detalle.DetalleRemito))
    q = Replace$(q, "'idEntrega'", detalle.DetalleRemitoId)
    q = Replace$(q, "'idFactura'", conectar.Escape(detalle.idFactura))
    q = Replace$(q, "'estado'", conectar.Escape(detalle.estado))
    q = Replace$(q, "'Valor'", conectar.Escape(detalle.Bruto))
    q = Replace$(q, "'detalle'", conectar.Escape(detalle.detalle))
    q = Replace$(q, "'Cantidad'", conectar.Escape(detalle.Cantidad))
    q = Replace$(q, "'iva'", conectar.Escape(detalle.IvaAplicado))
    q = Replace$(q, "'id'", conectar.Escape(detalle.Id))
    q = Replace$(q, "'IB'", conectar.Escape(detalle.IBAplicado))
    q = Replace$(q, "'observaciones'", conectar.Escape(detalle.Observacion))
    q = Replace$(q, "'porcentaje_descuento'", conectar.Escape(detalle.PorcentajeDescuento))
    q = Replace$(q, "'aplicada'", conectar.Escape(detalle.AplicadoARemito))
    q = Replace$(q, "'descuento_anticipo'", conectar.Escape(detalle.DescuentoAnticipo))
    q = Replace$(q, "'idOt_anticipo'", conectar.Escape(detalle.OtIdAnticipo))



    If Not conectar.execute(q) Then Err.Raise 4784, "Guardando detalle de factura", "Se produjo un error al guardar un detalle de factura"

    Exit Function
err1:
    Err.Raise Err.Number, Err.Source, Err.Description
    Guardar = False
End Function

Public Function Save(deta As FacturaDetalle) As Boolean
    Save = True
    conectar.BeginTransaction
    If Not Guardar(deta) Then GoTo err1
    conectar.CommitTransaction
    Exit Function
err1:
    conectar.RollBackTransaction
    Save = False
End Function
Public Function FindByFactura(idFactura As Long) As Collection
    Set FindByFactura = FindAll("idFactura= " & idFactura)
End Function



Public Function FindAll(Optional filtro As String = vbNullString, Optional withFacturaNoLazy As Boolean = False) As Collection
    Dim q As String
    Dim rs As Recordset
    Dim col As New Collection
    Dim F As FacturaDetalle
    Dim indice As Dictionary
    q = "SELECT * From sp.AdminFacturasDetalleNueva " _
      & "LEFT JOIN sp.entregas ON (AdminFacturasDetalleNueva.idEntrega = entregas.id) " _
      & " LEFT JOIN sp.AdminFacturas fc ON (fc.id = AdminFacturasDetalleNueva.idFactura) WHERE 1=1 "
      
    If LenB(filtro) > 0 Then
        q = q & " AND " & filtro
        
'        q = q & " AND fc.estado = 2 AND " & filtro
        
    End If

    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, indice

    While Not rs.EOF

        Set F = DAOFacturaDetalles.Map(rs, indice, "AdminFacturasDetalleNueva", "entregas")
        If withFacturaNoLazy Then
            Set F.Factura = DAOFactura.FindById(F.idFactura)
        End If
        col.Add F, CStr(F.Id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tablaFacturaDetalle As String, Optional tablaEntrega As String = vbNullString) As FacturaDetalle
    Dim Id As Long
    Dim F As FacturaDetalle
    Id = GetValue(rs, indice, tablaFacturaDetalle, "id")

    If Id > 0 Then
        Set F = New FacturaDetalle
        F.Id = Id
        F.AplicadoARemito = GetValue(rs, indice, tablaFacturaDetalle, "aplicada")
        F.IBAplicado = GetValue(rs, indice, tablaFacturaDetalle, "IB")
        F.IvaAplicado = GetValue(rs, indice, tablaFacturaDetalle, "iva")
        F.Bruto = GetValue(rs, indice, tablaFacturaDetalle, "Valor")

        F.detalle = GetValue(rs, indice, tablaFacturaDetalle, "detalle")

        F.estado = GetValue(rs, indice, tablaFacturaDetalle, "estado")
        F.idFactura = GetValue(rs, indice, tablaFacturaDetalle, "idFactura")
        F.Cantidad = GetValue(rs, indice, tablaFacturaDetalle, "Cantidad")
        F.DetalleRemitoId = GetValue(rs, indice, tablaFacturaDetalle, "idEntrega")
        F.Observacion = GetValue(rs, indice, tablaFacturaDetalle, "observaciones")
        F.PorcentajeDescuento = GetValue(rs, indice, tablaFacturaDetalle, "porcentaje_descuento")
        F.DescuentoAnticipo = GetValue(rs, indice, tablaFacturaDetalle, "descuento_anticipo")
        F.OtIdAnticipo = GetValue(rs, indice, tablaFacturaDetalle, "idOt_anticipo")

        If LenB(tablaEntrega) > 0 Then
            Set F.detalleRemito = DAORemitoSDetalle.Map(rs, indice, tablaEntrega)
        End If
    End If
    Set Map = F
End Function

Public Function Delete(filter As String) As Boolean
    Delete = conectar.execute("DELETE FROM AdminFacturasDetalleNueva WHERE " & filter)
End Function
