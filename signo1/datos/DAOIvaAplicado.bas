Attribute VB_Name = "DAOIvaAplicado"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection


Public Function listByIdFactura(id_factura As Long) As Collection
    Dim col As New Collection
    Dim B As clsAlicuotaAplicada
    Dim A As clsTipoIVA
    Set rs = conectar.RSFactory("select * from AdminComprasFacturasProveedoresIva where id_factura_proveedor=" & id_factura)
    While Not rs.EOF
        Set B = New clsAlicuotaAplicada
        Set A = New clsTipoIVA
        B.Monto = rs!Valor
        B.alicuota = DAOAlicuotas.GetById(rs!id_iva)
        col.Add B
        rs.MoveNext
    Wend

    Set listByIdFactura = col
End Function
Public Function Save(fc As clsFacturaProveedor) As Boolean
    Save = True
    Set cn = conectar.obternerConexion
    On Error GoTo er1:
    cn.execute "delete from AdminComprasFacturasProveedoresIva where id_factura_proveedor=" & fc.Id
  
    For K = 1 To fc.IvaAplicado.count

'22/08/2022
' EN ESTA FUNCION AGREGO EL CALCULO QUE VA EN LA NUEVA COLUMNA IVA_CALCULADO
' EN ESTA COLUMNA VA EL VALOR DEL MONTO * VALOR DE ALICUOTA /100

'        'debug.print ("-------------------")
'        'debug.print (fc.IvaAplicado(K).alicuota.Id)
'        'debug.print (fc.IvaAplicado(K).Monto)
'        'debug.print (fc.Id)
'        'debug.print (fc.IvaAplicado(K).alicuota.alicuota)
'        'debug.print (fc.IvaAplicado(K).Monto * (fc.IvaAplicado(K).alicuota.alicuota / 100))
'        'debug.print ("-------------------")
        
'22/08/2022
' COMENTO LA QUERY ANTERIOR
        'cn.execute "insert into AdminComprasFacturasProveedoresIva (id_iva, valor, id_factura_proveedor) values (" & fc.IvaAplicado(K).alicuota.Id & "," & fc.IvaAplicado(K).Monto & "," & fc.Id & ")"

'22/08/2022
' ACA AGREGO ESE CAMPO EN LA QUERY (iva_calculado) / (fc.IvaAplicado(K).Monto * (fc.IvaAplicado(K).alicuota.alicuota / 100)
        cn.execute "insert into AdminComprasFacturasProveedoresIva (id_iva, valor, id_factura_proveedor,iva_calculado) values (" & fc.IvaAplicado(K).alicuota.Id & "," & fc.IvaAplicado(K).Monto & "," & fc.Id & ", " & fc.IvaAplicado(K).Monto * funciones.FormatearDecimales(fc.IvaAplicado(K).alicuota.alicuota / 100) & " )"
    
    Next K

    Exit Function
er1:
    Save = False
    
End Function



Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, tablaAlicuota As String) As clsAlicuotaAplicada

    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim A As clsAlicuotaAplicada

    If Id > 0 Then
        Set A = New clsAlicuotaAplicada
        A.Id = Id
        A.alicuota = DAOAlicuotas.Map(rs, indice, tablaAlicuota)
        A.Monto = GetValue(rs, indice, tabla, "valor")
    End If
    Set Map = A
End Function
