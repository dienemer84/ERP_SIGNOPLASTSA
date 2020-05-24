Attribute VB_Name = "DAOCertificadoRetencionDetalles"
Option Explicit
Public Function FindAllByCertificadoId(id As Long) As Collection
    Set FindAllByCertificadoId = FindAll("id_certificado=" & id)
End Function
Public Function Save(cr As CertificadoRetencionDetalles) As Boolean

    conectar.BeginTransaction
    If Not Guardar(cr) Then GoTo err1
    conectar.CommitTransaction
    Exit Function
err1:
    conectar.RollBackTransaction
    Save = False

End Function
Public Function Guardar(cr As CertificadoRetencionDetalles) As Boolean
    Guardar = True
    Dim q As String
    Dim nuevo As Boolean
    If cr.id = 0 Then
        nuevo = True
        q = "INSERT INTO sp.certificados_retencion_detalles" _
            & "(id_factura_proveedor, Alicuota,id_certificado,comprobante,neto_gravado,id_moneda,total_factura) Values" _
            & "('id_factura_proveedor','alicuota','id_certificado','comprobante','neto_gravado','id_moneda','total_factura')"
    Else
        nuevo = False
        q = "Update sp.certificados_retencion_detalles  SET" _
            & "id = 'id' , " _
            & "id_factura_proveedor = 'id_factura_proveedor' , " _
            & "alicuota = 'alicuota', " _
            & "comprobante = 'comprobante', " _
            & "neto_gravado = 'neto_gravado', " _
            & "id_certificado = 'id_certificado' " _
            & "total_factura = 'total_factura' " _
            & "id_moneda = 'id_moneda' " _
            & " Where " _
            & "id = 'id' "
    End If

    q = Replace(q, "'id'", Escape(cr.id))
    q = Replace(q, "'id_factura_proveedor'", Escape(cr.FacturaProveedor.id))
    q = Replace(q, "'alicuota'", Escape(cr.Alicuota))
    q = Replace(q, "'comprobante'", Escape(cr.Comprobante))
    q = Replace(q, "'neto_gravado'", Escape(cr.NetoGravado))
    q = Replace(q, "'id_certificado'", Escape(cr.IdCertificado))
    q = Replace(q, "'id_moneda'", Escape(cr.IdMoneda))
    q = Replace(q, "'total_factura'", Escape(cr.TotalFactura))

    If Not conectar.execute(q) Then GoTo err1
    If nuevo Then cr.id = conectar.UltimoId2
    Guardar = True
    Exit Function
err1:

    Guardar = False

End Function

Public Function FindAll(Optional filtro As String = vbNullString, Optional WithFactura As Boolean = False) As Collection
    Dim rs As Recordset
    Dim q As String
    Dim idx As Dictionary
    Dim col As Collection
    q = "select * from certificados_retencion_detalles  where 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Dim cr As CertificadoRetencionDetalles
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, idx
    Set col = New Collection
    While Not rs.EOF And Not rs.BOF
        'Set cr = New CertificadoRetencionDetalles
        Set cr = Map(rs, idx, "certificados_retencion_detalles")
        col.Add cr, CStr(cr.id)
        rs.MoveNext
    Wend
    Set FindAll = col
End Function

Public Function Map(rs As Recordset, idx As Dictionary, tabla, Optional WithFactura As Boolean = False) As CertificadoRetencionDetalles
    Dim cr As CertificadoRetencionDetalles
    Dim id As Long
    id = GetValue(rs, idx, "certificados_retencion_detalles", "id")
    If id > 0 Then
        Dim idf As Long
        Set cr = New CertificadoRetencionDetalles
        idf = GetValue(rs, idx, "certificados_retencion_detalles", "id_factura_proveedor")
        If WithFactura Then Set cr.FacturaProveedor = DAOFacturaProveedor.FindById(idf)
        cr.id = id
        cr.IdCertificado = GetValue(rs, idx, "certificados_retencion_detalles", "id_certificado")
        cr.Alicuota = GetValue(rs, idx, "certificados_retencion_detalles", "alicuota")
        cr.NetoGravado = GetValue(rs, idx, "certificados_retencion_detalles", "neto_gravado")
        cr.Comprobante = GetValue(rs, idx, "certificados_retencion_detalles", "comprobante")
        cr.IdMoneda = GetValue(rs, idx, "certificados_retencion_detalles", "id_moneda")
        cr.TotalFactura = GetValue(rs, idx, "certificados_retencion_detalles", "total_factura")
    End If
    Set Map = cr
End Function




