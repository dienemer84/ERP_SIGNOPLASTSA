Attribute VB_Name = "DAOCertificadoRetencion"
Option Explicit


Public Function VerPosibleRetenciones(colFc As Collection, colret As Collection, Alicuota As Double, diferenciaDeCambio As Double, Optional TotalNGCompensatorios As Double = 0) As Dictionary
    'col es col de fcproveedor
    Dim F As clsFacturaProveedor
    Dim ret As Retencion
    Dim c As Double
    Dim dic As New Dictionary
    Dim difCambioFactura As Double
    Dim sumadorDeTotales As Double
    Dim totCambiong As Double
    Dim totdifcambio_hoy As Double
    Dim totdifcambio As Double
    For Each ret In colret
        c = 0
        For Each F In colFc

            If F.FormaPagoCuentaCorriente Then
'sumadorDeTotales = sumadorDeTotales + IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoDiaPago * -1, F.NetoGravadoDiaPago)
'fix 004
                sumadorDeTotales = sumadorDeTotales + IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, (F.NetoGravadoDiaPago) * -1, F.NetoGravadoDiaPago)

            End If

            sumadorDeTotales = sumadorDeTotales + diferenciaDeCambio

        Next F


        'cambie el priemr mas. Habia un menos antes
        '         If (sumadorDeTotales - diferenciaDeCambio - TotalNGCompensatorios) > ret.MinimoImponible Then
        '            sumadorDeTotales = (sumadorDeTotales - diferenciaDeCambio - TotalNGCompensatorios) * (Alicuota / 100)
        '        Else
        '            sumadorDeTotales = 0
        '        End If
        'cambiado el 22-10-12
        If (sumadorDeTotales - TotalNGCompensatorios) > ret.MinimoImponible Then
            sumadorDeTotales = (sumadorDeTotales - TotalNGCompensatorios) * (Alicuota / 100)
        Else
            sumadorDeTotales = 0
        End If




        dic.Add CStr(ret.id), funciones.RedondearDecimales(sumadorDeTotales, 2)
    Next ret

    Set VerPosibleRetenciones = dic
    'trae en un diccionario como clave la retencion (que sea agente) y el valor a retener de la lista
    'de facturas provista.
End Function

Public Function Create(op As OrdenPago, Optional Save As Boolean = False) As CertificadoRetencion
    Dim fac As clsFacturaProveedor
    Dim cer As CertificadoRetencion
    If op.StaticTotalRetenido > 0 Then   'op.FacturasProveedor.count > 0 Then
        Dim prov As clsProveedor
        Set prov = op.FacturasProveedor.item(1).Proveedor
        Set cer = New CertificadoRetencion
        cer.IdOrdenPago = op.id
        cer.Cuit = prov.Cuit
        cer.FEcha = Now
        cer.localidad = prov.Ciudad
        cer.cp = Val(prov.cp)
        cer.Domicilio = prov.direccion
        cer.IB = prov.IIBB
        Set cer.Retencion = DAORetenciones.FindAllEsAgente()(1)
        cer.RazonSocial = prov.RazonSocial
        Set cer.Detalles = New Collection
        Dim cerd As CertificadoRetencionDetalles


        If op.StaticTotalFacturasNG > cer.Retencion.MinimoImponible Then  'doble chequeo por las dudas
            For Each fac In op.FacturasProveedor
                Set fac = DAOFacturaProveedor.FindById(fac.id)
                Set cerd = New CertificadoRetencionDetalles
                cerd.Alicuota = op.Alicuota
                Set cerd.FacturaProveedor = fac
                'cerd.NetoGravado = IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.NetoGravado * -1, fac.NetoGravado)
                'fix 004
                cerd.NetoGravado = IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.NetoGravado * -1, fac.NetoGravado)
                cerd.Comprobante = fac.NumeroFormateado
                cerd.TotalFactura = IIf(fac.tipoDocumentoContable = tipoDocumentoContable.notaCredito, fac.Total * -1, fac.Total)
                cerd.IdCertificado = cer.id    'al pedo
                cer.Detalles.Add cerd
            Next
        End If


        If Save Then
            If Not DAOCertificadoRetencion.Guardar(cer, True) Then

                GoTo err1
            End If
        End If

    End If

    Set Create = cer
    Exit Function
err1:
    Create = Nothing
End Function

Public Function Save(Certificado As CertificadoRetencion, Optional cascada As Boolean = False) As Boolean
    conectar.BeginTransaction
    Save = Guardar(Certificado, cascada)
    conectar.CommitTransaction
    Exit Function
    Save = False
    conectar.RollBackTransaction
End Function

Public Function Guardar(Certificado As CertificadoRetencion, Optional cascada As Boolean = False) As Boolean
    Dim nuevo As Boolean
    Dim q As String
    If Certificado.id = 0 Then
        nuevo = True
        q = " INSERT INTO sp.certificados_retencion   (" _
            & "id_orden_pago, " _
            & "razon_social," _
            & "cuit," _
            & "ib, " _
            & "domicilio, " _
            & "localidad, " _
            & "id_retencion, " _
            & " FEcha )  Values " _
            & "('id_orden_pago', " _
            & "'razon_social', " _
            & "'cuit', " _
            & "'ib', " _
            & "'domicilio', " _
            & "'localidad', " _
            & "'id_retencion'," _
            & "'fecha' )"


    Else
        nuevo = False
        q = "Update sp.certificados_retencion  SET " _
            & "  id_orden_pago = 'id_orden_pago' , " _
            & "razon_social = 'razon_social' , " _
            & "cuit = 'cuit' , " _
            & "ib = 'ib' , " _
            & "domicilio = 'domicilio' , " _
            & "localidad = 'localidad' , " _
            & "cp = 'cp' , " _
            & "id_retencion = 'id_retencion' , " _
            & "fecha = 'fecha' " _
            & "Where   id = 'id' "

    End If
    q = Replace(q, "'id'", Escape(Certificado.id))
    q = Replace(q, "'id_orden_pago'", Certificado.IdOrdenPago)
    q = Replace(q, "'razon_social'", Escape(Certificado.RazonSocial))
    q = Replace(q, "'cuit'", Escape(Certificado.Cuit))

    q = Replace(q, "'id_retencion'", GetEntityId(Certificado.Retencion))
    q = Replace(q, "'fecha'", Escape(Certificado.FEcha))
    q = Replace(q, "'domicilio'", Escape(Certificado.Domicilio))
    q = Replace(q, "'localidad'", Escape(Certificado.localidad))
    q = Replace(q, "'cp'", Escape(Certificado.cp))
    q = Replace(q, "'ib'", Escape(Certificado.IB))

    If Not conectar.execute(q) Then GoTo err1
    If nuevo Then Certificado.id = conectar.UltimoId2

    Dim det As CertificadoRetencionDetalles
    If IsSomething(Certificado.Detalles) And cascada Then
        If Certificado.Detalles.count > 0 Then
            For Each det In Certificado.Detalles
                conectar.execute "delete from certificados_retencion_detalles where id=" & det.id
                det.id = 0    'fuerzo
                det.IdCertificado = Certificado.id
                If Not DAOCertificadoRetencionDetalles.Guardar(det) Then GoTo err1
            Next

        End If
    End If



    Guardar = True
    Exit Function
err1:
    Guardar = False

End Function

Public Function FindByOrdenPago(idOP As Long) As CertificadoRetencion
    Dim col As Collection
    Set col = FindAll("id_orden_pago=" & idOP, True)
    If col.count > 0 Then
        Set FindByOrdenPago = col(1)
    Else
        Set FindByOrdenPago = Nothing
    End If

End Function

Public Function FindAll(Optional filtro As String = vbNullString, Optional withDetalles As Boolean = False) As Collection
    Dim col As New Collection
    Dim rs As Recordset
    Dim q As String
    Dim idx As Dictionary
    q = "SELECT  * FROM  sp.certificados_retencion left join retenciones on certificados_retencion.id_retencion=retenciones.id where 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, idx
    While Not rs.EOF And Not rs.BOF
        Dim c As CertificadoRetencion
        Set c = Map(rs, idx, "certificados_retencion", "retenciones")

        If withDetalles Then
            Set c.Detalles = DAOCertificadoRetencionDetalles.FindAllByCertificadoId(c.id)
        End If

        col.Add c
        rs.MoveNext
    Wend

    Set FindAll = col
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaRet As String = vbNullString) As CertificadoRetencion
    Dim cer As CertificadoRetencion
    Dim id As Variant
    id = GetValue(rs, indice, tabla, "id")
    If id > 0 Then
        Set cer = New CertificadoRetencion
        cer.id = id
        cer.Cuit = GetValue(rs, indice, tabla, "cuit")
        cer.FEcha = GetValue(rs, indice, tabla, "fecha")
        cer.IB = GetValue(rs, indice, tabla, "ib")
        cer.RazonSocial = GetValue(rs, indice, tabla, "razon_social")
        cer.Domicilio = GetValue(rs, indice, tabla, "domicilio")
        cer.localidad = GetValue(rs, indice, tabla, "localidad")
        cer.cp = GetValue(rs, indice, tabla, "cp")
        cer.IdOrdenPago = GetValue(rs, indice, tabla, "id_orden_pago")
        If LenB(tablaRet) > 0 Then Set cer.Retencion = DAORetenciones.Map(rs, indice, tablaRet)
    End If

    Set Map = cer

End Function

Public Function VerCertificado(cr As CertificadoRetencion)
    If Not IsSomething(cr) Then
        MsgBox "No se pudo encontrar el certificado de retención.", vbExclamation
        Exit Function
    End If
    Dim mon As clsMoneda

    Set mon = DAOMoneda.FindFirstByPatronOrDefault

    dsrCertificadoIIBB.Sections("sec1").Controls("lblContribuyente").caption = "Contribuyente: " & cr.RazonSocial
    dsrCertificadoIIBB.Sections("sec1").Controls("lblDomicilio").caption = "Domicilio: " & cr.DomicilioFormateado
    dsrCertificadoIIBB.Sections("sec1").Controls("lblCuitIB").caption = cr.CuitIBFormateado
    dsrCertificadoIIBB.Sections("sec4").Controls("lblNumero").caption = "Certificado Nº " & Format(cr.id, "0000")
    dsrCertificadoIIBB.Sections("sec4").Controls("lblRetencion").caption = "Impuesto sobre " & cr.Retencion.nombre
    dsrCertificadoIIBB.Sections("sec5").Controls("lblTotalRetenido").caption = "Total Retenido " & mon.NombreCorto & " " & funciones.FormatearDecimales(cr.Total)


    Dim r_tmp As Recordset
    Set r_tmp = New Recordset
    With r_tmp
        .Fields.Append "nrocomprobante", adVarChar, 255, adFldUpdatable     ' And adFldIsNullable
        .Fields.Append "importe", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "importe_retenido", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
    End With
    r_tmp.Open

    Dim fac As clsFacturaProveedor
    Dim A As Double
    Dim det As CertificadoRetencionDetalles
    Dim fc As clsFacturaProveedor
    If cr.Detalles.count > 0 Then

        For Each det In cr.Detalles
            A = MonedaConverter.Convertir(det.TotalFactura, det.IdMoneda, mon.id)
            r_tmp.AddNew
            r_tmp!Importe = mon.NombreCorto & " " & funciones.FormatearDecimales(A)
            r_tmp!importe_retenido = mon.NombreCorto & " " & det.TotalRetenido(cr)
            r_tmp!nrocomprobante = det.Comprobante
            r_tmp.Update
        Next

        r_tmp.MoveFirst    'lo agregue yo

        Set dsrCertificadoIIBB.DataSource = r_tmp
        dsrCertificadoIIBB.Show vbModal

    Else
        MsgBox "El certificado de retención no tiene detalles.", vbExclamation
        Exit Function
    End If

End Function
