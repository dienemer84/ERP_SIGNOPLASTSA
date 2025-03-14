Attribute VB_Name = "ERPHelper"
Option Explicit

'Private Const srv As String = "http://192.168.0.187:8080/ERPHelper/erphelper/"

Private srv As String
Public Enum verbo
    POST_ = 0
    GET_ = 1
End Enum


Private Function ApiConnect(sUrl As String, verb As verbo, async As Boolean, Optional data As String = vbNullString, Optional Class As Object = Nothing) As String
    On Error GoTo err1
    Dim verbo As String
    verbo = "POST"
  
    srv = LeerIni(App.path & "\config.ini", "Configurar", "ERPHelperAddress", vbNullString)
    
    Dim xmlhttp As New MSXML2.xmlhttp


    If verb = GET_ Then verbo = "GET"
    sUrl = srv & sUrl
    xmlhttp.Open verbo, sUrl, async

    Dim c As New XmlSerializer

    Dim s As String

    If LenB(data) > 0 Then
        s = data
    Else
        If IsSomething(Class) Then
            s = c.Serialize(Class)
        End If
    End If


    Dim objXMLSendDoc As MSXML2.DOMDocument
    Set objXMLSendDoc = New MSXML2.DOMDocument

    xmlhttp.setRequestHeader "Content-Type", "text/plain"
    xmlhttp.send s    'objXMLSendDoc.XML
    Dim response As String
    response = xmlhttp.responseText
'    Debug.Print response

    ApiConnect = response
    Exit Function
err1:
    Debug.Print Err.Description
    Err.Raise 999, "Afip", "Imposible conectar con gateway de facturaci�n"

End Function


Private Function ApiConnectEXP(sUrl As String, verb As verbo, async As Boolean, Optional data As String = vbNullString, Optional Class As Object = Nothing) As String
    On Error GoTo err1
    Dim verbo As String
    verbo = "POST"
    
    srv = LeerIni(App.path & "\config.ini", "Configurar", "ERPHelperAddressEXP", vbNullString)
    
    Dim xmlhttp As New MSXML2.xmlhttp


    If verb = GET_ Then verbo = "GET"
    sUrl = srv & sUrl
    xmlhttp.Open verbo, sUrl, async

    Dim c As New XmlSerializer

    Dim s As String

    If LenB(data) > 0 Then
        s = data
    Else
        If IsSomething(Class) Then
            s = c.Serialize(Class)
        End If
    End If


    Dim objXMLSendDoc As MSXML2.DOMDocument
    Set objXMLSendDoc = New MSXML2.DOMDocument

    xmlhttp.setRequestHeader "Content-Type", "text/plain"
    xmlhttp.send s    'objXMLSendDoc.XML
    Dim response As String
    response = xmlhttp.responseText
'    Debug.Print response

    ApiConnectEXP = response
    Exit Function
err1:
    Debug.Print Err.Description
    Err.Raise 999, "Afip", "Imposible conectar con gateway de facturaci�n"

End Function


Public Function RecuperarCae(idPtoVta As String, tipoCbte As String, nroCbte As String) As String
    On Error GoTo err1
    Dim resp As String
    If CheckDummyAfip Then
        resp = ApiConnect("wsfe/FECompUltimoAutorizado/" & idPtoVta & "/" & tipoCbte & "/" & nroCbte, POST_, False)
    Else
        Err.Raise 1002, "Afip", "Imposible obtener el cae solicitado"
    End If
    Exit Function
err1:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'Desactivada el 17.07.20 -dnemer
Public Function GetUltimoAutorizado(idPtoVta As String, tipoComprobante As String, esCredito As Boolean) As String
    On Error GoTo err1
    Dim resp As String
    If CheckDummyAfip Then

        'Reestablecida esta linea de codigo el 17.07.20 -dnemer
        'resp = ApiConnect("wsfe/FECompUltimoAutorizado/" & idPtoVta & "/" & tipoComprobante, POST_, False)


        'Desactivada el 17.07.20 -dnemer
        If esCredito Then
            resp = ApiConnect("wsfe/FECompUltimoAutorizado/" & idPtoVta & "/" & tipoComprobante & "/true", POST_, False)
        Else
            resp = ApiConnect("wsfe/FECompUltimoAutorizado/" & idPtoVta & "/" & tipoComprobante, POST_, False)
        End If
    Else
        Err.Raise 1002, "Afip", "Imposible obtener ultimo comprobante autorizado"
    End If
    '    If resp = "0" Then
    'Err.Raise 1005, "Afip", "Error al obtener ultimo autorizado"
    '    End If
    GetUltimoAutorizado = resp

    Exit Function
err1:

    GetUltimoAutorizado = "-1"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function GetUltimoAutorizadoEXP(idPtoVta As String, tipoComprobante As String, esCredito As Boolean) As String
    On Error GoTo err1
    Dim resp As String
    If CheckDummyAfip Then

        'Reestablecida esta linea de codigo el 17.07.20 -dnemer
        'resp = ApiConnect("wsfe/FECompUltimoAutorizado/" & idPtoVta & "/" & tipoComprobante, POST_, False)

    resp = ApiConnectEXP("wsfe/FECompUltimoAutorizado/" & idPtoVta & "/" & tipoComprobante, POST_, False)

    Else
        Err.Raise 1002, "Afip", "Imposible obtener ultimo comprobante autorizado"
    End If
    '    If resp = "0" Then
    'Err.Raise 1005, "Afip", "Error al obtener ultimo autorizado"
    '    End If
    GetUltimoAutorizadoEXP = resp

    Exit Function
err1:

    GetUltimoAutorizadoEXP = "-1"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function CheckDummyAfip() As Boolean
    On Error GoTo err1
    Dim resp As String
    resp = ApiConnect("wsfe/FEDummy", POST_, False)
    If Not HasErrorMessage(resp) And resp = "1" Then
        CheckDummyAfip = True
    Else
        CheckDummyAfip = False
        Err.Raise 1001, "Afip", "Infraestructura no disponible"
    End If
    Exit Function
err1:
    Err.Raise Err.Number, Err.Source, Err.Description
    CheckDummyAfip = False
End Function


Public Function HasErrorMessage(resp As String) As Boolean
    Dim Text1 As String
    HasErrorMessage = False
    Text1 = "ERROR 404"
    If InStr(Text1, resp) > 0 Then
        GoTo ERR_404
    End If

    Exit Function
ERR_404:
    Err.Raise 100404, "Response", "Se produjo error HTTP 404"
    HasErrorMessage = True
End Function


Public Function ObtenerUltimoActual(F As Factura) As Integer
    Dim id

    id = CLng(ERPHelper.GetUltimoAutorizado(F.Tipo.PuntoVenta.PuntoVenta, F.Tipo.id, F.esCredito))

    ObtenerUltimoActual = id
End Function


Public Function ObtenerUltimoActualEXP(F As Factura) As Integer
    Dim id

    id = CLng(ERPHelper.GetUltimoAutorizadoEXP(F.Tipo.PuntoVenta.PuntoVenta, F.Tipo.id, F.esCredito))

    ObtenerUltimoActualEXP = id
End Function

Public Function CreateFECaeSolicitarRequest(F As Factura) As CAESolicitar
    On Error GoTo err1

    Dim body As String

    'Set f = DAOFactura.FindById(idFactura)
    Dim id

    'Desactivada el 17.07.20 -dnemer
    id = CLng(ERPHelper.GetUltimoAutorizado(F.Tipo.PuntoVenta.PuntoVenta, F.Tipo.id, F.esCredito))

    'Reestablecida esta linea de codigo el 17.07.20 -dnemer
    'id = CLng(ERPHelper.GetUltimoAutorizado(F.Tipo.PuntoVenta.PuntoVenta, F.Tipo.id))


    F.numero = id + 1

    Dim req As New FECAEDetRequest
    
    req.Concepto = F.ConceptoIncluir
    req.DocTipo = F.Cliente.TipoDocumento
    req.DocNro = F.Cliente.Cuit
    req.CbteDesde = F.numero
    req.CbteHasta = F.numero

    req.CbteFch = Format(F.FechaEmision, "yyyymmdd")
    req.ImpTotal = funciones.FormatearDecimales(F.TotalEstatico.total, 2)
    'req.MonCotiz = F.CambioAPatron

    'req.ImpTotal = funciones.FormatearDecimales(F.TotalEstatico.Total * F.CambioAPatron, 2)
    req.ImpTotConc = "0"    'no gavado+excento + gravado + iva + tributo
    req.ImpNeto = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado, 2)
    'req.ImpNeto = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado * F.CambioAPatron, 2)
    req.ImpOpEx = funciones.FormatearDecimales(F.TotalEstatico.TotalExento, 2)
    'req.ImpOpEx = funciones.FormatearDecimales(F.TotalEstatico.TotalExento * F.CambioAPatron, 2)

    'Reestablecida el 17.07.20 -dnemer
    'req.FchServDesde = ""    ' obligatorio para concepto de tipo 3 y 2
    'req.FchServHasta = ""    ' obligatorio para concepto de tipo 3 y 2

    'Desactivada el 17.07.20 -dnemer

    If F.ConceptoIncluir = ConceptoIncluir.ConceptoProductoServicio Or F.ConceptoIncluir = ConceptoIncluir.ConceptoServicio Then
        req.FchServDesde = Format(F.FechaVtoDesde, "yyyymmdd")   ' obligatorio para concepto de tipo 3 y 2
        req.FchServHasta = Format(F.FechaVtoHasta, "yyyymmdd")     ' obligatorio para concepto de tipo 3 y 2
    End If

    req.ImpTrib = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB, 2)
    'req.ImpTrib = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB * F.CambioAPatron, 2)

    req.ImpIVA = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo, 2)
    'req.ImpIVA = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo * F.CambioAPatron, 2)

    req.MonId = F.moneda.id
    'req.MonCotiz = F.Moneda.Cambio
    req.MonCotiz = F.CambioAPatron



    'Reestablecida el 17.07.20 -dnemer
    'req.FchVtoPago = ""    'obligatorio para concepto 2 y 3

    'Desactivada el 17.07.20 -dnemer

    If Not F.esCredito Then

        If F.ConceptoIncluir = ConceptoProductoServicio Or F.ConceptoIncluir = ConceptoProductoServicio Then
            req.FchVtoPago = Format(F.fechaPago, "yyyymmdd")       'obligatorio para concepto 2 y 3
        End If

    Else
        ' Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), el campo
        '�fecha de vencimiento para el pago� <FchVtoPago> no debe informarse si NO es
        'Factura de Cr�dito. En el caso de ser D�bito o Cr�dito, solo puede informarse si es de Anulaci�n.


        If F.TipoDocumento = tipoDocumentoContable.Factura Then
            req.FchVtoPago = Format(F.fechaPago, "yyyymmdd")
        Else
            If F.AnulacionAFIP = "S" Then
                req.FchVtoPago = Format(F.fechaPago, "yyyymmdd")
            End If
        End If
    End If

    'Desactivada el 17.07.20 -dnemer


    'NB: afip rechaza nc y nd que no sean mi pyme sin tener cbte asociado (err 10197 afip)
    'If F.esCredito And (F.TipoDocumento = tipoDocumentoContable.notaCredito Or F.TipoDocumento = tipoDocumentoContable.notaDebito) Then

    If F.TipoDocumento = tipoDocumentoContable.notaCredito Or F.TipoDocumento = tipoDocumentoContable.notaDebito Then
        Dim ftmp As Factura
        Set ftmp = DAOFactura.FindById(F.Cancelada)
        '23-8 NB: no puedo informar un comprobnte asociado que no est� previamente informado. (en caso que sea cr�dito o d�bido mipyme)

        If IsSomething(ftmp) Then
            If Not ftmp.AprobadaAFIP Then
                Err.Raise 22211, "El comprobante asociado debe estar informado previamente a AFIP"
            End If
        End If
        If IsSomething(ftmp) Then
            Dim cbt As CbteAsoc
            Set cbt = New CbteAsoc

            '       cbt.NRO = f.Cancelada

            cbt.NRO = ftmp.numero
            If ftmp.esCredito Then
                cbt.esCredito = "true"
            Else
                cbt.esCredito = "false"
            End If
            cbt.PtoVta = ftmp.Tipo.PuntoVenta.PuntoVenta
            cbt.Tipo = ftmp.Tipo.id
            cbt.CbteFch = Format(ftmp.FechaEmision, "yyyymmdd")
            cbt.Cuit = "30657604972"
            req.CbtesAsoc.Add cbt
        End If
    End If

    'Dim cbt As CbteAsoc
    'Set cbt = New CbteAsoc

    'cbt.nro = "2"
    'cbt.PtoVta = "2"
    'cbt.Tipo = "2"
    'req.CbtesAsoc.Add cbt

    ' Set cbt = New CbteAsoc
    '
    'cbt.nro = "1"
    'cbt.PtoVta = "1"
    'cbt.Tipo = "1"

    '0req.CbtesAsoc.Add cbt
    '


    Dim trib As Tributo
    Set trib = New Tributo

    Dim P As New clsPercepciones
    Set P = DAOPercepciones.GetById(5)

    trib.Alic = funciones.FormatearDecimales(((F.AlicuotaPercepcionesIIBB - 1) * 100), 2)
    'trib.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado * F.CambioAPatron, 2)    '"400" 'revisar con tulio,2)

    'bug #4
    trib.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado, 2)    '
    trib.Desc = P.Percepcion
    trib.idTributoCambiar = "1"


    'trib.Importe = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB * F.CambioAPatron, 2)

    'bug #4
    trib.importe = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB, 2)
    If F.TotalEstatico.TotalPercepcionesIB > 0 Then
        req.Tributos.Add trib
    Else
        Set req.Tributos = Nothing

    End If


    Dim Iva As New AlicIva
    Iva.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado, 2)
    'Iva.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado * F.CambioAPatron, 2)
    Iva.idAlicIvaCambiar = F.AlicuotaAplicada

    Iva.importe = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo, 2)
    'Iva.Importe = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo * F.CambioAPatron, 2)    'mapear en erphelper por F.TipoIVA.idIVA,2)

    If F.TotalEstatico.TotalIVADiscrimandoONo > 0 Then
        req.Iva.Add Iva
    Else
        Set req.Iva = Nothing

    End If


    If F.esCredito Then
        '4-4-21 #218
        'Si el tipo de comprobante que est� autorizando es Factura del tipo MiPyMEs (201, 206, 211), es
        'obligatorio informar <Opcionales> con id = 27. Los valores posiblesson SCA o ADC


        'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), es obligatorio informar al menos uno de los sig. c�digos
        '2101, 22.

        If F.TipoDocumento = tipoDocumentoContable.Factura Then

            '4-4-21 #218
            'Si el tipo de comprobante que est� autorizando es Factura del tipo MiPyMEs (201, 206, 211), es
            'obligatorio informar <Opcionales> con id = 27. Los valores posiblesson SCA o ADC

            Dim op27 As New Opcional
            op27.idOpcionalCambiar = "27"

            If F.Opcional27 = 1 Then
                op27.Valor = "SCA"
            ElseIf F.Opcional27 = 2 Then
                op27.Valor = "ADC"
            End If

            req.Opcionales.Add op27


            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), informa opcionales, el valor correcto para el c�digo
            '2101 es un CBU num�rico de 22 caracteres.

            'Si el tipo de comprobante que est� autorizando es Factura (201, 206, 211) del tipo MiPyMEs (FCE), informa opcionales, es
            'obligatorio informar CBU.

            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), Debito (202, 207, 212) o Cr�dito (203, 208, 213) No
            'no informar CBU y ALIAS.
            If Len(F.CBU) <> 22 Then
                Err.Raise 222101, "El cbu debe tener 22 caracteres"
            End If

            Dim op As New Opcional
            op.idOpcionalCambiar = "2101"
            op.Valor = F.CBU
            req.Opcionales.Add op
        End If
        
        If F.TipoDocumento <> tipoDocumentoContable.Factura Then
            ' Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), informa opcionales, el valor correcto para el c�digo 22
            ' es �S� o �N�: S = Es de Anulaci�n N = No es de Anulaci�n

            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), Factura (201, 206, 211), no informar C�digo de Anulaci�n

            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), Debito (202, 207, 212) o Cr�dito (203, 208, 213) informar
            'C�digo de Anulaci�n

            If Len(F.AnulacionAFIP) <> 1 Then
                Err.Raise 22221, "El campo anulacion afip debe informarse"
            End If


            If F.AnulacionAFIP <> "N" And F.AnulacionAFIP <> "S" Then
                Err.Raise 22222, "El campo anulacion afip debe ser S o N"
            End If

            Set op = New Opcional
            op.idOpcionalCambiar = "22"
            op.Valor = F.AnulacionAFIP
            req.Opcionales.Add op
        End If


    End If

    Dim FeDetReq As New FeDetReq

    Set FeDetReq.FECAEDetRequest = req
    Dim FeCabReq As New FeCabReq
    FeCabReq.CantReg = "1"
    FeCabReq.CbteTipo = F.Tipo.id
    
    If F.esCredito Then
        FeCabReq.esCredito = "true"
    Else
        FeCabReq.esCredito = "false"
    End If




    FeCabReq.PtoVta = F.Tipo.PuntoVenta.id
    Dim FeCAEReq As New FeCAEReq

    Set FeCAEReq.FeCabReq = FeCabReq
    Set FeCAEReq.FeDetReq = FeDetReq


    Dim msg As String
    Dim resp As New CAESolicitar


    msg = ApiConnect("wsfe/FECAESolicitar", POST_, False, AfipHelper.CrearXMLFromCaeSolicitar(FeCAEReq))
    Dim c As New XmlSerializer


    'If Not c.Deserialize(resp, msg) Then GoTo err1

    Dim m() As String
    Dim m2() As String
    Dim inty
    Dim intx

    m = Split(msg, "_", , vbBinaryCompare)


    For intx = 0 To UBound(m)
        m2 = Split(m(intx), "-")

        If m2(0) = "ESTADO" Then resp.Resultado = m2(1)
        If resp.Resultado = "APROBADO" Then

            If m2(0) = "CAEVTO" Then resp.CAEVencimiento = m2(1)
            If m2(0) = "CAE" Then resp.CAE = m2(1)
            If m2(0) = "CBTE" Then resp.Comprobante = m2(1)
            If m2(0) = "FCHEMISION" Then resp.FechaEmision = m2(1)
            If m2(0) = "FCHPROC" Then resp.FechaProceso = m2(1)
            If m2(0) = "OBS" Then resp.observaciones = m2(1)
        ElseIf resp.Resultado = "RECHAZADO" Then
            resp.Errores = m2(0) & " - " & m2(1)

        End If



    Next intx



    Set CreateFECaeSolicitarRequest = resp



    Exit Function
err1:
    Err.Raise 887766, Err.Source, Err.Description
End Function


Public Function CreateFECaeSolicitarRequestEXP(F As Factura) As CAESolicitar
    On Error GoTo err1

    Dim body As String

    'Set f = DAOFactura.FindById(idFactura)
    Dim id

    'Desactivada el 17.07.20 -dnemer
    id = CLng(ERPHelper.GetUltimoAutorizadoEXP(F.Tipo.PuntoVenta.PuntoVenta, F.Tipo.id, F.esCredito))

    'Reestablecida esta linea de codigo el 17.07.20 -dnemer
    'id = CLng(ERPHelper.GetUltimoAutorizado(F.Tipo.PuntoVenta.PuntoVenta, F.Tipo.id))


    F.numero = id + 1

    Dim req As New FECAEDetRequest
    
    req.Concepto = F.ConceptoIncluir
    req.DocTipo = F.Cliente.TipoDocumento
    req.DocNro = F.Cliente.Cuit
    req.CbteDesde = F.numero
    req.CbteHasta = F.numero

    req.CbteFch = Format(F.FechaEmision, "yyyymmdd")
    req.ImpTotal = funciones.FormatearDecimales(F.TotalEstatico.total, 2)
    'req.MonCotiz = F.CambioAPatron

    'req.ImpTotal = funciones.FormatearDecimales(F.TotalEstatico.Total * F.CambioAPatron, 2)
    req.ImpTotConc = "0"    'no gavado+excento + gravado + iva + tributo
    req.ImpNeto = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado, 2)
    'req.ImpNeto = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado * F.CambioAPatron, 2)
    req.ImpOpEx = funciones.FormatearDecimales(F.TotalEstatico.TotalExento, 2)
    'req.ImpOpEx = funciones.FormatearDecimales(F.TotalEstatico.TotalExento * F.CambioAPatron, 2)

    'Reestablecida el 17.07.20 -dnemer
    'req.FchServDesde = ""    ' obligatorio para concepto de tipo 3 y 2
    'req.FchServHasta = ""    ' obligatorio para concepto de tipo 3 y 2

    'Desactivada el 17.07.20 -dnemer

    If F.ConceptoIncluir = ConceptoIncluir.ConceptoProductoServicio Or F.ConceptoIncluir = ConceptoIncluir.ConceptoServicio Then
        req.FchServDesde = Format(F.FechaVtoDesde, "yyyymmdd")   ' obligatorio para concepto de tipo 3 y 2
        req.FchServHasta = Format(F.FechaVtoHasta, "yyyymmdd")     ' obligatorio para concepto de tipo 3 y 2
    End If

    req.ImpTrib = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB, 2)
    'req.ImpTrib = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB * F.CambioAPatron, 2)

    req.ImpIVA = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo, 2)
    'req.ImpIVA = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo * F.CambioAPatron, 2)

    req.MonId = F.moneda.id
    'req.MonCotiz = F.Moneda.Cambio
    req.MonCotiz = F.CambioAPatron



    'Reestablecida el 17.07.20 -dnemer
    'req.FchVtoPago = ""    'obligatorio para concepto 2 y 3

    'Desactivada el 17.07.20 -dnemer

    If Not F.esCredito Then

        If F.ConceptoIncluir = ConceptoProductoServicio Or F.ConceptoIncluir = ConceptoProductoServicio Then
            req.FchVtoPago = Format(F.fechaPago, "yyyymmdd")       'obligatorio para concepto 2 y 3
        End If

    Else
        ' Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), el campo
        '�fecha de vencimiento para el pago� <FchVtoPago> no debe informarse si NO es
        'Factura de Cr�dito. En el caso de ser D�bito o Cr�dito, solo puede informarse si es de Anulaci�n.


        If F.TipoDocumento = tipoDocumentoContable.Factura Then
            req.FchVtoPago = Format(F.fechaPago, "yyyymmdd")
        Else
            If F.AnulacionAFIP = "S" Then
                req.FchVtoPago = Format(F.fechaPago, "yyyymmdd")
            End If
        End If
    End If

    'Desactivada el 17.07.20 -dnemer


    'NB: afip rechaza nc y nd que no sean mi pyme sin tener cbte asociado (err 10197 afip)
    'If F.esCredito And (F.TipoDocumento = tipoDocumentoContable.notaCredito Or F.TipoDocumento = tipoDocumentoContable.notaDebito) Then

    If F.TipoDocumento = tipoDocumentoContable.notaCredito Or F.TipoDocumento = tipoDocumentoContable.notaDebito Then
        Dim ftmp As Factura
        Set ftmp = DAOFactura.FindById(F.Cancelada)
        '23-8 NB: no puedo informar un comprobnte asociado que no est� previamente informado. (en caso que sea cr�dito o d�bido mipyme)

        If IsSomething(ftmp) Then
            If Not ftmp.AprobadaAFIP Then
                Err.Raise 22211, "El comprobante asociado debe estar informado previamente a AFIP"
            End If
        End If
        If IsSomething(ftmp) Then
            Dim cbt As CbteAsoc
            Set cbt = New CbteAsoc

            '       cbt.NRO = f.Cancelada

            cbt.NRO = ftmp.numero
            If ftmp.esCredito Then
                cbt.esCredito = "true"
            Else
                cbt.esCredito = "false"
            End If
            cbt.PtoVta = ftmp.Tipo.PuntoVenta.PuntoVenta
            cbt.Tipo = ftmp.Tipo.id
            cbt.CbteFch = Format(ftmp.FechaEmision, "yyyymmdd")
            cbt.Cuit = "30657604972"
            req.CbtesAsoc.Add cbt
        End If
    End If

    'Dim cbt As CbteAsoc
    'Set cbt = New CbteAsoc

    'cbt.nro = "2"
    'cbt.PtoVta = "2"
    'cbt.Tipo = "2"
    'req.CbtesAsoc.Add cbt

    ' Set cbt = New CbteAsoc
    '
    'cbt.nro = "1"
    'cbt.PtoVta = "1"
    'cbt.Tipo = "1"

    '0req.CbtesAsoc.Add cbt
    '


    Dim trib As Tributo
    Set trib = New Tributo

    Dim P As New clsPercepciones
    Set P = DAOPercepciones.GetById(5)

    trib.Alic = funciones.FormatearDecimales(((F.AlicuotaPercepcionesIIBB - 1) * 100), 2)
    'trib.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado * F.CambioAPatron, 2)    '"400" 'revisar con tulio,2)

    'bug #4
    trib.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado, 2)    '
    trib.Desc = P.Percepcion
    trib.idTributoCambiar = "1"


    'trib.Importe = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB * F.CambioAPatron, 2)

    'bug #4
    trib.importe = funciones.FormatearDecimales(F.TotalEstatico.TotalPercepcionesIB, 2)
    If F.TotalEstatico.TotalPercepcionesIB > 0 Then
        req.Tributos.Add trib
    Else
        Set req.Tributos = Nothing

    End If


    Dim Iva As New AlicIva
    Iva.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado, 2)
    'Iva.BaseImp = funciones.FormatearDecimales(F.TotalEstatico.TotalNetoGravado * F.CambioAPatron, 2)
    Iva.idAlicIvaCambiar = F.AlicuotaAplicada

    Iva.importe = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo, 2)
    'Iva.Importe = funciones.FormatearDecimales(F.TotalEstatico.TotalIVADiscrimandoONo * F.CambioAPatron, 2)    'mapear en erphelper por F.TipoIVA.idIVA,2)

    If F.TotalEstatico.TotalIVADiscrimandoONo > 0 Then
        req.Iva.Add Iva
    Else
        Set req.Iva = Nothing

    End If


    If F.esCredito Then
        '4-4-21 #218
        'Si el tipo de comprobante que est� autorizando es Factura del tipo MiPyMEs (201, 206, 211), es
        'obligatorio informar <Opcionales> con id = 27. Los valores posiblesson SCA o ADC


        'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), es obligatorio informar al menos uno de los sig. c�digos
        '2101, 22.

        If F.TipoDocumento = tipoDocumentoContable.Factura Then

            '4-4-21 #218
            'Si el tipo de comprobante que est� autorizando es Factura del tipo MiPyMEs (201, 206, 211), es
            'obligatorio informar <Opcionales> con id = 27. Los valores posiblesson SCA o ADC

            Dim op27 As New Opcional
            op27.idOpcionalCambiar = "27"

            If F.Opcional27 = 1 Then
                op27.Valor = "SCA"
            ElseIf F.Opcional27 = 2 Then
                op27.Valor = "ADC"
            End If

            req.Opcionales.Add op27


            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), informa opcionales, el valor correcto para el c�digo
            '2101 es un CBU num�rico de 22 caracteres.

            'Si el tipo de comprobante que est� autorizando es Factura (201, 206, 211) del tipo MiPyMEs (FCE), informa opcionales, es
            'obligatorio informar CBU.

            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), Debito (202, 207, 212) o Cr�dito (203, 208, 213) No
            'no informar CBU y ALIAS.
            If Len(F.CBU) <> 22 Then
                Err.Raise 222101, "El cbu debe tener 22 caracteres"
            End If

            Dim op As New Opcional
            op.idOpcionalCambiar = "2101"
            op.Valor = F.CBU
            req.Opcionales.Add op
        End If
        
        If F.TipoDocumento <> tipoDocumentoContable.Factura Then
            ' Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), informa opcionales, el valor correcto para el c�digo 22
            ' es �S� o �N�: S = Es de Anulaci�n N = No es de Anulaci�n

            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), Factura (201, 206, 211), no informar C�digo de Anulaci�n

            'Si el tipo de comprobante que est� autorizando es MiPyMEs (FCE), Debito (202, 207, 212) o Cr�dito (203, 208, 213) informar
            'C�digo de Anulaci�n

            If Len(F.AnulacionAFIP) <> 1 Then
                Err.Raise 22221, "El campo anulacion afip debe informarse"
            End If


            If F.AnulacionAFIP <> "N" And F.AnulacionAFIP <> "S" Then
                Err.Raise 22222, "El campo anulacion afip debe ser S o N"
            End If

            Set op = New Opcional
            op.idOpcionalCambiar = "22"
            op.Valor = F.AnulacionAFIP
            req.Opcionales.Add op
        End If


    End If

    Dim FeDetReq As New FeDetReq

    Set FeDetReq.FECAEDetRequest = req
    Dim FeCabReq As New FeCabReq
    FeCabReq.CantReg = "1"
    FeCabReq.CbteTipo = F.Tipo.id
    
    If F.esCredito Then
        FeCabReq.esCredito = "true"
    Else
        FeCabReq.esCredito = "false"
    End If




    FeCabReq.PtoVta = F.Tipo.PuntoVenta.id
    Dim FeCAEReq As New FeCAEReq

    Set FeCAEReq.FeCabReq = FeCabReq
    Set FeCAEReq.FeDetReq = FeDetReq


    Dim msg As String
    Dim resp As New CAESolicitar


    msg = ApiConnectEXP("wsfe/FECAESolicitar", POST_, False, AfipHelper.CrearXMLFromCaeSolicitarEXP(FeCAEReq))
    Dim c As New XmlSerializer


    'If Not c.Deserialize(resp, msg) Then GoTo err1

    Dim m() As String
    Dim m2() As String
    Dim inty
    Dim intx

    m = Split(msg, "_", , vbBinaryCompare)


    For intx = 0 To UBound(m)
        m2 = Split(m(intx), "-")

        If m2(0) = "ESTADO" Then resp.Resultado = m2(1)
        If resp.Resultado = "APROBADO" Then

            If m2(0) = "CAEVTO" Then resp.CAEVencimiento = m2(1)
            If m2(0) = "CAE" Then resp.CAE = m2(1)
            If m2(0) = "CBTE" Then resp.Comprobante = m2(1)
            If m2(0) = "FCHEMISION" Then resp.FechaEmision = m2(1)
            If m2(0) = "FCHPROC" Then resp.FechaProceso = m2(1)
            If m2(0) = "OBS" Then resp.observaciones = m2(1)
        ElseIf resp.Resultado = "RECHAZADO" Then
            resp.Errores = m2(0) & " - " & m2(1)

        End If



    Next intx



    Set CreateFECaeSolicitarRequestEXP = resp



    Exit Function
err1:
    Err.Raise 887766, Err.Source, Err.Description
End Function

Public Function SendMail(asunto As String, mensaje As String, destino As String, Optional file As String, Optional Class As Object) As String
    On Error Resume Next
    Dim sUrl As String
    Dim verb As verbo
    Dim withfile As Boolean
    Dim baBuffer() As Byte
    srv = LeerIni(App.path & "\config.ini", "Configurar", "ERPHelperAddress", vbNullString)
    Dim de As String
    de = funciones.GetUserObj().Empleado.email

    Dim de_firma As String
    de_firma = funciones.GetUserObj().Empleado.NombreCompleto
    Dim async As Boolean
    async = True
    Dim verbo As String
    verbo = "POST"
    Dim xmlhttp As New MSXML2.xmlhttp
    verb = POST_
    If verb = GET_ Then verbo = "GET"

    If LenB(file) > 3 Then withfile = True Else withfile = False

    If Not Class Is Nothing Then
        Dim body As String
        Dim c As New XmlSerializer

        body = "<?xml version='1.0' encoding='utf-8'?>" & c.Serialize(Class)  'XMLProperties(Class, False)
    End If

    Dim sPostData As String

    If withfile Then

        Const STR_BOUNDARY As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
        Dim nFile As Integer

        '--- read file
        nFile = FreeFile
        Open file For Binary Access Read As nFile
        If LOF(nFile) > 0 Then
            ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
            Get nFile, , baBuffer
            sPostData = StrConv(baBuffer, vbUnicode)

        End If
        Close nFile
        'erphelper.SendMail "sdfdf","sdfdf","nbattaglia@signoplast.com.ar","C:\file.txt"
        ''--- prepare body
        sPostData = "--" & STR_BOUNDARY & vbCrLf & _
                    "Content-Disposition: form-data; name=""uploadfile""; filename=""" & Mid$(file, InStrRev(file, "\") + 1) & """" & vbCrLf & _
                    "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
                    sPostData & vbCrLf & _
                    "--" & STR_BOUNDARY & "--"

        sPostData = StrConv(baBuffer, vbUnicode)
        Dim filename As String
        'sPostData = ReadFile(file)
        filename = Mid$(file, InStrRev(file, "\") + 1)
        sUrl = "mailsender/sendfile?para=" & destino & "&asunto=" & asunto & "&msg=" & mensaje & "&filename=" & filename & " &de=" & de & " &de_firma=" & de_firma
        xmlhttp.Open verbo, srv + sUrl, async
        ' xmlhttp.setRequestHeader "Content-Type", "application/octet-stream; boundary=" & STR_BOUNDARY
        'xmlhttp.setRequestHeader "User-Agent", "Alalala"
        xmlhttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        xmlhttp.send pvToByteArray(sPostData)
    Else
        sUrl = "mailsender/send?para=" & destino & "&asunto=" & asunto & "&msg=" & mensaje & " &de=" & de & " &de_firma=" & de_firma
        xmlhttp.Open verbo, srv + sUrl, async
        xmlhttp.setRequestHeader "Content-Type", "text/plain"
        xmlhttp.send
    End If




    Dim response As String
    response = xmlhttp.responseText


    If LenB(xmlhttp.responseText) > 0 Then
        MsgBox xmlhttp.responseText
    End If


End Function


Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function


