Attribute VB_Name = "AfipHelper"
'---------------------------------------------------------------------------------------
' Module    : AfipHelper
' Author    : nicolasba
' Date      : 24/05/2024
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit



Public Function CrearXMLFromCaeSolicitar(C As FeCAEReq) As String
    Dim r As String
    r = "<FeCAEReq>"
    r = r & "<FeCabReq>"

    'Desactivado el 17.07.20 - dnemer
    r = r & "<EsCredito>" & C.FeCabReq.esCredito & "</EsCredito>"

    r = r & "<CantReg>" & C.FeCabReq.CantReg & "</CantReg>"
    r = r & "<PtoVta>" & C.FeCabReq.PtoVta & "</PtoVta>"
    r = r & "<CbteTipo>" & C.FeCabReq.CbteTipo & "</CbteTipo>"

    
    r = r & "</FeCabReq>"
    r = r & "<FeDetReq>"
    r = r & "<FECAEDetRequest>"
    r = r & "<Concepto>" & C.FeDetReq.FECAEDetRequest.Concepto & "</Concepto>"
    r = r & "<DocTipo>" & C.FeDetReq.FECAEDetRequest.DocTipo & "</DocTipo>"
    r = r & "<DocNro>" & C.FeDetReq.FECAEDetRequest.DocNro & "</DocNro>"
    r = r & "<CbteDesde>" & C.FeDetReq.FECAEDetRequest.CbteDesde & "</CbteDesde>"
    r = r & "<CbteHasta>" & C.FeDetReq.FECAEDetRequest.CbteHasta & "</CbteHasta>"
    r = r & "<CbteFch>" & C.FeDetReq.FECAEDetRequest.CbteFch & "</CbteFch>"
    r = r & "<ImpTotal>" & C.FeDetReq.FECAEDetRequest.ImpTotal & "</ImpTotal>"
    r = r & "<ImpTotConc>" & C.FeDetReq.FECAEDetRequest.ImpTotConc & "</ImpTotConc>"
    r = r & "<ImpNeto>" & C.FeDetReq.FECAEDetRequest.ImpNeto & "</ImpNeto>"
    r = r & "<ImpTrib>" & C.FeDetReq.FECAEDetRequest.ImpTrib & "</ImpTrib>"
    r = r & "<ImpOpEx>" & C.FeDetReq.FECAEDetRequest.ImpOpEx & "</ImpOpEx>"
    r = r & "<ImpIVA>" & C.FeDetReq.FECAEDetRequest.ImpIVA & "</ImpIVA>"
    If LenB(C.FeDetReq.FECAEDetRequest.FchServDesde) > 0 Then
        r = r & "<FchServDesde>" & C.FeDetReq.FECAEDetRequest.FchServDesde & "</FchServDesde>"
    End If
    If LenB(C.FeDetReq.FECAEDetRequest.FchServHasta) > 0 Then
        r = r & "<FchServHasta>" & C.FeDetReq.FECAEDetRequest.FchServHasta & "</FchServHasta>"
    End If

    If LenB(C.FeDetReq.FECAEDetRequest.FchVtoPago) > 0 Then

        r = r & "<FchVtoPago>" & C.FeDetReq.FECAEDetRequest.FchVtoPago & "</FchVtoPago>"

    End If
    r = r & "<MonId>" & C.FeDetReq.FECAEDetRequest.MonId & "</MonId>"
    r = r & "<MonCotiz>" & C.FeDetReq.FECAEDetRequest.MonCotiz & "</MonCotiz>"

    If C.FeDetReq.FECAEDetRequest.CbtesAsoc.count > 0 Then
        Dim ca As CbteAsoc
        r = r & "<CbtesAsoc>"
        For Each ca In C.FeDetReq.FECAEDetRequest.CbtesAsoc
            r = r & "<CbteAsoc>"
            Debug.Print ("Comprobante Asociado Tipo FC : " & ca.Tipo)
            'Desactivado el 17.07.20 - dnemer
            r = r & "<EsCredito>" & ca.esCredito & "</EsCredito>"
            r = r & "<CbteFch>" & ca.CbteFch & "</CbteFch>"
            r = r & "<Tipo>" & ca.Tipo & "</Tipo>"
            r = r & "<PtoVta>" & ca.PtoVta & "</PtoVta>"
            r = r & "<Nro>" & ca.NRO & "</Nro>"

            'Desactivado el 17.07.20 - dnemer
            r = r & "<Cuit>" & ca.Cuit & "</Cuit>"

            r = r & "</CbteAsoc>"
        Next
        r = r & "</CbtesAsoc>"
    End If

    If C.FeDetReq.FECAEDetRequest.Tributos.count > 0 Then
        Dim T As Tributo
        r = r & "<Tributos>"

        For Each T In C.FeDetReq.FECAEDetRequest.Tributos

            r = r & "<Tributo>"
            r = r & "<Id>" & T.idTributoCambiar & "</Id>"
            r = r & "<Desc>" & T.Desc & "</Desc>"
            r = r & "<Alic>" & T.Alic & "</Alic>"
            r = r & "<Importe>" & T.importe & "</Importe>"
            r = r & "<BaseImp>" & T.BaseImp & "</BaseImp>"
            r = r & "</Tributo>"
        Next
        r = r & "</Tributos>"
    End If

    If C.FeDetReq.FECAEDetRequest.Iva.count > 0 Then
        Dim i As AlicIva
        r = r & "<Iva>"
        For Each i In C.FeDetReq.FECAEDetRequest.Iva
            r = r & "<AlicIva>"
            r = r & "<Id>" & i.idAlicIvaCambiar & "</Id>"
            r = r & "<BaseImp>" & i.BaseImp & "</BaseImp>"
            r = r & "<Importe>" & i.importe & "</Importe>"
            r = r & "</AlicIva>"
        Next
        r = r & "</Iva>"
    End If
    If C.FeDetReq.FECAEDetRequest.Opcionales.count > 0 Then
        Dim o As Opcional
        r = r & "<Opcionales>"

        Dim ox As Opcional
        For Each ox In C.FeDetReq.FECAEDetRequest.Opcionales
            r = r & "<Opcional>"
            r = r & "<Id>" & ox.idOpcionalCambiar & "</Id>"
            r = r & "<Valor>" & ox.Valor & "</Valor>"
            r = r & "</Opcional>"
        Next

        '        For Each o In c.FeDetReq.FECAEDetRequest.Iva
        '            r = r & "<Opcional>"
        '            r = r & "<Id>" & o.idOpcionalCambiar & "</Id>"
        '            r = r & "<Valor>" & o.Valor & "</Valor>"
        '            r = r & "<Opcional>"
        '        Next
        r = r & "</Opcionales>"
    End If


    r = r & "</FECAEDetRequest>"
    r = r & "</FeDetReq>"
    r = r & "</FeCAEReq>"

    CrearXMLFromCaeSolicitar = r
End Function


Public Function CrearXMLFromCaeSolicitarEXP(C As FeCAEReq) As String
    Dim r As String
    r = "<FeCAEReq>"
    r = r & "<FeCabReq>"

    'Desactivado el 17.07.20 - dnemer
    r = r & "<EsCredito>" & C.FeCabReq.esCredito & "</EsCredito>"

    r = r & "<CantReg>" & C.FeCabReq.CantReg & "</CantReg>"
    r = r & "<PtoVta>" & C.FeCabReq.PtoVta & "</PtoVta>"
    r = r & "<CbteTipo>" & C.FeCabReq.CbteTipo & "</CbteTipo>"
    'Debug.Print ("Comprobante Tipo ND : " & C.FeCabReq.CbteTipo)
    
    r = r & "</FeCabReq>"
    r = r & "<FeDetReq>"
    r = r & "<FECAEDetRequest>"
    r = r & "<Concepto>" & C.FeDetReq.FECAEDetRequest.Concepto & "</Concepto>"
    r = r & "<DocTipo>" & C.FeDetReq.FECAEDetRequest.DocTipo & "</DocTipo>"
    r = r & "<DocNro>" & C.FeDetReq.FECAEDetRequest.DocNro & "</DocNro>"
    r = r & "<CbteDesde>" & C.FeDetReq.FECAEDetRequest.CbteDesde & "</CbteDesde>"
    r = r & "<CbteHasta>" & C.FeDetReq.FECAEDetRequest.CbteHasta & "</CbteHasta>"
    r = r & "<CbteFch>" & C.FeDetReq.FECAEDetRequest.CbteFch & "</CbteFch>"
    r = r & "<ImpTotal>" & C.FeDetReq.FECAEDetRequest.ImpTotal & "</ImpTotal>"
    r = r & "<ImpTotConc>" & C.FeDetReq.FECAEDetRequest.ImpTotConc & "</ImpTotConc>"
    r = r & "<ImpNeto>" & C.FeDetReq.FECAEDetRequest.ImpNeto & "</ImpNeto>"
    r = r & "<ImpTrib>" & C.FeDetReq.FECAEDetRequest.ImpTrib & "</ImpTrib>"
    r = r & "<ImpOpEx>" & C.FeDetReq.FECAEDetRequest.ImpOpEx & "</ImpOpEx>"
    r = r & "<ImpIVA>" & C.FeDetReq.FECAEDetRequest.ImpIVA & "</ImpIVA>"
    If LenB(C.FeDetReq.FECAEDetRequest.FchServDesde) > 0 Then
        r = r & "<FchServDesde>" & C.FeDetReq.FECAEDetRequest.FchServDesde & "</FchServDesde>"
    End If
    If LenB(C.FeDetReq.FECAEDetRequest.FchServHasta) > 0 Then
        r = r & "<FchServHasta>" & C.FeDetReq.FECAEDetRequest.FchServHasta & "</FchServHasta>"
    End If

    If LenB(C.FeDetReq.FECAEDetRequest.FchVtoPago) > 0 Then

        r = r & "<FchVtoPago>" & C.FeDetReq.FECAEDetRequest.FchVtoPago & "</FchVtoPago>"

    End If
    r = r & "<MonId>" & C.FeDetReq.FECAEDetRequest.MonId & "</MonId>"
    r = r & "<MonCotiz>" & C.FeDetReq.FECAEDetRequest.MonCotiz & "</MonCotiz>"

    If C.FeDetReq.FECAEDetRequest.CbtesAsoc.count > 0 Then
        Dim ca As CbteAsoc
        r = r & "<CbtesAsoc>"
        For Each ca In C.FeDetReq.FECAEDetRequest.CbtesAsoc
            r = r & "<CbteAsoc>"
            'Debug.Print ("Comprobante Asociado Tipo FC : " & ca.Tipo)
            'Desactivado el 17.07.20 - dnemer
            r = r & "<EsCredito>" & ca.esCredito & "</EsCredito>"
            r = r & "<CbteFch>" & ca.CbteFch & "</CbteFch>"
            r = r & "<Tipo>" & ca.Tipo & "</Tipo>"
            r = r & "<PtoVta>" & ca.PtoVta & "</PtoVta>"
            r = r & "<Nro>" & ca.NRO & "</Nro>"

            'Desactivado el 17.07.20 - dnemer
            r = r & "<Cuit>" & ca.Cuit & "</Cuit>"

            r = r & "</CbteAsoc>"
        Next
        r = r & "</CbtesAsoc>"
    End If

    If C.FeDetReq.FECAEDetRequest.Tributos.count > 0 Then
        Dim T As Tributo
        r = r & "<Tributos>"

        For Each T In C.FeDetReq.FECAEDetRequest.Tributos

            r = r & "<Tributo>"
            r = r & "<Id>" & T.idTributoCambiar & "</Id>"
            r = r & "<Desc>" & T.Desc & "</Desc>"
            r = r & "<Alic>" & T.Alic & "</Alic>"
            r = r & "<Importe>" & T.importe & "</Importe>"
            r = r & "<BaseImp>" & T.BaseImp & "</BaseImp>"
            r = r & "</Tributo>"
        Next
        r = r & "</Tributos>"
    End If

    If C.FeDetReq.FECAEDetRequest.Iva.count > 0 Then
        Dim i As AlicIva
        r = r & "<Iva>"
        For Each i In C.FeDetReq.FECAEDetRequest.Iva
            r = r & "<AlicIva>"
            r = r & "<Id>" & i.idAlicIvaCambiar & "</Id>"
            r = r & "<BaseImp>" & i.BaseImp & "</BaseImp>"
            r = r & "<Importe>" & i.importe & "</Importe>"
            r = r & "</AlicIva>"
        Next
        r = r & "</Iva>"
    End If
    If C.FeDetReq.FECAEDetRequest.Opcionales.count > 0 Then
        Dim o As Opcional
        r = r & "<Opcionales>"

        Dim ox As Opcional
        For Each ox In C.FeDetReq.FECAEDetRequest.Opcionales
            r = r & "<Opcional>"
            r = r & "<Id>" & ox.idOpcionalCambiar & "</Id>"
            r = r & "<Valor>" & ox.Valor & "</Valor>"
            r = r & "</Opcional>"
        Next

        '        For Each o In c.FeDetReq.FECAEDetRequest.Iva
        '            r = r & "<Opcional>"
        '            r = r & "<Id>" & o.idOpcionalCambiar & "</Id>"
        '            r = r & "<Valor>" & o.Valor & "</Valor>"
        '            r = r & "<Opcional>"
        '        Next
        r = r & "</Opcionales>"
    End If


    r = r & "</FECAEDetRequest>"
    r = r & "</FeDetReq>"
    r = r & "</FeCAEReq>"

    CrearXMLFromCaeSolicitarEXP = r
End Function
