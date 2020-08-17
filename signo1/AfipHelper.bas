Attribute VB_Name = "AfipHelper"
Option Explicit

Public Function CrearXMLFromCaeSolicitar(c As FeCAEReq) As String
    Dim r As String
    r = "<FeCAEReq>"
    r = r & "<FeCabReq>"
    
    'Desactivado el 17.07.20 - dnemer
     r = r & "<EsCredito>" & c.FeCabReq.esCredito & "</EsCredito>"
     
    r = r & "<CantReg>" & c.FeCabReq.CantReg & "</CantReg>"
    r = r & "<PtoVta>" & c.FeCabReq.PtoVta & "</PtoVta>"
    r = r & "<CbteTipo>" & c.FeCabReq.CbteTipo & "</CbteTipo>"
    r = r & "</FeCabReq>"
    r = r & "<FeDetReq>"
    r = r & "<FECAEDetRequest>"
    r = r & "<Concepto>" & c.FeDetReq.FECAEDetRequest.Concepto & "</Concepto>"
    r = r & "<DocTipo>" & c.FeDetReq.FECAEDetRequest.DocTipo & "</DocTipo>"
    r = r & "<DocNro>" & c.FeDetReq.FECAEDetRequest.DocNro & "</DocNro>"
    r = r & "<CbteDesde>" & c.FeDetReq.FECAEDetRequest.CbteDesde & "</CbteDesde>"
    r = r & "<CbteHasta>" & c.FeDetReq.FECAEDetRequest.CbteHasta & "</CbteHasta>"
    r = r & "<CbteFch>" & c.FeDetReq.FECAEDetRequest.CbteFch & "</CbteFch>"
    r = r & "<ImpTotal>" & c.FeDetReq.FECAEDetRequest.ImpTotal & "</ImpTotal>"
    r = r & "<ImpTotConc>" & c.FeDetReq.FECAEDetRequest.ImpTotConc & "</ImpTotConc>"
    r = r & "<ImpNeto>" & c.FeDetReq.FECAEDetRequest.ImpNeto & "</ImpNeto>"
    r = r & "<ImpTrib>" & c.FeDetReq.FECAEDetRequest.ImpTrib & "</ImpTrib>"
    r = r & "<ImpOpEx>" & c.FeDetReq.FECAEDetRequest.ImpOpEx & "</ImpOpEx>"
    r = r & "<ImpIVA>" & c.FeDetReq.FECAEDetRequest.ImpIVA & "</ImpIVA>"
    If LenB(c.FeDetReq.FECAEDetRequest.FchServDesde) > 0 Then
        r = r & "<FchServDesde>" & c.FeDetReq.FECAEDetRequest.FchServDesde & "</FchServDesde>"
    End If
    If LenB(c.FeDetReq.FECAEDetRequest.FchServHasta) > 0 Then
        r = r & "<FchServHasta>" & c.FeDetReq.FECAEDetRequest.FchServHasta & "</FchServHasta>"
    End If
    If LenB(c.FeDetReq.FECAEDetRequest.FchVtoPago) > 0 Then
        r = r & "<FchVtoPago>" & c.FeDetReq.FECAEDetRequest.FchVtoPago & "</FchVtoPago>"
    End If
    r = r & "<MonId>" & c.FeDetReq.FECAEDetRequest.MonId & "</MonId>"
    r = r & "<MonCotiz>" & c.FeDetReq.FECAEDetRequest.MonCotiz & "</MonCotiz>"

    If c.FeDetReq.FECAEDetRequest.CbtesAsoc.count > 0 Then
        Dim ca As CbteAsoc
        r = r & "<CbtesAsoc>"
        For Each ca In c.FeDetReq.FECAEDetRequest.CbtesAsoc
            r = r & "<CbteAsoc>"
            
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

    If c.FeDetReq.FECAEDetRequest.Tributos.count > 0 Then
        Dim T As Tributo
        r = r & "<Tributos>"

        For Each T In c.FeDetReq.FECAEDetRequest.Tributos

            r = r & "<Tributo>"
            r = r & "<Id>" & T.idTributoCambiar & "</Id>"
            r = r & "<Desc>" & T.Desc & "</Desc>"
            r = r & "<Alic>" & T.Alic & "</Alic>"
            r = r & "<Importe>" & T.Importe & "</Importe>"
            r = r & "<BaseImp>" & T.BaseImp & "</BaseImp>"
            r = r & "</Tributo>"
        Next
        r = r & "</Tributos>"
    End If

    If c.FeDetReq.FECAEDetRequest.Iva.count > 0 Then
        Dim i As AlicIva
        r = r & "<Iva>"
        For Each i In c.FeDetReq.FECAEDetRequest.Iva
            r = r & "<AlicIva>"
            r = r & "<Id>" & i.idAlicIvaCambiar & "</Id>"
            r = r & "<BaseImp>" & i.BaseImp & "</BaseImp>"
            r = r & "<Importe>" & i.Importe & "</Importe>"
            r = r & "</AlicIva>"
        Next
        r = r & "</Iva>"
    End If
    If c.FeDetReq.FECAEDetRequest.Opcionales.count > 0 Then
        Dim o As Opcional
        r = r & "<Opcionales>"
        
        Dim ox As Opcional
         For Each ox In c.FeDetReq.FECAEDetRequest.Opcionales
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
