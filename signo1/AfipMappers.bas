Attribute VB_Name = "AfipMappers"
Option Explicit


Function MapTiposMonedas(idMoneda As Long) As String

Select Case idMoneda
    Case 0: MapTiposMonedas = AfipEnums.GetTipoMoneda(Afip_TiposMoneda.Afip_TipoMoneda_PES)
End Select



End Function


