Attribute VB_Name = "AfipEnums"
Option Explicit
Dim arrTiposMoneda(1) As String

Public Enum Afip_TiposMoneda
    Afip_TipoMoneda_PES = 1
End Enum


Public Function GetTipoMoneda(indice)
    GetTipoMoneda = arrTiposMoneda(indice)
End Function

