Attribute VB_Name = "DAOProduccionAvanceSimpleDTO"
Option Explicit

Public Type AvanceSimpleDTO
    CantRecibida As Double
    CantFabricada As Double
    CantScrap As Double
    FechaInicio As Variant
    HoraInicio As Variant
    FechaFin As Variant
    HoraFin As Variant
    Recibio As Long
    Almacen As Long
    SiguienteProceso As String
    Observaciones As String
End Type
