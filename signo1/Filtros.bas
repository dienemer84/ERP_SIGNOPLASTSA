Attribute VB_Name = "Filtros"
Option Explicit

Public vFiltroBusquedaMaterial As New FilterDTObucarMateriaPrima

Public Function FiltroBusquedaMaterial(nombre As String, id_rubro As Integer)
vFiltroBusquedaMaterial.nombre = nombre
vFiltroBusquedaMaterial.rubro = id_rubro
End Function
