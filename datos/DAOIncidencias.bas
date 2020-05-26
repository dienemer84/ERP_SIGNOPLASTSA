Attribute VB_Name = "DAOIncidencias"
Option Explicit

Public Enum OrigenIncidencias
    OI_OrdenesTrabajo = 2
    OI_OrdenesTrabajoDetalles = 333
    OI_Presupuestos = 1
    OI_Piezas = 3
    OI_DetallePresupuesto = 33
    OI_Recibos = 4
End Enum


Public Function GetCantidadIncidenciasPorReferencia(ByRef Origen As OrigenIncidencias, Optional ByRef idReferencias As Variant) As Dictionary
    Dim diccionarioRetorno As New Dictionary

    Dim q As String
    q = "SELECT idReferencia, COUNT(0) AS cant FROM Incidencias WHERE origen = " & Origen
    If Not IsMissing(idReferencias) Then
        q = q & " AND idReferencia IN (" & Join(idReferencias, ", ") & ")"
    End If
    q = q & " GROUP BY idReferencia"

    Dim rs As New Recordset
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        diccionarioRetorno.Add rs.Fields("idReferencia").value, rs.Fields("cant").value
        rs.MoveNext
    Wend

    Set GetCantidadIncidenciasPorReferencia = diccionarioRetorno
End Function



