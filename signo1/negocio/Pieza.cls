VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pieza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public nombre As String
Public Cliente As clsCliente
Public CantidadStock As Long
Public Activa As Boolean
Public EsConjunto As Boolean
Public UbicacionStock As String
Public Precio As Double
Public FechaPrecio As Date
Public MonedaPrecio As clsMoneda
Public PiezasHijas As New Collection
Public Cantidad As Double     'cantidad de piezas que se necesitan para el conjunto
Public YaFabricada As Boolean
Public desarrollosManoObra As New Collection
Public DesarrollosMaterial As New Collection
Public Revision As String
Public IdPiezaUltimaRevision As Long
Public Complejidad As TipoComplejidad



Public Function MatreialValido() As Boolean
    Dim t As Boolean: t = False
    Dim m As DesarrolloMaterial
    For Each m In DesarrollosMaterial
        If Not m.material.Aprobado Then
            t = False
            Exit For
        End If
    Next m
End Function


Public Property Get UnidadMedida() As String
    UnidadMedida = IIf(Me.EsConjunto, "Conjunto", "Pieza")
End Property


Public Property Get Kilage() As Double
    Dim dmat As DesarrolloMaterial
    Dim tot As Double: tot = 0
    For Each dmat In Me.DesarrollosMaterial
        tot = tot + dmat.Kg
    Next
    Kilage = tot
End Property

Public Function LocatePiezaInPiezasHijas(piezaId As Long) As Pieza
    If Me.id = piezaId Then
        Set LocatePiezaInPiezasHijas = Me
    Else
        Dim tmpPieza As Pieza
        For Each tmpPieza In Me.PiezasHijas
            Set LocatePiezaInPiezasHijas = tmpPieza.LocatePiezaInPiezasHijas(piezaId)
            If Not LocatePiezaInPiezasHijas Is Nothing Then Exit For
        Next tmpPieza
    End If
End Function

Public Function CalcularValorMDO(ByVal MonedaDestino As clsMoneda, ByRef Fijo As Double, ByRef Cambio As Double, ByRef MDO As Double) As Boolean


    'cambiar el moneda origen a categoeria sueldo!!!!

    Dim clasea As New classAdministracion
    Fijo = 0
    Cambio = 0
    MDO = 0
    On Error GoTo err1
    CalcularValorMDO = True
    Dim desarrollo As DesarrolloManoObra
    For Each desarrollo In desarrollosManoObra
        If desarrollo.Tarea.CantPorProc = 0 Then
            Fijo = Fijo + (desarrollo.Cantidad * desarrollo.Tiempo * clasea.realizaCambio(desarrollo.Tarea.CategoriaSueldo.Valor, desarrollo.Tarea.CategoriaSueldo.Moneda.id, MonedaDestino.id))
        ElseIf desarrollo.Tarea.CantPorProc = -1 Then
            Cambio = Cambio + (clasea.realizaCambio(desarrollo.Tarea.CategoriaSueldo.Valor, desarrollo.Tarea.CategoriaSueldo.Moneda.id, MonedaDestino.id) * desarrollo.Cantidad * desarrollo.Tiempo)
        Else
            MDO = MDO + ((clasea.realizaCambio(desarrollo.Tarea.CategoriaSueldo.Valor, desarrollo.Tarea.CategoriaSueldo.Moneda.id, MonedaDestino.id) * desarrollo.Cantidad * desarrollo.Tiempo) / desarrollo.Tarea.CantPorProc)
        End If
    Next desarrollo
    Exit Function
err1:
    CalcularValorMDO = False
End Function

Public Function TotalCostoMateriales() As Double
    Dim tot As Double
    Dim dmat As DesarrolloMaterial

    For Each dmat In Me.DesarrollosMaterial
        tot = tot + funciones.RedondearDecimales(dmat.CalcularDatosMaterial(0).costo)    'siempre pido pesos porq spy groso
    Next dmat

    TotalCostoMateriales = tot
End Function


Public Function TotalKG() As Double
    Dim tot As Double
    Dim dmat As DesarrolloMaterial

    For Each dmat In Me.DesarrollosMaterial
        tot = tot + dmat.Kg
    Next dmat

    TotalKG = tot
End Function

Public Function TotalM2() As Double
    Dim tot As Double
    Dim dmat As DesarrolloMaterial

    For Each dmat In Me.DesarrollosMaterial
        tot = tot + dmat.m2
    Next dmat

    TotalM2 = tot
End Function

Public Function TotalCostoManoObra() As Double
    Dim tot As Double
    Dim dmdo As DesarrolloManoObra

    For Each dmdo In Me.desarrollosManoObra
        tot = tot + (dmdo.Cantidad * dmdo.Tiempo * dmdo.Tarea.CategoriaSueldo.Valor)
    Next dmdo

    TotalCostoManoObra = tot
End Function

