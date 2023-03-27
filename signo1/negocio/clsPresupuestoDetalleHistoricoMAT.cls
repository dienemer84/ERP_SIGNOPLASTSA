VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PresupuestoDetalleHistoricoMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Id As Long
Public Material As clsMaterial
Public Largo As Double
Public Ancho As Double
Public LargoPieza As Double
Public AnchoPieza As Double
Public Scrap As Double
Public Cantidad As Double
Public Valor As Double
Public moneda As clsMoneda
Private scrap_ As Double


Friend Function CalcularDatosMaterial(ByVal idMonedaDestino As Long) As DatosMaterialDTO
    Dim clasea As New classAdministracion
    Dim Kg As Double
    Dim m2ml As Double
    Dim Precio As Double
    Dim dimension_material As String
    Dim dimension_pieza As String
    Dim costo As Double
    Dim cxh As Double



    Dim dat As DatosMaterialDTO


    scrap_ = (Scrap / 100) + 1

    If Material.unidad = Unidades.Ml_ Then     'ml
        Kg = funciones.RedondearDecimales(Material.PesoXUnidad * (Me.LargoPieza / 1000) * Me.Cantidad)
        m2ml = funciones.RedondearDecimales(Me.LargoPieza / 1000 * Cantidad)
        Precio = Me.Valor * Me.Material.PesoXUnidad
        Precio = clasea.realizaCambio(Precio, moneda.Id, idMonedaDestino)
        Precio = ((Precio * LargoPieza * scrap_ * 2) / 1000) * Cantidad
        Precio = Precio / 2

        dimension_material = Cantidad & "X" & Me.Largo
        dimension_pieza = Cantidad & "X" & LargoPieza

    ElseIf Material.unidad = Unidades.kg_ Then
        Kg = funciones.RedondearDecimales(Material.PesoXUnidad * Cantidad)
        m2ml = 0
        Precio = Valor * Material.PesoXUnidad
        Precio = clasea.realizaCambio(Precio, moneda.Id, idMonedaDestino)
        Precio = (Precio * Cantidad)
        dimension_material = Cantidad
        dimension_pieza = Cantidad
    ElseIf Material.unidad = Unidades.m2_ Then     'm2

        Kg = funciones.RedondearDecimales(Material.PesoXUnidad * ((AnchoPieza * LargoPieza) / 1000000) * Cantidad)
        m2ml = funciones.RedondearDecimales((AnchoPieza * LargoPieza) / 1000000 * Cantidad)
        Precio = Valor * Material.PesoXUnidad
        Precio = clasea.realizaCambio(Precio, moneda.Id, idMonedaDestino)
        Precio = Precio * (Cantidad * AnchoPieza * LargoPieza * scrap_) / 1000000

        dimension_material = Cantidad & "x" & Largo & "x" & Ancho
        dimension_pieza = Cantidad & "x" & LargoPieza & "x" & AnchoPieza

    ElseIf Material.unidad = Unidades.un_ Then     'uni

        m2ml = 0
        Precio = Valor
        Precio = clasea.realizaCambio(Precio, moneda.Id, idMonedaDestino)
        Precio = Precio * Cantidad

        Kg = Cantidad
        dimension_material = Cantidad
        dimension_pieza = Cantidad
    End If
    cxh = funciones.cantxhoja(Ancho, Largo, AnchoPieza, LargoPieza)
    If cxh = 0 Then cxh = 1
    costo = Precio / cxh


    dat.costo = costo
    dat.DimensionMaterial = dimension_material
    dat.DimensionPieza = dimension_pieza
    dat.Kg = Kg
    dat.m2 = m2ml

    CalcularDatosMaterial = dat
End Function
