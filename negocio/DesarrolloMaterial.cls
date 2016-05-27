VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesarrolloMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public Pieza As Pieza
Public Scrap As Double
Public Largo As Double
Public Ancho As Double
Public LargoTerm As Double
Public AnchoTerm As Double
Public Material As clsMaterial
Public Cantidad As Double
Public detalle As String


Private iKg As Double
Private iMl As Double
Private iM2 As Double
Private iUn As Double


Public Property Get MedidaMLM2UN() As Double
    If Material.unidad = m2_ Then MedidaMLM2UN = iM2
    If Material.unidad = Ml_ Then MedidaMLM2UN = iMl
    If Material.unidad = litro_ Or Material.unidad = kg_ Then MedidaMLM2UN = iKg
    If Material.unidad = un_ Then MedidaMLM2UN = iUn


End Property
Public Property Get m2() As Double
    datosMaterial
    m2 = iM2
End Property
Public Property Get ML() As Double
    datosMaterial
    ML = iMl
End Property
Public Property Get UN() As Double
    datosMaterial
    UN = iUn
End Property
Public Property Get Kg() As Double
    datosMaterial
    Kg = iKg
End Property
Private Sub datosMaterial()
    If Me.Material.unidad = Ml_ Then
        iKg = Math.Round(Me.Material.PesoXUnidad * (Me.LargoTerm / 1000) * Me.Cantidad, 3)
        iMl = Math.Round(Me.LargoTerm / 1000 * Me.Cantidad, 3)
        iM2 = 0
        iUn = 0
    ElseIf Me.Material.unidad = kg_ Then
        iKg = Math.Round(Me.Material.PesoXUnidad * Me.Cantidad, 3)
        iM2 = 0
        iMl = 0
        iUn = 0
    ElseIf Me.Material.unidad = m2_ Then
        iKg = Math.Round(Me.Material.PesoXUnidad * ((Me.AnchoTerm * Me.LargoTerm) / 1000000) * Me.Cantidad, 3)
        iMl = 0
        iM2 = Math.Round((Me.AnchoTerm * Me.LargoTerm) / 1000000, 3)
        iUn = 0
    ElseIf Me.Material.unidad = un_ Then
        iM2 = 0
        iMl = 0
        iKg = 0
        iUn = Me.Cantidad
    End If
End Sub

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
    Dim scrap_ As Double

    scrap_ = (Scrap / 100) + 1

    If Material.unidad = Unidades.Ml_ Then     'ml
        Kg = funciones.RedondearDecimales(Material.PesoXUnidad * (Me.Largo / 1000) * Me.Cantidad)
        m2ml = funciones.RedondearDecimales(Me.Largo / 1000 * Cantidad)
        Precio = Me.Material.Valor * Me.Material.PesoXUnidad
        Precio = clasea.realizaCambio(Precio, Me.Material.Moneda.id, idMonedaDestino)
        Precio = ((Precio * Me.Largo * scrap_ * 2) / 1000) * Cantidad
        Precio = Precio / 2

        dimension_material = Cantidad & "X" & Me.LargoTerm
        dimension_pieza = Cantidad & "X" & Me.Largo

    ElseIf Material.unidad = Unidades.kg_ Then
        Kg = funciones.RedondearDecimales(Material.PesoXUnidad * Cantidad)
        m2ml = 0
        Precio = Me.Material.Valor * Material.PesoXUnidad
        Precio = clasea.realizaCambio(Precio, Me.Material.Moneda.id, idMonedaDestino)
        Precio = (Precio * Cantidad)
        dimension_material = Cantidad
        dimension_pieza = Cantidad
    ElseIf Material.unidad = Unidades.m2_ Then     'm2

        Kg = funciones.RedondearDecimales(Material.PesoXUnidad * ((Me.Ancho * Me.Largo) / 1000000) * Cantidad)
        m2ml = funciones.RedondearDecimales((Me.Ancho * Me.Largo) / 1000000 * Cantidad)
        Precio = Me.Material.Valor * Material.PesoXUnidad

        Precio = Precio * (Cantidad * Me.Ancho * Me.Largo * scrap_) / 1000000



        Precio = clasea.realizaCambio(Precio, Me.Material.Moneda.id, idMonedaDestino)
        dimension_material = Cantidad & "x" & Me.LargoTerm & "x" & Me.AnchoTerm
        dimension_pieza = Cantidad & "x" & Me.Largo & "x" & Me.Ancho

    ElseIf Material.unidad = Unidades.un_ Then     'uni

        m2ml = 0
        Precio = Me.Material.Valor
        Precio = clasea.realizaCambio(Precio, Me.Material.Moneda.id, idMonedaDestino)
        Precio = Precio * Cantidad

        Kg = Cantidad
        dimension_material = Cantidad
        dimension_pieza = Cantidad
    End If
    cxh = funciones.cantxhoja(Me.AnchoTerm, Me.LargoTerm, Me.Ancho, Me.Largo)
    If cxh = 0 Then cxh = 1
    costo = Precio / cxh


    dat.costo = funciones.FormatearDecimales(costo)
    dat.DimensionMaterial = dimension_material
    dat.DimensionPieza = dimension_pieza
    dat.Kg = Kg
    dat.m2 = m2ml

    CalcularDatosMaterial = dat
End Function
