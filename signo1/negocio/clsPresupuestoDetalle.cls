VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPresupuestoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public item As String
Public idPreuspuesto As Long
Public Pieza As Pieza
Public Cantidad As Long
Private vValorSistema As Double  'valorUnitario
Private vValorManual As Double  'ValorUnitarioManual
Public Detalles As String
Public entrega As Long
Public Amortizacion As Double
Public FormaCotizar As FormaCotizar
Public FechaPresupuesto As Date
Public Id As Long
Private PorcMO As Double
Private porcMAT As Double
Public presupuesto As clsPresupuesto
Public indiceAjuste As Double

Public Property Let ValorSistema(nvalue As Double)
    vValorSistema = funciones.RedondearDecimales(nvalue, 2)
End Property

Public Property Let ValorManual(nvalue As Double)
    vValorManual = funciones.RedondearDecimales(nvalue, 2)
End Property
Public Property Get ValorSistema() As Double
    ValorSistema = funciones.RedondearDecimales(vValorSistema, 2)
End Property

Public Property Get ValorManual() As Double
    ValorManual = funciones.RedondearDecimales(vValorManual, 2)
End Property
Public Property Get PorcentajeMDO() As Double
    PorcentajeMDO = PorcMO
End Property
Public Property Get PorcentajeMAT() As Double
    PorcentajeMAT = porcMAT
End Property


Public Property Get TotalMateriales() As Double
    TotalMateriales = RedondearDecimales((PorcentajeMAT / 100) * ValorManual * Cantidad)
End Property


Public Property Get TotalMDO() As Double
    TotalMDO = RedondearDecimales((PorcentajeMDO / 100) * ValorManual * Cantidad)
End Property
Public Function CalcularPrecioSistema(Optional ByRef costo As Double, Optional ByRef ManoObra As Double, Optional ByRef Materiales As Double, Optional ByRef TotalGastosGenerales As Double) As Double
    On Error GoTo err1
    Dim pMO As Double, PMA As Double, costo2 As Double
    Dim tot_con As Double
    Dim c As Long
    Dim tot_costo As Double, mo As Double, ma As Double
    Dim pie4 As Pieza
    Dim pie3 As Pieza
    Dim pie2 As Pieza
    Dim Pie As Pieza
    tot_costo = 0
    tot_con = 0
    c = 0
    mo = 0
    ma = 0

    tot_con = CalcularValorPieza(Pieza, pMO, PMA, costo2, Me.Cantidad, Pieza.EsConjunto)
    c = c + 1
    mo = mo + pMO
    ma = ma + PMA
    tot_costo = tot_costo + costo2
    PorcMO = funciones.RedondearDecimales(mo / c, 0)
    porcMAT = funciones.RedondearDecimales(ma / c, 0)
    vValorSistema = tot_con
    costo = tot_costo
    ManoObra = mo
    Materiales = ma

    TotalGastosGenerales = costo - (costo / (1 + (presupuesto.Gastos / 100)))
    If mo + ma = 0 Then c = c - 1
    If Not Pieza.EsConjunto Then
        'PorcMO = funciones.RedondearDecimales((mo * 100) / (mo + ma), 0)
        'porcMAT = funciones.RedondearDecimales((ma * 100) / (mo + ma), 0)
        '  Exit Function porq puse esto? 15-3-16
    End If

    For Each Pie In Pieza.PiezasHijas
        tot_con = tot_con + ((CalcularValorPieza(Pie, pMO, PMA, costo2, Pie.Cantidad * Me.Cantidad)) * Pie.Cantidad)
        c = c + 1
        mo = mo + pMO
        ma = ma + PMA
        tot_costo = tot_costo + costo2
        If mo + ma = 0 Then c = c - 1
        For Each pie2 In Pie.PiezasHijas
            tot_con = tot_con + (CalcularValorPieza(pie2, pMO, PMA, costo2, pie2.Cantidad * Pie.Cantidad * Me.Cantidad) * (pie2.Cantidad * Pie.Cantidad))
            c = c + 1
            mo = mo + pMO
            ma = ma + PMA
            tot_costo = tot_costo + costo2
            If mo + ma = 0 Then c = c - 1
            For Each pie3 In pie2.PiezasHijas
                tot_con = tot_con + (CalcularValorPieza(pie3, pMO, PMA, costo2, pie3.Cantidad * pie2.Cantidad * Pie.Cantidad * Me.Cantidad) * pie3.Cantidad * pie2.Cantidad * Pie.Cantidad)
                c = c + 1
                mo = mo + pMO
                ma = ma + PMA
                tot_costo = tot_costo + costo2
                If mo + ma = 0 Then c = c - 1
                For Each pie4 In pie3.PiezasHijas
                    tot_con = tot_con + (CalcularValorPieza(pie4, pMO, PMA, costo2, pie4.Cantidad * pie3.Cantidad * pie2.Cantidad * Pie.Cantidad * Me.Cantidad) * pie4.Cantidad * pie3.Cantidad * pie2.Cantidad * Pie.Cantidad)
                    c = c + 1
                    mo = mo + pMO
                    ma = ma + PMA
                    tot_costo = tot_costo + costo2
                    If mo + ma = 0 Then c = c - 1
                Next pie4
            Next pie3
        Next pie2
    Next Pie



    'importe mat e importe mo

    '18-12-15
    Dim T As Double
    T = mo + ma

    PorcMO = funciones.RedondearDecimales((mo * 100) / T, 0)
    porcMAT = funciones.RedondearDecimales((ma * 100) / T, 0)

    '18-12-15
    'se calcula el porcentaje en base al importe y no a los porcentajes de cada pieza dentro de un conjunto
    '    PorcMO = funciones.RedondearDecimales(mo / c, 0)
    '    porcMAT = funciones.RedondearDecimales(ma / c, 0)




    ManoObra = PorcMO
    Materiales = porcMAT
    vValorSistema = tot_con * (1 + (Me.indiceAjuste) / 100)
    costo = tot_costo


    TotalGastosGenerales = costo - (costo * (presupuesto.Gastos / 100))


    Exit Function

err1:
    CalcularPrecioSistema = -1
End Function
Private Function CalcularValorPieza(piez As Pieza, ByRef mo As Double, ByRef ma As Double, ByRef vcosto As Double, ByVal Cantidad As Double, Optional conj As Boolean = True, Optional ByRef TotalGastosGenerales) As Double
    On Error GoTo err1
    Dim Precio As Double
    Dim Kg As Double
    Dim claseS As New classStock
    Dim Total As Double
    Dim unitario As Double
    Dim amort As Long
    Dim MDO As Double, Fijo As Double, Cambio As Double, muMater As Double, mCosto As Double, materiale As Double

    Dim mocosto As Double
    Dim manodeobra As Double
    Dim Materiales As Double
    Cantidad = Cantidad
    FormaCotizar = FormaCotizar




    amort = funciones.amortizaV2(piez.Id, Cantidad, FormaCotizar, Amortizacion, conj)



    claseS.calcular_valor_materiales piez.Id, Kg, Precio, presupuesto.moneda.Id  '--> VER DE REEMPLAZAR


    'Precio = piez.TotalCostoMateriales  'LO AGREGE 29-7-11
    piez.CalcularValorMDO presupuesto.moneda, Fijo, Cambio, MDO

    Kg = piez.TotalKG


    muMater = valorPorPeso(Kg, presupuesto.PorcMen10, presupuesto.PorcMen15, presupuesto.PorcMas15)
    mCosto = Precio



    Materiales = ((Precio * ((presupuesto.Gastos / 100) + 1)) * muMater)



    manodeobra = (MDO + (Cambio / Cantidad) + (Fijo / amort))
    mocosto = manodeobra * ((presupuesto.PorcentajeManoObraMuerta / 100) + 1)

    manodeobra = manodeobra * ((presupuesto.Gastos / 100) + 1) * ((presupuesto.PorcentajeManoObraMuerta / 100) + 1)
    manodeobra = manodeobra * ((presupuesto.PorcMDO / 100) + 1)
    vcosto = ((funciones.RedondearDecimales((mocosto + mCosto), 2) * Cantidad) + vcosto)



    TotalGastosGenerales = vcosto - (vcosto * (presupuesto.Gastos / 100))


    '18-12-15 --------
    'cambio pedido por marceolo tadiotto, que enlugar de calculaar los porcentajes en base a los porcentajes de
    'los conjuntos, lo haga ne base a los importes por cada grupo
    'mo = funciones.RedondearDecimales((mocosto * 100) / (mocosto + mCosto))
    ' ma = funciones.RedondearDecimales((mCosto * 100) / (mocosto + mCosto))
    mo = mocosto
    ma = mCosto
    '------------------------------------------------
    CalcularValorPieza = funciones.RedondearDecimales(((manodeobra + Materiales)), 2)
    Exit Function
err1:
    CalcularValorPieza = 0
End Function
Public Function costo() As Double
    Me.CalcularPrecioSistema costo
End Function



Public Function TotalGastosGenerales() As Double
    Me.CalcularPrecioSistema , , , TotalGastosGenerales

End Function
