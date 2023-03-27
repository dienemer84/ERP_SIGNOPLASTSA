VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Id As Long
Public EstadoPresupuesto As EstadoPresupuesto
Public FechaEntrega As Long
Public PorcMDO As Double
Public PorcMen15 As Double
Public PorcMen10 As Double
Public PorcMas15 As Double
Public Gastos As Double
Public manteOferta As Integer
Public detalle As String
Public FechaProcesado As Date
Public fechaCreado As Date
Public moneda As clsMoneda
Public UsuarioCreado As clsUsuario
Public FechaModificado As Date
Public UsuarioModificado As clsUsuario
Public FechaFinalizado As Date
Public UsuarioFinalizado As clsUsuario
Public DetallePresupuesto As Collection    'of clsPresupuestoDetalle
Public cliente As clsCliente
Public VencimientoPresupuesto As Date
Public Descuento As Double
Public Anticipo As Double
Public FormaPagoAnticipo As String
Public FormaPagoSaldo As String
Public PorcentajeManoObraMuerta As Double
Public DiasPagoSaldo As Long
Public DiasPagoAnticipo As Long
Private Detalle_ As clsPresupuestoDetalle
Enum SistemaManual
    Sistema = 1
    Manual = 2
    SMCosto = 3
    SMGG = 4
    SMUtilidad = 5
End Enum

Public Property Get CondicionesComerciales() As String
    CondicionesComerciales = "Anticipo: " & Me.CondicionesComercialesAnticipo & ", Saldo: " & Me.CondicionesComercialesSaldo
End Property


Public Property Get CondicionesComercialesAnticipo() As String
    If Anticipo = 0 Then
        CondicionesComercialesAnticipo = "Sin Anticipo"
    Else

        If LenB(Trim(FormaDePagoAnticipo)) = 0 Then
            If CantDiasAnticipo > 0 Then
                CondicionesComercialesAnticipo = Anticipo & "%, a " & CantDiasAnticipo & " días  "
            Else
                CondicionesComercialesAnticipo = Anticipo & "%, Contado"
            End If
        Else
            If CantDiasAnticipo > 0 Then
                CondicionesComercialesAnticipo = Anticipo & "%, a " & CantDiasAnticipo & " días, F.Pago: " & FormaDePagoAnticipo
            Else
                CondicionesComercialesAnticipo = Anticipo & "&, F.Pago: " & FormaDePagoAnticipo
            End If
        End If
    End If
End Property
Public Property Get CondicionesComercialesSaldo() As String
    If FormaDePagoSaldo <> Empty Then
        If CantDiasSaldo > 0 Then
            CondicionesComercialesSaldo = CantDiasSaldo & " días,  F.Pago: " & FormaDePagoSaldo
        Else
            CondicionesComercialesSaldo = "Contado,  F.Pago: " & FormaDePagoSaldo
        End If
    Else
        If CantDiasSaldo > 0 Then
            CondicionesComercialesSaldo = CantDiasSaldo & " días"
        Else
            CondicionesComercialesSaldo = "Contado"
        End If
    End If
End Property




Public Property Get IdFormateada() As String
    IdFormateada = Format("0000", Id)
End Property

Public Function SubTotal(A As SistemaManual) As Double
    Dim tot As Double
    tot = 0
    For Each Detalle_ In DetallePresupuesto
        If A = Sistema Then
            tot = tot + Detalle_.ValorSistema * Detalle_.Cantidad
        ElseIf A = Manual Then
            tot = tot + Detalle_.ValorManual * Detalle_.Cantidad
        End If
    Next
    SubTotal = funciones.RedondearDecimales(tot)
End Function





Public Function Total(A As SistemaManual) As Double
    Dim tot As Double
    tot = 0
    For Each Detalle_ In DetallePresupuesto
        If A = Sistema Then
            tot = tot + Detalle_.ValorSistema * Detalle_.Cantidad
        ElseIf A = Manual Then
            tot = tot + Detalle_.ValorManual * Detalle_.Cantidad
        ElseIf A = SMCosto Then
            tot = tot + Detalle_.costo    '* detalle_.cantidad
        ElseIf A = SMGG Then
            tot = tot + Detalle_.TotalGastosGenerales

        End If


    Next
    If A = SMUtilidad Then
        tot = Me.Total(Manual) - Me.Total(SMCosto)
    End If

    Total = funciones.RedondearDecimales(tot * (1 - Me.Descuento / 100))
End Function


Public Function TotalMateriales()
    Dim A As clsPresupuestoDetalle
    Dim c As Double
    c = 0
    For Each A In Me.DetallePresupuesto
        c = c + (A.TotalMateriales)
    Next A

    TotalMateriales = c

End Function



Public Function TotalMDO()
    Dim A As clsPresupuestoDetalle
    Dim c As Double
    c = 0
    For Each A In Me.DetallePresupuesto
        c = c + (A.TotalMDO)
    Next A

    TotalMDO = c

End Function




Public Function TotalConDescuento(A As SistemaManual) As Double
    TotalConDescuento = Total(A) * (1 - (Descuento / 100))
End Function
