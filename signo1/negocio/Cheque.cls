VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public numero As String
Public FechaRecibido As Date
Public FechaVencimiento As Date
Public FechaRecibo As Date
Public Monto As Double
Public entro As Boolean
Public FechaIngresoBanco As Date
Public IdChequera As Long
Public chequera As chequera
Public Banco As Banco
Public OrigenDestino As String
Public EnCartera As Boolean
Public Propio As Boolean
Public moneda As clsMoneda
Public Observaciones As String
Public TercerosPropio As Boolean
Public FechaEmision As Date

Public IdOrdenPagoOrigen As Long
Public NumeroLiquidacionCaja As Long
Public NumeroPagoACuenta As Long
Public NumeroMovimiento As Long

Public estado As EstadoCheque
Public Depositado As Boolean
Public destino As String
Public Recibo As String

Private mIdLiquidacionCajaOrigen As Long

Public Property Get IdLiquidacionCajaOrigen() As Long
    IdLiquidacionCajaOrigen = mIdLiquidacionCajaOrigen
End Property

Public Property Let IdLiquidacionCajaOrigen(ByVal value As Long)
    mIdLiquidacionCajaOrigen = value
End Property

Public Property Get OrigenCheque() As String

    If Propio Then
        OrigenCheque = "PROPIO"
    ElseIf TercerosPropio Then
        OrigenCheque = "TERCERO PROPIO"
    Else
        OrigenCheque = "TERCERO"
    End If

End Property

Public Property Get Utilizado() As Boolean

    Utilizado = _
        (IdOrdenPagoOrigen > 0) Or _
        (IdLiquidacionCajaOrigen > 0) Or _
        (NumeroLiquidacionCaja > 0) Or _
        (NumeroPagoACuenta > 0) Or _
        (NumeroMovimiento > 0)

End Property


Public Sub LimpiarOrigenUso()

    IdOrdenPagoOrigen = 0
    IdLiquidacionCajaOrigen = 0
    NumeroLiquidacionCaja = 0
    NumeroPagoACuenta = 0
    NumeroMovimiento = 0

End Sub
