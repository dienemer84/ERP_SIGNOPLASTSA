VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPresupuestoDetalleHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Id As Long
Public NombrePieza As String
Public Pieza As Pieza
Public HistoricoHijos As New Collection
Public historicoMDO As New Collection
Public HistoricoMAT As New Collection
Public FEcha As Date
Public DetallePresupuesto As clsPresupuestoDetalle


Public Property Get TotalKGMateriales() As Double
    Dim T As PresupuestoDetalleHistoricoMAT
    Dim tot As Double

    For Each T In Me.HistoricoMAT
        tot = tot + T.CalcularDatosMaterial(DetallePresupuesto.presupuesto.moneda.Id).Kg
    Next T

    TotalKGMateriales = tot
End Property
Public Property Get TotalM2Materiales() As Double
    Dim T As PresupuestoDetalleHistoricoMAT
    Dim tot As Double

    For Each T In Me.HistoricoMAT
        tot = tot + T.CalcularDatosMaterial(DetallePresupuesto.presupuesto.moneda.Id).m2
    Next T

    TotalM2Materiales = tot
End Property
Public Property Get TotalCostoMateriales() As Double
    Dim T As PresupuestoDetalleHistoricoMAT
    Dim tot As Double

    For Each T In Me.HistoricoMAT
        tot = tot + T.CalcularDatosMaterial(DetallePresupuesto.presupuesto.moneda.Id).costo
    Next T

    TotalCostoMateriales = tot
End Property

Public Property Get TotalCostoMDO() As Double
    Dim T As PresupuestoDetalleHistoricoMDO
    Dim tot As Double

    For Each T In Me.historicoMDO
        tot = tot + (T.Valor * T.CantOperarios * T.Tiempo)
    Next T

    TotalCostoMDO = tot
End Property




