VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConciliacionBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long

Public IdCuentaBancaria As Long

Public FechaDesde As Date
Public FechaHasta As Date

Public FechaCierre As Date
Public IdUsuarioCierre As Long

Public SaldoInicial As Double
Public TotalIngresos As Double
Public TotalEgresos As Double
Public SaldoFinal As Double

Public CantidadMovimientos As Long

'1 = cerrada
'0 = anulada
Public estado As Integer

Public Observaciones As String

Private m_Detalles As Collection


Private Sub Class_Initialize()
    Set m_Detalles = New Collection
End Sub


Public Property Get Detalles() As Collection
    Set Detalles = m_Detalles
End Property


Public Property Set Detalles(ByVal value As Collection)
    Set m_Detalles = value
End Property
