VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiquidacionSubdiarioVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public nombre As String
Public desde As Date
Public hasta As Date
Private m_detalles As Collection    'SubdiarioVentasDetalles
Public EsDeVenta As Boolean    'venta o compra


Public Property Get Detalles() As Collection
    If Not IsSomething(m_detalles) And Me.Id <> 0 Then
        Set m_detalles = DAOSubdiarios.FindAllDetallesLiquiVentaByLiquiVenta(Me.Id, Me.EsDeVenta)
    End If
    Set Detalles = m_detalles
End Property

Public Property Set Detalles(value As Collection)
    Set m_detalles = value
End Property




