VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CertificadoRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public RazonSocial As String
Public IdOrdenPago As Long
Public Cuit As String
Public IB As String
Public FEcha As Date
Public Domicilio As String
Public localidad As String
Public Retencion As Retencion
Public Detalles As Collection

Public cp As Long


Public Property Get Total() As Double
    Dim det As CertificadoRetencionDetalles
    Dim c As Double
    c = 0
    For Each det In Detalles
        c = c + det.TotalRetenido(Me)
    Next
    Total = c

End Property

Public Property Get TotalRetenido() As Double
    Dim det As CertificadoRetencionDetalles
    Dim a As Double
    For Each det In Detalles
        a = a + det.TotalRetenido(Me)
    Next
    TotalRetenido = a
End Property
Public Property Get DomicilioFormateado() As String
    DomicilioFormateado = Me.Domicilio & "  |  " & Me.localidad & " (" & Me.cp & ")"
End Property

Public Property Get CuitIBFormateado() As String
    CuitIBFormateado = "CUIT: " & Me.Cuit & "  |  " & " IBCM: " & Me.IB
End Property
