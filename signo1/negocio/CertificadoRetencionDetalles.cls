VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CertificadoRetencionDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public Comprobante As String
Public NetoGravado As Double
Public alicuota As Double
Public FacturaProveedor As clsFacturaProveedor
Public IdMoneda As Long
Public IdCertificado As Long
Public TotalFactura As Double
Public NCertificado As Long

Public Property Get TotalRetenido(cr As CertificadoRetencion)
    TotalRetenido = funciones.RedondearDecimales(IIf(Me.NetoGravado > cr.Retencion.MinimoImponible, Me.NetoGravado * (Me.alicuota / 100), 0))
End Property

