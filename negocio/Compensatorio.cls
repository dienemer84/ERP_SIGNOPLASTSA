VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Compensatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Id As Long
Public Comprobante As clsFacturaProveedor
Public Monto As Double
Public Tipo As TipoCompensatorio
Public Observacion As String
Public FechaCancelacion As Date
Public IdOrdenPago As Long

Public Cancelado As Boolean
Public NetoGravadoCompensado As Double
Public alicuotaPercepcion As Double
Public MontoAPercibir As Double
