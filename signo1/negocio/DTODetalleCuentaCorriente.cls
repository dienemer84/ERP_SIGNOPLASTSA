VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTODetalleCuentaCorriente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FEcha As Date
Public Comprobante As String
Public Debe As Double
Public Haber As Double
Public AtributoExtra As Boolean    'en cta cte prov, marco si la factura esta en alguna orden, ver si tiene utilidad en cta cte cliente
Public saldo As Double
Public tipoComprobante As TipoComprobanteUsado
Public IdComprobante As Long
Public TmpId As Long
