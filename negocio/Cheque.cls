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

Public id As Long
Public numero As String
Public FechaRecibido As Date
Public FechaVencimiento As Date
Public Monto As Double
Public entro As Boolean
Public IdChequera As Long
Public chequera As chequera
Public Banco As Banco
Public OrigenDestino As String
Public EnCartera As Boolean
Public Propio As Boolean
Public Moneda As clsMoneda
Public Observaciones As String
Public TercerosPropio As Boolean
Public FechaEmision As Date
Public IdOrdenPagoOrigen As Long
Public estado As EstadoCheque
Public Depositado As Boolean

Public Property Get OrigenCheque()
    If Propio Then
        OrigenCheque = "Propio"
    Else
        If TercerosPropio Then
            OrigenCheque = "Terceros propio"
        Else
            OrigenCheque = "Terceros"
        End If
    End If
End Property

Public Property Get Utilizado() As Boolean
    Utilizado = (CDbl(FechaVencimiento) <> 0 And Not Me.EnCartera)
End Property
