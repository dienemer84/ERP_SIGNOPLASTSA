VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoletaDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Id As Long
Public numero As Long
Public Monto As Double
Public fechaDeposito As Date
Public CuentaDestino As CuentaBancaria
Public Cheques As New Collection
Public OperacionesCaja As New Collection

Public CantidadCheques As Long

Public OperacionOrigen As operacion
Public TipoDeposito As TipoDeposito
