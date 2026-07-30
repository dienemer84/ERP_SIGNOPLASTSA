VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransferenciaBcariaCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public Monto As Double

Public moneda As clsMoneda

Public FechaCarga As Date
Public FechaOperacion As Date
Public IdPertenencia As Long
Public Pertenencia As OrigenOperacion
Public EntradaSalida As OperacionEntradaSalida
Public Comprobante As String

Public ReciboID As String
Public Recibo As Recibo

Public ClienteRazon As String
Public Cliente As clsCliente

Public NombreBanco As String
Public CuentaBancaria As String
Public IdCtaBancaria As Long

Public ReciboAplicado As Long

Private Sub Class_Initialize()
    Me.EntradaSalida = OPEntrada
End Sub


