VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TotalEstaticoRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TotalReciboEstatico As Double

Public TotalChequesEstatico As Double
Public TotalEfectivoEstatico As Double
Public TotalDepositosEstatico As Double

'Public Retenciones As New Collection 'uso la que ya tiene el recibo

Public Property Get TotalRecibidoEstatico() As Double
    TotalRecibidoEstatico = Me.TotalChequesEstatico + Me.TotalEfectivoEstatico + Me.TotalDepositosEstatico
End Property
