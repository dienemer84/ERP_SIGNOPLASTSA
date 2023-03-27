VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCuentaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim vMonto As Double
Dim vCuentas As clsCuentaContable

Public Id As Long
Public Property Let Monto(nvalue As Double)
    vMonto = nvalue
End Property

Public Property Get Monto() As Double
    Monto = vMonto

End Property
Public Property Let cuentas(nvalue As clsCuentaContable)
    Set vCuentas = nvalue
End Property
Public Property Get cuentas() As clsCuentaContable
    Set cuentas = vCuentas
End Property

