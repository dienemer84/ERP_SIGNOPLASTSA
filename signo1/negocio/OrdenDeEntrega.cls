VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenDeEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public id As Long
Public referencia As String
Public FEcha As Date
Public usuarioCreador As clsUsuario
Public usuarioAprobador As clsUsuario
Public fechaCreado As Date
Public fechaAprobado As Date
Public cliente As clsCliente
Public estado As EstadoOrdenEntrega
Public moneda As clsMoneda

