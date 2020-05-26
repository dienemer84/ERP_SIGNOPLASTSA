VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum EstadoProveedor
    Activo = 1
    Inactivo = 2
End Enum


Public Enum TipoOperacionProveedor
    Alta = 1
    Modificacion = 2
End Enum



Public Enum EstadoFacturaProveedor
    EnProceso = 1
    Aprobada = 2
End Enum

