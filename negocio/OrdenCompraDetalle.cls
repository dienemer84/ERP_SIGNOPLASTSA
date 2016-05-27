VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdenCompraDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public IdOrdenCompra As Long
Public IdPeticionOfertaDetalle As Long
Public Valor As Double
Public Cantidad As Long
Public descripcion As String
Public Entregas As New Collection    '<OrdenCompraDetalleEntrega>
Public DetallesRemitos As New Collection    '<RemitoProveedorDetalle>

