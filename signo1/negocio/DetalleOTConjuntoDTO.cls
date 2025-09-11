VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleOTConjuntoDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public idDetallePedido As Long
Public idPedido As Long
Public Pieza As Pieza
Public Cantidad As Double
Public estado As EstadoProcesoDetalleOrdenTrabajo
Public DetalleRaiz As DetalleOrdenTrabajo
Public CantidadTotalSegunConjunto As Double
Public CantidadTotalStatic As Double 'Cantidad Pedida
Public IdentificadorPosicion As String

Public CantidadRecibida As Double
Public CantidadFabricada As Double
Public CantidadScrap As Double

Public FechaInicio As Date
Public FechaFin As Date

Public Recibio As Long
Public SiguienteProceso As String
