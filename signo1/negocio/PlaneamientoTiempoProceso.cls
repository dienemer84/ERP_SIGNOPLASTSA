VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlaneamientoTiempoProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Id As Long
Public idpedido As Long
Public idPieza As Long
Public idDetallePedido As Long
Public idDetallePedidoConj As Long
Public Tarea As clsTarea
Public FechaFin As Date
Public Detalles As New Collection
Public OperariosCotizado As Long
Public TiempoCotizado As Double
Public TiempoTotalReal As Double
Public item As String
Public EsConjunto As Boolean
Public DetalleOt As DetalleOrdenTrabajo
Public DetalleOtConjunto As DetalleOTConjuntoDTO
Public Observacion As String

Public Planificacion As New TiempoProcesoPlanificado

Public Property Get FINALIZADO() As Boolean
    FINALIZADO = (CDbl(Me.FechaFin) <> 0)
End Property

